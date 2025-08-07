"""
MOPS Sheets Uploader - Main Orchestrator
Main class that orchestrates the complete pipeline from PDF scanning to Google Sheets upload.
"""

import os
import time
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path

from .config import MOPSConfig, load_config
from .models import ProcessingResult, CoverageStats, StockListChanges
from .pdf_scanner import PDFScanner
from .stock_data_loader import StockDataLoader
from .matrix_builder import MatrixBuilder
from .sheets_connector import SheetsUploadManager
from .report_analyzer import ReportAnalyzer

logger = logging.getLogger(__name__)

class MOPSSheetsUploader:
    """
    MOPS PDF files to Google Sheets matrix uploader
    
    Main orchestrator class that coordinates all components to:
    1. Scan PDF files in downloads directory
    2. Load stock data from CSV
    3. Build company × quarter matrix
    4. Upload to Google Sheets with formatting
    5. Generate comprehensive analysis report
    """
    
    def __init__(self, 
                 downloads_dir: str = "downloads",
                 stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                 config: Optional[MOPSConfig] = None,
                 **kwargs):
        """
        Initialize MOPS Sheets Uploader
        
        Args:
            downloads_dir: Path to downloads directory
            stock_csv_path: Path to stock CSV file
            config: Optional configuration object
            **kwargs: Additional configuration parameters
        """
        # Load configuration
        if config is None:
            self.config = load_config()
        else:
            self.config = config
        
        # Override config with provided parameters
        if downloads_dir != "downloads":
            self.config.downloads_dir = downloads_dir
        if stock_csv_path != "StockID_TWSE_TPEX.csv":
            self.config.stock_csv_path = stock_csv_path
        
        # Apply additional configuration
        for key, value in kwargs.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
        
        # Initialize components
        self.pdf_scanner = PDFScanner(self.config)
        self.stock_loader = StockDataLoader(self.config)
        self.matrix_builder = MatrixBuilder(self.config)
        self.upload_manager = SheetsUploadManager(self.config)
        self.analyzer = ReportAnalyzer(self.config)
        
        # Processing state
        self._last_processing_result: Optional[ProcessingResult] = None
        
        # Setup logging
        self._setup_logging()
        
        logger.info(f"📊 MOPS Sheets Uploader v1.0.0 初始化完成")
        logger.info(f"   下載目錄: {self.config.downloads_dir}")
        logger.info(f"   股票清單: {self.config.stock_csv_path}")
        logger.info(f"   工作表名稱: {self.config.worksheet_name}")
    
    def run(self) -> ProcessingResult:
        """
        Main execution method - runs the complete pipeline
        
        Returns:
            ProcessingResult with detailed execution information
        """
        start_time = time.time()
        logger.info("🚀 開始執行 MOPS Sheets Uploader 流程")
        logger.info("=" * 50)
        
        try:
            # Stage 1: Validate configuration
            self._validate_configuration()
            
            # Stage 2: Scan PDF files
            pdf_data = self._scan_pdf_files()
            
            # Stage 3: Load stock data
            stock_df, stock_changes = self._load_stock_data()
            
            # Stage 4: Build matrix
            matrix_df = self._build_matrix(stock_df, pdf_data)
            
            # Stage 5: Analyze data
            coverage_stats = self._analyze_coverage(matrix_df, pdf_data)
            
            # Stage 6: Upload to Sheets and/or export CSV
            upload_result = self._upload_matrix(matrix_df, coverage_stats, stock_changes)
            
            # Stage 7: Generate comprehensive report
            analysis_report = self._generate_analysis_report(matrix_df, pdf_data, coverage_stats, stock_changes)
            
            # Create final result
            processing_time = time.time() - start_time
            result = ProcessingResult(
                success=True,
                matrix_uploaded=upload_result.get('sheets_success', False),
                csv_exported=upload_result.get('csv_exported', False),
                csv_path=upload_result.get('csv_path'),
                sheets_url=upload_result.get('sheets_url'),
                coverage_stats=coverage_stats,
                stock_changes=stock_changes,
                processing_time=processing_time,
                total_files_processed=sum(len(pdfs) for pdfs in pdf_data.values())
            )
            
            # Log final summary
            self._log_final_summary(result)
            
            # Store result for later access
            self._last_processing_result = result
            
            return result
            
        except Exception as e:
            processing_time = time.time() - start_time
            logger.error(f"❌ 處理失敗: {e}")
            
            result = ProcessingResult(
                success=False,
                error_message=str(e),
                processing_time=processing_time
            )
            
            self._last_processing_result = result
            return result
    
    def scan_pdf_files(self) -> Dict[str, List]:
        """
        Scan downloads directory and return PDF file data
        
        Returns:
            Dictionary mapping company_id to list of PDFFile objects
        """
        return self.pdf_scanner.scan_downloads_directory()
    
    def load_stock_data(self) -> Tuple[Any, Optional[StockListChanges]]:
        """
        Load and validate stock data from CSV
        
        Returns:
            Tuple of (DataFrame, StockListChanges)
        """
        stock_df = self.stock_loader.load_stock_csv()
        validation_info = self.stock_loader.validate_stock_data(stock_df)
        
        if not validation_info['validation_passed']:
            logger.warning("⚠️ 股票資料驗證發現問題，但繼續處理")
        
        # Detect changes if enabled
        stock_changes = None
        if self.config.detect_stock_changes:
            stock_changes = self.stock_loader.detect_stock_list_changes(stock_df)
        
        return stock_df, stock_changes
    
    def build_matrix(self, stock_df=None, pdf_data=None) -> Any:
        """
        Build company × quarter matrix
        
        Returns:
            DataFrame with matrix data
        """
        if stock_df is None:
            stock_df, _ = self.load_stock_data()
        if pdf_data is None:
            pdf_data = self.scan_pdf_files()
        
        # Build base matrix
        matrix_df = self.matrix_builder.build_base_matrix(stock_df, pdf_data)
        
        # Populate with PDF status
        matrix_df = self.matrix_builder.populate_pdf_status(matrix_df, pdf_data)
        
        # Apply priority rules for conflicts
        matrix_df = self.matrix_builder.apply_priority_rules(matrix_df, pdf_data)
        
        # Add summary columns if enabled
        if self.config.include_summary_sheet:
            matrix_df = self.matrix_builder.add_summary_columns(matrix_df)
        
        return matrix_df
    
    def upload_to_sheets(self, matrix_df=None, coverage_stats=None, stock_changes=None) -> Dict[str, Any]:
        """
        Upload matrix to Google Sheets
        
        Returns:
            Dictionary with upload results
        """
        if matrix_df is None:
            stock_df, stock_changes = self.load_stock_data()
            pdf_data = self.scan_pdf_files()
            matrix_df = self.build_matrix(stock_df, pdf_data)
            coverage_stats = self.analyzer.analyze_coverage(matrix_df, pdf_data)
        
        return self.upload_manager.upload_with_fallback(matrix_df, coverage_stats, stock_changes)
    
    def export_to_csv(self, matrix_df=None, coverage_stats=None) -> str:
        """
        Export matrix to CSV file
        
        Returns:
            Path to exported CSV file
        """
        if matrix_df is None:
            stock_df, _ = self.load_stock_data()
            pdf_data = self.scan_pdf_files()
            matrix_df = self.build_matrix(stock_df, pdf_data)
            coverage_stats = self.analyzer.analyze_coverage(matrix_df, pdf_data)
        
        return self.upload_manager.connector.export_csv_backup(matrix_df, coverage_stats)
    
    def generate_report(self, include_analysis: bool = True) -> Dict[str, Any]:
        """
        Generate comprehensive status report
        
        Args:
            include_analysis: Whether to include detailed analysis
            
        Returns:
            Comprehensive report dictionary
        """
        # Get or generate data
        stock_df, stock_changes = self.load_stock_data()
        pdf_data = self.scan_pdf_files()
        matrix_df = self.build_matrix(stock_df, pdf_data)
        coverage_stats = self.analyzer.analyze_coverage(matrix_df, pdf_data)
        
        base_report = {
            'generation_time': datetime.now().isoformat(),
            'configuration': {
                'downloads_dir': self.config.downloads_dir,
                'stock_csv_path': self.config.stock_csv_path,
                'max_years': self.config.max_years,
                'worksheet_name': self.config.worksheet_name
            },
            'summary': coverage_stats.get_summary() if hasattr(coverage_stats, 'get_summary') else {
                'total_companies': coverage_stats.total_companies,
                'companies_with_pdfs': coverage_stats.companies_with_pdfs,
                'coverage_percentage': coverage_stats.coverage_percentage,
                'total_reports': coverage_stats.total_actual_reports
            },
            'stock_changes': stock_changes,
            'matrix_dimensions': {
                'companies': len(matrix_df),
                'quarters': len([col for col in matrix_df.columns if 'Q' in col])
            }
        }
        
        if include_analysis:
            analysis_report = self.analyzer.generate_comprehensive_report(
                matrix_df, pdf_data, coverage_stats, stock_changes
            )
            base_report.update(analysis_report)
        
        return base_report
    
    def test_connection(self) -> bool:
        """Test Google Sheets connection"""
        return self.upload_manager.connector.test_connection()
    
    def get_last_result(self) -> Optional[ProcessingResult]:
        """Get the result of the last processing run"""
        return self._last_processing_result
    
    # Private methods for pipeline stages
    
    def _validate_configuration(self) -> None:
        """Validate configuration before processing"""
        logger.info("🔧 驗證設定...")
        
        errors = self.config.validate()
        if errors:
            error_msg = "Configuration validation failed: " + "; ".join(errors)
            logger.error(f"❌ {error_msg}")
            raise ValueError(error_msg)
        
        logger.info("✅ 設定驗證通過")
    
    def _scan_pdf_files(self) -> Dict[str, List]:
        """Stage 2: Scan PDF files"""
        logger.info("🔍 第二階段: 掃描 PDF 檔案")
        
        pdf_data = self.pdf_scanner.scan_downloads_directory()
        
        if not pdf_data:
            logger.warning("⚠️ 未發現任何 PDF 檔案")
        
        return pdf_data
    
    def _load_stock_data(self) -> Tuple[Any, Optional[StockListChanges]]:
        """Stage 3: Load stock data"""
        logger.info("📋 第三階段: 載入股票資料")
        
        stock_df, stock_changes = self.load_stock_data()
        
        return stock_df, stock_changes
    
    def _build_matrix(self, stock_df, pdf_data) -> Any:
        """Stage 4: Build matrix"""
        logger.info("🏗️ 第四階段: 建構矩陣")
        
        matrix_df = self.build_matrix(stock_df, pdf_data)
        
        logger.info(f"✅ 矩陣建構完成: {len(matrix_df)} × {len(matrix_df.columns)}")
        
        return matrix_df
    
    def _analyze_coverage(self, matrix_df, pdf_data) -> CoverageStats:
        """Stage 5: Analyze coverage"""
        logger.info("📊 第五階段: 分析涵蓋率")
        
        coverage_stats = self.matrix_builder.generate_coverage_stats(matrix_df, pdf_data)
        
        return coverage_stats
    
    def _upload_matrix(self, matrix_df, coverage_stats, stock_changes) -> Dict[str, Any]:
        """Stage 6: Upload matrix"""
        logger.info("🚀 第六階段: 上傳矩陣")
        
        upload_result = self.upload_manager.upload_with_fallback(
            matrix_df, coverage_stats, stock_changes
        )
        
        return upload_result
    
    def _generate_analysis_report(self, matrix_df, pdf_data, coverage_stats, stock_changes) -> Dict[str, Any]:
        """Stage 7: Generate analysis report"""
        logger.info("📋 第七階段: 產生分析報告")
        
        if self.config.auto_suggest_downloads:
            analysis_report = self.analyzer.generate_comprehensive_report(
                matrix_df, pdf_data, coverage_stats, stock_changes
            )
            
            # Log key suggestions
            suggestions = analysis_report.get('download_suggestions', [])
            if suggestions:
                logger.info("💡 下載建議:")
                for suggestion in suggestions[:5]:  # Show first 5
                    if suggestion.startswith('   •'):
                        logger.info(suggestion)
            
            return analysis_report
        
        return {}
    
    def _setup_logging(self) -> None:
        """Setup logging configuration"""
        # Create logs directory if it doesn't exist
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)
        
        # Configure logging format
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        
        # Set up file handler
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = logs_dir / f"mops_sheets_uploader_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        # Reduce noise from external libraries
        logging.getLogger('urllib3').setLevel(logging.WARNING)
        logging.getLogger('requests').setLevel(logging.WARNING)
        logging.getLogger('gspread').setLevel(logging.WARNING)
    
    def _log_final_summary(self, result: ProcessingResult) -> None:
        """Log final processing summary"""
        logger.info("=" * 50)
        logger.info("📊 處理完成摘要")
        logger.info("=" * 50)
        
        if result.success:
            logger.info("✅ 整體狀態: 成功")
        else:
            logger.info("❌ 整體狀態: 失敗")
            if result.error_message:
                logger.info(f"   錯誤訊息: {result.error_message}")
            return
        
        # Success details
        if result.matrix_uploaded:
            logger.info("✅ Google Sheets 上傳: 成功")
            if result.sheets_url:
                logger.info(f"🔗 試算表網址: {result.sheets_url}")
        else:
            logger.info("⚠️ Google Sheets 上傳: 跳過或失敗")
        
        if result.csv_exported:
            logger.info("✅ CSV 匯出: 成功")
            if result.csv_path:
                logger.info(f"💾 CSV 檔案: {result.csv_path}")
        
        # Statistics
        if result.coverage_stats:
            stats = result.coverage_stats
            logger.info(f"📈 統計資訊:")
            logger.info(f"   • 總公司數: {stats.total_companies}")
            logger.info(f"   • 有PDF公司: {stats.companies_with_pdfs}")
            logger.info(f"   • 涵蓋率: {stats.coverage_percentage:.1f}%")
            logger.info(f"   • 總報告數: {stats.total_actual_reports}")
        
        # Changes
        if result.stock_changes and result.stock_changes.has_changes:
            changes = result.stock_changes
            logger.info(f"🔄 股票清單變更:")
            if changes.added_companies:
                logger.info(f"   • 新增: {len(changes.added_companies)} 家公司")
            if changes.removed_companies:
                logger.info(f"   • 移除: {len(changes.removed_companies)} 家公司")
        
        # Performance
        logger.info(f"⏱️ 處理時間: {result.processing_time:.1f} 秒")
        logger.info(f"📁 處理檔案: {result.total_files_processed} 個 PDF")
        
        logger.info("=" * 50)

class QuickStart:
    """Quick start helper for common use cases"""
    
    @staticmethod
    def upload_to_existing_sheet(sheet_id: str, worksheet_name: str = "MOPS下載狀態") -> ProcessingResult:
        """
        Quick upload to existing Google Sheet
        
        Args:
            sheet_id: Google Sheets ID
            worksheet_name: Worksheet name to create/update
            
        Returns:
            ProcessingResult
        """
        config = load_config()
        config.google_sheet_id = sheet_id
        config.worksheet_name = worksheet_name
        
        uploader = MOPSSheetsUploader(config=config)
        return uploader.run()
    
    @staticmethod
    def csv_only_export(output_dir: str = "data/reports") -> ProcessingResult:
        """
        Quick CSV-only export (no Google Sheets)
        
        Args:
            output_dir: Directory for CSV output
            
        Returns:
            ProcessingResult
        """
        config = load_config()
        config.csv_backup = True
        config.google_sheet_id = None  # Disable Sheets upload
        config.output_dir = output_dir
        
        uploader = MOPSSheetsUploader(config=config)
        return uploader.run()
    
    @staticmethod
    def analyze_only(include_suggestions: bool = True) -> Dict[str, Any]:
        """
        Quick analysis without uploading
        
        Args:
            include_suggestions: Include download suggestions
            
        Returns:
            Analysis report dictionary
        """
        config = load_config()
        config.auto_suggest_downloads = include_suggestions
        
        uploader = MOPSSheetsUploader(config=config)
        return uploader.generate_report(include_analysis=True)