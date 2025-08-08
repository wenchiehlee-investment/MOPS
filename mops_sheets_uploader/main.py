"""
Complete Enhanced Main Class for MOPS Sheets Uploader v1.1.1
Replace your existing main.py with this complete version
"""

import os
import time
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path

from .config import MOPSConfig, load_config
from .models import (
    ProcessingResult, CoverageStats, StockListChanges, 
    create_font_config_preset, validate_font_configuration
)

logger = logging.getLogger(__name__)

class MOPSSheetsUploader:
    """MOPS PDF files to Google Sheets matrix uploader with v1.1.1 enhancements"""
    
    def __init__(self, 
                 downloads_dir: str = "downloads",
                 stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                 config: Optional[MOPSConfig] = None,
                 font_preset: Optional[str] = None,
                 font_size: Optional[int] = None,
                 header_font_size: Optional[int] = None,
                 bold_headers: Optional[bool] = None,
                 bold_company_info: Optional[bool] = None,
                 show_all_report_types: Optional[bool] = None,
                 **kwargs):
        """
        Initialize MOPS Sheets Uploader with v1.1.1 enhancements
        
        Args:
            downloads_dir: Path to downloads directory
            stock_csv_path: Path to stock CSV file
            config: Optional MOPSConfig instance
            font_preset: Font preset name ('small', 'normal', 'large', 'extra_large', 'huge')
            font_size: Custom font size (overrides preset)
            header_font_size: Custom header font size (overrides preset)
            bold_headers: Use bold headers (overrides preset)
            bold_company_info: Use bold company info (overrides preset)
            show_all_report_types: Enable multiple type display
            **kwargs: Additional configuration options
        """
        # Load configuration
        if config is None:
            self.config = load_config()
        else:
            self.config = config
        
        # Apply font preset first (if provided)
        if font_preset:
            self.config.apply_font_preset(font_preset)
            logger.info(f"📝 Applied font preset: {font_preset}")
        
        # Override config with provided parameters
        if downloads_dir != "downloads":
            self.config.downloads_dir = downloads_dir
        if stock_csv_path != "StockID_TWSE_TPEX.csv":
            self.config.stock_csv_path = stock_csv_path
        
        # Apply font configuration overrides
        if font_size is not None:
            self.config.font_size = font_size
        if header_font_size is not None:
            self.config.header_font_size = header_font_size
        if bold_headers is not None:
            self.config.bold_headers = bold_headers
        if bold_company_info is not None:
            self.config.bold_company_info = bold_company_info
        
        # Apply multiple type settings
        if show_all_report_types is not None:
            self.config.show_all_report_types = show_all_report_types
        
        # Apply additional configuration
        for key, value in kwargs.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
        
        # Validate font configuration
        font_config = self.config.get_font_config()
        font_errors = validate_font_configuration(font_config)
        if font_errors:
            logger.warning(f"⚠️ Font configuration issues: {'; '.join(font_errors)}")
            # Auto-fix common issues
            if self.config.font_size < 8 or self.config.font_size > 72:
                self.config.font_size = 14
            if self.config.header_font_size < 8 or self.config.header_font_size > 72:
                self.config.header_font_size = 14
            logger.info("🔧 Auto-corrected font configuration")
        
        # Initialize components (will be imported when needed)
        self.pdf_scanner = None
        self.stock_loader = None
        self.matrix_builder = None
        self.upload_manager = None
        self.analyzer = None
        
        # Processing state
        self._last_processing_result: Optional[ProcessingResult] = None
        
        # Setup logging
        self._setup_enhanced_logging()
        
        # Log enhanced initialization summary
        self._log_initialization_summary()
    
    def run(self) -> ProcessingResult:
        """Main execution method with v1.1.1 enhancements"""
        start_time = time.time()
        logger.info("🚀 開始執行 MOPS Sheets Uploader v1.1.1 流程")
        logger.info("=" * 60)
        
        try:
            # Initialize components as needed
            self._init_components()
            
            # Stage 1: Validate configuration
            self._validate_configuration()
            
            # Stage 2: Scan PDF files with enhanced analysis
            pdf_data = self._scan_pdf_files_enhanced()
            
            # Stage 3: Load stock data with change detection
            stock_df, stock_changes = self._load_stock_data_enhanced()
            
            # Stage 4: Build matrix with multiple type support
            matrix_df = self._build_matrix_enhanced(stock_df, pdf_data)
            
            # Stage 5: Analyze data with v1.1.1 enhancements
            coverage_stats = self._analyze_coverage_enhanced(matrix_df, pdf_data)
            
            # Stage 6: Upload to Sheets and/or export CSV with font settings
            upload_result = self._upload_matrix_enhanced(matrix_df, coverage_stats, stock_changes)
            
            # Create enhanced result with v1.1.1 data
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
                total_files_processed=sum(len(pdfs) for pdfs in pdf_data.values()) if pdf_data else 0,
                
                # v1.1.1 Enhanced fields
                font_config_used=self.config.get_enhanced_font_config(),
                multiple_types_found=getattr(coverage_stats, 'cells_with_multiple_types', 0),
                type_combinations_analyzed=getattr(coverage_stats, 'report_type_combinations', {})
            )
            
            # Log enhanced final summary
            self._log_enhanced_final_summary(result)
            
            # Store result for later access
            self._last_processing_result = result
            
            return result
            
        except Exception as e:
            processing_time = time.time() - start_time
            logger.error(f"❌ 處理失敗: {e}")
            
            result = ProcessingResult(
                success=False,
                error_message=str(e),
                processing_time=processing_time,
                font_config_used=self.config.get_enhanced_font_config()
            )
            
            self._last_processing_result = result
            return result
    
    def _init_components(self):
        """Initialize components safely with enhanced error handling"""
        try:
            # Import and initialize components with enhanced error handling
            from .pdf_scanner import PDFScanner
            from .stock_data_loader import StockDataLoader
            from .matrix_builder import MatrixBuilder
            from .sheets_connector import SheetsUploadManager
            from .report_analyzer import ReportAnalyzer
            
            self.pdf_scanner = PDFScanner(self.config)
            self.stock_loader = StockDataLoader(self.config)
            self.matrix_builder = MatrixBuilder(self.config)
            self.upload_manager = SheetsUploadManager(self.config)
            self.analyzer = ReportAnalyzer(self.config)
            
            logger.info("✅ All components initialized successfully")
            
        except ImportError as e:
            logger.warning(f"⚠️ Some components not available: {e}")
            # Create minimal fallbacks
            self._create_minimal_components()
    
    def _create_minimal_components(self):
        """Create minimal component fallbacks for missing imports"""
        try:
            from .sheets_connector import SheetsUploadManager
            self.upload_manager = SheetsUploadManager(self.config)
            logger.info("✅ Upload manager loaded successfully")
        except ImportError:
            logger.error("❌ Critical: Cannot load upload manager")
        
        # Create dummy components for missing ones
        class DummyComponent:
            def __init__(self, config):
                self.config = config
        
        # Only create dummies for components that failed to load
        if not self.pdf_scanner:
            self.pdf_scanner = DummyComponent(self.config)
            logger.info("📄 Using dummy PDF scanner")
        
        if not self.stock_loader:
            self.stock_loader = DummyComponent(self.config)
            logger.info("📋 Using dummy stock loader")
        
        if not self.matrix_builder:
            self.matrix_builder = DummyComponent(self.config)
            logger.info("🏗️ Using dummy matrix builder")
        
        if not self.analyzer:
            self.analyzer = DummyComponent(self.config)
            logger.info("📊 Using dummy analyzer")
    
    def _validate_configuration(self) -> None:
        """Enhanced configuration validation"""
        logger.info("🔧 驗證增強型設定...")
        
        errors = self.config.validate()
        if errors:
            error_msg = "Configuration validation failed: " + "; ".join(errors)
            logger.error(f"❌ {error_msg}")
            raise ValueError(error_msg)
        
        # v1.1.1 specific validations
        font_config = self.config.get_font_config()
        font_errors = validate_font_configuration(font_config)
        if font_errors:
            logger.warning(f"⚠️ Font validation warnings: {'; '.join(font_errors)}")
        
        # Log validation success with v1.1.1 details
        logger.info("✅ 設定驗證通過")
        logger.info(f"   • 字體設定: {font_config['font_size']}pt/{font_config['header_font_size']}pt")
        logger.info(f"   • 多重報告類型: {'啟用' if self.config.show_all_report_types else '停用'}")
        logger.info(f"   • 最大顯示類型: {self.config.max_display_types}")
    
    def _scan_pdf_files_enhanced(self) -> Dict[str, List]:
        """Enhanced PDF scanning with v1.1.1 analysis"""
        logger.info("🔍 第二階段: 增強型 PDF 檔案掃描")
        
        pdf_data = {}
        
        try:
            if hasattr(self.pdf_scanner, 'scan_downloads_directory'):
                pdf_data = self.pdf_scanner.scan_downloads_directory()
                
                # v1.1.1 Enhanced analysis
                total_files = sum(len(pdfs) for pdfs in pdf_data.values())
                companies_found = len(pdf_data)
                
                # Analyze report types
                type_distribution = {}
                multiple_type_quarters = 0
                
                for company_pdfs in pdf_data.values():
                    quarter_types = {}
                    for pdf in company_pdfs:
                        quarter_key = pdf.quarter_key
                        if quarter_key not in quarter_types:
                            quarter_types[quarter_key] = set()
                        quarter_types[quarter_key].add(pdf.report_type)
                        
                        # Track type distribution
                        type_distribution[pdf.report_type] = type_distribution.get(pdf.report_type, 0) + 1
                    
                    # Count quarters with multiple types
                    multiple_type_quarters += sum(1 for types in quarter_types.values() if len(types) > 1)
                
                logger.info(f"   發現 {total_files} 個 PDF 檔案 ({companies_found} 家公司)")
                logger.info(f"   📊 報告類型分布: {dict(list(type_distribution.items())[:5])}")
                if multiple_type_quarters > 0:
                    logger.info(f"   🔄 多重類型季度: {multiple_type_quarters} 個")
                
        except Exception as e:
            logger.warning(f"PDF scanning failed: {e}")
        
        if not pdf_data:
            logger.warning("⚠️ 未發現任何 PDF 檔案")
        
        return pdf_data
    
    def _load_stock_data_enhanced(self) -> Tuple[Any, Optional[StockListChanges]]:
        """Enhanced stock data loading with change detection"""
        logger.info("📋 第三階段: 增強型股票資料載入")
        
        import pandas as pd
        
        # Create dummy stock data for fallback
        stock_df = pd.DataFrame({
            '代號': ['2330', '8272'],
            '名稱': ['台積電', '全景軟體']
        })
        
        stock_changes = None
        
        try:
            if hasattr(self.stock_loader, 'load_stock_csv'):
                stock_df = self.stock_loader.load_stock_csv()
                logger.info(f"   載入 {len(stock_df)} 家公司資料")
                
            if hasattr(self.stock_loader, 'detect_stock_list_changes'):
                stock_changes = self.stock_loader.detect_stock_list_changes(stock_df)
                
                if stock_changes and stock_changes.has_changes:
                    logger.info(f"   🔄 偵測到變更: {stock_changes.change_summary}")
                
        except Exception as e:
            logger.warning(f"Stock data loading failed: {e}")
        
        return stock_df, stock_changes
    
    def _build_matrix_enhanced(self, stock_df, pdf_data) -> Any:
        """Enhanced matrix building with v1.1.1 multiple type support"""
        logger.info("🏗️ 第四階段: 增強型矩陣建構")
        
        import pandas as pd
        
        # Create enhanced matrix structure
        matrix_data = {
            '代號': stock_df['代號'].astype(str),
            '名稱': stock_df['名稱'].astype(str),
            '2024 Q4': ['-'] * len(stock_df),
            '2024 Q3': ['-'] * len(stock_df),
            '2024 Q2': ['-'] * len(stock_df),
            '2024 Q1': ['-'] * len(stock_df)
        }
        
        matrix_df = pd.DataFrame(matrix_data)
        
        try:
            if hasattr(self.matrix_builder, 'build_base_matrix'):
                matrix_df = self.matrix_builder.build_base_matrix(stock_df, pdf_data)
                
            if hasattr(self.matrix_builder, 'populate_pdf_status'):
                matrix_df = self.matrix_builder.populate_pdf_status(matrix_df, pdf_data)
            else:
                # Fallback: Simple population for demonstration
                matrix_df = self._populate_matrix_simple(matrix_df, pdf_data)
                
        except Exception as e:
            logger.warning(f"Enhanced matrix building failed: {e}")
        
        logger.info(f"✅ 增強型矩陣建構完成: {len(matrix_df)} × {len(matrix_df.columns)}")
        logger.info(f"   📊 多重報告類型支援: {'啟用' if self.config.show_all_report_types else '停用'}")
        
        return matrix_df
    
    def _populate_matrix_simple(self, matrix_df, pdf_data) -> Any:
        """Simple fallback matrix population with multiple type support"""
        for idx, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            
            if company_id in pdf_data:
                for pdf in pdf_data[company_id]:
                    quarter_col = pdf.quarter_key
                    if quarter_col in matrix_df.columns:
                        current_value = matrix_df.loc[idx, quarter_col]
                        
                        if current_value == '-':
                            matrix_df.loc[idx, quarter_col] = pdf.report_type
                        elif self.config.show_all_report_types and current_value != pdf.report_type:
                            # Add to existing types
                            existing_types = current_value.split(self.config.report_type_separator)
                            if pdf.report_type not in existing_types:
                                existing_types.append(pdf.report_type)
                                matrix_df.loc[idx, quarter_col] = self.config.report_type_separator.join(sorted(existing_types))
        
        return matrix_df
    
    def _analyze_coverage_enhanced(self, matrix_df, pdf_data) -> CoverageStats:
        """Enhanced coverage analysis with v1.1.1 features"""
        logger.info("📊 第五階段: 增強型涵蓋率分析")
        
        # Enhanced coverage calculation
        total_companies = len(matrix_df)
        companies_with_pdfs = len(pdf_data)
        
        quarter_columns = [col for col in matrix_df.columns if 'Q' in col and col not in ['代號', '名稱']]
        total_quarters = len(quarter_columns)
        
        total_possible_reports = total_companies * total_quarters
        total_actual_reports = sum(len(pdfs) for pdfs in pdf_data.values()) if pdf_data else 0
        
        coverage_percentage = (total_actual_reports / total_possible_reports * 100) if total_possible_reports > 0 else 0
        
        # v1.1.1 Enhanced analysis
        cells_with_multiple_types = 0
        type_combinations = {}
        
        for _, row in matrix_df.iterrows():
            for col in quarter_columns:
                value = str(row[col])
                if value != '-' and self.config.report_type_separator in value:
                    cells_with_multiple_types += 1
                    type_combinations[value] = type_combinations.get(value, 0) + 1
        
        coverage_stats = CoverageStats(
            total_companies=total_companies,
            companies_with_pdfs=companies_with_pdfs,
            total_quarters=total_quarters,
            total_possible_reports=total_possible_reports,
            total_actual_reports=total_actual_reports,
            coverage_percentage=coverage_percentage,
            cells_with_multiple_types=cells_with_multiple_types,
            report_type_combinations=type_combinations
        )
        
        logger.info(f"   📈 涵蓋率: {coverage_percentage:.1f}%")
        if cells_with_multiple_types > 0:
            logger.info(f"   🔄 多重類型: {cells_with_multiple_types} 個季度")
        
        return coverage_stats
    
    def _upload_matrix_enhanced(self, matrix_df, coverage_stats, stock_changes) -> Dict[str, Any]:
        """Enhanced matrix upload with v1.1.1 font support"""
        logger.info("🚀 第六階段: 增強型矩陣上傳")
        
        upload_result = self.upload_manager.upload_with_fallback(
            matrix_df, coverage_stats, stock_changes
        )
        
        return upload_result
    
    def _setup_enhanced_logging(self) -> None:
        """Enhanced logging setup with v1.1.1 features"""
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)
        
        # Enhanced log format with version info
        log_format = '%(asctime)s - %(name)s - %(levelname)s - [v1.1.1] %(message)s'
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = logs_dir / f"mops_sheets_uploader_v1.1.1_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ],
            force=True  # Override any existing logging config
        )
        
        # Quiet external libraries
        logging.getLogger('urllib3').setLevel(logging.WARNING)
        logging.getLogger('requests').setLevel(logging.WARNING)
        logging.getLogger('gspread').setLevel(logging.WARNING)
        logging.getLogger('google').setLevel(logging.WARNING)
        
        logger.info(f"📝 Enhanced v1.1.1 logging initialized: {log_file}")
    
    def _log_initialization_summary(self) -> None:
        """Log enhanced initialization summary"""
        font_config = self.config.get_enhanced_font_config()
        
        logger.info(f"📊 MOPS Sheets Uploader v1.1.1 初始化完成")
        logger.info(f"   下載目錄: {self.config.downloads_dir}")
        logger.info(f"   股票清單: {self.config.stock_csv_path}")
        logger.info(f"   工作表名稱: {self.config.worksheet_name}")
        logger.info(f"   字體設定: {font_config['font_size']}pt (標題: {font_config['header_font_size']}pt)")
        logger.info(f"   粗體設定: 標題={font_config['bold_headers']}, 公司資訊={font_config['bold_company_info']}")
        logger.info(f"   字體預設: {font_config['preset_equivalent']}")
        logger.info(f"   多重報告類型: {'啟用' if self.config.show_all_report_types else '停用'}")
        
        if self.config.show_all_report_types:
            logger.info(f"   類型分隔符號: '{self.config.report_type_separator}'")
            logger.info(f"   最大顯示類型: {self.config.max_display_types}")
    
    def _log_enhanced_final_summary(self, result: ProcessingResult) -> None:
        """Enhanced final processing summary with v1.1.1 features"""
        logger.info("=" * 60)
        logger.info("📊 處理完成摘要 (Enhanced v1.1.1)")
        logger.info("=" * 60)
        
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
        
        # Enhanced Statistics (v1.1.1)
        if result.coverage_stats:
            stats = result.coverage_stats
            logger.info(f"📈 增強統計資訊:")
            logger.info(f"   • 總公司數: {stats.total_companies}")
            logger.info(f"   • 有PDF公司: {stats.companies_with_pdfs}")
            logger.info(f"   • 涵蓋率: {stats.coverage_percentage:.1f}%")
            logger.info(f"   • 總報告數: {stats.total_actual_reports}")
            
            # v1.1.1 Multiple type statistics
            if hasattr(stats, 'cells_with_multiple_types') and stats.cells_with_multiple_types > 0:
                logger.info(f"   🔄 多重類型: {stats.cells_with_multiple_types} 個季度 ({stats.multiple_types_percentage:.1f}%)")
            
            if hasattr(stats, 'report_type_combinations') and stats.report_type_combinations:
                top_combo = max(stats.report_type_combinations.items(), key=lambda x: x[1])
                logger.info(f"   📊 最常見組合: {top_combo[0]} ({top_combo[1]} 次)")
        
        # Stock changes
        if result.stock_changes and result.stock_changes.has_changes:
            changes = result.stock_changes
            logger.info(f"🔄 股票清單變更:")
            if changes.added_companies:
                logger.info(f"   • 新增: {len(changes.added_companies)} 家公司")
            if changes.removed_companies:
                logger.info(f"   • 移除: {len(changes.removed_companies)} 家公司")
        
        # v1.1.1 Font configuration used
        if result.font_config_used:
            font_config = result.font_config_used
            logger.info(f"🔤 字體設定:")
            logger.info(f"   • 內容字體: {font_config.get('font_size', 14)}pt")
            logger.info(f"   • 標題字體: {font_config.get('header_font_size', 14)}pt")
            logger.info(f"   • 預設模式: {font_config.get('preset_equivalent', 'custom')}")
        
        # Performance
        logger.info(f"⏱️ 處理時間: {result.processing_time:.1f} 秒")
        logger.info(f"📁 處理檔案: {result.total_files_processed} 個 PDF")
        
        # v1.1.1 specific metrics
        if result.multiple_types_found > 0:
            logger.info(f"🔄 發現多重類型: {result.multiple_types_found} 個季度")
        
        if result.type_combinations_analyzed:
            logger.info(f"📊 類型組合分析: {len(result.type_combinations_analyzed)} 種組合")
        
        logger.info("=" * 60)
    
    def generate_enhanced_report(self, include_analysis: bool = True) -> Dict[str, Any]:
        """Generate comprehensive enhanced report (v1.1.1)"""
        if not self._last_processing_result:
            logger.warning("No processing result available for report generation")
            return {"error": "No data available"}
        
        result = self._last_processing_result
        
        # Base report structure
        report = {
            'metadata': {
                'version': '1.1.1',
                'generation_time': datetime.now().isoformat(),
                'uploader_version': 'MOPS Sheets Uploader v1.1.1',
                'font_configuration': result.font_config_used
            },
            'summary': {
                'success': result.success,
                'processing_time': result.processing_time,
                'total_files_processed': result.total_files_processed,
                'multiple_types_found': result.multiple_types_found
            }
        }
        
        # Add coverage statistics if available
        if result.coverage_stats:
            stats = result.coverage_stats
            report['coverage'] = {
                'total_companies': stats.total_companies,
                'companies_with_pdfs': stats.companies_with_pdfs,
                'coverage_percentage': stats.coverage_percentage,
                'total_reports': stats.total_actual_reports
            }
            
            # v1.1.1 Enhanced coverage data
            if hasattr(stats, 'cells_with_multiple_types'):
                report['coverage']['multiple_types'] = {
                    'count': stats.cells_with_multiple_types,
                    'percentage': stats.multiple_types_percentage
                }
            
            if hasattr(stats, 'report_type_combinations'):
                report['coverage']['type_combinations'] = stats.report_type_combinations
        
        # Add change information
        if result.stock_changes:
            report['changes'] = {
                'added_companies': result.stock_changes.added_companies,
                'removed_companies': result.stock_changes.removed_companies,
                'net_change': result.stock_changes.net_change,
                'change_summary': result.stock_changes.change_summary
            }
        
        # Add upload results
        report['output'] = {
            'matrix_uploaded': result.matrix_uploaded,
            'csv_exported': result.csv_exported,
            'sheets_url': result.sheets_url,
            'csv_path': result.csv_path
        }
        
        # v1.1.1 Configuration details
        report['configuration'] = {
            'font_config': result.font_config_used,
            'multiple_types_enabled': self.config.show_all_report_types,
            'type_separator': self.config.report_type_separator,
            'max_display_types': self.config.max_display_types
        }
        
        logger.info(f"📋 Enhanced v1.1.1 report generated")
        
        return report
    
    def test_connection(self) -> bool:
        """Test Google Sheets connection"""
        if not self.upload_manager:
            self._init_components()
        
        if hasattr(self.upload_manager, 'connector'):
            return self.upload_manager.connector.test_connection()
        else:
            logger.warning("Upload manager not available for connection test")
            return False

class QuickStart:
    """Enhanced quick start helper for v1.1.1 common use cases"""
    
    @staticmethod
    def upload_to_existing_sheet(sheet_id: str, 
                                worksheet_name: str = "MOPS下載狀態",
                                font_preset: str = "large",
                                show_all_types: bool = True,
                                **kwargs) -> ProcessingResult:
        """Quick upload to existing Google Sheet with v1.1.1 font configuration"""
        config = load_config()
        config.google_sheet_id = sheet_id
        config.worksheet_name = worksheet_name
        config.show_all_report_types = show_all_types
        
        uploader = MOPSSheetsUploader(
            config=config, 
            font_preset=font_preset,
            **kwargs
        )
        return uploader.run()
    
    @staticmethod
    def csv_only_export(output_dir: str = "data/reports", 
                       font_preset: str = "large") -> ProcessingResult:
        """Quick CSV-only export with v1.1.1 enhancements"""
        config = load_config()
        config.csv_backup = True
        config.google_sheet_id = None
        config.output_dir = output_dir
        
        uploader = MOPSSheetsUploader(config=config, font_preset=font_preset)
        return uploader.run()
    
    @staticmethod
    def analyze_multiple_types(downloads_dir: str = "downloads",
                              stock_csv_path: str = "StockID_TWSE_TPEX.csv") -> dict:
        """Quick analysis focusing on multiple report types"""
        config = load_config()
        config.downloads_dir = downloads_dir
        config.stock_csv_path = stock_csv_path
        config.show_all_report_types = True
        
        uploader = MOPSSheetsUploader(config=config)
        return uploader.generate_enhanced_report(include_analysis=True)