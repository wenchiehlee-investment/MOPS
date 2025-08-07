"""
MOPS Sheets Uploader - Package Initialization
A Python package for uploading MOPS PDF download status to Google Sheets matrix.
"""

__version__ = '1.0.0'
__author__ = 'MOPS Analysis Team'
__description__ = 'Upload MOPS PDF download status to Google Sheets matrix'

# Import main classes for easy access
from .main import MOPSSheetsUploader, QuickStart
from .config import MOPSConfig, load_config
from .models import (
    PDFFile, 
    MatrixCell, 
    CoverageStats, 
    ProcessingResult, 
    StockListChanges,
    FutureQuarterAnalysis
)

# Import component classes for advanced usage
from .pdf_scanner import PDFScanner
from .stock_data_loader import StockDataLoader
from .matrix_builder import MatrixBuilder
from .sheets_connector import MOPSSheetsConnector, SheetsUploadManager
from .report_analyzer import ReportAnalyzer

# Import CLI for programmatic access
from .cli import main as cli_main

# Package metadata
__all__ = [
    # Main classes
    'MOPSSheetsUploader',
    'QuickStart',
    
    # Configuration
    'MOPSConfig',
    'load_config',
    
    # Data models
    'PDFFile',
    'MatrixCell', 
    'CoverageStats',
    'ProcessingResult',
    'StockListChanges',
    'FutureQuarterAnalysis',
    
    # Component classes
    'PDFScanner',
    'StockDataLoader',
    'MatrixBuilder',
    'MOPSSheetsConnector',
    'SheetsUploadManager',
    'ReportAnalyzer',
    
    # CLI
    'cli_main',
    
    # Metadata
    '__version__',
    '__author__',
    '__description__'
]

# Quick access functions for common use cases
def upload_to_sheets(sheet_id: str, 
                    downloads_dir: str = "downloads",
                    stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                    worksheet_name: str = "MOPS下載狀態",
                    **kwargs) -> ProcessingResult:
    """
    Quick function to upload MOPS matrix to Google Sheets
    
    Args:
        sheet_id: Google Sheets ID
        downloads_dir: Path to downloads directory
        stock_csv_path: Path to stock CSV file
        worksheet_name: Name of worksheet to create/update
        **kwargs: Additional configuration options
        
    Returns:
        ProcessingResult with upload status and details
        
    Example:
        >>> from mops_sheets_uploader import upload_to_sheets
        >>> result = upload_to_sheets('1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0')
        >>> if result.success:
        ...     print(f"✅ 上傳成功: {result.sheets_url}")
    """
    config = load_config()
    config.google_sheet_id = sheet_id
    config.downloads_dir = downloads_dir
    config.stock_csv_path = stock_csv_path
    config.worksheet_name = worksheet_name
    
    # Apply additional configuration
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
    
    uploader = MOPSSheetsUploader(config=config)
    return uploader.run()

def export_to_csv(output_dir: str = "data/reports",
                 downloads_dir: str = "downloads", 
                 stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                 **kwargs) -> ProcessingResult:
    """
    Quick function to export MOPS matrix to CSV
    
    Args:
        output_dir: Directory for CSV output
        downloads_dir: Path to downloads directory
        stock_csv_path: Path to stock CSV file
        **kwargs: Additional configuration options
        
    Returns:
        ProcessingResult with export status and file path
        
    Example:
        >>> from mops_sheets_uploader import export_to_csv
        >>> result = export_to_csv('./reports')
        >>> if result.success:
        ...     print(f"💾 CSV 檔案: {result.csv_path}")
    """
    config = load_config()
    config.output_dir = output_dir
    config.downloads_dir = downloads_dir
    config.stock_csv_path = stock_csv_path
    config.csv_backup = True
    config.google_sheet_id = None  # Disable Sheets upload
    
    # Apply additional configuration
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
    
    uploader = MOPSSheetsUploader(config=config)
    return uploader.run()

def analyze_coverage(downloads_dir: str = "downloads",
                    stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                    include_suggestions: bool = True,
                    **kwargs) -> dict:
    """
    Quick function to analyze PDF coverage without uploading
    
    Args:
        downloads_dir: Path to downloads directory
        stock_csv_path: Path to stock CSV file
        include_suggestions: Whether to include download suggestions
        **kwargs: Additional configuration options
        
    Returns:
        Dictionary with comprehensive analysis report
        
    Example:
        >>> from mops_sheets_uploader import analyze_coverage
        >>> report = analyze_coverage()
        >>> coverage = report['summary']['coverage_percentage']
        >>> print(f"📊 涵蓋率: {coverage:.1f}%")
    """
    config = load_config()
    config.downloads_dir = downloads_dir
    config.stock_csv_path = stock_csv_path
    config.auto_suggest_downloads = include_suggestions
    
    # Apply additional configuration
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
    
    uploader = MOPSSheetsUploader(config=config)
    return uploader.generate_report(include_analysis=True)

def test_google_sheets_connection(sheet_id: str = None) -> bool:
    """
    Test Google Sheets connection
    
    Args:
        sheet_id: Optional Google Sheets ID to test specific sheet access
        
    Returns:
        True if connection successful, False otherwise
        
    Example:
        >>> from mops_sheets_uploader import test_google_sheets_connection
        >>> if test_google_sheets_connection():
        ...     print("✅ Google Sheets 連線正常")
        ... else:
        ...     print("❌ Google Sheets 連線失敗")
    """
    config = load_config()
    if sheet_id:
        config.google_sheet_id = sheet_id
    
    uploader = MOPSSheetsUploader(config=config)
    return uploader.test_connection()

# Add quick access functions to __all__
__all__.extend([
    'upload_to_sheets',
    'export_to_csv', 
    'analyze_coverage',
    'test_google_sheets_connection'
])

# Package information for help and documentation
def get_package_info() -> dict:
    """Get package information"""
    return {
        'name': 'mops_sheets_uploader',
        'version': __version__,
        'author': __author__,
        'description': __description__,
        'main_class': 'MOPSSheetsUploader',
        'quick_functions': [
            'upload_to_sheets',
            'export_to_csv',
            'analyze_coverage', 
            'test_google_sheets_connection'
        ],
        'supported_formats': ['Google Sheets', 'CSV'],
        'required_files': [
            'downloads/ (directory with PDF files)',
            'StockID_TWSE_TPEX.csv (stock list)'
        ],
        'environment_variables': [
            'GOOGLE_SHEETS_CREDENTIALS',
            'GOOGLE_SHEET_ID'
        ]
    }

def print_quick_help():
    """Print quick help information"""
    print(f"MOPS Sheets Uploader v{__version__}")
    print("=" * 40)
    print("📊 將 MOPS PDF 下載狀態上傳至 Google Sheets 矩陣")
    print()
    print("🚀 快速開始:")
    print("from mops_sheets_uploader import upload_to_sheets")
    print("result = upload_to_sheets('YOUR_SHEET_ID')")
    print()
    print("💡 其他功能:")
    print("• export_to_csv() - 匯出至 CSV")
    print("• analyze_coverage() - 分析涵蓋率")
    print("• test_google_sheets_connection() - 測試連線")
    print()
    print("📚 完整文件: 請參考 instructions-sheets-uploader.md")

# CLI entry point for python -m mops_sheets_uploader
def main():
    """Entry point for python -m mops_sheets_uploader"""
    from .cli import main as cli_main
    import sys
    sys.exit(cli_main())

# Version check and compatibility
import sys
if sys.version_info < (3, 9):
    import warnings
    warnings.warn(
        f"MOPS Sheets Uploader {__version__} requires Python 3.9+, "
        f"but you are running Python {sys.version_info.major}.{sys.version_info.minor}. "
        "Some features may not work correctly.",
        RuntimeWarning
    )