"""
MOPS Sheets Uploader - Enhanced Package Initialization v1.1.1
Complete package interface with multiple report type and font support
"""

import os  # FIX: Add this import at module level
import sys
from pathlib import Path

__version__ = '1.1.1'
__author__ = 'MOPS Analysis Team'
__description__ = 'Upload MOPS PDF download status to Google Sheets matrix with multiple report type and font support'

# Import main classes for easy access
from .main import MOPSSheetsUploader, QuickStart
from .config import MOPSConfig, load_config, create_font_config_preset
from .models import (
    PDFFile, 
    MatrixCell, 
    CoverageStats, 
    ProcessingResult, 
    StockListChanges,
    FutureQuarterAnalysis,
    create_font_config_preset as model_font_preset,
    get_default_display_config
)

# Import component classes for advanced usage
try:
    from .pdf_scanner import PDFScanner
except ImportError:
    PDFScanner = None

try:
    from .stock_data_loader import StockDataLoader
except ImportError:
    StockDataLoader = None

try:
    from .matrix_builder import MatrixBuilder
except ImportError:
    MatrixBuilder = None

try:
    from .sheets_connector import MOPSSheetsConnector, SheetsUploadManager
except ImportError:
    MOPSSheetsConnector = None
    SheetsUploadManager = None

try:
    from .report_analyzer import ReportAnalyzer
except ImportError:
    ReportAnalyzer = None

# Import CLI for programmatic access
try:
    from .cli import main as cli_main
except ImportError:
    cli_main = None

# Package metadata
__all__ = [
    # Main classes
    'MOPSSheetsUploader',
    'QuickStart',
    
    # Configuration
    'MOPSConfig',
    'load_config',
    'create_font_config_preset',
    
    # Data models
    'PDFFile',
    'MatrixCell', 
    'CoverageStats',
    'ProcessingResult',
    'StockListChanges',
    'FutureQuarterAnalysis',
    
    # Component classes (if available)
    'PDFScanner',
    'StockDataLoader',
    'MatrixBuilder',
    'MOPSSheetsConnector',
    'SheetsUploadManager',
    'ReportAnalyzer',
    
    # CLI
    'cli_main',
    
    # v1.1.1 Enhanced functions
    'upload_to_sheets',
    'upload_with_font_preset',
    'export_to_csv', 
    'analyze_coverage',
    'analyze_with_multiple_types',
    'test_google_sheets_connection',
    
    # Metadata
    '__version__',
    '__author__',
    '__description__'
]

# Enhanced quick access functions for common use cases

def upload_to_sheets(sheet_id: str, 
                    downloads_dir: str = "downloads",
                    stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                    worksheet_name: str = "MOPS下載狀態",
                    font_preset: str = "large",
                    show_all_types: bool = True,
                    **kwargs) -> ProcessingResult:
    """
    Enhanced quick function to upload MOPS matrix to Google Sheets (v1.1.1)
    
    Args:
        sheet_id: Google Sheets ID
        downloads_dir: Path to downloads directory
        stock_csv_path: Path to stock CSV file
        worksheet_name: Name of worksheet to create/update
        font_preset: Font preset ('small', 'normal', 'large', 'extra_large', 'huge')
        show_all_types: Enable multiple report type display
        **kwargs: Additional configuration options
        
    Returns:
        ProcessingResult with upload status and enhanced details
        
    Example:
        >>> from mops_sheets_uploader import upload_to_sheets
        >>> result = upload_to_sheets(
        ...     '1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0',
        ...     font_preset="large",
        ...     show_all_types=True
        ... )
        >>> if result.success:
        ...     print(f"✅ 上傳成功: {result.sheets_url}")
        ...     print(f"🔄 多重類型: {result.multiple_types_found} 個")
    """
    config = load_config()
    config.google_sheet_id = sheet_id
    config.downloads_dir = downloads_dir
    config.stock_csv_path = stock_csv_path
    config.worksheet_name = worksheet_name
    config.show_all_report_types = show_all_types
    
    # Apply additional configuration
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
    
    uploader = MOPSSheetsUploader(config=config, font_preset=font_preset)
    return uploader.run()

def upload_with_font_preset(sheet_id: str,
                           font_preset: str = "large", 
                           worksheet_name: str = "MOPS下載狀態",
                           **kwargs) -> ProcessingResult:
    """
    Quick upload with specific font preset (v1.1.1 convenience function)
    
    Args:
        sheet_id: Google Sheets ID
        font_preset: Font preset name ('small', 'normal', 'large', 'extra_large', 'huge')
        worksheet_name: Worksheet name
        **kwargs: Additional configuration
        
    Returns:
        ProcessingResult with enhanced v1.1.1 data
        
    Example:
        >>> from mops_sheets_uploader import upload_with_font_preset
        >>> result = upload_with_font_preset(
        ...     "your_sheet_id",
        ...     font_preset="extra_large"  # 16pt for better readability
        ... )
        >>> print(f"Font used: {result.font_config_used}")
    """
    return upload_to_sheets(
        sheet_id=sheet_id,
        font_preset=font_preset,
        worksheet_name=worksheet_name,
        **kwargs
    )

def export_to_csv(output_dir: str = "data/reports",
                 downloads_dir: str = "downloads", 
                 stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                 font_preset: str = "large",
                 show_all_types: bool = True,
                 **kwargs) -> ProcessingResult:
    """
    Enhanced quick function to export MOPS matrix to CSV (v1.1.1)
    
    Args:
        output_dir: Directory for CSV output
        downloads_dir: Path to downloads directory
        stock_csv_path: Path to stock CSV file
        font_preset: Font preset for consistency (stored in metadata)
        show_all_types: Enable multiple report type display
        **kwargs: Additional configuration options
        
    Returns:
        ProcessingResult with export status and file path
        
    Example:
        >>> from mops_sheets_uploader import export_to_csv
        >>> result = export_to_csv(
        ...     './reports',
        ...     show_all_types=True,
        ...     font_preset="large"
        ... )
        >>> if result.success:
        ...     print(f"💾 CSV 檔案: {result.csv_path}")
        ...     print(f"🔄 多重類型: {result.multiple_types_found} 個")
    """
    config = load_config()
    config.output_dir = output_dir
    config.downloads_dir = downloads_dir
    config.stock_csv_path = stock_csv_path
    config.csv_backup = True
    config.google_sheet_id = None  # Disable Sheets upload
    config.show_all_report_types = show_all_types
    
    # Apply additional configuration
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
    
    uploader = MOPSSheetsUploader(config=config, font_preset=font_preset)
    return uploader.run()

def analyze_coverage(downloads_dir: str = "downloads",
                    stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                    include_suggestions: bool = True,
                    show_all_types: bool = True,
                    **kwargs) -> dict:
    """
    Enhanced quick function to analyze PDF coverage without uploading (v1.1.1)
    
    Args:
        downloads_dir: Path to downloads directory
        stock_csv_path: Path to stock CSV file
        include_suggestions: Whether to include download suggestions
        show_all_types: Enable multiple report type analysis
        **kwargs: Additional configuration options
        
    Returns:
        Dictionary with comprehensive analysis report including v1.1.1 features
        
    Example:
        >>> from mops_sheets_uploader import analyze_coverage
        >>> report = analyze_coverage(show_all_types=True)
        >>> coverage = report['summary']['coverage_percentage']
        >>> multiple_types = report['summary']['multiple_types_found']
        >>> print(f"📊 涵蓋率: {coverage:.1f}%")
        >>> print(f"🔄 多重類型: {multiple_types} 個")
    """
    config = load_config()
    config.downloads_dir = downloads_dir
    config.stock_csv_path = stock_csv_path
    config.auto_suggest_downloads = include_suggestions
    config.show_all_report_types = show_all_types
    
    # Apply additional configuration
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
    
    uploader = MOPSSheetsUploader(config=config)
    return uploader.generate_enhanced_report(include_analysis=True)

def analyze_with_multiple_types(downloads_dir: str = "downloads",
                               stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                               font_preset: str = "large",
                               **kwargs) -> dict:
    """
    Enhanced analysis focusing on multiple report types (v1.1.1)
    
    Args:
        downloads_dir: Path to downloads directory
        stock_csv_path: Path to stock CSV file
        font_preset: Font preset for configuration consistency
        **kwargs: Additional configuration options
        
    Returns:
        Dictionary with detailed multiple type analysis
        
    Example:
        >>> from mops_sheets_uploader import analyze_with_multiple_types
        >>> report = analyze_with_multiple_types()
        >>> combinations = report['coverage']['type_combinations']
        >>> print(f"📊 最常見組合: {max(combinations.items(), key=lambda x: x[1])}")
    """
    config = load_config()
    config.downloads_dir = downloads_dir
    config.stock_csv_path = stock_csv_path
    config.show_all_report_types = True  # Force enable for this analysis
    config.auto_suggest_downloads = True
    
    # Apply font preset for consistency
    if font_preset:
        config.apply_font_preset(font_preset)
    
    # Apply additional configuration
    for key, value in kwargs.items():
        if hasattr(config, key):
            setattr(config, key, value)
    
    uploader = MOPSSheetsUploader(config=config)
    return uploader.generate_enhanced_report(include_analysis=True)

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

# Enhanced v1.1.1 utility functions

def get_available_font_presets() -> dict:
    """Get available font presets with descriptions"""
    presets = {}
    preset_names = ['small', 'normal', 'large', 'extra_large', 'huge']
    
    for preset_name in preset_names:
        preset_config = create_font_config_preset(preset_name)
        presets[preset_name] = {
            'font_size': preset_config['font_size'],
            'header_font_size': preset_config['header_font_size'],
            'description': preset_config['description']
        }
    
    return presets

def create_custom_font_config(font_size: int = 14,
                             header_font_size: int = 16,
                             bold_headers: bool = True,
                             bold_company_info: bool = True) -> dict:
    """Create custom font configuration"""
    return {
        'font_size': font_size,
        'header_font_size': header_font_size,
        'bold_headers': bold_headers,
        'bold_company_info': bold_company_info
    }

def validate_environment() -> dict:
    """Validate v1.1.1 environment setup"""
    validation = {
        'directories': {},
        'environment_variables': {},
        'files': {},
        'overall_status': True
    }
    
    # Check directories
    required_dirs = ['downloads', 'logs', 'data/reports']
    for directory in required_dirs:
        exists = Path(directory).exists()
        validation['directories'][directory] = exists
        if not exists:
            validation['overall_status'] = False
    
    # Check environment variables
    required_env = ['GOOGLE_SHEETS_CREDENTIALS', 'GOOGLE_SHEET_ID']
    for env_var in required_env:
        exists = bool(os.getenv(env_var))
        validation['environment_variables'][env_var] = exists
        if not exists:
            validation['overall_status'] = False
    
    # Check files
    required_files = ['StockID_TWSE_TPEX.csv']
    for file_name in required_files:
        exists = Path(file_name).exists()
        validation['files'][file_name] = exists
    
    return validation

def setup_environment():
    """Setup v1.1.1 environment with required directories"""
    # Create required directories
    directories = ['downloads', 'logs', 'data/reports', 'data/exports']
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
    
    # Set default environment variables if not present
    env_defaults = {
        'MOPS_SHOW_ALL_TYPES': 'true',
        'MOPS_FONT_SIZE': '14',
        'MOPS_HEADER_FONT_SIZE': '16',
        'MOPS_BOLD_HEADERS': 'true',
        'MOPS_BOLD_COMPANY_INFO': 'true',
        'MOPS_WORKSHEET_NAME': 'MOPS下載狀態'
    }
    
    set_vars = []
    for key, value in env_defaults.items():
        if key not in os.environ:
            os.environ[key] = value
            set_vars.append(key)
    
    return {
        'directories_created': directories,
        'environment_variables_set': set_vars
    }

# Package information for help and documentation
def get_package_info() -> dict:
    """Get enhanced package information"""
    return {
        'name': 'mops_sheets_uploader',
        'version': __version__,
        'author': __author__,
        'description': __description__,
        'main_class': 'MOPSSheetsUploader',
        'quick_functions': [
            'upload_to_sheets',
            'upload_with_font_preset',
            'export_to_csv',
            'analyze_coverage',
            'analyze_with_multiple_types', 
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
        ],
        'v1_1_1_features': [
            'Multiple report type display (A12/A13/AI1)',
            'Enhanced font configuration with presets',
            'Professional 14pt default styling',
            'Type combination analysis',
            'Configurable display separators'
        ],
        'font_presets': list(get_available_font_presets().keys())
    }

def print_quick_help():
    """Print enhanced quick help information for v1.1.1"""
    print(f"MOPS Sheets Uploader v{__version__} - Enhanced")
    print("=" * 50)
    print("📊 將 MOPS PDF 下載狀態上傳至 Google Sheets 矩陣")
    print("🆕 v1.1.1 新功能: 多重報告類型 + 字體設定")
    print()
    print("🚀 快速開始 (Enhanced):")
    print("from mops_sheets_uploader import upload_with_font_preset")
    print("result = upload_with_font_preset('YOUR_SHEET_ID', font_preset='large')")
    print()
    print("💡 v1.1.1 功能:")
    print("• upload_with_font_preset() - 字體預設上傳")
    print("• analyze_with_multiple_types() - 多重類型分析") 
    print("• export_to_csv() - 增強 CSV 匯出")
    print("• test_google_sheets_connection() - 連線測試")
    print()
    print("🎨 字體預設:")
    presets = get_available_font_presets()
    for preset_name, config in presets.items():
        print(f"• {preset_name}: {config['font_size']}pt - {config['description']}")
    print()
    print("📚 完整文件: 請參考 instructions-sheets-uploader.md v1.1.1")

# CLI entry point for python -m mops_sheets_uploader
def main():
    """Entry point for python -m mops_sheets_uploader"""
    if cli_main:
        sys.exit(cli_main())
    else:
        print("❌ CLI module not available")
        print("Please check that cli.py is present and properly configured")
        return 1

# Version check and compatibility
if sys.version_info < (3, 9):
    import warnings
    warnings.warn(
        f"MOPS Sheets Uploader {__version__} requires Python 3.9+, "
        f"but you are running Python {sys.version_info.major}.{sys.version_info.minor}. "
        "Some features may not work correctly.",
        RuntimeWarning
    )

# Auto-setup on import (optional)
def _auto_setup():
    """Automatically setup environment if needed"""
    # Only auto-create directories, don't modify environment
    required_dirs = ['downloads', 'logs']
    for directory in required_dirs:
        if not Path(directory).exists():
            Path(directory).mkdir(parents=True, exist_ok=True)

# Run auto-setup if enabled
if os.getenv('MOPS_AUTO_SETUP', 'false').lower() == 'true':
    _auto_setup()

# Add enhanced functions to __all__
__all__.extend([
    'upload_with_font_preset',
    'analyze_with_multiple_types',
    'get_available_font_presets',
    'create_custom_font_config',
    'validate_environment',
    'setup_environment',
    'get_package_info',
    'print_quick_help'
])