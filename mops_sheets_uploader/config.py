"""
Complete Enhanced Configuration for MOPS Sheets Uploader v1.1.1
Replace your existing config.py with this complete version
"""

import os
import json
from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional

@dataclass
class MOPSConfig:
    """Complete configuration for MOPS Sheets Uploader v1.1.1"""
    
    # Directory settings
    downloads_dir: str = "downloads"
    stock_csv_path: str = "StockID_TWSE_TPEX.csv"
    output_dir: str = "data/reports"
    
    # Matrix settings
    max_years: int = 3
    include_current_year: bool = True
    quarter_order: str = "desc"  # newest first
    dynamic_columns: bool = True
    dynamic_rows: bool = True
    
    # v1.1.1 Enhanced Multiple Report Type Display Settings
    show_all_report_types: bool = True
    report_type_separator: str = "/"
    category_separator: str = " → "
    max_display_types: int = 5
    use_categorized_display: bool = False
    priority_display_mode: str = "all"  # 'all', 'best', 'category_best'
    show_type_counts: bool = False
    truncate_indicator: str = "+"
    exclude_english_reports: bool = True
    individual_reports_priority: bool = True
    
    # v1.1.1 Enhanced Font and Display Settings
    font_size: int = 14
    header_font_size: int = 14
    bold_headers: bool = True
    bold_company_info: bool = True
    
    # Stock list change handling
    detect_stock_changes: bool = True
    highlight_new_companies: bool = True
    archive_orphaned_pdfs: bool = False
    warn_on_removals: bool = True
    auto_suggest_downloads: bool = True
    change_threshold_warning: int = 5
    
    # Future quarter handling
    include_future_quarters: bool = True
    warn_threshold_months: int = 6
    mark_suspicious: bool = True
    max_future_quarters: int = 8
    generate_future_report: bool = True
    
    # Google Sheets settings
    worksheet_name: str = "MOPS下載狀態"
    include_summary_sheet: bool = True
    auto_resize_columns: bool = True
    highlight_future_quarters: bool = True
    highlight_multiple_types: bool = True
    freeze_company_columns: bool = True
    apply_conditional_formatting: bool = True
    
    # CSV export settings
    csv_backup: bool = True
    csv_filename_pattern: str = "mops_matrix_{timestamp}.csv"
    include_future_analysis: bool = True
    include_type_analysis: bool = True
    
    # Report type settings
    preferred_types: List[str] = field(default_factory=lambda: ["A12", "A13"])
    acceptable_types: List[str] = field(default_factory=lambda: ["AI1", "A1L"])
    excluded_types: List[str] = field(default_factory=lambda: ["AIA", "AE2"])
    
    # Display settings
    use_symbols: bool = True
    show_file_counts: bool = False
    highlight_missing: bool = True
    future_quarter_symbol: str = "⏭️"
    multiple_types_symbol: str = "🔄"
    mixed_categories_symbol: str = "📊"
    
    # Google Sheets Integration
    google_sheet_id: Optional[str] = None
    google_credentials: Optional[str] = None
    
    @classmethod
    def from_env(cls) -> 'MOPSConfig':
        """Load configuration from environment variables with v1.1.1 enhancements"""
        config = cls()
        
        # Basic settings
        config.downloads_dir = os.getenv('MOPS_DOWNLOADS_DIR', config.downloads_dir)
        config.stock_csv_path = os.getenv('MOPS_STOCK_CSV_PATH', config.stock_csv_path)
        config.output_dir = os.getenv('MOPS_OUTPUT_DIR', config.output_dir)
        config.max_years = int(os.getenv('MOPS_MAX_YEARS', str(config.max_years)))
        config.worksheet_name = os.getenv('MOPS_WORKSHEET_NAME', config.worksheet_name)
        
        # v1.1.1 Enhanced Multiple Report Type Settings
        config.show_all_report_types = os.getenv('MOPS_SHOW_ALL_TYPES', 'true').lower() == 'true'
        config.report_type_separator = os.getenv('MOPS_TYPE_SEPARATOR', config.report_type_separator)
        config.category_separator = os.getenv('MOPS_CATEGORY_SEPARATOR', config.category_separator)
        config.max_display_types = int(os.getenv('MOPS_MAX_DISPLAY_TYPES', str(config.max_display_types)))
        config.use_categorized_display = os.getenv('MOPS_CATEGORIZED_DISPLAY', 'false').lower() == 'true'
        config.priority_display_mode = os.getenv('MOPS_PRIORITY_MODE', config.priority_display_mode)
        
        # v1.1.1 Enhanced Font Settings
        config.font_size = int(os.getenv('MOPS_FONT_SIZE', str(config.font_size)))
        config.header_font_size = int(os.getenv('MOPS_HEADER_FONT_SIZE', str(config.header_font_size)))
        config.bold_headers = os.getenv('MOPS_BOLD_HEADERS', 'true').lower() == 'true'
        config.bold_company_info = os.getenv('MOPS_BOLD_COMPANY_INFO', 'true').lower() == 'true'
        
        # Google Sheets settings
        config.google_sheet_id = os.getenv('GOOGLE_SHEET_ID')
        config.google_credentials = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
        
        # Boolean settings
        config.csv_backup = os.getenv('MOPS_CSV_BACKUP', 'true').lower() == 'true'
        config.detect_stock_changes = os.getenv('MOPS_DETECT_CHANGES', 'true').lower() == 'true'
        config.include_future_quarters = os.getenv('MOPS_INCLUDE_FUTURE', 'true').lower() == 'true'
        config.highlight_multiple_types = os.getenv('MOPS_HIGHLIGHT_MULTIPLE', 'true').lower() == 'true'
        config.auto_suggest_downloads = os.getenv('MOPS_AUTO_SUGGEST', 'true').lower() == 'true'
        
        return config
    
    @classmethod
    def from_file(cls, config_path: str) -> 'MOPSConfig':
        """Load configuration from YAML/JSON file"""
        try:
            import yaml
        except ImportError:
            yaml = None
        
        with open(config_path, 'r', encoding='utf-8') as f:
            if config_path.endswith('.json'):
                data = json.load(f)
            elif yaml and config_path.endswith(('.yaml', '.yml')):
                data = yaml.safe_load(f)
            else:
                raise ValueError("Unsupported config file format. Use .json or .yaml")
        
        config = cls()
        for key, value in data.items():
            if hasattr(config, key):
                setattr(config, key, value)
        
        return config
    
    def validate(self) -> List[str]:
        """Validate configuration and return list of errors"""
        errors = []
        
        # Directory validation
        if not os.path.exists(self.downloads_dir):
            errors.append(f"Downloads directory not found: {self.downloads_dir}")
        
        if not os.path.exists(self.stock_csv_path):
            errors.append(f"Stock CSV file not found: {self.stock_csv_path}")
        
        # Google Sheets validation
        if not self.csv_backup and not self.google_credentials:
            errors.append("Google Sheets credentials required when CSV backup disabled")
        
        # Range validations
        if self.max_years < 1 or self.max_years > 10:
            errors.append("max_years must be between 1 and 10")
        
        # v1.1.1 Font validation
        if self.font_size < 8 or self.font_size > 72:
            errors.append("font_size must be between 8 and 72 pt")
        
        if self.header_font_size < 8 or self.header_font_size > 72:
            errors.append("header_font_size must be between 8 and 72 pt")
        
        # v1.1.1 Multiple type validation
        if self.max_display_types < 1 or self.max_display_types > 20:
            errors.append("max_display_types must be between 1 and 20")
        
        if len(self.report_type_separator) > 5:
            errors.append("report_type_separator should be 5 characters or less")
        
        if self.priority_display_mode not in ['all', 'best', 'category_best']:
            errors.append("priority_display_mode must be 'all', 'best', or 'category_best'")
        
        return errors
    
    def get_quarter_columns(self, discovered_quarters: List[str]) -> List[str]:
        """Generate quarter column headers based on configuration"""
        from datetime import datetime
        
        current_year = datetime.now().year
        current_quarter = (datetime.now().month - 1) // 3 + 1
        
        # Start with discovered quarters
        all_quarters = set(discovered_quarters)
        
        # Add expected quarters within max_years range
        for year in range(current_year, current_year - self.max_years, -1):
            for quarter in [4, 3, 2, 1]:
                if year < current_year or (year == current_year and quarter <= current_quarter):
                    all_quarters.add(f"{year} Q{quarter}")
        
        # Sort based on quarter_order setting
        def quarter_sort_key(q):
            try:
                year, quarter = q.split(' Q')
                return (int(year), int(quarter))
            except:
                return (0, 0)
        
        sorted_quarters = sorted(all_quarters, 
                               key=quarter_sort_key, 
                               reverse=(self.quarter_order == "desc"))
        
        return sorted_quarters
    
    def get_font_config(self) -> Dict[str, Any]:
        """Get font configuration for Google Sheets formatting"""
        return {
            'font_size': self.font_size,
            'header_font_size': self.header_font_size,
            'bold_headers': self.bold_headers,
            'bold_company_info': self.bold_company_info
        }
    
    def get_enhanced_font_config(self) -> Dict[str, Any]:
        """Get enhanced font configuration with metadata (v1.1.1)"""
        return {
            'font_size': self.font_size,
            'header_font_size': self.header_font_size,
            'bold_headers': self.bold_headers,
            'bold_company_info': self.bold_company_info,
            'preset_equivalent': self._detect_preset(),
            'is_default': self._is_default_font_config()
        }
    
    def _detect_preset(self) -> str:
        """Detect which preset best matches current configuration"""
        from .models import create_font_config_preset
        
        current_config = {
            'font_size': self.font_size,
            'header_font_size': self.header_font_size,
            'bold_headers': self.bold_headers,
            'bold_company_info': self.bold_company_info
        }
        
        presets = ['small', 'normal', 'large', 'extra_large', 'huge']
        
        for preset_name in presets:
            preset = create_font_config_preset(preset_name)
            if all(current_config[key] == preset[key] 
                   for key in ['font_size', 'header_font_size']):
                return preset_name
        
        return 'custom'
    
    def _is_default_font_config(self) -> bool:
        """Check if using default font configuration"""
        return (self.font_size == 14 and 
                self.header_font_size == 14 and
                self.bold_headers == True and
                self.bold_company_info == True)
    
    def apply_font_preset(self, preset_name: str) -> None:
        """Apply font preset to this configuration (v1.1.1)"""
        from .models import create_font_config_preset
        
        preset = create_font_config_preset(preset_name)
        self.font_size = preset['font_size']
        self.header_font_size = preset['header_font_size']
        self.bold_headers = preset['bold_headers']
        self.bold_company_info = preset['bold_company_info']
    
    def get_display_config(self) -> Dict[str, Any]:
        """Get display configuration for multiple report types (v1.1.1)"""
        return {
            'show_all_report_types': self.show_all_report_types,
            'report_type_separator': self.report_type_separator,
            'category_separator': self.category_separator,
            'max_display_types': self.max_display_types,
            'use_categorized_display': self.use_categorized_display,
            'priority_display_mode': self.priority_display_mode,
            'truncate_indicator': self.truncate_indicator
        }
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary for serialization"""
        return {
            # Directory settings
            'downloads_dir': self.downloads_dir,
            'stock_csv_path': self.stock_csv_path,
            'output_dir': self.output_dir,
            
            # Matrix settings
            'max_years': self.max_years,
            'quarter_order': self.quarter_order,
            
            # v1.1.1 Multiple Report Type Settings
            'show_all_report_types': self.show_all_report_types,
            'report_type_separator': self.report_type_separator,
            'max_display_types': self.max_display_types,
            'use_categorized_display': self.use_categorized_display,
            
            # v1.1.1 Font Settings
            'font_size': self.font_size,
            'header_font_size': self.header_font_size,
            'bold_headers': self.bold_headers,
            'bold_company_info': self.bold_company_info,
            
            # Google Sheets settings
            'worksheet_name': self.worksheet_name,
            'csv_backup': self.csv_backup,
            'include_future_quarters': self.include_future_quarters,
            'auto_suggest_downloads': self.auto_suggest_downloads,
            
            # Report type preferences
            'preferred_types': self.preferred_types,
            'acceptable_types': self.acceptable_types,
            'excluded_types': self.excluded_types
        }

# Enhanced utility functions for v1.1.1

def create_font_config_preset(preset_name: str) -> Dict[str, Any]:
    """
    Create font configuration preset (v1.1.1 enhancement)
    
    Args:
        preset_name: One of 'small', 'normal', 'large', 'extra_large', 'huge'
    
    Returns:
        Dictionary with font configuration
    """
    presets = {
        'small': {
            'font_size': 10,
            'header_font_size': 11,
            'bold_headers': True,
            'bold_company_info': False,
            'description': 'Compact display for smaller screens'
        },
        'normal': {
            'font_size': 12,
            'header_font_size': 13,
            'bold_headers': True,
            'bold_company_info': False,
            'description': 'Standard font sizing'
        },
        'large': {
            'font_size': 14,
            'header_font_size': 16,
            'bold_headers': True,
            'bold_company_info': True,
            'description': 'Professional display (recommended)'
        },
        'extra_large': {
            'font_size': 16,
            'header_font_size': 18,
            'bold_headers': True,
            'bold_company_info': True,
            'description': 'Enhanced readability'
        },
        'huge': {
            'font_size': 20,
            'header_font_size': 24,
            'bold_headers': True,
            'bold_company_info': True,
            'description': 'Presentation mode'
        }
    }
    
    return presets.get(preset_name, presets['large'])  # Default to large

def apply_font_preset_to_config(config: MOPSConfig, preset_name: str) -> MOPSConfig:
    """Apply font preset to configuration object"""
    config.apply_font_preset(preset_name)
    return config

# Default configuration instance
DEFAULT_CONFIG = MOPSConfig()

def load_config(config_path: Optional[str] = None) -> MOPSConfig:
    """
    Load configuration from file or environment
    
    Args:
        config_path: Optional path to configuration file
        
    Returns:
        MOPSConfig instance with v1.1.1 enhancements
    """
    if config_path and os.path.exists(config_path):
        return MOPSConfig.from_file(config_path)
    else:
        return MOPSConfig.from_env()

def get_google_sheets_config() -> Dict[str, str]:
    """Get Google Sheets configuration from environment"""
    return {
        'credentials': os.getenv('GOOGLE_SHEETS_CREDENTIALS'),
        'sheet_id': os.getenv('GOOGLE_SHEET_ID'),
        'worksheet_name': os.getenv('MOPS_WORKSHEET_NAME', 'MOPS下載狀態')
    }

def validate_v1_1_1_environment() -> Dict[str, bool]:
    """Validate v1.1.1 environment setup"""
    validation = {}
    
    # Check required directories
    required_dirs = ['downloads', 'logs', 'data/reports']
    for directory in required_dirs:
        validation[f'dir_{directory.replace("/", "_")}'] = os.path.exists(directory)
    
    # Check environment variables
    required_env = ['GOOGLE_SHEETS_CREDENTIALS', 'GOOGLE_SHEET_ID']
    optional_env = ['MOPS_FONT_SIZE', 'MOPS_SHOW_ALL_TYPES']
    
    for env_var in required_env:
        validation[f'env_{env_var}'] = bool(os.getenv(env_var))
    
    for env_var in optional_env:
        validation[f'opt_env_{env_var}'] = bool(os.getenv(env_var))
    
    # Check if stock CSV exists
    stock_csv_path = os.getenv('MOPS_STOCK_CSV_PATH', 'StockID_TWSE_TPEX.csv')
    validation['stock_csv_exists'] = os.path.exists(stock_csv_path)
    
    return validation

def create_default_config_file(config_path: str = 'config.yaml') -> str:
    """Create a default configuration file with v1.1.1 settings"""
    import yaml
    
    default_config = {
        'downloads_dir': 'downloads',
        'stock_csv_path': 'StockID_TWSE_TPEX.csv',
        'output_dir': 'data/reports',
        'max_years': 3,
        
        # v1.1.1 Enhanced settings
        'show_all_report_types': True,
        'report_type_separator': '/',
        'max_display_types': 5,
        'use_categorized_display': False,
        
        # Font settings
        'font_size': 14,
        'header_font_size': 16,
        'bold_headers': True,
        'bold_company_info': True,
        
        # Google Sheets
        'worksheet_name': 'MOPS下載狀態',
        'csv_backup': True,
        'highlight_multiple_types': True,
        
        # Report preferences
        'preferred_types': ['A12', 'A13'],
        'acceptable_types': ['AI1', 'A1L'],
        'excluded_types': ['AIA', 'AE2']
    }
    
    with open(config_path, 'w', encoding='utf-8') as f:
        yaml.dump(default_config, f, allow_unicode=True, default_flow_style=False)
    
    return config_path