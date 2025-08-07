"""
MOPS Sheets Uploader - Configuration Management
Handles configuration loading from environment variables and config files.
"""

import os
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field
import json

@dataclass
class MOPSConfig:
    """Configuration for MOPS Sheets Uploader"""
    
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
    
    # CSV export settings
    csv_backup: bool = True
    csv_filename_pattern: str = "mops_matrix_{timestamp}.csv"
    include_future_analysis: bool = True
    
    # Report type settings
    preferred_types: List[str] = field(default_factory=lambda: ["A12", "A13"])
    acceptable_types: List[str] = field(default_factory=lambda: ["AI1", "A1L"])
    excluded_types: List[str] = field(default_factory=lambda: ["AIA", "AE2"])
    
    # Display settings
    use_symbols: bool = True
    show_file_counts: bool = False
    highlight_missing: bool = True
    future_quarter_symbol: str = "⚠️"
    
    # Google Sheets Integration
    google_sheet_id: Optional[str] = None
    google_credentials: Optional[str] = None
    
    @classmethod
    def from_env(cls) -> 'MOPSConfig':
        """Load configuration from environment variables"""
        config = cls()
        
        # Load from environment variables
        config.downloads_dir = os.getenv('MOPS_DOWNLOADS_DIR', config.downloads_dir)
        config.stock_csv_path = os.getenv('MOPS_STOCK_CSV_PATH', config.stock_csv_path)
        config.output_dir = os.getenv('MOPS_OUTPUT_DIR', config.output_dir)
        config.max_years = int(os.getenv('MOPS_MAX_YEARS', str(config.max_years)))
        config.worksheet_name = os.getenv('MOPS_WORKSHEET_NAME', config.worksheet_name)
        
        # Google Sheets settings
        config.google_sheet_id = os.getenv('GOOGLE_SHEET_ID')
        config.google_credentials = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
        
        # Boolean settings
        config.csv_backup = os.getenv('MOPS_CSV_BACKUP', 'true').lower() == 'true'
        config.detect_stock_changes = os.getenv('MOPS_DETECT_CHANGES', 'true').lower() == 'true'
        config.include_future_quarters = os.getenv('MOPS_INCLUDE_FUTURE', 'true').lower() == 'true'
        
        return config
    
    @classmethod
    def from_file(cls, config_path: str) -> 'MOPSConfig':
        """Load configuration from YAML/JSON file"""
        import yaml
        
        with open(config_path, 'r', encoding='utf-8') as f:
            if config_path.endswith('.json'):
                data = json.load(f)
            else:
                data = yaml.safe_load(f)
        
        # Create config with defaults, then update with file data
        config = cls()
        for key, value in data.items():
            if hasattr(config, key):
                setattr(config, key, value)
        
        return config
    
    def validate(self) -> List[str]:
        """Validate configuration and return list of errors"""
        errors = []
        
        # Check required directories
        if not os.path.exists(self.downloads_dir):
            errors.append(f"Downloads directory not found: {self.downloads_dir}")
        
        if not os.path.exists(self.stock_csv_path):
            errors.append(f"Stock CSV file not found: {self.stock_csv_path}")
        
        # Check Google Sheets settings if needed
        if not self.csv_backup and not self.google_credentials:
            errors.append("Google Sheets credentials required when CSV backup disabled")
        
        # Check numeric ranges
        if self.max_years < 1 or self.max_years > 10:
            errors.append("max_years must be between 1 and 10")
        
        if self.change_threshold_warning < 0:
            errors.append("change_threshold_warning must be non-negative")
        
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
            year, quarter = q.split(' Q')
            return (int(year), int(quarter))
        
        sorted_quarters = sorted(all_quarters, 
                               key=quarter_sort_key, 
                               reverse=(self.quarter_order == "desc"))
        
        return sorted_quarters

# Default configuration instance
DEFAULT_CONFIG = MOPSConfig()

# Target report configurations
TARGET_INDIVIDUAL_REPORTS = [
    "IFRSs個別財報",     # True Individual Financial Report - Type 1 (A12.pdf)
    "IFRSs個體財報",     # True Individual Financial Report - Type 2 (A13.pdf)
    "IFRSs合併財報"      # Consolidated reports (A1L.pdf, AI1.pdf) - accepted as secondary
]

TARGET_FILE_PATTERNS = [
    r"A12\.pdf$",       # True individual (company 8272 style)
    r"A13\.pdf$",       # True individual (company 2330 historical)
    r"AI1\.pdf$",       # Consolidated accepted as secondary (company 2330 current)
    r"A1[0-9]\.pdf$"    # Generic patterns (mixed individual/consolidated)
]

FLEXIBLE_TARGETS = [
    "IFRSs合併財報",     # Consolidated reports (fallback when no individual available)
    "財務報告書"         # General financial reports (broadest fallback)
]

EXCLUDED_KEYWORDS = [
    "英文版",            # English versions
    "AIA\.pdf",         # English consolidated
    "AE2\.pdf"          # English parent-subsidiary
]

def load_config(config_path: Optional[str] = None) -> MOPSConfig:
    """
    Load configuration from file or environment
    
    Args:
        config_path: Optional path to configuration file
        
    Returns:
        MOPSConfig instance
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