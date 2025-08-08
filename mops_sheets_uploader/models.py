"""
MOPS Sheets Uploader - Enhanced Data Models v1.1.1
Complete data models with multiple report type support and font configuration
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, List, Dict, Any, Set
import re
from collections import defaultdict

# Enhanced Report Type Categories for v1.1.1
REPORT_TYPE_CATEGORIES = {
    'individual': ['A12', 'A13'],           # Individual reports (highest priority)
    'consolidated': ['AI1', 'A1L'],         # Consolidated reports (secondary)
    'generic': ['A10', 'A11'],              # Generic reports (lower priority)
    'english': ['AIA', 'AE2'],              # English reports (excluded)
    'other': []                             # Any other types discovered
}

# Enhanced Report Type Priorities with v1.1.1 updates
REPORT_TYPE_PRIORITY = {
    'A12': 1,  # IFRSs個別財報 - Highest priority
    'A13': 1,  # IFRSs個體財報 - Highest priority  
    'AI1': 2,  # IFRSs合併財報 - Secondary priority
    'A1L': 2,  # IFRSs合併財報 - Secondary priority
    'A10': 3,  # Generic financial reports - Lower priority
    'A11': 3,  # Generic financial reports - Lower priority
    'AIA': 9,  # English consolidated - Excluded
    'AE2': 9   # English parent-subsidiary - Excluded
}

# Enhanced Status Mapping for Display v1.1.1
STATUS_MAPPING = {
    1: '✅',  # Individual reports (preferred)
    2: '🟡',  # Consolidated reports (acceptable)
    3: '⚠️',  # Generic reports (caution)
    9: '❌',  # English or excluded reports
    None: '-',  # No reports found
    'multiple': '🔄',  # Multiple types available (NEW in v1.1.1)
    'mixed': '📊'      # Mixed categories available (NEW in v1.1.1)
}

@dataclass
class PDFFile:
    """Enhanced PDF file representation with v1.1.1 categorization support"""
    company_id: str
    year: int
    quarter: int
    report_type: str
    filename: str
    file_path: str
    file_size: int
    modified_date: datetime
    
    # v1.1.1 Enhanced fields
    category: str = field(default="")  # individual, consolidated, generic, english
    priority: int = field(default=9)   # Priority score
    
    def __post_init__(self):
        """Auto-categorize and set priority after initialization"""
        if not self.category:
            self.category = self.get_category()
        if self.priority == 9:
            self.priority = self.get_priority()
    
    @property
    def quarter_key(self) -> str:
        """Return quarter key for matrix (e.g., '2024 Q1')"""
        return f"{self.year} Q{self.quarter}"
    
    @property
    def status_symbol(self) -> str:
        """Return status symbol based on priority"""
        return STATUS_MAPPING.get(self.priority, '-')
    
    def get_category(self) -> str:
        """Get report type category (v1.1.1 enhancement)"""
        for category, types in REPORT_TYPE_CATEGORIES.items():
            if self.report_type in types:
                return category
        return 'other'
    
    def get_priority(self) -> int:
        """Get report type priority"""
        return REPORT_TYPE_PRIORITY.get(self.report_type, 9)
    
    @property
    def is_individual(self) -> bool:
        """Check if this is an individual report type (v1.1.1)"""
        return self.category == "individual"
    
    @property
    def is_consolidated(self) -> bool:
        """Check if this is a consolidated report type (v1.1.1)"""
        return self.category == "consolidated"
    
    @property
    def is_english(self) -> bool:
        """Check if this is an English report type (v1.1.1)"""
        return self.category == "english"
    
    def is_future_quarter(self, reference_date: datetime = None) -> bool:
        """Check if this PDF represents a future quarter"""
        if reference_date is None:
            reference_date = datetime.now()
            
        current_year = reference_date.year
        current_quarter = (reference_date.month - 1) // 3 + 1
        
        return (self.year > current_year or 
                (self.year == current_year and self.quarter > current_quarter))

@dataclass
class PDFMetadata:
    """Enhanced metadata extracted from PDF filename"""
    year: int
    quarter: int
    company_id: str
    report_type: str
    filename: str
    category: str = field(default="")
    priority: int = field(default=9)
    
    def __post_init__(self):
        """Auto-categorize after initialization"""
        if not self.category:
            self.category = self._get_category()
        if self.priority == 9:
            self.priority = REPORT_TYPE_PRIORITY.get(self.report_type, 9)
    
    def _get_category(self) -> str:
        """Get report type category"""
        for category, types in REPORT_TYPE_CATEGORIES.items():
            if self.report_type in types:
                return category
        return 'other'
    
    @classmethod
    def from_filename(cls, filename: str) -> Optional['PDFMetadata']:
        """Parse PDF filename to extract metadata with enhanced validation"""
        # Pattern: YYYYQQ_COMPANYID_TYPE.pdf
        pattern = r"^(\d{4})(\d{2})_(\d{4})_([A-Z0-9]+)\.pdf$"
        match = re.match(pattern, filename)
        
        if not match:
            return None
            
        year_str, quarter_str, company_id, report_type = match.groups()
        
        # Parse year and quarter
        year = int(year_str)
        quarter = int(quarter_str)
        
        # Enhanced validation for v1.1.1
        if not (2020 <= year <= 2030):
            return None
        if not (1 <= quarter <= 4):
            return None
        if len(company_id) != 4:
            return None
        if not company_id.isdigit():
            return None
            
        return cls(
            year=year,
            quarter=quarter,
            company_id=company_id,
            report_type=report_type,
            filename=filename
        )

@dataclass
class MatrixCell:
    """Enhanced matrix cell with v1.1.1 multiple report type support"""
    has_pdf: bool = False
    report_types: Set[str] = field(default_factory=set)  # All report types available
    categories: Set[str] = field(default_factory=set)    # All categories available
    file_paths: List[str] = field(default_factory=list)
    priorities: Set[int] = field(default_factory=set)    # All priorities present
    best_priority: int = 9  # Best (lowest) priority
    file_count: int = 0
    
    def add_pdf(self, pdf_file: PDFFile) -> None:
        """Add a PDF file to this cell with enhanced tracking"""
        self.has_pdf = True
        self.file_count += 1
        self.file_paths.append(pdf_file.file_path)
        self.report_types.add(pdf_file.report_type)
        self.categories.add(pdf_file.category)
        self.priorities.add(pdf_file.priority)
        
        # Update best priority
        if pdf_file.priority < self.best_priority:
            self.best_priority = pdf_file.priority
    
    @property
    def has_multiple_types(self) -> bool:
        """Check if cell has multiple report types (v1.1.1)"""
        return len(self.report_types) > 1
    
    @property
    def has_mixed_categories(self) -> bool:
        """Check if cell has mixed categories (v1.1.1)"""
        return len(self.categories) > 1
    
    @property
    def status_symbol(self) -> str:
        """Get enhanced status symbol (v1.1.1)"""
        if not self.has_pdf:
            return '-'
        elif self.has_mixed_categories:
            return STATUS_MAPPING['mixed']
        elif self.has_multiple_types:
            return STATUS_MAPPING['multiple']
        else:
            return STATUS_MAPPING.get(self.best_priority, '-')
    
    def get_display_value(self, show_all_types: bool = True, 
                         separator: str = "/", max_types: int = 5,
                         use_categorized: bool = False,
                         category_separator: str = " → ") -> str:
        """Enhanced display value with v1.1.1 multiple type support"""
        if not self.has_pdf:
            return '-'
        
        if not show_all_types:
            # Return only the best priority report type
            best_types = [rt for rt in self.report_types 
                         if REPORT_TYPE_PRIORITY.get(rt, 9) == self.best_priority]
            return best_types[0] if best_types else list(self.report_types)[0]
        
        if use_categorized:
            return self._get_categorized_display(separator, category_separator)
        
        # Show all report types, sorted by priority then alphabetically
        sorted_types = sorted(self.report_types, 
                            key=lambda x: (REPORT_TYPE_PRIORITY.get(x, 9), x))
        
        if len(sorted_types) <= max_types:
            return separator.join(sorted_types)
        else:
            # Truncate with indicator
            visible_types = sorted_types[:max_types-1]
            remaining_count = len(sorted_types) - len(visible_types)
            return separator.join(visible_types) + f"+{remaining_count}"
    
    def _get_categorized_display(self, separator: str = "/", 
                               category_separator: str = " → ") -> str:
        """Get categorized display string (v1.1.1)"""
        # Group types by category
        category_groups = defaultdict(list)
        for report_type in self.report_types:
            category = self._get_report_type_category(report_type)
            category_groups[category].append(report_type)
        
        # Build display string by category priority
        category_order = ['individual', 'consolidated', 'generic', 'other']
        display_parts = []
        
        for category in category_order:
            if category in category_groups:
                types_in_category = sorted(category_groups[category])
                display_parts.append(separator.join(types_in_category))
        
        return category_separator.join(display_parts)
    
    def _get_report_type_category(self, report_type: str) -> str:
        """Get category for a report type"""
        for category, types in REPORT_TYPE_CATEGORIES.items():
            if report_type in types:
                return category
        return 'other'
    
    def get_best_type(self) -> str:
        """Get the best (highest priority) report type"""
        if not self.report_types:
            return '-'
        
        best_types = [rt for rt in self.report_types 
                     if REPORT_TYPE_PRIORITY.get(rt, 9) == self.best_priority]
        return best_types[0] if best_types else list(self.report_types)[0]

@dataclass
class StockInfo:
    """Stock information from CSV"""
    stock_id: str
    company_name: str
    
    def __post_init__(self):
        """Validate stock info after initialization"""
        if not self.stock_id or len(str(self.stock_id)) != 4:
            raise ValueError(f"Invalid stock ID: {self.stock_id}")
        if not self.company_name:
            raise ValueError(f"Missing company name for {self.stock_id}")

@dataclass
class CoverageStats:
    """Enhanced coverage statistics with v1.1.1 multiple type analysis"""
    total_companies: int
    companies_with_pdfs: int
    total_quarters: int
    total_possible_reports: int
    total_actual_reports: int
    coverage_percentage: float
    
    # v1.1.1 Enhanced fields
    report_type_distribution: Dict[str, int] = field(default_factory=dict)
    missing_quarters_by_company: Dict[str, List[str]] = field(default_factory=dict)
    future_quarters: List[str] = field(default_factory=list)
    cells_with_multiple_types: int = 0
    report_type_combinations: Dict[str, int] = field(default_factory=dict)
    category_distribution: Dict[str, int] = field(default_factory=dict)
    
    @property
    def multiple_types_percentage(self) -> float:
        """Calculate percentage of cells with multiple types"""
        if self.total_actual_reports == 0:
            return 0.0
        return (self.cells_with_multiple_types / self.total_actual_reports) * 100

@dataclass
class StockListChanges:
    """Enhanced stock list changes detection"""
    added_companies: List[str] = field(default_factory=list)
    removed_companies: List[str] = field(default_factory=list)
    total_companies: int = 0
    previous_total: int = 0
    
    @property
    def has_changes(self) -> bool:
        """Check if there are any changes"""
        return len(self.added_companies) > 0 or len(self.removed_companies) > 0
    
    @property
    def net_change(self) -> int:
        """Calculate net change in companies"""
        return len(self.added_companies) - len(self.removed_companies)
    
    @property
    def change_summary(self) -> str:
        """Get human-readable change summary"""
        if not self.has_changes:
            return "無變更"
        
        parts = []
        if self.added_companies:
            parts.append(f"新增 {len(self.added_companies)} 家")
        if self.removed_companies:
            parts.append(f"移除 {len(self.removed_companies)} 家")
        
        return ", ".join(parts)

@dataclass
class FutureQuarterAnalysis:
    """Analysis of PDFs with future quarter dates"""
    future_pdfs: List[PDFFile] = field(default_factory=list)
    suspicious_files: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    def has_future_pdfs(self) -> bool:
        """Check if there are any future PDFs"""
        return len(self.future_pdfs) > 0
    
    def add_warning(self, warning: str) -> None:
        """Add a warning message"""
        self.warnings.append(warning)

@dataclass
class ProcessingResult:
    """Enhanced processing result with v1.1.1 features"""
    success: bool = False
    matrix_uploaded: bool = False
    csv_exported: bool = False
    csv_path: Optional[str] = None
    sheets_url: Optional[str] = None
    coverage_stats: Optional[CoverageStats] = None
    stock_changes: Optional[StockListChanges] = None
    error_message: Optional[str] = None
    processing_time: float = 0.0
    total_files_processed: int = 0
    
    # v1.1.1 Enhanced fields
    font_config_used: Dict[str, Any] = field(default_factory=dict)
    multiple_types_found: int = 0
    type_combinations_analyzed: Dict[str, int] = field(default_factory=dict)

# v1.1.1 Enhanced Analysis Functions

def analyze_report_type_combinations(matrix_cells: Dict[str, MatrixCell]) -> Dict[str, int]:
    """Analyze report type combinations across matrix cells"""
    combinations = defaultdict(int)
    
    for cell in matrix_cells.values():
        if cell.has_pdf and len(cell.report_types) > 1:
            # Create combination string sorted by priority
            sorted_types = sorted(cell.report_types, 
                                key=lambda x: (REPORT_TYPE_PRIORITY.get(x, 9), x))
            combination = "/".join(sorted_types)
            combinations[combination] += 1
    
    return dict(combinations)

def get_report_type_category_stats(matrix_cells: Dict[str, MatrixCell]) -> Dict[str, int]:
    """Get statistics on report type categories"""
    stats = {
        'individual_only': 0,
        'consolidated_only': 0,
        'mixed_types': 0,
        'generic_only': 0,
        'other_only': 0
    }
    
    for cell in matrix_cells.values():
        if not cell.has_pdf:
            continue
            
        categories = cell.categories
        
        if len(categories) > 1:
            stats['mixed_types'] += 1
        elif 'individual' in categories:
            stats['individual_only'] += 1
        elif 'consolidated' in categories:
            stats['consolidated_only'] += 1
        elif 'generic' in categories:
            stats['generic_only'] += 1
        else:
            stats['other_only'] += 1
    
    return stats

def create_font_config_preset(preset_name: str) -> Dict[str, Any]:
    """Create font configuration preset (v1.1.1)"""
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

# v1.1.1 Enhanced Validation Functions

def validate_font_configuration(font_config: Dict[str, Any]) -> List[str]:
    """Validate font configuration settings"""
    errors = []
    
    font_size = font_config.get('font_size', 14)
    header_font_size = font_config.get('header_font_size', 14)
    
    if not isinstance(font_size, int) or font_size < 8 or font_size > 72:
        errors.append("Font size must be an integer between 8 and 72")
    
    if not isinstance(header_font_size, int) or header_font_size < 8 or header_font_size > 72:
        errors.append("Header font size must be an integer between 8 and 72")
    
    return errors

def get_default_display_config() -> Dict[str, Any]:
    """Get default display configuration for v1.1.1"""
    return {
        'show_all_report_types': True,
        'report_type_separator': '/',
        'category_separator': ' → ',
        'max_display_types': 5,
        'use_categorized_display': False,
        'font_size': 14,
        'header_font_size': 14,
        'bold_headers': True,
        'bold_company_info': True
    }

# Compatibility functions for backward compatibility
def create_matrix_cell_from_pdfs(pdfs: List[PDFFile]) -> MatrixCell:
    """Create a matrix cell from a list of PDF files"""
    cell = MatrixCell()
    for pdf in pdfs:
        cell.add_pdf(pdf)
    return cell