"""
MOPS Sheets Uploader - Data Models
Defines data structures for PDF files, matrix cells, and analysis results.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, List, Dict, Any
import re

# Report Type Priorities (based on MOPS design)
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

# Status Mapping for Display
STATUS_MAPPING = {
    1: '✅',  # Individual reports (preferred)
    2: '🟡',  # Consolidated reports (acceptable)
    3: '⚠️',  # Generic reports (caution)
    9: '❌',  # English or excluded reports
    None: '-'  # No reports found
}

@dataclass
class PDFFile:
    """Represents a discovered PDF file"""
    company_id: str
    year: int
    quarter: int
    report_type: str
    filename: str
    file_path: str
    file_size: int
    modified_date: datetime
    
    @property
    def quarter_key(self) -> str:
        """Return quarter key for matrix (e.g., '2024 Q1')"""
        return f"{self.year} Q{self.quarter}"
    
    @property
    def priority(self) -> int:
        """Return report type priority"""
        return REPORT_TYPE_PRIORITY.get(self.report_type, 9)
    
    @property
    def status_symbol(self) -> str:
        """Return status symbol based on priority"""
        return STATUS_MAPPING.get(self.priority, '-')
    
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
    """Metadata extracted from PDF filename"""
    year: int
    quarter: int
    company_id: str
    report_type: str
    filename: str
    
    @classmethod
    def from_filename(cls, filename: str) -> Optional['PDFMetadata']:
        """Parse PDF filename to extract metadata"""
        # Pattern: YYYYQQ_COMPANYID_TYPE.pdf
        pattern = r"^(\d{4})(\d{2})_(\d{4})_([A-Z0-9]+)\.pdf$"
        match = re.match(pattern, filename)
        
        if not match:
            return None
            
        year_str, quarter_str, company_id, report_type = match.groups()
        
        # Parse year and quarter
        year = int(year_str)
        quarter = int(quarter_str)
        
        # Validate ranges
        if not (2020 <= year <= 2030):
            return None
        if not (1 <= quarter <= 4):
            return None
        if len(company_id) != 4:
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
    """Represents a single cell in the matrix"""
    has_pdf: bool = False
    report_type: Optional[str] = None
    status_symbol: str = '-'
    file_count: int = 0
    file_paths: List[str] = field(default_factory=list)
    priority: int = 9
    
    def add_pdf(self, pdf_file: PDFFile) -> None:
        """Add a PDF file to this cell"""
        self.has_pdf = True
        self.file_count += 1
        self.file_paths.append(pdf_file.file_path)
        
        # Update with best priority report
        if pdf_file.priority < self.priority:
            self.priority = pdf_file.priority
            self.report_type = pdf_file.report_type
            self.status_symbol = pdf_file.status_symbol
    
    def get_display_value(self) -> str:
        """Get display value for matrix cell"""
        if not self.has_pdf:
            return '-'
        elif self.file_count > 1:
            return f"{self.report_type}({self.file_count})"
        else:
            return self.report_type or self.status_symbol

@dataclass
class StockInfo:
    """Stock information from CSV"""
    stock_id: str
    company_name: str
    
    def __post_init__(self):
        """Validate stock info"""
        if not self.stock_id or not self.stock_id.isdigit():
            raise ValueError(f"Invalid stock ID: {self.stock_id}")
        if len(self.stock_id) != 4:
            raise ValueError(f"Stock ID must be 4 digits: {self.stock_id}")

@dataclass
class StockListChanges:
    """Represents changes in stock list"""
    added_companies: List[str] = field(default_factory=list)
    removed_companies: List[str] = field(default_factory=list)
    total_companies: int = 0
    previous_total: int = 0
    
    @property
    def has_changes(self) -> bool:
        """Check if there are any changes"""
        return bool(self.added_companies or self.removed_companies)
    
    @property
    def change_summary(self) -> str:
        """Human-readable summary of changes"""
        if not self.has_changes:
            return "No changes detected"
            
        parts = []
        if self.added_companies:
            parts.append(f"Added {len(self.added_companies)} companies")
        if self.removed_companies:
            parts.append(f"Removed {len(self.removed_companies)} companies")
            
        return f"{', '.join(parts)} (Total: {self.previous_total} → {self.total_companies})"

@dataclass
class CoverageStats:
    """Statistics about PDF coverage"""
    total_companies: int
    companies_with_pdfs: int
    total_quarters: int
    total_possible_reports: int
    total_actual_reports: int
    coverage_percentage: float
    report_type_distribution: Dict[str, int] = field(default_factory=dict)
    missing_quarters_by_company: Dict[str, List[str]] = field(default_factory=dict)
    future_quarters: List[str] = field(default_factory=list)
    
    @property
    def companies_without_pdfs(self) -> int:
        """Number of companies with no PDFs"""
        return self.total_companies - self.companies_with_pdfs

@dataclass
class ProcessingResult:
    """Result of the complete processing pipeline"""
    success: bool
    matrix_uploaded: bool = False
    csv_exported: bool = False
    csv_path: Optional[str] = None
    sheets_url: Optional[str] = None
    coverage_stats: Optional[CoverageStats] = None
    stock_changes: Optional[StockListChanges] = None
    error_message: Optional[str] = None
    processing_time: float = 0.0
    total_files_processed: int = 0
    
    def get_summary(self) -> Dict[str, Any]:
        """Get summary of processing result"""
        return {
            'success': self.success,
            'matrix_uploaded': self.matrix_uploaded,
            'csv_exported': self.csv_exported,
            'coverage_percentage': self.coverage_stats.coverage_percentage if self.coverage_stats else 0,
            'total_companies': self.coverage_stats.total_companies if self.coverage_stats else 0,
            'processing_time': self.processing_time,
            'has_stock_changes': self.stock_changes.has_changes if self.stock_changes else False,
            'error': self.error_message
        }

@dataclass
class FutureQuarterAnalysis:
    """Analysis of PDFs with future quarter dates"""
    future_pdfs: List[PDFFile] = field(default_factory=list)
    suspicious_files: List[str] = field(default_factory=list)
    max_future_months: int = 0
    warnings: List[str] = field(default_factory=list)
    
    def add_warning(self, message: str) -> None:
        """Add a warning message"""
        self.warnings.append(message)
    
    def has_future_pdfs(self) -> bool:
        """Check if any future PDFs were found"""
        return bool(self.future_pdfs)
    
    def get_analysis_summary(self) -> Dict[str, Any]:
        """Get summary of future quarter analysis"""
        return {
            'future_pdf_count': len(self.future_pdfs),
            'suspicious_file_count': len(self.suspicious_files),
            'max_future_months': self.max_future_months,
            'warning_count': len(self.warnings),
            'quarters_affected': list(set(pdf.quarter_key for pdf in self.future_pdfs))
        }