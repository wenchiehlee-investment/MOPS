# MOPS Sheets Uploader - Design Document v1.1.1 (Complete)

[![Version](https://img.shields.io/badge/Version-1.1.1-blue)](https://github.com/your-repo/mops-downloader)
[![Python](https://img.shields.io/badge/Python-3.9%2B-green)](https://python.org)
[![Integration](https://img.shields.io/badge/Integration-MOPS_Downloader-orange)](https://github.com/your-repo/mops-downloader)
[![Enhancement](https://img.shields.io/badge/Enhancement-Multiple_Report_Types-purple)](https://github.com/your-repo/mops-downloader)

## 📋 Project Overview

The MOPS Sheets Uploader is a Python-based companion tool for the MOPS Downloader System that automatically scans downloaded PDF files and creates a comprehensive matrix view in Google Sheets. **Enhanced with multiple report type support** - now displays all available report types per quarter (e.g., "A12/A13/AI1") instead of just the highest priority report, with configurable font sizing for optimal readability.

## 🆕 Version 1.1.1 Enhancements

### **Multiple Report Type Display**
- **Show All Available Types**: Instead of resolving conflicts to one report type, displays all available types for each quarter
- **Intelligent Categorization**: Groups report types by category (Individual/Consolidated/Generic)
- **Flexible Separators**: Configurable separators between types and categories
- **Enhanced Statistics**: Detailed analysis of report type combinations and coverage patterns

### **Enhanced Font and Display Control**
- **Configurable Font Size**: Default 14pt font size for optimal readability
- **Separate Header Formatting**: Independent font size control for headers (default 14pt)
- **Bold Formatting Options**: Configurable bold formatting for headers and company info columns
- **Enhanced Google Sheets Formatting**: Improved visual presentation with professional styling

### **Key New Features**
- **Comprehensive Type Display**: Shows "A12/A13/AI1" when multiple reports available
- **Category-Based Grouping**: Optional grouping by Individual → Consolidated → Generic
- **Enhanced Coverage Analysis**: Statistics on multiple type availability and combinations
- **Flexible Configuration**: Control display format, separators, and prioritization
- **Smart Truncation**: Handles cases with many report types (A12/A13+3 more)
- **Professional Styling**: 14pt font with bold headers for clear readability

## 🎯 Goals and Objectives

### Primary Goal
Create a Google Sheets matrix showing **comprehensive PDF file availability** with companies as rows and year-quarters as columns, **displaying all available report types** to provide complete visibility into download status and report coverage diversity.

### Enhanced Key Features
- **🔄 Multiple Report Type Display**: Show all available report types per quarter (A12/A13/AI1)
- **📊 Intelligent Categorization**: Group by Individual/Consolidated/Generic categories
- **🎯 Flexible Display Modes**: Choose between showing all types, best only, or category-based
- **📈 Enhanced Analytics**: Detailed analysis of report type combinations and patterns
- **⚙️ Configurable Formatting**: Customizable separators, truncation, and display options
- **🔍 Advanced Statistics**: Track multiple type availability and coverage diversity
- **🎨 Professional Styling**: 14pt font with configurable formatting for optimal readability
- **Automated PDF Discovery**: Scan `downloads/` folder structure to identify available reports
- **Stock Information Integration**: Load company details from `StockID_TWSE_TPEX.csv`
- **Report Type Classification**: Identify and display PDF report types (A12, A13, AI1, etc.)
- **Progress Tracking**: Visual indicators for download completeness
- **Flexible Output**: Support both Google Sheets upload and CSV backup
- **Error Recovery**: Robust handling of missing files and API limitations

## 📊 Enhanced Matrix Specifications

### Concrete Dimensions
Based on the provided StockID_TWSE_TPEX.csv:

**Matrix Size**: 116+ companies × (2 + quarters) columns
- **Rows**: 116+ companies (from CSV file, auto-expands with additions)
- **Base Columns**: 2 (代號, 名稱) 
- **Quarter Columns**: Configurable (default: 12 quarters = 3 years × 4 quarters)
- **Total Columns**: 14 (2 base + 12 quarters) for default 3-year view

**Memory Requirements**: 
- Estimated CSV size: ~50KB for full matrix with enhanced text data
- Google Sheets: Well within limits (116+ rows × 14+ columns)
- Processing: Lightweight - can handle in memory easily

**API Efficiency**:
- Google Sheets batch upload: Single API call for entire matrix
- Processing time: < 5 seconds for full scan and upload
- Rate limits: No concern with this dataset size

### Enhanced Cell Value Types

**New Multiple Report Type Display**:
- **Multiple Individual**: `A12/A13` - Multiple individual report types available
- **Mixed Categories**: `A12/AI1` - Both individual and consolidated reports available  
- **Full Combination**: `A12/A13/AI1/A1L` - All report types available for the quarter
- **Categorized Display**: `A12/A13 → AI1/A1L` - Grouped by category with separator
- **Truncated Display**: `A12/A13/AI1+2` - Truncated when too many types (configurable)

**Traditional Single Type Display** (when `show_all_report_types: false`):
- **PDF Type Codes**: `A12`, `A13`, `AI1`, `A1L` (highest priority PDF type found)
- **Status Indicators**: 
  - `✅` - Individual report available (A12/A13)
  - `🟡` - Consolidated report available (AI1/A1L)
  - `⚠️` - Generic reports available
  - `❌` - No reports found
  - `-` - Not attempted/out of scope

**Enhanced Status Indicators**:
- **Multiple Types**: `🔄` - Multiple report types available (new)
- **Mixed Categories**: `📊` - Both individual and consolidated available (new)
- **Future Quarter**: `⏭️` - Future quarter with reports (enhanced)

### Google Sheets Matrix Layout (Enhanced Examples)

**Scenario 1: Enhanced Multiple Report Types Display (116+ Companies)**
```
代號    | 名稱       | 2025 Q1 | 2024 Q4 | 2024 Q3 | 2024 Q2 | 2024 Q1 | ...
--------|------------|---------|---------|---------|---------|---------|----
2330    | 台積電     | AI1     | AI1     | AI1     | AI1     | A13/AI1 | ...
8272    | 全景軟體   | A12     | A12     | A12/A13 | A12     | -       | ...
2382    | 廣達       | -       | A13     | A13     | -       | -       | ...
1234    | 混合範例   | A12/AI1 | A12/A13/AI1 | A12/A13+2 | AI1/A1L | A12 | ...
...     | ...        | ...     | ...     | ...     | ...     | ...     | ...
(116+ total companies with enhanced type display)
```

**Scenario 2: After Adding 2 New Companies (Total: 118)**
```
代號    | 名稱       | 2025 Q1 | 2024 Q4 | 2024 Q3 | 2024 Q2 | 2024 Q1 | ...
--------|------------|---------|---------|---------|---------|---------|----
2330    | 台積電     | AI1     | AI1     | AI1     | AI1     | A13/AI1 | ...
8272    | 全景軟體   | A12     | A12     | A12/A13 | A12     | -       | ...
2382    | 廣達       | -       | A13     | A13     | -       | -       | ...
...     | ...        | ...     | ...     | ...     | ...     | ...     | ...
2334    | 聯發科     | -       | -       | -       | -       | -       | ... (🆕 new)
1234    | 測試公司   | -       | -       | -       | -       | -       | ... (🆕 new)
(118 total companies - auto-expanded matrix)
```

**Scenario 3: Categorized Display Mode**
```
代號    | 名稱       | 2025 Q1 | 2024 Q4 | 2024 Q3 | 2024 Q2 | 2024 Q1 | ...
--------|------------|---------|---------|---------|---------|---------|----
2330    | 台積電     | AI1     | AI1     | AI1     | AI1     | A13→AI1 | ...
8272    | 全景軟體   | A12     | A12     | A12/A13 | A12     | -       | ...
2382    | 廣達       | -       | A13     | A13     | -       | -       | ...
1234    | 混合範例   | A12→AI1 | A12/A13→AI1 | A12/A13→AI1/A1L | AI1/A1L | A12 | ...
...     | ...        | ...     | ...     | ...     | ...     | ...     | ...
(Shows categories: Individual → Consolidated → Generic)
```

**Scenario 4: Enhanced Summary Statistics**
```
代號    | 名稱   | 總報告數 | 涵蓋率 | 個別報告 | 合併報告 | 多重類型 | 類型多樣性 | ...
--------|--------|----------|--------|----------|----------|----------|------------|----
2330    | 台積電 | 6        | 100%   | 1        | 6        | 1        | 0.33       | ...
8272    | 全景軟體| 4        | 66.7%  | 4        | 0        | 1        | 0.75       | ...
1234    | 混合範例| 8        | 100%   | 4        | 4        | 3        | 0.63       | ...
```

**New Company Indicators:**
- `-` = No PDFs downloaded yet (normal for new companies)
- Companies are automatically sorted by stock code (代號)
- Matrix dynamically expands both vertically (companies) and horizontally (quarters)
- Google Sheets worksheet auto-resizes to fit new dimensions

## 🏗️ Enhanced System Architecture

### Component Structure (Enhanced)
```
mops_sheets_uploader/
├── __init__.py                 # Enhanced with multiple type functions
├── pdf_scanner.py             # Enhanced PDF discovery and classification  
├── matrix_builder.py          # 🆕 Enhanced with multiple type support
├── sheets_connector.py        # Enhanced formatting for multiple types and fonts
├── stock_data_loader.py       # Stock CSV handling with dynamic expansion
├── report_analyzer.py         # 🆕 Enhanced with combination analysis
├── models.py                  # 🆕 Enhanced MatrixCell for multiple types
└── config.py                  # 🆕 Enhanced with display and font configuration
```

### Enhanced Data Flow Architecture
```
Stage 1: Stock Data Loading
    ↓
Stage 2: Enhanced PDF File Discovery & Classification
    ↓
Stage 3: Multiple Type Matrix Data Construction
    ↓  
Stage 4: Enhanced Categorization & Display Logic with Font Settings
    ↓
Stage 5: Google Sheets Upload / CSV Export with Multiple Type Formatting
    ↓
Stage 6: Enhanced Analysis & Combination Statistics
```

## 📁 File Structure Analysis

### Expected Directory Structure
```
project_root/
├── StockID_TWSE_TPEX.csv          # 116+ companies, 2 columns (代號, 名稱)
├── downloads/
│   ├── 2330/
│   │   ├── 202401_2330_AI1.pdf
│   │   ├── 202402_2330_AI1.pdf
│   │   ├── 202403_2330_AI1.pdf
│   │   ├── 202403_2330_A13.pdf    # Multiple types per quarter
│   │   └── metadata.json
│   ├── 8272/
│   │   ├── 202401_8272_A12.pdf
│   │   ├── 202402_8272_A12.pdf
│   │   ├── 202402_8272_A13.pdf    # Multiple individual types
│   │   └── metadata.json
│   └── 2382/
│       ├── 202404_2382_A13.pdf
│       └── metadata.json
└── mops_sheets_uploader.py
```

### PDF Filename Pattern Recognition
```python
# Pattern: YYYYQQ_COMPANYID_TYPE.pdf
FILENAME_PATTERN = r"^(\d{4})(\d{2})_(\d{4})_([A-Z0-9]+)\.pdf$"

# Examples:
# 202401_2330_AI1.pdf → Year: 2024, Quarter: 1, Company: 2330, Type: AI1
# 202403_8272_A12.pdf → Year: 2024, Quarter: 3, Company: 8272, Type: A12
# 202403_8272_A13.pdf → Year: 2024, Quarter: 3, Company: 8272, Type: A13
# → Result for 8272 Q3: "A12/A13" (multiple individual types)
```

## 🔧 Enhanced Technical Specifications

### Enhanced Class Interface Design

```python
class MOPSSheetsUploader:
    """MOPS PDF files to Google Sheets matrix uploader with multiple type support"""
    
    def __init__(self, 
                 downloads_dir: str = "downloads",
                 stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                 sheet_id: str = None,
                 max_years: int = 3,
                 csv_backup: bool = True,
                 # 🆕 Enhanced multiple type parameters
                 show_all_report_types: bool = True,
                 report_type_separator: str = "/",
                 use_categorized_display: bool = False,
                 max_display_types: int = 5,
                 # 🆕 Font and formatting parameters
                 font_size: int = 14,
                 header_font_size: int = 14,
                 bold_headers: bool = True,
                 bold_company_info: bool = True):
        """Initialize uploader with enhanced multiple type and font configuration"""
        
    def scan_pdf_files(self) -> Dict[str, Dict[str, str]]:
        """Scan downloads directory and return enhanced PDF file matrix"""
        
    def load_stock_data(self) -> pd.DataFrame:
        """Load and validate StockID_TWSE_TPEX.csv with dynamic size support"""
        
    def build_enhanced_matrix(self) -> pd.DataFrame:
        """🆕 Build company × quarter matrix with multiple report type support"""
        
    def upload_to_sheets(self, matrix_df: pd.DataFrame) -> bool:
        """Upload matrix to Google Sheets with enhanced multiple type and font formatting"""
        
    def export_to_csv(self, matrix_df: pd.DataFrame) -> str:
        """Export matrix to CSV with enhanced type combination analysis"""
        
    def generate_enhanced_report(self) -> Dict[str, Any]:
        """🆕 Generate comprehensive status report with multiple type analysis"""
        
    def run(self) -> bool:
        """Main execution method with enhanced multiple type and font processing"""
```

### Enhanced Multiple Report Type System

```python
# Enhanced Report Type Categories
REPORT_TYPE_CATEGORIES = {
    'individual': ['A12', 'A13'],           # Individual reports (highest priority)
    'consolidated': ['AI1', 'A1L'],         # Consolidated reports (secondary)
    'generic': ['A10', 'A11'],              # Generic reports (lower priority)
    'english': ['AIA', 'AE2'],              # English reports (excluded)
    'other': []                             # Any other types discovered
}

# Enhanced Display Configuration
DISPLAY_MODES = {
    'all_types': True,          # Show all available report types
    'best_only': False,         # Show only highest priority type
    'category_grouped': False   # Group by category with separators
}

# Enhanced PDF Classification System
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

# Enhanced Status Mapping for Display
STATUS_MAPPING = {
    1: '✅',  # Individual reports (preferred)
    2: '🟡',  # Consolidated reports (acceptable)
    3: '⚠️',  # Generic reports (caution)
    9: '❌',  # English or excluded reports
    None: '-', # No reports found
    'multiple': '🔄',  # Multiple types available (new)
    'mixed': '📊'      # Mixed categories available (new)
}

# Enhanced Cell Display Logic
class MatrixCell:
    def get_display_value(self, 
                         show_all_types: bool = True, 
                         separator: str = "/", 
                         max_types: int = 5,
                         use_categorized: bool = False) -> str:
        """
        🆕 Enhanced display value with multiple type support
        
        Examples:
        - show_all_types=True: "A12/A13/AI1"
        - show_all_types=False: "A12" (highest priority)
        - use_categorized=True: "A12/A13 → AI1"
        - max_types=2: "A12/A13+2" (truncated)
        """
        if not show_all_types:
            return self._get_best_type()
        
        if use_categorized:
            return self._get_categorized_display(separator)
            
        return self._get_all_types_display(separator, max_types)
```

### Quarter Column Generation (Enhanced)
```python
def generate_quarter_columns(pdf_data: Dict[str, List[PDFFile]], max_years: int = 3) -> List[str]:
    """
    Generate quarter column headers based on discovered PDFs + expected quarters
    
    Strategy:
    1. Scan all discovered PDFs to find actual quarters
    2. Add expected quarters within max_years range
    3. Sort in reverse chronological order (newest first)
    4. Handle future quarters gracefully
    5. Support multiple report types per quarter
    """
    discovered_quarters = set()
    current_year = datetime.now().year
    current_quarter = (datetime.now().month - 1) // 3 + 1
    
    # Scan discovered PDFs for actual quarters (enhanced for multiple types)
    for company_pdfs in pdf_data.values():
        for pdf in company_pdfs:
            discovered_quarters.add(f"{pdf.year} Q{pdf.quarter}")
    
    # Add expected quarters (current and past)
    expected_quarters = set()
    for year in range(current_year, current_year - max_years, -1):
        for quarter in [4, 3, 2, 1]:
            # Only add if not in future relative to current date
            if year < current_year or (year == current_year and quarter <= current_quarter):
                expected_quarters.add(f"{year} Q{quarter}")
    
    # Combine discovered and expected quarters
    all_quarters = discovered_quarters.union(expected_quarters)
    
    # Sort chronologically (newest first)
    def quarter_sort_key(q):
        year, quarter = q.split(' Q')
        return (int(year), int(quarter))
    
    sorted_quarters = sorted(all_quarters, key=quarter_sort_key, reverse=True)
    
    # Log warnings for unexpected future quarters
    future_quarters = [q for q in discovered_quarters 
                      if quarter_sort_key(q) > (current_year, current_quarter)]
    if future_quarters:
        logger.warning(f"⚠️ 發現未來季度PDF: {', '.join(future_quarters)}")
        logger.warning("   這可能是檔案命名錯誤或系統時間問題")
    
    return sorted_quarters
```

## 📋 Enhanced Component Specifications

### 1. Enhanced PDF Scanner Module (`pdf_scanner.py`)

**Purpose**: Discover and catalog all PDF files in the downloads directory with multiple type support

```python
class PDFScanner:
    def scan_downloads_directory(self, downloads_dir: str) -> Dict[str, List[PDFFile]]:
        """
        Scan downloads directory structure and return organized PDF files
        
        Enhanced to handle:
        - Multiple report types per quarter (A12 + A13 + AI1)
        - Current and past quarters (expected)
        - Future quarters (with warnings)
        - Invalid filenames (logged and skipped)
        - Type categorization and priority analysis
        
        Returns:
            Dict mapping company_id to list of PDFFile objects with enhanced metadata
        """
        
    def discover_available_quarters(self, pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """
        Discover all quarters that have at least one PDF file
        
        Enhanced to handle multiple types per quarter
        
        Returns:
            Sorted list of quarters (e.g., ['2025 Q4', '2025 Q3', '2025 Q1', '2024 Q4'])
        """
        
    def parse_pdf_filename(self, filename: str) -> Optional[PDFMetadata]:
        """
        Parse PDF filename to extract metadata with validation
        
        Enhanced with type categorization and priority scoring
        
        Args:
            filename: PDF filename (e.g., "202501_2330_AI1.pdf" for 2025 Q1)
            
        Returns:
            Enhanced PDFMetadata object with category and priority info
            
        Validation:
        - Year range: 2020-2030 (prevents obvious errors)
        - Quarter range: 1-4 
        - Company ID: 4-digit format
        - Report type categorization (Individual/Consolidated/Generic/English)
        - Warns about future quarters beyond current date
        """
        
    def categorize_report_type(self, report_type: str) -> Dict[str, Any]:
        """🆕 Categorize report type and assign priority"""
        
    def validate_pdf_file(self, file_path: str) -> bool:
        """Validate that file exists and is a valid PDF with size check"""
```

### 2. Enhanced Stock Data Loader (`stock_data_loader.py`)

**Purpose**: Load and process company information from StockID_TWSE_TPEX.csv (dynamically sized)

```python
class StockDataLoader:
    def load_stock_csv(self, csv_path: str) -> pd.DataFrame:
        """
        Load StockID_TWSE_TPEX.csv with proper encoding and validation
        
        Expected format:
        - Rows: Variable number of companies (116+ as it grows)
        - Columns: 代號 (Integer), 名稱 (String)
        - Encoding: UTF-8
        
        Handles:
        - Dynamic company list size (auto-detects current count)
        - New company additions (gracefully expands matrix)
        - Company removals (logs warnings, continues processing)
        - Duplicate company codes (validates and deduplicates)
        """
        
    def validate_stock_data(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Validate stock data completeness and format
        
        Returns validation info including:
        - Current company count (was 116, now dynamic)
        - New companies detected since last run
        - Removed companies detected since last run  
        - Duplicate detection and resolution
        """
        
    def detect_stock_list_changes(self, df: pd.DataFrame, 
                                 previous_matrix: pd.DataFrame = None) -> Dict[str, Any]:
        """
        🆕 Detect changes in stock list since last matrix generation
        
        Returns:
        - added_companies: List of new company codes
        - removed_companies: List of removed company codes  
        - total_companies: Current total count
        - change_summary: Human-readable summary of changes
        """
        
    def get_company_name(self, stock_id: str) -> str:
        """Get company name by stock ID from the current company list"""
```

### 3. Enhanced Matrix Builder (`matrix_builder.py`)

**Purpose**: Construct the main data matrix combining stock data and PDF information with multiple type support

```python
class MatrixBuilder:
    def __init__(self, stock_df: pd.DataFrame, pdf_data: Dict[str, List[PDFFile]],
                 show_all_types: bool = True, font_size: int = 14):
        """
        Initialize with stock data and PDF scan results
        
        Enhanced Args:
            stock_df: DataFrame with N companies (代號, 名稱) - dynamically sized
            pdf_data: Dictionary mapping company_id to PDFFile list
            show_all_types: Enable multiple type display (new)
            font_size: Font size for Google Sheets display (new)
        """
        
    def build_base_matrix(self, max_years: int = 3) -> pd.DataFrame:
        """
        Create base matrix with companies and dynamically discovered quarter columns
        
        Enhanced Process:
        1. Start with N companies from StockID_TWSE_TPEX.csv (auto-detects count)
        2. Discover quarters from actual PDF files (including multiple types)
        3. Add expected quarters within max_years range  
        4. Create matrix with dynamic row and column count
        5. Prepare for multiple report type display
        
        Returns:
            Enhanced DataFrame with shape (N, 2 + dynamic_quarters) where:
            - N rows: current number of companies in CSV (116, 118, etc.)
            - Columns: 代號, 名稱, [discovered quarters in reverse chronological order]
            - Enhanced cells ready for multiple type display
            
        Examples:
            Original: (116, 14) → 116 companies, 12 quarters + 2 base columns
            After adding 2: (118, 14) → 118 companies, 12 quarters + 2 base columns
            With future PDFs: (118, 18) → 118 companies, 16 quarters + 2 base columns
        """
        
    def populate_enhanced_pdf_status(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """
        🆕 Fill matrix with enhanced PDF availability status including multiple types
        
        Enhanced Handles:
        - Multiple report types per quarter (A12/A13/AI1)
        - Category-based grouping and display
        - Existing companies with PDFs (normal processing)
        - New companies without PDFs (shows "-" for all quarters)  
        - Companies with partial PDF coverage
        - Type combination analysis and statistics
        """
        
    def identify_new_companies(self, matrix_df: pd.DataFrame, 
                              pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """
        🆕 Identify companies in stock list but not in PDF data
        
        Returns:
            List of company codes that have no PDFs downloaded yet
        """
        
    def apply_enhanced_priority_rules(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """🆕 Apply enhanced report type display rules for multiple types"""
        
    def add_enhanced_summary_columns(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """🆕 Add enhanced summary columns with multiple type analysis"""
```

### 4. Enhanced Sheets Connector (`sheets_connector.py`)

**Purpose**: Handle Google Sheets integration with multiple type and font support

```python
from sheets_uploader import SheetsUploader

class MOPSSheetsConnector(SheetsUploader):
    """
    Enhanced sheets connector for MOPS matrix data with multiple type and font support
    
    Extends your existing SheetsUploader class to reuse:
    - Google Sheets authentication (GOOGLE_SHEETS_CREDENTIALS)
    - Connection management (_setup_connection)
    - Error handling and retry logic
    - CSV fallback functionality
    
    Enhanced with:
    - Multiple report type formatting
    - Configurable font size (14pt default)
    - Enhanced visual styling
    
    Target Sheet: ID 1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0
    Target Worksheet: "MOPS下載狀態" (will be created if doesn't exist)
    """
    
    def __init__(self, worksheet_name: str = "MOPS下載狀態",
                 font_size: int = 14, header_font_size: int = 14,
                 bold_headers: bool = True, bold_company_info: bool = True):
        """Initialize using existing Google Sheets configuration with enhanced formatting"""
        super().__init__()  # Inherits your existing setup
        self.mops_worksheet_name = worksheet_name
        self.font_size = font_size
        self.header_font_size = header_font_size
        self.bold_headers = bold_headers
        self.bold_company_info = bold_company_info
        
    def upload_enhanced_matrix(self, matrix_df: pd.DataFrame) -> bool:
        """
        🆕 Upload matrix data to Google Sheets with enhanced multiple type formatting
        
        Enhanced Process:
        1. Use inherited _setup_connection() method
        2. Create or find "MOPS下載狀態" worksheet  
        3. Auto-resize worksheet for current matrix dimensions
        4. Upload matrix data with enhanced multiple type formatting
        5. Apply enhanced MOPS-specific styling with 14pt font
        6. Format multiple report type cells with special highlighting
        7. Apply font size settings to all content and headers
        8. Fallback to CSV if API limits hit (inherited behavior)
        
        Handles:
        - Dynamic matrix sizes (116→118→N companies)
        - Auto-worksheet resizing when companies added/removed
        - Multiple report type display formatting
        - Enhanced font sizing and bold formatting
        - Maintains formatting consistency regardless of size changes
        """
        
    def format_enhanced_matrix_worksheet(self, worksheet, data_rows: int, 
                                       future_quarters: List[str] = None,
                                       new_companies: List[str] = None,
                                       multiple_type_cells: List[str] = None) -> None:
        """
        🆕 Apply enhanced MOPS-specific formatting for matrix display
        
        Enhanced Features:
        - 14pt font size for all content (configurable)
        - 14pt font size for headers (configurable)  
        - Bold formatting for headers (configurable)
        - Bold formatting for company info columns (configurable)
        - Header row styling (company info vs quarters)
        - Future quarter column highlighting (orange background)
        - New company row highlighting (light blue background)
        - Multiple type cell highlighting (purple background) 🆕
        - Status symbol color coding (✅🟡⚠️❌🔄📊)
        - Auto-resize columns for readability
        - Freeze first two columns (代號, 名稱)
        - Dynamic range formatting based on current matrix size
        """
        
    def auto_resize_worksheet(self, worksheet, required_rows: int, 
                             required_cols: int) -> None:
        """
        🆕 Automatically resize worksheet to fit matrix dimensions
        
        Enhanced Args:
            required_rows: Current company count + header (e.g., 118 + 1 = 119)
            required_cols: Base columns + quarter columns (e.g., 2 + 12 = 14)
            
        Enhanced Handles:
        - Expanding when companies added
        - Shrinking when companies removed (with safety buffer)
        - Column additions for new quarters discovered
        - Font size adjustments for cell dimensions
        """
        
    def create_enhanced_summary_statistics(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """
        🆕 Create enhanced summary statistics for separate worksheet upload
        
        Enhanced Statistics:
        - Total companies (116+)
        - Companies with PDFs vs without
        - Coverage by quarter
        - Enhanced report type distribution
        - Multiple type analysis and combinations 🆕
        - Type diversity scoring 🆕
        - Future quarter warnings
        """
        
    def upload_with_csv_fallback(self, matrix_df: pd.DataFrame) -> bool:
        """
        Upload with automatic CSV fallback (inherits from parent)
        
        Uses inherited CSV-only mode when:
        - API limits exceeded 
        - Authentication failures
        - Worksheet conflicts
        """
```

### 5. Enhanced Report Analyzer (`report_analyzer.py`)

**Purpose**: Analyze PDF files and generate insights with multiple type support

```python
class ReportAnalyzer:
    def __init__(self, show_all_types: bool = True):
        """Initialize with enhanced multiple type analysis support"""
        self.show_all_types = show_all_types
        
    def analyze_enhanced_coverage(self, matrix_df: pd.DataFrame) -> Dict[str, Any]:
        """🆕 Analyze report coverage statistics with multiple type support"""
        
    def analyze_report_combinations(self, matrix_df: pd.DataFrame) -> Dict[str, Any]:
        """🆕 Analyze report type combinations and patterns"""
        
    def calculate_type_diversity(self, matrix_df: pd.DataFrame) -> Dict[str, float]:
        """🆕 Calculate type diversity scores for companies"""
        
    def identify_enhanced_missing_reports(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """Generate list of missing reports with enhanced priority analysis"""
        
    def generate_download_suggestions(self, matrix_df: pd.DataFrame) -> List[str]:
        """Suggest which companies/quarters to download next with type analysis"""
```

## 📊 Enhanced Data Models

### Enhanced PDFFile Data Class
```python
@dataclass
class PDFFile:
    """Represents a discovered PDF file with enhanced metadata"""
    company_id: str
    year: int
    quarter: int
    report_type: str
    filename: str
    file_path: str
    file_size: int
    modified_date: datetime
    # 🆕 Enhanced fields
    category: str = field(default="")  # individual, consolidated, generic, english
    priority: int = field(default=9)   # Priority score
    
    @property
    def quarter_key(self) -> str:
        """Return quarter key for matrix (e.g., '2024 Q1')"""
        return f"{self.year} Q{self.quarter}"
    
    @property
    def is_individual(self) -> bool:
        """🆕 Check if this is an individual report type"""
        return self.category == "individual"
    
    @property
    def is_consolidated(self) -> bool:
        """🆕 Check if this is a consolidated report type"""
        return self.category == "consolidated"

@dataclass
class EnhancedMatrixCell:
    """🆕 Represents a single cell in the matrix with multiple type support"""
    report_types: Set[str] = field(default_factory=set)
    file_paths: List[str] = field(default_factory=list)
    categories: Set[str] = field(default_factory=set)
    priorities: List[int] = field(default_factory=list)
    
    @property
    def has_pdf(self) -> bool:
        """Check if cell has any PDFs"""
        return len(self.report_types) > 0
    
    @property
    def has_multiple_types(self) -> bool:
        """🆕 Check if cell has multiple report types"""
        return len(self.report_types) > 1
    
    @property
    def has_mixed_categories(self) -> bool:
        """🆕 Check if cell has mixed categories (individual + consolidated)"""
        return len(self.categories) > 1
    
    def get_display_value(self, separator: str = "/", max_types: int = 5, 
                         use_categorized: bool = False) -> str:
        """🆕 Get formatted display value with multiple type support"""
        if not self.report_types:
            return "-"
            
        if use_categorized:
            return self._get_categorized_display(separator)
        
        sorted_types = sorted(self.report_types, 
                            key=lambda x: REPORT_TYPE_PRIORITY.get(x, 9))
        
        if len(sorted_types) <= max_types:
            return separator.join(sorted_types)
        else:
            visible_types = sorted_types[:max_types-1]
            remaining_count = len(sorted_types) - len(visible_types)
            return separator.join(visible_types) + f"+{remaining_count}"
```

## 📋 Enhanced Configuration System

### Enhanced Configuration File (`config.yaml`)
```yaml
# MOPS Sheets Uploader Configuration v1.1.1

# Directory settings
downloads_dir: "downloads"
stock_csv_path: "StockID_TWSE_TPEX.csv"
output_dir: "data/reports"

# Matrix settings
max_years: 3
include_current_year: true
quarter_order: "desc"  # newest first
dynamic_columns: true  # adjust columns based on discovered PDFs
dynamic_rows: true     # 🆕 adjust rows based on stock list changes

# 🆕 Enhanced Multiple Report Type Settings
multiple_report_types:
  show_all_report_types: true          # Enable multiple type display (default: true)
  report_type_separator: "/"           # Separator between types (default: /)
  category_separator: " → "            # Separator between categories (default: " → ")
  max_display_types: 5                 # Max types before truncation (default: 5)
  use_categorized_display: false       # Group by category (default: false)
  priority_display_mode: "all"        # 'all', 'best', 'category_best' (default: all)
  show_type_counts: false              # Show counts like A12(3) (default: false)
  truncate_indicator: "+"              # Truncation indicator (default: +)
  exclude_english_reports: true       # Exclude English types (default: true)
  individual_reports_priority: true   # Prioritize individual over consolidated

# 🆕 Font and Display Settings
display_formatting:
  font_size: 14                        # Font size for all content in pt (default: 14)
  header_font_size: 14                 # Header font size in pt (default: 14)
  bold_headers: true                   # Bold formatting for headers (default: true)
  bold_company_info: true              # Bold formatting for company columns (default: true)
  highlight_multiple_types: true      # Color-code multiple type cells (default: true)
  use_enhanced_symbols: true          # Use enhanced status symbols (default: true)
  include_type_analysis: true         # Include combination analysis (default: true)

# 🆕 Stock list change handling
stock_list_changes:
  detect_changes: true         # Enable change detection
  highlight_new_companies: true  # Highlight new companies in sheets
  archive_orphaned_pdfs: false   # Move PDFs of removed companies to archive
  warn_on_removals: true         # Warn when companies are removed
  auto_suggest_downloads: true   # Suggest downloads for new companies
  change_threshold_warning: 5    # Warn if >5 companies added/removed at once

# 🆕 Future quarter handling
future_quarter_handling:
  include_in_matrix: true      # Show future quarters in output
  warn_threshold_months: 6     # Warn if more than 6 months in future
  mark_suspicious: true        # Mark with ⚠️ indicator
  max_future_quarters: 8       # Limit how far future to include
  generate_future_report: true # Create separate report of future PDFs

# Google Sheets settings
google_sheets:
  worksheet_name: "MOPS下載狀態"
  include_summary_sheet: true
  auto_resize_columns: true
  highlight_future_quarters: true  # 🆕 Color-code future quarter cells
  freeze_company_columns: true     # Freeze first two columns
  apply_conditional_formatting: true

# CSV export settings
csv_export:
  csv_backup: true
  csv_filename_pattern: "mops_matrix_{timestamp}.csv"
  include_future_analysis: true   # 🆕 Add future quarter analysis to CSV
  include_type_combinations: true # 🆕 Add type combination analysis

# Report type settings (enhanced)
report_types:
  preferred_types: ["A12", "A13"]
  acceptable_types: ["AI1", "A1L"]
  excluded_types: ["AIA", "AE2"]
  category_definitions:
    individual: ["A12", "A13"]
    consolidated: ["AI1", "A1L"]
    generic: ["A10", "A11"]
    english: ["AIA", "AE2"]

# Display settings (enhanced)
display_settings:
  use_symbols: true
  show_file_counts: false
  highlight_missing: true
  future_quarter_symbol: "⏭️"    # 🆕 Symbol for future quarters
  multiple_types_symbol: "🔄"    # 🆕 Symbol for multiple types
  mixed_categories_symbol: "📊"  # 🆕 Symbol for mixed categories
```

### Enhanced Environment Variables
```bash
# Google Sheets Integration (from your updated .env)
GOOGLE_SHEETS_CREDENTIALS={"type":"service_account","project_id":"coherent-vision-463514-n3",...}
GOOGLE_SHEET_ID=1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0

# Enhanced Multiple Report Type Settings
MOPS_SHOW_ALL_TYPES=true                   # Show all report types (default: true)
MOPS_TYPE_SEPARATOR=/                      # Type separator (default: /)
MOPS_CATEGORY_SEPARATOR=" → "              # Category separator (default: " → ")
MOPS_MAX_DISPLAY_TYPES=5                   # Max types before truncation (default: 5)
MOPS_CATEGORIZED_DISPLAY=false             # Use categorized display (default: false)
MOPS_PRIORITY_MODE=all                     # Priority mode: all/best/category_best

# 🆕 Font and Display Settings
MOPS_FONT_SIZE=14                          # Font size in pt (default: 14)
MOPS_HEADER_FONT_SIZE=14                   # Header font size in pt (default: 14)
MOPS_BOLD_HEADERS=true                     # Bold headers (default: true)
MOPS_BOLD_COMPANY_INFO=true                # Bold company info columns (default: true)

# Enhanced Analysis Settings  
MOPS_HIGHLIGHT_MULTIPLE=true               # Highlight multiple type cells
MOPS_INCLUDE_TYPE_ANALYSIS=true            # Include type combination analysis

# Optional: Custom configuration
MOPS_CONFIG_PATH=path/to/config.yaml
MOPS_MAX_YEARS=3
MOPS_DOWNLOADS_DIR=downloads
MOPS_WORKSHEET_NAME=MOPS下載狀態  # Will create this worksheet in your existing sheet
```

**Integration Note**: The MOPS Sheets Uploader will use your existing Google Sheets service account and spreadsheet, creating a new worksheet called "MOPS下載狀態" within your current spreadsheet (ID: 1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0).

## 🚦 Enhanced Error Handling Strategy

### Enhanced Error Categories
1. **File System Errors**: Missing directories, permission issues, corrupted files
2. **Data Validation Errors**: Invalid stock CSV format, malformed PDF filenames
3. **Google Sheets Errors**: API limits, authentication failures, worksheet conflicts
4. **Processing Errors**: Memory issues with large datasets, encoding problems
5. **Temporal Anomalies**: Future quarter PDFs, invalid date ranges, system clock issues
6. **🆕 Stock List Changes**: Company additions, removals, duplicate codes
7. **🆕 Multiple Type Processing**: Type combination parsing, display formatting errors
8. **🆕 Font and Display Errors**: Google Sheets formatting failures, cell overflow

### Enhanced Recovery Mechanisms
```python
class EnhancedErrorHandler:
    def handle_missing_stock_csv(self) -> pd.DataFrame:
        """Generate default stock list from discovered companies"""
        
    def handle_sheets_api_failure(self, matrix_df: pd.DataFrame) -> str:
        """Fallback to CSV export when Sheets API fails"""
        
    def handle_parsing_errors(self, invalid_files: List[str]) -> None:
        """Log and skip files that don't match expected patterns"""
        
    def handle_future_quarters(self, future_pdfs: List[PDFFile]) -> Dict[str, Any]:
        """Handle PDFs with future quarter dates (enhanced logic)"""
        
    def handle_multiple_type_errors(self, problematic_cells: List[str]) -> Dict[str, Any]:
        """🆕 Handle multiple type processing errors"""
        
    def handle_font_formatting_errors(self, worksheet, error_cells: List[str]) -> bool:
        """🆕 Handle Google Sheets font formatting failures"""
        
    def handle_stock_list_changes(self, added: List[str], removed: List[str], 
                                 pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, Any]:
        """
        🆕 Enhanced handling of stock list changes
        
        Enhanced Strategy for additions:
        1. Add new companies to matrix with empty cells
        2. Log them as "needs download" in enhanced statistics  
        3. Generate download suggestions with type preferences
        4. Highlight in Google Sheets with special formatting and proper fonts
        
        Enhanced Strategy for removals:
        1. Remove companies from matrix
        2. Log orphaned PDF files (PDFs for removed companies)
        3. Optionally move orphaned PDFs to archive folder
        4. Update enhanced statistics to reflect new baseline
        
        Returns:
            Dict with enhanced change analysis and recommendations
        """
        
    def validate_quarter_range(self, year: int, quarter: int) -> bool:
        """Validate if quarter/year combination is reasonable (enhanced logic)"""
```

## 📊 Enhanced Usage Examples

### Basic Usage with Enhanced Features
```python
from mops_sheets_uploader import MOPSSheetsUploader

# Initialize uploader with enhanced defaults (uses .env configuration automatically)
uploader = MOPSSheetsUploader(
    downloads_dir="downloads",
    stock_csv_path="StockID_TWSE_TPEX.csv",
    max_years=3,
    # 🆕 Enhanced defaults
    show_all_report_types=True,      # Show all available types
    font_size=14,                    # 14pt font for readability
    bold_headers=True                # Bold headers
)

# Run complete process with enhanced multiple type support
success = uploader.run()

if success:
    print("✅ Enhanced matrix uploaded to Google Sheets successfully")
    print("🔄 Multiple report types displayed per quarter")
    print("🎨 Professional 14pt font styling applied")
    print("🔗 View at: https://docs.google.com/spreadsheets/d/1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0")
    print("📊 Worksheet: MOPS下載狀態")
else:
    print("⚠️ Upload failed, but CSV backup is available")
```

### Advanced Configuration with Font Settings
```python
# Enhanced configuration with custom font and display settings
uploader = MOPSSheetsUploader(
    downloads_dir="./data/downloads",
    stock_csv_path="./data/StockID_TWSE_TPEX.csv",
    sheet_id="1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0",  # Your specific sheet
    max_years=5,
    csv_backup=True,
    worksheet_name="MOPS財報矩陣",  # Custom worksheet name
    # 🆕 Enhanced multiple type configuration
    show_all_report_types=True,
    use_categorized_display=True,    # Group Individual → Consolidated
    category_separator=" → ",        # Use arrow for categories
    report_type_separator="/",       # Use slash within categories  
    max_display_types=8,             # Allow more types before truncation
    # 🆕 Enhanced font and formatting settings
    font_size=16,                    # Larger 16pt font for all cells
    header_font_size=18,             # Even larger font for headers
    bold_headers=True,               # Bold headers
    bold_company_info=True           # Bold company info columns
)

# Test connection using existing credentials
if uploader.test_connection():
    print("✅ Google Sheets 連線成功")
    
    # Get enhanced analysis including change detection and type combinations
    matrix_df = uploader.build_enhanced_matrix()
    enhanced_stats = uploader.generate_enhanced_report()

    print(f"📊 Enhanced Statistics:")
    print(f"   Total companies: {enhanced_stats['total_companies']}")  # Could be 116, 118, etc.
    print(f"   Coverage rate: {enhanced_stats['coverage_percentage']:.1f}%")
    print(f"   Missing reports: {enhanced_stats['missing_count']}")
    print(f"   Multiple types percentage: {enhanced_stats['multiple_types_percentage']:.1f}%")
    print(f"   Most common combination: {enhanced_stats['top_combination']}")
    
    # Check for stock list changes
    if enhanced_stats.get('stock_changes'):
        changes = enhanced_stats['stock_changes']
        if changes['added_companies']:
            print(f"🆕 新增公司: {', '.join(changes['added_companies'])}")
        if changes['removed_companies']:
            print(f"🗑️ 移除公司: {', '.join(changes['removed_companies'])}")

    # Display enhanced type combination analysis
    combinations = enhanced_stats.get('report_type_combinations', {})
    print(f"🔄 Report Type Combinations:")
    for combo, count in sorted(combinations.items(), key=lambda x: x[1], reverse=True)[:5]:
        print(f"   • {combo}: {count} occurrences")

    # Export for further analysis
    csv_path = uploader.export_to_csv(matrix_df)
    print(f"Enhanced matrix exported to: {csv_path}")
else:
    print("❌ Google Sheets 連線失敗，請檢查 .env 設定")
```

### CLI Interface with Enhanced Options
```bash
# Basic usage with enhanced multiple type display and 14pt font
python mops_sheets_uploader.py --show-all-types --font-size=14

# Categorized display mode with large custom font
python mops_sheets_uploader.py \
    --show-all-types \
    --categorized-display \
    --type-separator="/" \
    --category-separator=" → " \
    --font-size=16 \
    --header-font-size=18 \
    --bold-headers \
    --bold-company-info

# Best-only mode (traditional behavior) with large font
python mops_sheets_uploader.py \
    --priority-mode=best \
    --no-multiple-types \
    --font-size=16 \
    --header-font-size=16

# Enhanced analysis with detailed type combinations and custom formatting
python mops_sheets_uploader.py \
    --analyze-only \
    --include-type-analysis \
    --font-size=14 \
    --output enhanced_analysis.json

# Custom configuration with specific font preferences
python mops_sheets_uploader.py \
    --config custom_config.yaml \
    --font-size=15 \
    --header-font-size=17 \
    --max-display-types=6
```

## 📊 Enhanced Output Examples

### Console Output (Enhanced v1.1.1)

**Scenario 1: Enhanced Processing with Multiple Types and Font Settings**
```
📊 MOPS Sheets Uploader v1.1.1 (Complete Enhanced)
============================================================

🔍 掃描 PDF 檔案...
   發現 68 個 PDF 檔案 (32 家公司)
   季度範圍: 2024 Q1 到 2025 Q1 (5 個季度)
   📊 報告類型分布: A12(23), A13(15), AI1(18), A1L(12)
   
📋 載入股票清單...
   已載入 116 家公司資料 (StockID_TWSE_TPEX.csv)
   
🏗️ 建構增強型矩陣資料...
   矩陣大小: 116 × 7 (公司 × 季度)
   📊 支援多重報告類型顯示: 是
   🎨 字體設定: 內容 14pt, 標題 14pt, 粗體標題
   動態季度欄位: 2025 Q1, 2024 Q4, 2024 Q3, 2024 Q2, 2024 Q1
   
📊 填入增強型 PDF 狀態資料...
   填入狀態: 156/580 個儲存格 (26.9%)
   🔄 多重類型: 23/156 個儲存格 (14.7%)
   📊 混合類別: 8/156 個儲存格 (5.1%)
   常見組合: A12/A13 (8次), AI1/A1L (5次), A12/AI1 (3次)
   
📈 上傳統計:
   ✅ 個別財報: 43 份 (63.2%)
   🟡 合併財報: 25 份 (36.8%)  
   🔄 多重類型: 23 個季度 (14.7%)
   📊 類型組合: 12 種不同組合模式
   🎯 類型多樣性平均分數: 6.8/10
   
🚀 上傳到 Google Sheets...
   ✅ 增強型矩陣上傳成功 (116 公司 × 5 季度)
   🔄 多重報告類型格式化完成
   🎨 14pt 字體和粗體標題套用完成
   📊 多重類型儲存格特殊標示完成
   
💡 增強分析結果:
   • 最常見組合: A12/A13 (個別報告雙重覆蓋)
   • 混合類型: 8 家公司同時有個別和合併報告
   • 類型多樣性分數: 7.2/10 (良好的報告類型覆蓋)
   • 字體設定: 內容與標題皆為 14pt，標題為粗體
```

**Scenario 2: New Companies Added with Enhanced Display**
```
📊 MOPS Sheets Uploader v1.1.1 (Complete Enhanced)
============================================================

🔍 掃描 PDF 檔案...
   發現 68 個 PDF 檔案 (32 家公司)
   季度範圍: 2024 Q1 到 2025 Q1 (5 個季度)
   
📋 載入股票清單...
   已載入 118 家公司資料 (StockID_TWSE_TPEX.csv)
   🆕 偵測到股票清單變更:
      • 新增公司: 2334 (聯發科), 1234 (測試公司)
      • 移除公司: 無
      • 總公司數: 116 → 118 (+2)
   
🏗️ 建構增強型矩陣資料...
   矩陣大小: 118 × 7 (公司 × 季度)  
   📊 支援多重報告類型顯示: 是
   🎨 字體設定: 內容 14pt, 標題 14pt, 粗體標題和公司資訊
   動態季度欄位: 2025 Q1, 2024 Q4, 2024 Q3, 2024 Q2, 2024 Q1
   
📊 填入增強型 PDF 狀態資料...
   填入狀態: 156/590 個儲存格 (26.4%)
   🔄 多重類型: 23/156 個儲存格 (14.7%)
   📈 新公司狀態: 2 家公司無PDF檔案 (需要下載)
   
🚀 上傳到 Google Sheets...
   ✅ 主矩陣上傳成功 (118 公司 × 5 季度)
   📊 工作表已自動調整大小以容納新公司
   🆕 新公司列特殊標示完成 (淡藍色背景)
   🎨 14pt 字體和粗體格式套用完成
   
💡 建議: 請為以下新公司下載PDF:
   • 2334 (聯發科): 缺少所有季度，建議優先下載 A12/A13 個別報告
   • 1234 (測試公司): 缺少所有季度，建議優先下載 A12/A13 個別報告
```

### Enhanced CSV Output Sample
```csv
代號,名稱,2025 Q2,2025 Q1,2024 Q4,2024 Q3,2024 Q2,2024 Q1,總報告數,涵蓋率,個別報告,合併報告,多重類型,類型多樣性,最常見組合,混合類別
2330,台積電,AI1,AI1,AI1,AI1,A13/AI1,AI1,6,100.0%,1,6,1,0.33,AI1,是
8272,全景軟體,A12,A12,A12/A13,A12,-,-,4,66.7%,4,0,1,0.75,A12/A13,否
2382,廣達,-,-,A13,-,-,-,1,16.7%,1,0,0,1.00,A13,否
1234,混合範例,A12/AI1,A12/A13/AI1,A12/A13+2,AI1/A1L,A12,-,8,83.3%,4,4,3,0.80,A12/AI1,是
2334,聯發科,-,-,-,-,-,-,0,0.0%,0,0,0,0.00,-,否
...,(114 more companies with enhanced type display),...
```

### Enhanced Google Sheets Formatting Preview
```
Applied Formatting:
✅ 14pt font size for all content cells
✅ 14pt font size for headers (or custom size if configured)
✅ Bold formatting for header row
✅ Bold formatting for 代號 and 名稱 columns  
✅ Multiple report type cells highlighted with light purple background
✅ New company rows highlighted with light blue background
✅ Future quarter columns highlighted with light orange background
✅ Auto-resized columns for optimal readability
✅ Frozen first two columns (代號, 名稱) for easy scrolling
✅ Conditional formatting for status symbols (✅🟡⚠️❌🔄📊)
```

## 🆕 Enhanced Features Detail

### 1. Enhanced Multiple Report Type Display Engine

**Complete Enhancement**: The system now provides comprehensive multiple type support:

1. **Collects All Available Types**: Gathers all PDF report types for each company-quarter
2. **Enhanced Display Logic**: Formats multiple types with configurable separators and font sizing
3. **Category-Based Grouping**: Optional grouping by Individual → Consolidated → Generic
4. **Smart Truncation**: Handles cases with many types (A12/A13/AI1+2 more)
5. **Professional Font Styling**: Configurable 14pt font with bold headers for optimal readability

### 2. Enhanced Font and Display Control System

**New Professional Styling Features**:
- **Configurable Font Size**: Default 14pt for optimal readability in Google Sheets
- **Independent Header Sizing**: Separate font size control for headers (default 14pt)
- **Bold Formatting Options**: Configurable bold for headers and company info columns
- **Enhanced Visual Hierarchy**: Clear distinction between different content types
- **Cross-Platform Consistency**: Ensures consistent appearance across different devices

### 3. Enhanced Configuration Template System

**Configuration Templates for Different Use Cases**:

```python
# Template 1: Professional (default)
professional_config = {
    'font_size': 14,
    'header_font_size': 14,
    'bold_headers': True,
    'bold_company_info': True,
    'show_all_report_types': True
}

# Template 2: Large Display (for presentations)
large_display_config = {
    'font_size': 16,
    'header_font_size': 18,
    'bold_headers': True,
    'bold_company_info': True,
    'show_all_report_types': True
}

# Template 3: Compact (for smaller screens)
compact_config = {
    'font_size': 12,
    'header_font_size': 13,
    'bold_headers': True,
    'bold_company_info': False,
    'show_all_report_types': True,
    'max_display_types': 3
}
```

## 🧪 Enhanced Testing Strategy

### Unit Tests (Enhanced)
```python
# Test enhanced multiple report type display logic
def test_enhanced_multiple_type_display():
    cell = EnhancedMatrixCell()
    cell.add_pdf(create_pdf("A12"))
    cell.add_pdf(create_pdf("A13"))
    cell.add_pdf(create_pdf("AI1"))
    
    # Test different display modes
    assert cell.get_display_value(show_all_types=True) == "A12/A13/AI1"
    assert cell.get_display_value(show_all_types=False) == "A12"  # Best only
    assert cell.get_categorized_display() == "A12/A13 → AI1"
    assert cell.has_multiple_types == True
    assert cell.has_mixed_categories == True

# Test enhanced matrix building with fonts
def test_enhanced_matrix_building_with_fonts():
    matrix = build_enhanced_test_matrix(font_size=14)
    assert "A12/A13" in matrix.values  # Multiple types present
    assert matrix.shape == (116, 14)   # Same dimensions as before
    assert matrix.font_config['size'] == 14

# Test font configuration
def test_font_configuration():
    config = EnhancedMOPSConfig(font_size=16, header_font_size=18)
    assert config.font_size == 16
    assert config.header_font_size == 18
    assert config.bold_headers == True  # Default

# Test display templates
def test_enhanced_display_templates():
    config = apply_display_template(EnhancedMOPSConfig(), "professional")
    assert config.font_size == 14
    assert config.header_font_size == 14
    assert config.bold_headers == True
```

### Integration Tests (Enhanced)
- **End-to-end multiple type processing** with enhanced font support
- **Google Sheets formatting** with 14pt font and multiple type cells
- **CSV export validation** with enhanced columns and font metadata
- **Configuration template testing** for different display modes and font sizes
- **Font rendering validation** across different Google Sheets environments

### Performance Tests (Enhanced)
- **Enhanced Processing Efficiency**: Parallel type analysis with font formatting
- **Font Memory Management**: Efficient font configuration storage and application
- **Large Dataset Handling**: Test with 200+ companies and complex type combinations
- **Google Sheets API Efficiency**: Batch formatting operations with font settings

## 📈 Performance Optimizations (Enhanced)

### Enhanced Processing Efficiency
- **Parallel Type Analysis**: Concurrent processing of type combinations with font metadata
- **Smart Caching**: Cache type categorization and font formatting results
- **Optimized Display Logic**: Efficient string formatting for multiple types with font info
- **Memory Management**: Efficient storage of type sets and font configurations per cell
- **Batch Google Sheets Operations**: Single API call for content and font formatting

## 🔄 Migration Guide

### From v1.0.0 to v1.1.1
```python
# v1.0.0 behavior (legacy mode)
uploader = MOPSSheetsUploader(
    show_all_report_types=False,  # Disable enhanced features
    priority_display_mode="best"  # Show only best type
    # No font configuration needed
)

# v1.1.1 enhanced behavior (recommended)
uploader = MOPSSheetsUploader(
    show_all_report_types=True,   # Enable multiple type display
    report_type_separator="/",    # Configure separator
    max_display_types=5,          # Configure truncation
    font_size=14,                 # Professional font size
    bold_headers=True             # Enhanced readability
)
```

### Configuration File Migration
```yaml
# Add these sections to existing config.yaml
multiple_report_types:
  show_all_report_types: true
  
display_formatting:
  font_size: 14
  header_font_size: 14
  bold_headers: true
```

---

## 🔄 Version History (Complete)

### v1.1.1 (2025-08-08) - Complete Enhanced Version
- **📋 Complete**: Merged comprehensive v1.0.0 foundation with v1.1.0 enhancements
- **🎨 Font Control**: Added configurable font sizing with 14pt default for optimal readability
- **📊 Complete Integration**: Full integration of all multiple report type features
- **📋 Documentation**: Complete specification with all features documented
- **🔧 Configuration**: Enhanced configuration system with font and display templates

### v1.1.0 (2025-08-06) - Multiple Report Types Enhancement  
- **🆕 Major Feature**: Added comprehensive multiple report type display support
- **🆕 Enhancement**: Intelligent categorization and grouping of report types
- **🆕 Feature**: Configurable display modes (all types, best only, categorized)
- **🆕 Addition**: Enhanced analytics with type combination analysis
- **🆕 Improvement**: Flexible formatting with customizable separators
- **🆕 Feature**: Smart truncation for cases with many report types

### v1.0.0 (2025-08-05) - Initial Release
- Initial specification with comprehensive foundation
- Stock data loading and PDF discovery
- Basic Google Sheets integration and CSV export
- Error handling and logging framework
- Single report type per cell display

---

*Document Version: 1.1.1*  
*Last Updated: 2025-08-08*  
*Next Review: 2025-11-08*

## 🔗 Enhanced Integration Benefits

### **Complete Multiple Report Type Workflow**:
```
MOPS Downloader → Downloads multiple PDFs per quarter (A12, A13, AI1, etc.)
         ↓
MOPS Sheets Uploader (v1.1.1) → Creates comprehensive matrix with 14pt font
         ↓
Enhanced Google Sheets Dashboard → Complete view: "A12/A13/AI1" with professional styling
         ↓
FactSet Pipeline → Analyzes all available report types for comprehensive insights
```

### **Reuses Your Proven Google Sheets Setup**:
- ✅ **Same Service Account**: Uses your existing service account credentials
- ✅ **Same Spreadsheet**: Creates new worksheet in your current sheet (ID: 1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0)
- ✅ **Same Error Handling**: Inherits CSV fallback, retry logic, API limit protection
- ✅ **Same Authentication**: No additional setup required

### **Enhanced with Professional Features**:
- 📊 **Multiple Type Matrix**: Shows all available report types per quarter
- 🎨 **Professional Styling**: 14pt font with bold headers for optimal readability
- ⚙️ **Flexible Configuration**: Customizable display modes and font settings
- 📈 **Enhanced Analytics**: Detailed type combination and diversity analysis
- 🔄 **Dynamic Adaptation**: Handles company list changes and new report types

**Worksheet Structure in Your Spreadsheet**:
```
Your Google Spreadsheet (1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0)
└── 🆕 MOPS下載狀態 (enhanced MOPS matrix worksheet with 14pt professional styling)
```

This enhanced system provides **complete transparency and professional presentation** into report availability, showing exactly which report types are available for each company and quarter with optimal 14pt font readability, enabling better decision-making for financial analysis workflows.