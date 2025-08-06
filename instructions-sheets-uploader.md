# MOPS Sheets Uploader - Design Document v1.0.0

[![Version](https://img.shields.io/badge/Version-1.0.0-blue)](https://github.com/your-repo/mops-downloader)
[![Python](https://img.shields.io/badge/Python-3.9%2B-green)](https://python.org)
[![Integration](https://img.shields.io/badge/Integration-MOPS_Downloader-orange)](https://github.com/your-repo/mops-downloader)

## 📋 Project Overview

The MOPS Sheets Uploader is a Python-based companion tool for the MOPS Downloader System that automatically scans downloaded PDF files and creates a comprehensive matrix view in Google Sheets. The tool provides a clear overview of financial report availability across companies and quarters, making it easy to track download progress and identify missing reports.

## 🎯 Goals and Objectives

### Primary Goal
Create a Google Sheets matrix showing PDF file availability with companies as rows and year-quarters as columns, providing instant visibility into download status and report coverage.

### Key Features
- **Automated PDF Discovery**: Scan `downloads/` folder structure to identify available reports
- **Matrix Visualization**: Create company × quarter matrix in Google Sheets
- **Stock Information Integration**: Load company details from `StockID_TWSE_TPEX.csv`
- **Report Type Classification**: Identify and display PDF report types (A12, A13, AI1, etc.)
- **Progress Tracking**: Visual indicators for download completeness
- **Flexible Output**: Support both Google Sheets upload and CSV backup
- **Error Recovery**: Robust handling of missing files and API limitations

## 📊 Matrix Specifications

### Concrete Dimensions
Based on the provided StockID_TWSE_TPEX.csv:

**Matrix Size**: 116 companies × (2 + quarters) columns
- **Rows**: Exactly 116 companies (from CSV file)
- **Base Columns**: 2 (代號, 名稱) 
- **Quarter Columns**: Configurable (default: 12 quarters = 3 years × 4 quarters)
- **Total Columns**: 14 (2 base + 12 quarters) for default 3-year view

**Memory Requirements**: 
- Estimated CSV size: ~50KB for full matrix with text data
- Google Sheets: Well within limits (116 rows × 14 columns = 1,624 cells)
- Processing: Lightweight - can handle in memory easily

**API Efficiency**:
- Google Sheets batch upload: Single API call for entire matrix
- Processing time: < 5 seconds for full scan and upload
- Rate limits: No concern with this dataset size

### Google Sheets Matrix Layout

**Scenario 1: Normal Case (Current Quarter = 2025 Q1)**
```
代號    | 名稱       | 2025 Q1 | 2024 Q4 | 2024 Q3 | 2024 Q2 | 2024 Q1 | 2023 Q4 | ...
--------|------------|---------|---------|---------|---------|---------|---------|----
2330    | 台積電     | AI1     | AI1     | AI1     | AI1     | AI1     | AI1     | ...
8272    | 全景軟體   | A12     | A12     | A12     | A12     | -       | -       | ...
2382    | 廣達       | -       | -       | A13     | -       | -       | -       | ...
```

**Scenario 2: With Future Quarter PDFs (2025 Q2, Q3, Q4 exist)**
```
代號    | 名稱       | 2025 Q4 | 2025 Q3 | 2025 Q2 | 2025 Q1 | 2024 Q4 | 2024 Q3 | ...
--------|------------|---------|---------|---------|---------|---------|---------|----
2330    | 台積電     | AI1     | AI1     | AI1     | AI1     | AI1     | AI1     | ...
8272    | 全景軟體   | A12     | A12     | A12     | A12     | A12     | A12     | ...
2382    | 廣達       | ⚠️      | -       | -       | -       | A13     | -       | ...
```

**Future Quarter Indicators:**
- ⚠️ = PDF exists but quarter is in future (possibly test data or mislabeled)
- Regular codes (AI1, A12) = Valid PDFs regardless of quarter timing

### Cell Value Types
- **PDF Type Codes**: `A12`, `A13`, `AI1`, `A1L` (actual PDF type found)
- **Status Indicators**: 
  - `✅` - Individual report available (A12/A13)
  - `🟡` - Consolidated report available (AI1/A1L)
  - `⚠️` - Multiple report types available
  - `❌` - No reports found
  - `-` - Not attempted/out of scope

## 🏗️ System Architecture

### Component Structure
```
mops_sheets_uploader/
├── __init__.py
├── pdf_scanner.py          # PDF file discovery and analysis
├── matrix_builder.py       # Matrix data structure creation
├── sheets_connector.py     # Google Sheets integration (based on existing)
├── stock_data_loader.py    # StockID_TWSE_TPEX.csv handling
└── report_analyzer.py      # PDF filename parsing and classification
```

### Data Flow Architecture
```
Stage 1: Stock Data Loading
    ↓
Stage 2: PDF File Discovery & Analysis
    ↓
Stage 3: Matrix Data Construction
    ↓
Stage 4: Google Sheets Upload / CSV Export
    ↓
Stage 5: Status Reporting & Validation
```

## 📁 File Structure Analysis

### Expected Directory Structure
```
project_root/
├── StockID_TWSE_TPEX.csv          # 116 companies, 2 columns (代號, 名稱)
├── downloads/
│   ├── 2330/
│   │   ├── 202401_2330_AI1.pdf
│   │   ├── 202402_2330_AI1.pdf
│   │   ├── 202403_2330_AI1.pdf
│   │   └── metadata.json
│   ├── 8272/
│   │   ├── 202401_8272_A12.pdf
│   │   ├── 202402_8272_A12.pdf
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
```

## 🔧 Technical Specifications

### Class Interface Design

```python
class MOPSSheetsUploader:
    """MOPS PDF files to Google Sheets matrix uploader"""
    
    def __init__(self, 
                 downloads_dir: str = "downloads",
                 stock_csv_path: str = "StockID_TWSE_TPEX.csv",
                 sheet_id: str = None,
                 max_years: int = 3,
                 csv_backup: bool = True):
        """Initialize uploader with configuration"""
        
    def scan_pdf_files(self) -> Dict[str, Dict[str, str]]:
        """Scan downloads directory and return PDF file matrix"""
        
    def load_stock_data(self) -> pd.DataFrame:
        """Load and validate StockID_TWSE_TPEX.csv"""
        
    def build_matrix(self) -> pd.DataFrame:
        """Build company × quarter matrix with PDF status"""
        
    def upload_to_sheets(self, matrix_df: pd.DataFrame) -> bool:
        """Upload matrix to Google Sheets"""
        
    def export_to_csv(self, matrix_df: pd.DataFrame) -> str:
        """Export matrix to CSV as backup"""
        
    def generate_report(self) -> Dict[str, Any]:
        """Generate comprehensive status report"""
        
    def run(self) -> bool:
        """Main execution method"""
```

### PDF Classification System

```python
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
```

### Quarter Column Generation
```python
def generate_quarter_columns(pdf_data: Dict[str, List[PDFFile]], max_years: int = 3) -> List[str]:
    """
    Generate quarter column headers based on discovered PDFs + expected quarters
    
    Strategy:
    1. Scan all discovered PDFs to find actual quarters
    2. Add expected quarters within max_years range
    3. Sort in reverse chronological order (newest first)
    4. Handle future quarters gracefully
    """
    discovered_quarters = set()
    current_year = datetime.now().year
    current_quarter = (datetime.now().month - 1) // 3 + 1
    
    # Scan discovered PDFs for actual quarters
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

# Example outputs based on different scenarios:
# Scenario 1 (Normal): ['2025 Q1', '2024 Q4', '2024 Q3', '2024 Q2', ...]
# Scenario 2 (Future PDFs): ['2025 Q4', '2025 Q3', '2025 Q2', '2025 Q1', '2024 Q4', ...]
```

## 📋 Component Specifications

### 1. PDF Scanner Module (`pdf_scanner.py`)

**Purpose**: Discover and catalog all PDF files in the downloads directory

```python
class PDFScanner:
    def scan_downloads_directory(self, downloads_dir: str) -> Dict[str, List[PDFFile]]:
        """
        Scan downloads directory structure and return organized PDF files
        
        Handles:
        - Current and past quarters (expected)
        - Future quarters (with warnings)
        - Invalid filenames (logged and skipped)
        - Multiple report types per quarter
        
        Returns:
            Dict mapping company_id to list of PDFFile objects
        """
        
    def discover_available_quarters(self, pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """
        Discover all quarters that have at least one PDF file
        
        Returns:
            Sorted list of quarters (e.g., ['2025 Q4', '2025 Q3', '2025 Q1', '2024 Q4'])
        """
        
    def parse_pdf_filename(self, filename: str) -> Optional[PDFMetadata]:
        """
        Parse PDF filename to extract metadata with validation
        
        Args:
            filename: PDF filename (e.g., "202501_2330_AI1.pdf" for 2025 Q1)
            
        Returns:
            PDFMetadata object or None if parsing fails
            
        Validation:
        - Year range: 2020-2030 (prevents obvious errors)
        - Quarter range: 1-4 
        - Company ID: 4-digit format
        - Warns about future quarters beyond current date
        """
        
    def validate_pdf_file(self, file_path: str) -> bool:
        """Validate that file exists and is a valid PDF"""
```

### 2. Stock Data Loader (`stock_data_loader.py`)

**Purpose**: Load and process company information from StockID_TWSE_TPEX.csv (116 companies)

```python
class StockDataLoader:
    def load_stock_csv(self, csv_path: str) -> pd.DataFrame:
        """
        Load StockID_TWSE_TPEX.csv with proper encoding and validation
        
        Expected format:
        - Rows: 116 companies (plus header)
        - Columns: 代號 (Integer), 名稱 (String)
        - Encoding: UTF-8
        """
        
    def validate_stock_data(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Validate stock data completeness and format
        Expected: exactly 116 companies with valid stock codes
        """
        
    def get_company_name(self, stock_id: str) -> str:
        """Get company name by stock ID from the 116 company list"""
```

### 3. Matrix Builder (`matrix_builder.py`)

**Purpose**: Construct the main data matrix combining stock data and PDF information

```python
class MatrixBuilder:
    def __init__(self, stock_df: pd.DataFrame, pdf_data: Dict[str, List[PDFFile]]):
        """
        Initialize with stock data and PDF scan results
        
        Args:
            stock_df: DataFrame with 116 companies (代號, 名稱)
            pdf_data: Dictionary mapping company_id to PDFFile list
        """
        
    def build_base_matrix(self, max_years: int = 3) -> pd.DataFrame:
        """
        Create base matrix with companies and dynamically discovered quarter columns
        
        Process:
        1. Start with 116 companies from StockID_TWSE_TPEX.csv
        2. Discover quarters from actual PDF files
        3. Add expected quarters within max_years range  
        4. Create matrix with dynamic column count
        
        Returns:
            DataFrame with shape (116, 2 + dynamic_quarters) where:
            - 116 rows: one per company in StockID_TWSE_TPEX.csv
            - Columns: 代號, 名稱, [discovered quarters in reverse chronological order]
            
        Examples:
            Normal case: (116, 14) → 12 quarters + 2 base columns
            With future PDFs: (116, 18) → 16 quarters + 2 base columns
        """
        
    def populate_pdf_status(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """Fill matrix with PDF availability status for all 116 companies"""
        
    def apply_priority_rules(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """Apply report type priority rules for conflicting files"""
        
    def add_summary_columns(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """Add summary columns like total reports, coverage percentage"""
```

### 4. Sheets Connector (`sheets_connector.py`)

**Purpose**: Handle Google Sheets integration (extends existing sheets_uploader.py)

```python
from sheets_uploader import SheetsUploader

class MOPSSheetsConnector(SheetsUploader):
    """
    Extended sheets connector for MOPS matrix data
    
    Inherits from your existing SheetsUploader class to reuse:
    - Google Sheets authentication (GOOGLE_SHEETS_CREDENTIALS)
    - Connection management (_setup_connection)
    - Error handling and retry logic
    - CSV fallback functionality
    
    Target Sheet: ID 1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0
    Target Worksheet: "MOPS下載狀態" (will be created if doesn't exist)
    """
    
    def __init__(self, worksheet_name: str = "MOPS下載狀態"):
        """Initialize using existing Google Sheets configuration"""
        super().__init__()  # Inherits your existing setup
        self.mops_worksheet_name = worksheet_name
        
    def upload_matrix(self, matrix_df: pd.DataFrame) -> bool:
        """
        Upload matrix data to Google Sheets with formatting
        
        Process:
        1. Use inherited _setup_connection() method
        2. Create or find "MOPS下載狀態" worksheet  
        3. Upload matrix data with proper formatting
        4. Apply MOPS-specific styling (future quarter highlighting)
        5. Fallback to CSV if API limits hit (inherited behavior)
        """
        
    def format_matrix_worksheet(self, worksheet, data_rows: int, 
                               future_quarters: List[str] = None) -> None:
        """
        Apply MOPS-specific formatting for matrix display
        
        Features:
        - Header row styling (company info vs quarters)
        - Future quarter column highlighting (orange background)
        - Status symbol color coding (✅🟡⚠️❌)
        - Auto-resize columns for readability
        - Freeze first two columns (代號, 名稱)
        """
        
    def create_summary_statistics(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """
        Create summary statistics for separate worksheet upload
        
        Statistics:
        - Total companies (116)
        - Companies with PDFs vs without
        - Coverage by quarter
        - Report type distribution
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

### 5. Report Analyzer (`report_analyzer.py`)

**Purpose**: Analyze PDF files and generate insights

```python
class ReportAnalyzer:
    def analyze_coverage(self, matrix_df: pd.DataFrame) -> Dict[str, Any]:
        """Analyze report coverage statistics"""
        
    def identify_missing_reports(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """Generate list of missing reports by priority"""
        
    def generate_download_suggestions(self, matrix_df: pd.DataFrame) -> List[str]:
        """Suggest which companies/quarters to download next"""
```

## 📊 Data Models

### PDFFile Data Class
```python
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

@dataclass
class MatrixCell:
    """Represents a single cell in the matrix"""
    has_pdf: bool
    report_type: Optional[str]
    status_symbol: str
    file_count: int
    file_paths: List[str]
```

## 🔧 Configuration System

### Configuration File (`config.yaml`)
```yaml
# MOPS Sheets Uploader Configuration

# Directory settings
downloads_dir: "downloads"
stock_csv_path: "StockID_TWSE_TPEX.csv"
output_dir: "data/reports"

# Matrix settings
max_years: 3
include_current_year: true
quarter_order: "desc"  # newest first
dynamic_columns: true  # adjust columns based on discovered PDFs

# 🆕 Future quarter handling
future_quarter_handling:
  include_in_matrix: true      # Show future quarters in output
  warn_threshold_months: 6     # Warn if more than 6 months in future
  mark_suspicious: true        # Mark with ⚠️ indicator
  max_future_quarters: 8       # Limit how far future to include
  generate_future_report: true # Create separate report of future PDFs

# Google Sheets settings
worksheet_name: "MOPS下載狀態"
include_summary_sheet: true
auto_resize_columns: true
highlight_future_quarters: true  # 🆕 Color-code future quarter cells

# CSV export settings
csv_backup: true
csv_filename_pattern: "mops_matrix_{timestamp}.csv"
include_future_analysis: true   # 🆕 Add future quarter analysis to CSV

# Report type settings
preferred_types: ["A12", "A13"]
acceptable_types: ["AI1", "A1L"]
excluded_types: ["AIA", "AE2"]

# Display settings
use_symbols: true
show_file_counts: false
highlight_missing: true
future_quarter_symbol: "⚠️"    # 🆕 Symbol for future quarters
```

### Environment Variables
```bash
# Google Sheets Integration (from your updated .env)
GOOGLE_SHEETS_CREDENTIALS={"type":"service_account","project_id":"coherent-vision-463514-n3",...}
GOOGLE_SHEET_ID=1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0

# Optional: Custom configuration
MOPS_CONFIG_PATH=path/to/config.yaml
MOPS_MAX_YEARS=3
MOPS_DOWNLOADS_DIR=downloads
MOPS_WORKSHEET_NAME=MOPS下載狀態  # Will create this worksheet in your existing sheet
```

**Integration Note**: The MOPS Sheets Uploader will use your existing Google Sheets service account and spreadsheet, creating a new worksheet called "MOPS下載狀態" within your current spreadsheet (ID: 1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0).

## 🚦 Error Handling Strategy

### Error Categories
1. **File System Errors**: Missing directories, permission issues, corrupted files
2. **Data Validation Errors**: Invalid stock CSV format, malformed PDF filenames
3. **Google Sheets Errors**: API limits, authentication failures, worksheet conflicts
4. **Processing Errors**: Memory issues with large datasets, encoding problems
5. **🆕 Temporal Anomalies**: Future quarter PDFs, invalid date ranges, system clock issues

### Recovery Mechanisms
```python
class ErrorHandler:
    def handle_missing_stock_csv(self) -> pd.DataFrame:
        """Generate default stock list from discovered companies"""
        
    def handle_sheets_api_failure(self, matrix_df: pd.DataFrame) -> str:
        """Fallback to CSV export when Sheets API fails"""
        
    def handle_parsing_errors(self, invalid_files: List[str]) -> None:
        """Log and skip files that don't match expected patterns"""
        
    def handle_future_quarters(self, future_pdfs: List[PDFFile]) -> Dict[str, Any]:
        """
        Handle PDFs with future quarter dates
        
        Strategy:
        1. Include in matrix but mark with warning indicator
        2. Log detailed warning with file paths
        3. Suggest verification steps for user
        4. Generate separate report of suspicious files
        
        Returns:
            Dict with future quarter analysis and recommendations
        """
        
    def validate_quarter_range(self, year: int, quarter: int) -> bool:
        """
        Validate if quarter/year combination is reasonable
        
        Validation rules:
        - Year: 2020-2030 (prevents obvious errors)
        - Quarter: 1-4
        - Not more than 2 quarters in future from current date
        """
```

## 📈 Usage Examples

### Basic Usage
```python
from mops_sheets_uploader import MOPSSheetsUploader

# Initialize uploader (uses .env configuration automatically)
uploader = MOPSSheetsUploader(
    downloads_dir="downloads",
    stock_csv_path="StockID_TWSE_TPEX.csv",
    max_years=3
)

# Run complete process
success = uploader.run()

if success:
    print("✅ Matrix uploaded to Google Sheets successfully")
    print("🔗 View at: https://docs.google.com/spreadsheets/d/1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0")
    print("📊 Worksheet: MOPS下載狀態")
else:
    print("⚠️ Upload failed, but CSV backup is available")
```

### Advanced Configuration
```python
# Custom configuration with specific options
uploader = MOPSSheetsUploader(
    downloads_dir="./data/downloads",
    stock_csv_path="./data/StockID_TWSE_TPEX.csv",
    sheet_id="1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0",  # Your specific sheet
    max_years=5,
    csv_backup=True,
    worksheet_name="MOPS財報矩陣"  # Custom worksheet name
)

# Test connection using existing credentials
if uploader.test_connection():
    print("✅ Google Sheets 連線成功")
    
    # Get detailed analysis
    matrix_df = uploader.build_matrix()
    coverage_stats = uploader.generate_report()

    print(f"Total companies: {coverage_stats['total_companies']}")  # Expected: 116
    print(f"Coverage rate: {coverage_stats['coverage_percentage']:.1f}%")
    print(f"Missing reports: {coverage_stats['missing_count']}")

    # Export for further analysis
    csv_path = uploader.export_to_csv(matrix_df)
    print(f"Matrix exported to: {csv_path}")
else:
    print("❌ Google Sheets 連線失敗，請檢查 .env 設定")
```

### CLI Interface
```bash
# Basic usage
python mops_sheets_uploader.py

# With custom parameters
python mops_sheets_uploader.py \
    --downloads-dir ./data/downloads \
    --stock-csv ./data/stocks.csv \
    --max-years 5 \
    --csv-only

# Generate report only (no upload)
python mops_sheets_uploader.py --report-only --output report.json
```

## 📊 Output Examples

### Console Output

**Scenario 1: Normal Case**
```
📊 MOPS Sheets Uploader v1.0.0
=======================================

🔍 掃描 PDF 檔案...
   發現 68 個 PDF 檔案 (32 家公司)
   季度範圍: 2024 Q1 到 2025 Q1 (5 個季度)
   
📋 載入股票清單...
   已載入 116 家公司資料 (StockID_TWSE_TPEX.csv)
   
🏗️ 建構矩陣資料...
   矩陣大小: 116 × 7 (公司 × 季度)
   動態季度欄位: 2025 Q1, 2024 Q4, 2024 Q3, 2024 Q2, 2024 Q1
   
📈 上傳統計:
   ✅ 個別財報: 43 份 (63.2%)
   🟡 合併財報: 25 份 (36.8%)
   
🚀 上傳到 Google Sheets...
   ✅ 主矩陣上傳成功 (116 公司 × 5 季度)
```

**Scenario 2: With Future Quarter PDFs**
```
📊 MOPS Sheets Uploader v1.0.0
=======================================

🔍 掃描 PDF 檔案...
   發現 94 個 PDF 檔案 (35 家公司)
   ⚠️ 發現未來季度PDF: 2025 Q2, 2025 Q3, 2025 Q4
      這可能是檔案命名錯誤或系統時間問題
   季度範圍: 2024 Q1 到 2025 Q4 (8 個季度)
   
📋 載入股票清單...
   已載入 116 家公司資料 (StockID_TWSE_TPEX.csv)
   
🏗️ 建構矩陣資料...
   矩陣大小: 116 × 10 (公司 × 季度)  
   動態季度欄位: 2025 Q4, 2025 Q3, 2025 Q2, 2025 Q1, 2024 Q4, 2024 Q3, 2024 Q2, 2024 Q1
   
📈 上傳統計:
   ✅ 個別財報: 58 份 (61.7%)
   🟡 合併財報: 36 份 (38.3%)
   ⚠️ 未來季度報告: 26 份 (建議檢查檔案日期)
   
🚀 上傳到 Google Sheets...
   ✅ 主矩陣上傳成功 (116 公司 × 8 季度)
   📋 建議: 請檢查 2025 Q2-Q4 的PDF檔案命名是否正確
```

### CSV Output Sample
```csv
代號,名稱,2025 Q2,2025 Q1,2024 Q4,2024 Q3,2024 Q2,2024 Q1,總報告數,涵蓋率
2330,台積電,AI1,AI1,AI1,AI1,AI1,AI1,6,100.0%
8272,全景軟體,A12,A12,A12,A12,-,-,4,66.7%
2382,廣達,-,-,A13,-,-,-,1,16.7%
6462,神盾,✅,⚠️,❌,✅,-,-,2,50.0%
...,(112 more companies),...
```

## 🔒 Security Considerations

### Data Privacy
- PDF file contents are not read, only metadata is analyzed
- Company information is sourced from public CSV files
- Google Sheets credentials handled securely (existing pattern)

### Access Control
- Read-only access to downloads directory
- Configurable Google Sheets permissions
- Option to generate CSV-only output for air-gapped environments

## 🧪 Testing Strategy

### Unit Tests
```python
# Test PDF filename parsing
def test_pdf_filename_parsing():
    assert parse_filename("202401_2330_AI1.pdf") == PDFMetadata(2024, 1, "2330", "AI1")

# Test matrix construction
def test_matrix_building():
    matrix = build_test_matrix()
    assert matrix.shape == (116, 14)  # 116 companies, 12 quarters + 代號 + 名稱

# Test priority rules
def test_report_priority():
    assert get_best_report(["AI1", "A12"]) == "A12"  # Individual over consolidated
```

### Integration Tests
- End-to-end processing with sample data
- Google Sheets API integration testing
- CSV export validation

### Performance Tests
- Dataset processing with 116 companies (realistic load)
- Memory usage optimization for matrix operations
- API rate limiting compliance with Google Sheets
- Large PDF directory scanning (hundreds of files)

## 📋 Implementation Checklist

### Phase 1: Core Functionality
- [ ] PDF scanner implementation
- [ ] Stock data loader
- [ ] Basic matrix builder
- [ ] CSV export functionality
- [ ] Unit test coverage

### Phase 2: Google Sheets Integration  
- [ ] Adapt existing sheets_uploader.py
- [ ] Matrix formatting and styling
- [ ] Error handling and fallback
- [ ] Integration testing

### Phase 3: Advanced Features
- [ ] Summary statistics generation
- [ ] Missing report analysis
- [ ] CLI interface
- [ ] Configuration system

### Phase 4: Polish & Documentation
- [ ] Comprehensive error messages
- [ ] Usage examples and tutorials
- [ ] Performance optimization
- [ ] Production deployment guide

## 🔄 Future Enhancements

### Phase 2 Features
- **Interactive Dashboard**: Web-based interface for matrix visualization
- **Automated Alerts**: Email notifications for missing critical reports
- **Historical Tracking**: Track coverage trends over time
- **Batch Processing**: Handle multiple download directories

### Phase 3 Features
- **ML-Based Analysis**: Predict optimal download timing
- **Integration APIs**: Connect with financial analysis tools
- **Mobile App**: Mobile dashboard for executives
- **Multi-Market Support**: Extend beyond Taiwan markets

## 📞 Support and Maintenance

### Monitoring Points
- PDF discovery success rates
- Google Sheets API usage and limits  
- Matrix accuracy and completeness
- Processing performance metrics

### Maintenance Schedule
- Weekly: Monitor API usage and error rates
- Monthly: Review PDF pattern changes and new report types
- Quarterly: Update stock list and validate company mappings
- Annually: Architectural review and performance optimization

---

*Document Version: 1.0.0*  
*Last Updated: 2025-08-06*  
*Next Review: 2025-11-06*

## 🔗 Integration Benefits

### Leveraging Existing Infrastructure

**Reuses Your Proven Google Sheets Setup**:
- ✅ **Same Service Account**: Uses your existing `googlesearch@coherent-vision-463514-n3.iam.gserviceaccount.com`
- ✅ **Same Spreadsheet**: Creates new worksheet in your current sheet (ID: 1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0)
- ✅ **Same Error Handling**: Inherits CSV fallback, retry logic, API limit protection
- ✅ **Same Authentication**: No additional setup required

**Enhanced with MOPS-Specific Features**:
- 📊 **Matrix Layout**: Optimized for company × quarter visualization
- ⚠️ **Future Quarter Handling**: Smart detection and warning system
- 🎯 **PDF Type Display**: Shows A12, A13, AI1 codes directly in cells
- 📈 **Dynamic Columns**: Adapts to discovered quarters automatically

**Worksheet Structure in Your Spreadsheet**:
```
Your Google Spreadsheet (1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0)
├── 投資組合摘要 (existing from FactSet Pipeline)
├── 詳細報告 (existing from FactSet Pipeline) 
├── 驗證摘要 (existing from FactSet Pipeline)
├── 查詢模式摘要 (existing from FactSet Pipeline)
└── 🆕 MOPS下載狀態 (new MOPS matrix worksheet)
```

This approach provides a **unified dashboard** where you can see both FactSet analysis results and MOPS download status in the same spreadsheet.

## 🔄 Integration with MOPS Downloader & FactSet Pipeline

This tool is designed as a companion to both your MOPS Downloader System and FactSet Pipeline, creating a unified dashboard experience:

**MOPS Downloader Integration**:
1. **File Structure Compatibility**: Works with existing downloads/ directory structure
2. **Metadata Consistency**: Uses same PDF naming conventions and metadata format
3. **Error Handling**: Compatible error recovery strategies
4. **Configuration**: Shares environment variables and configuration patterns
5. **Logging**: Integrates with existing logging framework

**FactSet Pipeline Integration**:
1. **Same Google Sheets**: Uses your existing spreadsheet and credentials
2. **Complementary Data**: MOPS download status complements FactSet analysis results
3. **Unified Dashboard**: All financial data insights in one spreadsheet
4. **Shared Infrastructure**: Leverages proven sheets_uploader.py codebase
5. **Consistent UX**: Same CSV fallback and error handling patterns

**Complete Workflow**:
```
MOPS Downloader → Downloads PDFs to downloads/ folder
         ↓
MOPS Sheets Uploader → Creates download status matrix  
         ↓
FactSet Pipeline → Analyzes PDFs and generates investment insights
         ↓
Unified Google Sheets Dashboard → Complete financial analysis view
```

This combination provides a complete solution: **MOPS Downloader** for acquisition, **MOPS Sheets Uploader** for download tracking, and **FactSet Pipeline** for financial analysis - all integrated into your existing Google Sheets dashboard.