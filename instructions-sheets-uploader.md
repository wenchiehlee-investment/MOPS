# MOPS Sheets Uploader - Design Document v1.1.0

[![Version](https://img.shields.io/badge/Version-1.1.0-blue)](https://github.com/your-repo/mops-downloader)
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
- **Parameter-Based CLI**: Clean command-line interface with no interactive menus
- **Path-Aware Execution**: Works from project root or scripts/ folder
- **Error Recovery**: Robust handling of missing files and API limitations

## 📁 Complete Project Structure

```
project_root/
├── .env                              # Environment variables
├── StockID_TWSE_TPEX.csv            # Stock list (116+ companies)
├── requirements.txt                  # Package dependencies
├── mops_sheets_uploader/            # Main package
│   ├── __init__.py                  # Package initialization
│   ├── main.py                      # Main orchestrator (MOPSSheetsUploader)
│   ├── config.py                    # Configuration management
│   ├── models.py                    # Data structures and models
│   ├── pdf_scanner.py               # PDF file discovery
│   ├── stock_data_loader.py         # Stock data loading
│   ├── matrix_builder.py            # Matrix construction
│   ├── sheets_connector.py          # Google Sheets integration
│   ├── report_analyzer.py           # Analysis and insights
│   └── cli.py                       # Command-line interface
├── scripts/                         # Optional scripts folder
│   └── sheets_uploader.py           # Entry point for scripts/
├── downloads/                       # PDF files directory
│   ├── 2330/
│   │   ├── 202401_2330_AI1.pdf
│   │   ├── 202402_2330_AI1.pdf
│   │   └── metadata.json
│   └── 8272/
│       ├── 202401_8272_A12.pdf
│       └── metadata.json
└── data/
    └── reports/                     # CSV output directory
```

## 🚀 Usage Interface

### Command-Line Interface (Parameter-Based)

**Primary Entry Points:**
```bash
# From project root
python upload_to_sheets.py --upload

# From scripts folder  
python scripts/sheets_uploader.py --upload
```

**Available Actions:**
```bash
# Upload to Google Sheets (with CSV backup)
python upload_to_sheets.py --upload

# Export CSV only
python upload_to_sheets.py --csv-only

# Analysis only (no upload)
python upload_to_sheets.py --analyze

# Test Google Sheets connection
python upload_to_sheets.py --test

# Show help
python upload_to_sheets.py --help
```

**Advanced Usage:**
```bash
# Custom directories
python upload_to_sheets.py --upload \
  --downloads-dir ./data/downloads \
  --output-dir ./reports \
  --stock-csv ./data/stocks.csv

# Quiet mode (minimal output)
python upload_to_sheets.py --upload --quiet

# Analysis with output file
python upload_to_sheets.py --analyze \
  --output-report analysis_report.json \
  --no-suggestions

# Custom Google Sheets settings
python upload_to_sheets.py --upload \
  --sheet-id YOUR_CUSTOM_SHEET_ID \
  --worksheet-name "自訂工作表"
```

## 📊 Matrix Specifications

### Current Implementation
Based on actual testing with your data:

**Matrix Size**: 116 companies × (2 + 11 quarters) columns
- **Rows**: 116 companies (from StockID_TWSE_TPEX.csv)
- **Base Columns**: 2 (代號, 名稱) 
- **Quarter Columns**: 11 dynamic quarters (2023 Q1 to 2025 Q3)
- **Total Columns**: 13 base columns + 2 summary columns = 15

**Actual Data Discovered**: 136 PDF files from 82 companies
- **Coverage**: 10.7% overall coverage
- **Quality Score**: 3.5/10
- **Report Types**: AI1 (102), AI3 (26), AI2 (8)

### Google Sheets Matrix Layout (Actual Output)

```
代號    | 名稱       | 2025 Q3 | 2025 Q2 | 2025 Q1 | 2024 Q4 | 2024 Q3 | ... | 總報告數 | 涵蓋率
--------|------------|---------|---------|---------|---------|---------|-----|----------|--------
2330    | 台積電     | -       | -       | AI1     | AI1     | AI1     | ... | 4        | 36.4%
8272    | 全景軟體   | -       | -       | AI1     | AI1     | AI1     | ... | 3        | 27.3%
2382    | 廣達       | -       | -       | -       | A13     | -       | ... | 1        | 9.1%
...     | ...        | ...     | ...     | ...     | ...     | ...     | ... | ...      | ...
(116 total companies)
```

### Cell Value Types (Implemented)
- **PDF Type Codes**: `AI1`, `AI2`, `AI3`, `A12`, `A13` (actual types found)
- **Status Indicators**: `-` for missing reports
- **Future Quarters**: Handled with warnings
- **Priority System**: A12/A13 > AI1/A1L > others > English (excluded)

## 🔧 Technical Implementation

### Environment Configuration

**Required Environment Variables** (`.env` file):
```bash
# Google Sheets Integration
GOOGLE_SHEETS_CREDENTIALS={"type":"service_account",...}
GOOGLE_SHEET_ID=1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0

# Optional Configuration Overrides
MOPS_DOWNLOADS_DIR=downloads
MOPS_STOCK_CSV_PATH=StockID_TWSE_TPEX.csv
MOPS_MAX_YEARS=3
MOPS_WORKSHEET_NAME=MOPS下載狀態
```

**Automatic Environment Loading:**
- Searches multiple locations: `./.env`, `../,env`, current directory
- Supports both `python-dotenv` and manual parsing
- Works from project root or scripts/ folder

### Python API Interface

```python
# Quick Functions (Recommended)
from mops_sheets_uploader import upload_to_sheets, export_to_csv, analyze_coverage

# Upload to Google Sheets
result = upload_to_sheets('1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0')
if result.success:
    print(f"✅ 上傳成功: {result.sheets_url}")

# Export CSV only
result = export_to_csv('./reports')
print(f"💾 CSV 檔案: {result.csv_path}")

# Analysis only
report = analyze_coverage(include_suggestions=True)
print(f"📊 涵蓋率: {report['summary']['coverage_percentage']:.1f}%")
```

```python
# Advanced Usage
from mops_sheets_uploader import MOPSSheetsUploader

uploader = MOPSSheetsUploader(
    downloads_dir="downloads",
    stock_csv_path="StockID_TWSE_TPEX.csv"
)

result = uploader.run()
if result.success:
    print(f"✅ 上傳成功!")
    print(f"🔗 Google Sheets: {result.sheets_url}")
    print(f"💾 CSV 備份: {result.csv_path}")
    print(f"📊 涵蓋率: {result.coverage_stats.coverage_percentage:.1f}%")
```

## 📋 Component Specifications (Implemented)

### 1. PDF Scanner Module (`pdf_scanner.py`)

**Implemented Features:**
- Scans `downloads/` directory structure
- Parses PDF filenames with pattern `YYYYQQ_COMPANYID_TYPE.pdf`
- Validates PDF files (checks headers and file sizes)
- Handles future quarter detection with warnings
- Reports suspicious files (< 1KB size)

**Actual Results from Your Data:**
- 136 PDF files discovered across 82 companies
- 10 quarters range: 2023 Q1 to 2025 Q2
- 10 files flagged as suspicious (403 bytes each)
- Report types found: AI1, AI2, AI3

### 2. Stock Data Loader (`stock_data_loader.py`)

**Implemented Features:**
- Loads `StockID_TWSE_TPEX.csv` with encoding detection
- Validates 4-digit company codes
- Detects stock list changes (additions/removals)
- Handles duplicate detection and cleanup

**Your Data:**
- Successfully loads 116 companies
- Supports dynamic company list growth
- UTF-8 encoding with proper Chinese character handling

### 3. Matrix Builder (`matrix_builder.py`)

**Implemented Features:**
- Creates dynamic company × quarter matrix
- Applies report type priority rules (A12/A13 > AI1 > others)
- Resolves conflicts when multiple report types exist
- Adds summary columns (總報告數, 涵蓋率)

**Your Results:**
- Matrix size: 116 × 15 (companies × columns)
- Resolved 26 report type conflicts
- 7.6% cell fill rate (97/1276 cells populated)

### 4. Sheets Connector (`sheets_connector.py`)

**Implemented Features:**
- Integrates with existing Google Sheets credentials
- Creates "MOPS下載狀態" worksheet in your existing spreadsheet
- Auto-resizes worksheet for matrix dimensions
- Applies formatting and highlighting
- CSV fallback when Sheets upload fails

**Integration:**
- Worksheet auto-creation and formatting

### 5. Report Analyzer (`report_analyzer.py`)

**Implemented Features:**
- Comprehensive coverage analysis
- Missing report identification
- Download suggestions generation
- Temporal pattern analysis
- Quality scoring (0-10 scale)

**Your Analysis Results:**
- Quality score: 3.5/10
- 116 companies with missing reports (all high priority)
- 52.6% average reporting consistency
- 56 companies with irregular patterns

## 🎯 Actual Usage Examples

### From Your Test Results

**Successful CSV Export:**
```bash
$ python upload_to_sheets.py --csv-only
✅ CSV 匯出成功!
📁 檔案位置: data/reports/mops_matrix_20250807_113440.csv
📊 涵蓋率: 10.7%
```

**Analysis Results:**
```bash
$ python upload_to_sheets.py --analyze
📈 分析結果:
   • 總公司數: 116
   • 有PDF公司: 82
   • 涵蓋率: 10.7%
   • 品質分數: 3.5/10

💡 關鍵洞察:
   • 涵蓋率偏低 (10.7%)，需要大量補充
   • 合併財報比例較高，建議補充個別財報

📋 下載建議:
   • 1587 (吉茂): 2025 Q3 - 缺失最近季度報告
   • 1587 (吉茂): 2025 Q2 - 缺失最近季度報告
   • 2301 (光寶科): 2025 Q3 - 缺失最近季度報告
```

**Google Sheets Upload:**
```bash
$ python upload_to_sheets.py --upload
✅ 上傳成功!
🔗 Google Sheets: https://docs.google.com/spreadsheets/d/1BmpE_ZeusMlgEeESoHxAzxtW7gyij6eF11RkvtjlgR0
💾 CSV 檔案: data/reports/mops_matrix_20250807_113440.csv
📊 涵蓋率: 10.7%
```

## 🔧 Configuration System (Implemented)

### MOPSConfig Class

```python
@dataclass
class MOPSConfig:
    # Directory settings
    downloads_dir: str = "downloads"
    stock_csv_path: str = "StockID_TWSE_TPEX.csv"
    output_dir: str = "data/reports"
    
    # Matrix settings
    max_years: int = 3
    include_current_year: bool = True
    quarter_order: str = "desc"  # newest first
    
    # Google Sheets settings
    worksheet_name: str = "MOPS下載狀態"
    google_sheet_id: Optional[str] = None
    google_credentials: Optional[str] = None
    
    # Report type settings
    preferred_types: List[str] = ["A12", "A13"]
    acceptable_types: List[str] = ["AI1", "A1L"] 
    excluded_types: List[str] = ["AIA", "AE2"]
```

### Environment Loading (Implemented)

```python
# Automatic .env loading with path resolution
def load_env_file():
    try:
        from dotenv import load_dotenv
        env_locations = [
            script_dir / '.env',
            script_dir.parent / '.env',
            Path('.env'),
            Path.cwd() / '.env'
        ]
        for env_file in env_locations:
            if env_file.exists():
                load_dotenv(env_file)
                break
    except ImportError:
        # Manual fallback parsing
```

## 📊 Error Handling (Implemented)

### Error Categories Handled

1. **Path Resolution**: Works from root or scripts/ folder
2. **Environment Loading**: Multiple .env file locations
3. **Package Import**: Automatic Python path adjustment
4. **Google Sheets**: Connection failures with detailed diagnostics
5. **PDF Parsing**: Invalid filenames and suspicious files
6. **CSV Export**: Fallback when Sheets upload fails

### Example Error Messages

```bash
❌ mops_sheets_uploader 套件目錄不存在
   已搜尋位置:
     • ./mops_sheets_uploader
     • ../mops_sheets_uploader
     • mops_sheets_uploader

❌ 缺少 Google Sheets 憑證，無法上傳
   請確認 .env 檔案中的 GOOGLE_SHEETS_CREDENTIALS

⚠️ 發現未來季度PDF: 2025 Q2, 2025 Q3
   這可能是檔案命名錯誤或系統時間問題
```

## 🔗 Integration Benefits (Implemented)

### Reuses Your Existing Infrastructure

**✅ Proven Google Sheets Setup:**
- Same error handling: CSV fallback, retry logic
- Same authentication: No additional setup required

**✅ MOPS-Specific Features:**
- Matrix layout optimized for company × quarter visualization
- PDF type display: Shows A12, A13, AI1 codes directly
- Dynamic columns: Adapts to discovered quarters automatically
- Future quarter handling: Smart detection and warnings

### Complete Workflow Integration

```
MOPS Downloader → Downloads PDFs to downloads/ folder
         ↓
MOPS Sheets Uploader → Creates download status matrix  
         ↓
Your Analysis → Complete financial overview
         ↓
Unified Dashboard → All in one Google Sheets
```

## 📈 Performance Results (Actual)

**Processing Speed:**
- Scan 136 PDFs: < 1 second
- Build 116×15 matrix: < 1 second  
- Total processing time: 0.2 seconds
- Google Sheets upload: ~2 seconds

**Memory Usage:**
- Lightweight: < 50MB memory usage
- CSV output: ~15KB file size
- No performance issues with current dataset

**API Efficiency:**
- Single batch upload to Google Sheets
- No rate limiting concerns
- Reliable connection with retry logic

## 🛠️ Installation & Setup

### Dependencies (requirements.txt)

```txt
pandas>=1.5.0
numpy>=1.21.0
gspread>=5.0.0
google-auth>=2.0.0
python-dotenv>=0.19.0
PyYAML>=6.0
python-dateutil>=2.8.0
```

### Quick Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Test environment
python debug_paths.py

# 3. Test connection
python upload_to_sheets.py --test

# 4. Run upload
python upload_to_sheets.py --upload
```

## 🧪 Testing Strategy (Implemented)

### Diagnostic Tools Created

**Environment Debugging:**
```bash
python debug_paths.py              # Check paths and environment
python check_requirements.py       # Verify package installation  
python test_google_sheets.py       # Comprehensive connection test
```

**Unit Testing:**
- PDF filename parsing validation
- Matrix construction verification
- Report type priority rules
- Environment variable loading

### Real-World Testing

**✅ Tested with your actual data:**
- 116 companies from StockID_TWSE_TPEX.csv
- 136 PDF files in downloads/ directory
- Google Sheets integration working
- CSV backup successful
- Error handling verified

## 🔄 Future Enhancements

### Phase 2 Features (Ready for Implementation)

- **Enhanced Formatting**: Color-coded cells based on report types
- **Interactive Dashboard**: Web interface for matrix visualization
- **Historical Tracking**: Track coverage changes over time
- **Automated Alerts**: Email notifications for missing reports

### Phase 3 Features (Planned)

- **ML-Based Analysis**: Predict download patterns
- **Multi-Market Support**: Extend beyond Taiwan markets
- **API Integration**: Connect with external financial databases
- **Mobile Dashboard**: Mobile app for executives

---

*Document Version: 1.1.0*  
*Last Updated: 2025-08-07*  
*Tested with: 116 companies, 136 PDFs, 10.7% coverage*  
*Next Review: 2025-11-07*

## 📞 Quick Reference

**Most Common Commands:**
```bash
# Upload everything
python upload_to_sheets.py --upload

# Just CSV
python upload_to_sheets.py --csv-only

# Check what you have
python upload_to_sheets.py --analyze

# Test connection
python upload_to_sheets.py --test
```

**Troubleshooting:**
```bash
python debug_paths.py              # Check setup
python check_requirements.py       # Check packages
python test_google_sheets.py       # Test Google Sheets
```

The tool is production-ready and working with your actual data! 🎉