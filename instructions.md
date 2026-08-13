# MOPS Downloader System - Design Document v2.0.0

[![Version](https://img.shields.io/badge/Version-2.0.0-blue)](https://github.com/your-repo/mops-downloader)
[![Python](https://img.shields.io/badge/Python-3.9%2B-green)](https://python.org)
[![Architecture](https://img.shields.io/badge/Architecture-Clean_Pipeline-orange)](https://github.com/your-repo/mops-downloader)

## 📋 Project Overview

The MOPS Downloader System is a Python-based tool designed to automatically download quarterly financial reports from Taiwan's Market Observation Post System (MOPS). The system provides a clean, automated pipeline for retrieving IFRSs financial reports in Chinese format with flexible targeting criteria to handle real-world variations in report availability.

## 📁 Complete Project Structure

```
.
├── .gitignore
├── DownloadAll.py    <--- Batch all stocks via skill-company-mops-financialreport-pdf-md
├── Get觀察名單.py     <--- Update StockID_TWSE_TPEX.csv to newest
├── instructions.md   <--- this file
├── LICENSE
├── README.md
├── requirements.txt
├── StockID_TWSE_TPEX.csv
├── .github/workflows
│   └── Download.yaml  <--- Github Actions for automated downloads
├── downloads/
├── logs/
├── data/reports/      <--- CSV matrix backups
├── mops_downloader/
│   ├── cli.py
│   ├── config.py
│   ├── exceptions.py
│   ├── main.py
│   ├── models.py
│   ├── __init__.py
│   ├── downloads/
│   │   ├── download_manager.py
│   │   └── __init__.py
│   ├── parsers/
│   │   ├── document_parser.py
│   │   └── __init__.py
│   ├── storage/
│   │   ├── file_manager.py
│   │   └── __init__.py
│   ├── utils/
│   │   ├── logging_config.py
│   │   ├── year_converter.py
│   │   └── __init__.py
│   ├── validators/
│   │   ├── input_validator.py
│   │   └── __init__.py
│   └── web/
│       ├── navigator.py
│       └── __init__.py
└── scripts/
    └── mops_downloader.py
```

## 🤖 Automated Downloads via GitHub Actions

The system includes automated quarterly downloads via GitHub Actions that align with Taiwan's MOPS filing deadlines.

### MOPS Filing Deadlines

Taiwan's MOPS requires companies to file quarterly reports by these deadlines:

| Quarter | Period | Filing Deadline | Auto-Download Window |
|---------|--------|----------------|---------------------|
| **Q1** | Jan-Mar | **May 15** | May 15-19 (5-day retry) |
| **Q2** | Apr-Jun | **Aug 14** | Aug 14-18 (5-day retry) |
| **Q3** | Jul-Sep | **Nov 14** | Nov 14-18 (5-day retry) |
| **Q4** | Oct-Dec | **March 31** (next year) | March 31 - April 4 (5-day retry) |

### GitHub Actions Workflow

**Location**: `.github/workflows/Download.yaml`

**Scheduled Triggers**:
- Runs automatically on filing deadline dates at 02:00 UTC
- 5-day retry window for each quarter
- Downloads previous quarter's reports (reports published AFTER quarter ends)
- Uses `--only-missing-files` flag to skip existing downloads

**Example Schedule**:
```yaml
# Q1 reports (Jan-Mar) - Available by May 15
- cron: '0 2 15 5 *'      # May 15 - First attempt
- cron: '0 2 16 5 *'      # May 16 - Retry 1
- cron: '0 2 17 5 *'      # May 17 - Retry 2
- cron: '0 2 18 5 *'      # May 18 - Retry 3
- cron: '0 2 19 5 *'      # May 19 - Final retry

# Q4 reports (Oct-Dec previous year) - Available by March 31
- cron: '0 2 31 3 *'      # March 31 - First attempt
- cron: '0 2 1 4 *'       # April 1 - Retry 1
- cron: '0 2 2 4 *'       # April 2 - Retry 2
- cron: '0 2 3 4 *'       # April 3 - Retry 3
- cron: '0 2 4 4 *'       # April 4 - Final retry
```

**Manual Trigger**:
The workflow can also be triggered manually via GitHub Actions UI with custom parameters:
- `year`: Target year (e.g., 2025)
- `quarter`: Target quarter (1, 2, 3, or 4)
- `delay`: Delay between downloads (seconds)
- `start_from`: Start from specific company ID
- `only_missing_files`: Skip existing files (default: true)
- `upload_to_sheets`: Upload matrix to Google Sheets (default: true)

**Automated Features**:
- ✅ Downloads all companies from StockID_TWSE_TPEX.csv
- ✅ Automatically commits downloaded PDFs to repository
- ✅ Uploads matrix view to Google Sheets (if configured)
- ✅ Creates CSV backup of matrix data in `data/reports/`
- ✅ Generates comprehensive status reports
- ✅ 5-day retry window for maximum success rate

### Quarter Download Timing

**Important**: The system downloads reports AFTER the filing deadline, not during the quarter:

```
Q1 2025 (Jan-Mar 2025)
├── Quarter Period: January 1 - March 31, 2025
├── Quarter Ends: March 31, 2025
├── Filing Deadline: May 15, 2025
└── Auto-Download: May 15-19, 2025 ✅ (reports are available)

Q4 2025 (Oct-Dec 2025)
├── Quarter Period: October 1 - December 31, 2025
├── Quarter Ends: December 31, 2025
├── Filing Deadline: March 31, 2026 (next year!)
└── Auto-Download: March 31 - April 4, 2026 ✅ (reports are available)
```

**Why This Timing?**
1. Companies need time after quarter ends to prepare financial statements
2. Taiwan law gives companies 45 days (Q1-Q3) or 90 days (Q4) to file
3. Most companies file near the deadline
4. The 5-day retry window catches late filers
5. This ensures reports are actually available when download runs

**Retry Logic**:
Each quarter has a 5-day retry window:
- **Day 1**: First attempt downloads all available reports
- **Days 2-5**: Only download files that failed or weren't available on previous days
- Smart skip logic prevents re-downloading existing files
- Commit messages show retry attempt: "Retry 1/5", "Retry 2/5", etc.

## 🎯 Goals and Objectives

### Primary Goal
Download MOPS quarterly financial reports based on user-specified criteria and organize them systematically with intelligent fallback mechanisms.

### Key Features
- Automated PDF download from MOPS system with two-step process
- Intelligent file organization by company ID (flattened structure)
- Support for multiple years and companies
- **Flexible report targeting** with primary and fallback criteria
- **Enhanced report analysis** and categorization
- Error handling and retry mechanisms
- Progress tracking and comprehensive logging
- **Real-time report availability analysis**
- **GitHub Actions automation** aligned with MOPS filing deadlines

## 📥 User Requirements

### Input Parameters
| Parameter | Type | Description | Example | Required |
|-----------|------|-------------|---------|----------|
| `year` | Integer | Reporting year (Western YYYY format) | 2024 | Yes |
| `company_id` | String | Taiwan stock company ID | "8272" | Yes |
| `quarter` | Integer/String | Quarter (1-4) or "all" | 1, 2, 3, 4, "all" | Optional (default: "all") |
| `strict_mode` | Boolean | Use strict targeting criteria only | True/False | Optional (default: False) |

### Enhanced Targeting System
The system now uses **flexible targeting** to handle real-world variations in MOPS report availability:

#### **Primary Targets** (Always preferred):
- **IFRSs個別財報** - Individual Financial Reports
- **IFRSs個體財報** - Individual Financial Reports (alternative naming)

#### **Secondary Targets** (Used when primary not available):
- **IFRSs合併財報** - Consolidated Financial Reports (may be individual for some companies)
- **財務報告書** - General Financial Reports

#### **File Pattern Targets**:
- **A12.pdf** - True individual reports (IFRSs個別財報 - company 8272 style)
- **A13.pdf** - True individual reports (IFRSs個體財報 - company 2330 historical style)
- **AI1.pdf** - Consolidated reports accepted as secondary target (IFRSs合併財報 - company 2330 current style)
- **A1[0-9].pdf** - Generic financial report patterns (mixed individual/consolidated)

#### **Exclusion Criteria** (Always rejected):
- **英文版** - English versions
- **AIA.pdf** - English consolidated reports
- **AE2.pdf** - English parent-subsidiary reports

### Operating Modes

#### **Flexible Mode** (Default)
- Uses primary targets first, falls back to secondary targets
- Provides comprehensive report analysis
- Maximizes download success rate
- Suitable for most users

#### **Strict Mode** (Optional)
- Only downloads primary individual reports
- Rejects all consolidated reports
- Follows original specification exactly
- For users requiring only pure individual reports

### User Input vs API Parameter Conversion
- **User Input**: Western year (e.g., 2024, 2023)
- **API Call**: ROC year (e.g., 113, 112)
- **Conversion**: Automatic conversion by system (Western year - 1911)

### Quarter Behavior
- **Specific Quarter** (1-4): Downloads only the specified quarter's report via `seamon={quarter}`
- **All Quarters** ("all" or default): Makes separate requests for each quarter (seamon=1,2,3,4) to ensure all available reports are downloaded

### Expected Behavior
1. Validate input parameters
2. Navigate to MOPS website with specified criteria
3. **Analyze all available reports** with categorization
4. **Apply flexible targeting logic** to maximize success
5. Locate and download target financial reports using **two-step process**
6. Save files with organized structure and consistent naming
7. **Provide comprehensive download summary** with missing quarter analysis

## 🗏️ System Architecture

### 5-Stage Enhanced Pipeline Architecture

```
Stage 1: Input Validation & Configuration
    ↓
Stage 2: Web Navigation & Authentication  
    ↓
Stage 3: Enhanced Document Discovery & Flexible Filtering
    ↓
Stage 4: Two-Step Download & File Management
    ↓
Stage 5: Verification, Analysis & Cleanup
```

## 🔧 Technical Specifications

### URL Pattern
```
Base URL: https://doc.twse.com.tw/server-java/t57sb01
Parameters:
- step=1
- colorchg=1
- seamon={quarter}  # 1=Q1, 2=Q2, 3=Q3, 4=Q4, empty=all
- mtype=A
- co_id={company_id}
- year={roc_year}   # ROC year format (Western year - 1911)
```

### Enhanced Download Process

#### **Two-Step Download Process**
1. **Step 1**: Fetch HTML page with report metadata and JavaScript download links
2. **Step 2**: Extract actual PDF URL from HTML response using pattern matching
3. **Step 3**: Download PDF file using extracted direct URL

#### **PDF URL Extraction**
The system extracts PDF URLs from HTML using multiple strategies:
- Direct href extraction from `<a>` tags
- JavaScript `readfile2()` parameter parsing
- Pattern matching for `/pdf/` paths
- Fallback URL construction

### Year Conversion (Critical!)
Taiwan uses Republic of China (ROC) calendar system:
```
ROC Year = Western Year - 1911

Examples:
- 2024 → 113 (2024 - 1911 = 113)
- 2023 → 112 (2023 - 1911 = 112)  
- 2022 → 111 (2022 - 1911 = 111)
```

### Enhanced File Organization Structure
```
downloads/
├── {company_id}/
│   ├── 202401_{company_id}_AXX.pdf  # Q1 (direct in company folder)
│   ├── 202402_{company_id}_AXX.pdf  # Q2 (flattened structure)
│   ├── 202403_{company_id}_AXX.pdf  # Q3 (original MOPS filename)
│   ├── 202404_{company_id}_AXX.pdf  # Q4 (original MOPS filename)
│   └── metadata.json                # Enhanced metadata with analysis
├── logs/
│   └── mops_downloader_{timestamp}.log
└── data/reports/
    └── mops_matrix_{timestamp}.csv  # Matrix backup from sheets uploader

Example for company 2330 (current AI1 pattern):
downloads/
├── 2330/
│   ├── 202401_2330_AI1.pdf
│   ├── 202402_2330_AI1.pdf  
│   ├── 202403_2330_AI1.pdf
│   ├── 202404_2330_AI1.pdf
│   └── metadata.json
```

### Report Availability Patterns - Enhanced Understanding

**The system now handles these real-world patterns**:

**Company 8272** (全景軟體):
- ✅ Q1, Q2, Q3, Q4: "IFRSs個別財報" (A12.pdf) - Primary target

**Company 2330** (台積電 TSMC - Current):
- ✅ Q1, Q2, Q3, Q4: "IFRSs合併財報" (AI1.pdf) - **Now accepted as secondary target**

**Company 2330** (台積電 TSMC - Historical):  
- ✅ Q1, Q2, Q3, Q4: "IFRSs個體財報" (A13.pdf) - Primary target

**Company 2382** (廣達 Quanta):
- ❌ Q1, Q2, Q3: Only consolidated reports available
- ✅ Q4: "IFRSs個體財報" (A13.pdf) - Primary target

### System Behavior for Mixed Report Availability

#### **Enhanced Analysis Phase**
Before downloading, the system analyzes and categorizes ALL available reports:

```
📊 Report Analysis:
   ✅ Target reports found: 2
      • IFRSs個別財報 → 202401_8272_A12.pdf (Matched primary target)
      • IFRSs合併財報 → 202401_2330_AI1.pdf (Matched flexible target)
   📋 Consolidated reports: 3
      • IFRSs合併財報 → 202401_XXXX_A1L.pdf
   🌍 English reports: 2
      • IFRSs英文版-合併財報 → 202401_XXXX_AIA.pdf
```

#### **Flexible Targeting Logic**
```python
# Primary targets (highest priority)
TARGET_INDIVIDUAL_REPORTS = [
    "IFRSs個別財報",     # True Individual Financial Report - Type 1 (A12.pdf)
    "IFRSs個體財報",     # True Individual Financial Report - Type 2 (A13.pdf)
    "IFRSs合併財報"      # Consolidated reports (A1L.pdf, AI1.pdf) - accepted as secondary
]

# File pattern targets
TARGET_FILE_PATTERNS = [
    r"A12\.pdf$",       # True individual (company 8272 style)
    r"A13\.pdf$",       # True individual (company 2330 historical)
    r"AI1\.pdf$",       # Consolidated accepted as secondary (company 2330 current)
    r"A1[0-9]\.pdf$"    # Generic patterns (mixed individual/consolidated)
]

# Flexible fallbacks (when primary targets not found)
FLEXIBLE_TARGETS = [
    "IFRSs合併財報",     # Consolidated reports (fallback when no individual available)
    "財務報告書"         # General financial reports (broadest fallback)
]

# Always excluded
EXCLUDED_KEYWORDS = [
    "英文版",            # English versions
    "AIA\.pdf",         # English consolidated
    "AE2\.pdf"          # English parent-subsidiary
]
```

## 🔍 Enhanced Document Discovery Specifications

### Advanced HTML Parsing
The `DocumentParser` now provides:

#### **Comprehensive Report Analysis**
- **Target Classification**: Identifies reports matching primary/secondary criteria
- **Report Categorization**: Groups reports by type (individual, consolidated, English, other)
- **Availability Assessment**: Determines which quarters have which report types
- **Matching Reason Tracking**: Records why each report was selected/rejected

#### **Enhanced Filtering Logic**
```python
def _is_target_report(self, description: str, filename: str = "") -> tuple[bool, str]:
    """Check if report matches target criteria with detailed reasoning."""
    
    # 1. Check exclusions first
    for keyword in EXCLUDED_KEYWORDS:
        if keyword in description or keyword in filename:
            return False, f"Excluded by keyword: {keyword}"
    
    # 2. Check primary targets  
    for target in TARGET_INDIVIDUAL_REPORTS:
        if target in description:
            return True, f"Matched primary target: {target}"
    
    # 3. Check file patterns
    for pattern in TARGET_FILE_PATTERNS:
        if re.search(pattern, filename):
            return True, f"Matched file pattern: {pattern}"
    
    # 4. Flexible fallbacks (if not in strict mode)
    if not self.strict_mode:
        for target in FLEXIBLE_TARGETS:
            if target in description:
                return True, f"Matched flexible target: {target}"
    
    return False, "No matching criteria"
```

### Enhanced File Naming Pattern Recognition
```
Current Patterns Supported:
- 202401_8272_A12.pdf (IFRSs個別財報 - True Individual)
- 202401_2330_A13.pdf (IFRSs個體財報 - True Individual, Historical)  
- 202401_2330_AI1.pdf (IFRSs合併財報 - Consolidated, accepted as secondary)
- 202401_XXXX_A1L.pdf (IFRSs合併財報 - Standard Consolidated, fallback)

Pattern: YYYYQQ_COMPANYID_TYPE.pdf
Type Classification:
- A12, A13 = True Individual Financial Reports (Primary Targets)
- AI1, A1L = Consolidated Financial Reports (Secondary Targets)
- AIA, A1A = English Consolidated Reports (Excluded)
- A1C = English Individual Reports (Excluded)
```

## 🔍 Detailed Component Specifications

### 1. Enhanced Input Validation Module
**Interface**: `InputValidator`
**New Features**:
- Support for `strict_mode` parameter
- Enhanced company ID validation with Taiwan stock code verification
- Flexible quarter input handling
- Comprehensive error reporting

### 2. Advanced Web Navigation Module  
**Interface**: `WebNavigator`
**Enhanced Features**:
- **Big5 encoding support** for Taiwan websites
- **SSL certificate flexibility** for MOPS compatibility
- **Intelligent content validation** with MOPS page detection
- **Enhanced retry logic** with exponential backoff
- **Proper User-Agent and header management**

### 3. Enhanced Document Discovery Module
**Interface**: `DocumentParser`
**Major Enhancements**:
- **Comprehensive report analysis** before filtering
- **Flexible vs strict mode** operation
- **Detailed reason tracking** for all decisions
- **Real-time report availability assessment**
- **Enhanced HTML parsing** with BeautifulSoup
- **JavaScript download URL extraction**

### 4. Advanced Download Manager Module
**Interface**: `DownloadManager`  
**New Two-Step Process**:
1. **HTML Page Retrieval**: Fetch report page with metadata
2. **PDF URL Extraction**: Parse HTML to find actual PDF download links
3. **PDF Download**: Download actual PDF files with verification
4. **File Validation**: Verify PDF headers and reasonable file sizes
5. **Smart Retry Logic**: Handle partial downloads and connection issues

### 5. Enhanced File Management Module
**Interface**: `FileManager`
**Enhanced Features**:
- **Flattened directory structure** for easier access
- **Enhanced metadata tracking** with download session history
- **Comprehensive download summaries**
- **Intelligent cleanup** of failed/partial downloads
- **JSON metadata** with detailed download information

## 📋 Enhanced Configuration Requirements

### Expanded Configuration Parameters
```python
# Enhanced Target Configuration
TARGET_INDIVIDUAL_REPORTS = [
    "IFRSs個別財報", "IFRSs個體財報", "IFRSs合併財報"
]

TARGET_FILE_PATTERNS = [
    r"A12\.pdf$", r"A13\.pdf$", r"AI1\.pdf$", r"A1[0-9]\.pdf$"
]

EXCLUDED_KEYWORDS = [
    "英文版", "AIA\.pdf", "AE2\.pdf"
]

FLEXIBLE_TARGETS = [
    "IFRSs合併財報", "財務報告書"
]

# Network Configuration  
VERIFY_SSL = False  # Disabled for MOPS compatibility
RATE_LIMIT_DELAY = 1.0
MAX_RETRIES = 3
TIMEOUT = 30

# Download Configuration
CHUNK_SIZE = 8192
MAX_CONCURRENT_DOWNLOADS = 1  # Sequential for reliability
```

## 🚦 Enhanced Error Handling Strategy

### Error Categories with Enhanced Handling
1. **Network Errors**: SSL issues, encoding problems, connection timeouts
2. **Parsing Errors**: HTML structure changes, missing elements, encoding issues
3. **Download Errors**: PDF extraction failures, invalid URLs, file corruption
4. **Validation Errors**: Invalid parameters, unsupported quarters/years
5. **Partial Success**: Mixed availability scenarios (not errors, logged as info)

### Enhanced Recovery Mechanisms
- **Encoding fallback chain**: Big5 → UTF-8 → GB2312 → Latin-1
- **SSL compatibility mode** for Taiwan government websites
- **Flexible PDF URL extraction** with multiple parsing strategies
- **Graceful degradation** for missing reports
- **Comprehensive logging** with detailed error context

## 📊 Enhanced Logging and Monitoring

### Enhanced Log Levels with Context
```
[INFO] Stage 1: Validating inputs...
[INFO] 📊 Report Analysis:
[INFO]    ✅ Target reports found: 2
[INFO]       • IFRSs個別財報 → 202401_8272_A12.pdf (Matched primary target)
[INFO]    📋 Consolidated reports: 1
[INFO]    🌍 English reports: 2
[INFO] ✅ Successfully downloaded 202401_8272_A12.pdf (1,234,567 bytes)
```

### Enhanced Metrics Tracking
- **Report availability patterns** by company and quarter
- **Target matching statistics** (primary vs flexible vs file pattern)
- **Download success rates** by report type and company
- **File size distributions** and validation statistics
- **Parsing success rates** and HTML structure changes
- **Network performance** and retry statistics

## 🧪 Enhanced Testing Strategy

### Unit Tests
- **Flexible targeting logic** with various company scenarios
- **Two-step download process** with mock HTML responses
- **Enhanced parsing** with real MOPS HTML samples
- **PDF URL extraction** with various JavaScript patterns
- **Error handling** for all network and parsing scenarios

### Integration Tests  
- **End-to-end flexible download** scenarios
- **Mixed report availability** handling
- **Real MOPS website** compatibility testing
- **Encoding and SSL** compatibility verification

### Performance Tests
- **Large file download** efficiency
- **Memory usage** during HTML parsing and PDF downloads
- **Network retry logic** performance impact

## 🔄 Enhanced Usage Examples

### Command Line Interface - Enhanced
```bash
# Download with flexible targeting (default mode)
python ../skills/common/skill-company-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py 2330 2024 all --only-missing-files

# Download with strict targeting (individual reports only)
python ../skills/common/skill-company-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py 2330 2024 all --only-missing-files

# Download specific quarter with detailed logging
python ../skills/common/skill-company-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py 8272 2023 2 --log-level DEBUG --only-missing-files

# Custom output directory
python ../skills/common/skill-company-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py 2382 2023 all --only-missing-files
```

### Enhanced Python API Interface
```python
from mops_downloader import MOPSDownloader

# Initialize with options
downloader = MOPSDownloader(
    download_dir="./reports",
    strict_mode=False,  # Use flexible targeting
    log_level="INFO"
)

# Download with enhanced result information
result = downloader.download("2330", 2024, "all")

# Access detailed results
print(f"Success: {result.success}")
print(f"Downloaded: {result.downloaded_files}")
print(f"Missing quarters: {result.missing_quarters}")
print(f"Total size: {result.total_size:,} bytes")
print(f"File paths: {result.file_paths}")

# Enhanced result data structure
result.success              # Overall success boolean
result.downloaded_files     # List of successful filenames  
result.missing_quarters     # Quarters without target reports
result.error_details        # Detailed error information
result.file_paths          # Full paths to downloaded files
result.total_files         # Count of successful downloads
result.total_size          # Total bytes downloaded
```

### Enhanced Result Examples
```
Scenario 1 - Flexible Success (Company 2330):
✅ Downloaded: ['202401_2330_AI1.pdf', '202402_2330_AI1.pdf', '202403_2330_AI1.pdf', '202404_2330_AI1.pdf']
📊 Report Analysis showed consolidated reports were used as secondary targets
Result: 4/4 reports downloaded (flexible targeting successful)

Scenario 2 - Mixed Success (Company 2382):  
✅ Downloaded: ['202404_2382_A13.pdf']
❌ Missing: Q1, Q2, Q3 (only consolidated reports available)
📊 Report Analysis found individual reports only for Q4
Result: 1/4 reports downloaded (partial success with clear explanation)

Scenario 3 - Strict Mode (Company 2330):
❌ Downloaded: []
📊 Report Analysis found only consolidated reports (excluded in strict mode)
Result: 0/4 reports downloaded (strict mode rejected available consolidated reports)
```

## 📈 Future Enhancements

### Phase 2 Features - Enhanced
- **Machine learning-based** report type classification
- **Automatic company categorization** by report availability patterns
- **Enhanced metadata extraction** from PDF content
- **Batch processing optimization** with intelligent scheduling
- **Web dashboard** with real-time download monitoring

### Phase 3 Features - Advanced
- **Multi-market support** (Hong Kong, Singapore exchanges)
- **Historical trend analysis** of report availability changes
- **Automated data extraction** from downloaded PDFs
- **API integration** with financial databases and analysis tools
- **Cloud deployment** with distributed download capabilities

## 🔒 Enhanced Security Considerations

### Enhanced Best Practices
- **SSL flexibility** for Taiwan government website compatibility
- **Encoding security** with proper character set handling
- **Rate limiting compliance** with respectful server interaction
- **User agent rotation** and session management
- **Input sanitization** with Taiwan-specific validation rules
- **Content validation** to ensure downloaded files are legitimate PDFs

## 📝 Implementation Notes

### Enhanced Development Priorities
1. **Flexible targeting system** with comprehensive analysis
2. **Two-step download process** with robust PDF extraction
3. **Enhanced error handling** with graceful degradation
4. **Real-world compatibility** with various MOPS response formats
5. **Comprehensive logging** and monitoring capabilities
6. **User-friendly results** with detailed success/failure reporting
7. **GitHub Actions integration** with MOPS deadline alignment

### Performance Optimizations - Enhanced
- **Sequential downloads** for reliability (avoiding concurrent complications)
- **Intelligent caching** of HTML parsing results
- **Memory-efficient** PDF downloading with streaming
- **Smart retry logic** with exponential backoff
- **Enhanced SSL context** management for government websites

---

## 📞 Support and Maintenance

### Issue Reporting - Enhanced
- **Detailed logging** makes troubleshooting easier
- **Report analysis output** helps identify MOPS website changes
- **Enhanced error messages** provide actionable guidance
- **Comprehensive test coverage** with real-world scenarios

### Maintenance Schedule - Enhanced
- **Monthly compatibility testing** with MOPS website
- **Quarterly report targeting review** based on new company patterns
- **Semi-annual dependency updates** with Taiwan-specific testing
- **Annual architecture assessment** and enhancement planning

---

*Document Version: 2.0.0*  
*Last Updated: 2025-10-15*  
*Next Review: 2026-01-15*

## 🔄 Version History

### v2.0.0 (2025-10-15)
- **Major Enhancement**: Added flexible targeting system
- **New Feature**: Two-step download process with PDF URL extraction
- **Enhancement**: Comprehensive report analysis and categorization
- **Change**: Flattened file organization structure
- **Addition**: Support for modern MOPS file patterns (AI1.pdf)
- **Improvement**: Enhanced error handling and logging
- **Feature**: Strict vs flexible mode operation
- **Integration**: GitHub Actions automation with MOPS filing deadline alignment
- **Automation**: 5-day retry window for each quarter
- **Enhancement**: Google Sheets matrix upload integration

### v1.0.0 (2025-08-04)
- Initial specification with basic individual report targeting
- Simple download process and file organization
- Basic error handling and logging framework