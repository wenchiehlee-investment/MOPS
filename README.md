# MOPS Downloader System

[![Version](https://img.shields.io/badge/Version-2.0.0-blue)](https://github.com/your-repo/mops-downloader)
[![Python](https://img.shields.io/badge/Python-3.9%2B-green)](https://python.org)
[![Architecture](https://img.shields.io/badge/Architecture-Clean_Pipeline-orange)](https://github.com/your-repo/mops-downloader)

A Python-based tool for automatically downloading quarterly financial reports from Taiwan's Market Observation Post System (MOPS). Designed to handle real-world variations in report availability with intelligent fallback mechanisms.

## 🎯 What This Tool Does

- **Automates MOPS Downloads**: Fetches IFRSs financial reports in Chinese format from Taiwan's official MOPS system
- **Smart Report Detection**: Uses flexible targeting to find the best available reports (individual reports preferred, consolidated as fallback)
- **Organized File Management**: Downloads are systematically organized by company with consistent naming
- **Handles Real-World Complexity**: Different companies have different report types available - this tool adapts automatically

## ✨ Key Features

- 🔍 **Flexible Targeting System**: Prioritizes individual financial reports but falls back to consolidated reports when needed
- 📁 **Clean Organization**: Files saved in `downloads/{company_id}/` with standardized naming
- 🛡️ **Robust Error Handling**: Handles SSL issues, encoding problems, and missing reports gracefully  
- 📊 **Comprehensive Analysis**: Shows exactly what reports were found and why they were selected/rejected
- 🔄 **Two Operating Modes**: Flexible mode (default) for maximum success, strict mode for individual reports only
- 📝 **Detailed Logging**: Complete audit trail of all operations and decisions

## 🚀 Quick Start

### Installation

1. **Clone the repository**:
```bash
git clone https://github.com/your-repo/mops-downloader.git
cd mops-downloader
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

### Basic Usage

**Download all quarters for a company (recommended)**:
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024
```

**Download specific quarter**:
```bash
python scripts/mops_downloader.py --company_id 8272 --year 2023 --quarter 2
```

**Use strict mode (individual reports only)**:
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024 --strict_mode
```

### Batch Processing

**Update stock list and download all companies**:
```bash
# First, update the stock list
python Get觀察名單.py

# Then download reports for all companies in the list
python DownloadAll.py
```

## 📋 Input Parameters

| Parameter | Type | Description | Example | Default |
|-----------|------|-------------|---------|---------|
| `company_id` | String | Taiwan stock company ID | "2330", "8272" | Required |
| `year` | Integer | Reporting year (Western format) | 2024, 2023 | Required |
| `quarter` | Integer/String | Quarter (1-4) or "all" | 1, 2, 3, 4, "all" | "all" |
| `strict_mode` | Boolean | Only download individual reports | True/False | False |
| `output` | String | Output directory | "./reports" | "./downloads" |

## 🎯 Understanding Report Types

The system intelligently handles different types of financial reports:

### Primary Targets (Preferred)
- **IFRSs個別財報** - Individual Financial Reports (A12.pdf)
- **IFRSs個體財報** - Individual Financial Reports (A13.pdf)

### Secondary Targets (Fallback)
- **IFRSs合併財報** - Consolidated Financial Reports (AI1.pdf, A1L.pdf)
- **財務報告書** - General Financial Reports

### Always Excluded
- **英文版** - English versions
- **AIA.pdf**, **AE2.pdf** - English consolidated reports

## 📂 Output Structure

```
downloads/
├── 2330/                           # Company folder
│   ├── 202401_2330_AI1.pdf        # Q1 2024
│   ├── 202402_2330_AI1.pdf        # Q2 2024  
│   ├── 202403_2330_AI1.pdf        # Q3 2024
│   ├── 202404_2330_AI1.pdf        # Q4 2024
│   └── metadata.json              # Download metadata
├── 8272/
│   ├── 202401_8272_A12.pdf
│   └── metadata.json
└── logs/
    └── mops_downloader_20240805_143022.log
```

**File Naming**: `YYYYQQ_{company_id}_{report_type}.pdf`
- `YYYY`: Year (2024)
- `QQ`: Quarter (01, 02, 03, 04)
- `{company_id}`: Company stock ID
- `{report_type}`: A12, A13, AI1, etc.

## 💡 Usage Examples

### Example 1: Taiwan Semiconductor (TSMC) - Company 2330
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024
```

**Expected Result**: Downloads consolidated reports (AI1.pdf) as individual reports aren't available
```
✅ Downloaded: 202401_2330_AI1.pdf, 202402_2330_AI1.pdf, 202403_2330_AI1.pdf, 202404_2330_AI1.pdf
📊 Used consolidated reports as fallback (no individual reports available)
```

### Example 2: Systex Corporation - Company 8272
```bash
python scripts/mops_downloader.py --company_id 8272 --year 2024
```

**Expected Result**: Downloads individual reports (A12.pdf) - preferred type
```
✅ Downloaded: 202401_8272_A12.pdf, 202402_8272_A12.pdf, 202403_8272_A12.pdf, 202404_8272_A12.pdf
📊 Used individual reports (primary target achieved)
```

### Example 3: Mixed Availability - Company 2382
```bash
python scripts/mops_downloader.py --company_id 2382 --year 2023
```

**Expected Result**: Partial success with clear explanation
```
✅ Downloaded: 202304_2382_A13.pdf
❌ Missing: Q1, Q2, Q3 (only consolidated reports available, individual reports found for Q4 only)
```

## 🔧 Python API Usage

```python
from mops_downloader import MOPSDownloader

# Initialize downloader
downloader = MOPSDownloader(
    download_dir="./financial_reports",
    strict_mode=False,  # Use flexible targeting
    log_level="INFO"
)

# Download reports
result = downloader.download("2330", 2024, "all")

# Check results
if result.success:
    print(f"✅ Successfully downloaded {result.total_files} files")
    print(f"📁 Files: {result.downloaded_files}")
    print(f"💾 Total size: {result.total_size:,} bytes")
else:
    print(f"❌ Download failed: {result.error_details}")

# Handle partial success
if result.missing_quarters:
    print(f"⚠️ Missing quarters: {', '.join(result.missing_quarters)}")
```

## 📊 Understanding the Output

### Successful Download
```
[INFO] 📊 Report Analysis:
[INFO]    ✅ Target reports found: 4
[INFO]       • IFRSs個別財報 → 202401_8272_A12.pdf (Matched primary target)
[INFO]       • IFRSs個別財報 → 202402_8272_A12.pdf (Matched primary target)
[INFO]    📋 Consolidated reports: 0
[INFO]    🌍 English reports: 0
[INFO] ✅ Download completed successfully: 4/4 files (12.3 MB total)
```

### Partial Success with Explanation
```
[INFO] 📊 Report Analysis:
[INFO]    ✅ Target reports found: 1
[INFO]       • IFRSs個體財報 → 202304_2382_A13.pdf (Matched primary target)
[INFO]    📋 Consolidated reports: 3 (excluded in flexible mode preference)
[INFO] ⚠️ Download completed with missing quarters: 1/4 files
[INFO] ❌ Q1, Q2, Q3: No individual reports available
```

## 🛠️ Configuration

### Environment Setup
```bash
# Optional: Create virtual environment
python -m venv venv
source venv/bin/activate  # Linux/Mac
# or
venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt
```

### Common Configuration Options
```python
# In your script or config file
DOWNLOAD_CONFIG = {
    'verify_ssl': False,           # Needed for MOPS compatibility
    'rate_limit_delay': 1.0,       # Seconds between requests
    'max_retries': 3,              # Retry attempts for failed downloads
    'timeout': 30,                 # Request timeout in seconds
    'strict_mode': False           # Use flexible targeting by default
}
```

## 🔍 Troubleshooting

### Common Issues

**SSL Certificate Errors**:
```
Solution: SSL verification is automatically disabled for MOPS compatibility
```

**Encoding Issues**:
```
Solution: The system automatically handles Big5/UTF-8 encoding conversion
```

**No Reports Found**:
```
Check: 1) Company ID is correct 2) Year/quarter has data 3) Try flexible mode
```

**Partial Downloads**:
```
This is normal - not all companies have all report types for all quarters
Check the detailed log output for explanation
```

### Debug Mode
```bash
python scripts/mops_downloader.py --company_id 2330 --year 2024 --log_level DEBUG
```

## 📁 Project Structure

```
mops-downloader/
├── mops_downloader/           # Main package
│   ├── downloads/             # Download management
│   ├── parsers/              # HTML/document parsing
│   ├── storage/              # File management
│   ├── validators/           # Input validation
│   └── web/                  # Web navigation
├── scripts/
│   └── mops_downloader.py    # Main CLI script
├── DownloadAll.py            # Batch download all companies
├── Get觀察名單.py             # Update stock list
├── StockID_TWSE_TPEX.csv    # Taiwan stock company list
├── downloads/                # Downloaded files (created automatically)
├── logs/                     # Log files (created automatically)
└── requirements.txt          # Python dependencies
```

## 📈 Requirements

- **Python**: 3.9 or higher
- **Dependencies**: See `requirements.txt`
- **Network**: Internet connection for MOPS access
- **Disk Space**: Varies by usage (PDFs are typically 1-5 MB each)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Make your changes
4. Add tests if applicable
5. Commit your changes (`git commit -am 'Add new feature'`)
6. Push to the branch (`git push origin feature/new-feature`)
7. Create a Pull Request

## 📞 Support

- **Documentation**: See `instructions.md` for detailed technical specifications
- **Issues**: Report bugs or request features via GitHub issues
- **Logs**: Check `logs/` directory for detailed error information

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔄 Version History

### v2.0.0 (Current)
- ✅ Flexible targeting system with intelligent fallbacks
- ✅ Two-step download process for improved reliability  
- ✅ Comprehensive report analysis and categorization
- ✅ Enhanced error handling and logging
- ✅ Support for modern MOPS file patterns

### v1.0.0
- Basic individual report downloading
- Simple file organization
- Core functionality

---

**Note**: This tool is designed to work with Taiwan's MOPS system and handles the complexities of real-world financial report availability. The flexible targeting system ensures maximum download success while providing clear explanations for any missing reports.