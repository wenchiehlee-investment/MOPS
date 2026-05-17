# MOPS Downloader System

[![Version](https://img.shields.io/badge/Version-2.0.0-blue)](https://github.com/your-repo/mops-downloader)
[![Python](https://img.shields.io/badge/Python-3.9%2B-green)](https://python.org)
[![Architecture](https://img.shields.io/badge/Architecture-Clean_Pipeline-orange)](https://github.com/your-repo/mops-downloader)

A Python-based tool for automatically downloading quarterly financial reports from Taiwan's Market Observation Post System (MOPS). Designed to handle real-world variations in report availability with intelligent fallback mechanisms.

<!-- BEGIN_STATUS -->

## 📊 Current Download Status

> **Last Updated**: 2026-05-17 | **Source**: `mops_matrix_20260517_060128.csv`

**118 companies tracked**

### 季財報 概況

| Quarter | 季財報 | Coverage | Notes |
|---------|--------|----------|-------|
| 2026 Q2 | 0 / 118 | — |  |
| 2026 Q1 | 109 / 118 | 92% | Filing deadline: May 15 |
| 2025 Q4 | 109 / 118 | 92% | Filing deadline: Mar 31 (next year) |
| 2025 Q3 | 109 / 118 | 92% |  |
| 2025 Q2 | 111 / 118 | 94% |  |
| 2025 Q1 | 105 / 118 | 89% |  |
| 2024 Q4 | 90 / 118 | 76% |  |
| 2024 Q3 | 78 / 118 | 66% |  |
| 2024 Q2 | 17 / 118 | 14% |  |
| 2024 Q1 | 15 / 118 | 13% |  |
| 2023 Q4 | 12 / 118 | 10% |  |
| 2023 Q3 | 9 / 118 | 8% |  |
| 2023 Q2 | 11 / 118 | 9% |  |
| 2023 Q1 | 9 / 118 | 8% |  |
| 2020 Q4 | 1 / 118 | 1% |  |
| 2020 Q3 | 1 / 118 | 1% |  |
| 2020 Q2 | 1 / 118 | 1% |  |
| 2020 Q1 | 1 / 118 | 1% |  |

---

### 📂 法說會 & 新聞 資料庫

#### 完成度概況

> 各公司法說會簡報、逐字稿、新聞收錄數量。點擊公司名稱展開詳細連結。

| 公司 | 法說會 PDF/MD | 逐字稿 | 新聞 |
|------|:------------:|:------:|:----:|
| [2357 華碩](#2357-華碩) | 11 季 | 2 | 2 |
| [2382 廣達](#2382-廣達) | — | 1 | — |

**覆蓋率**：2 / 118 companies

---

#### 2357 華碩

<details>
<summary>法說會（季度）</summary>

| Quarter | 法說會 PDF/MD | 逐字稿 |
|---------|:------------:|:------:|
| 2025 Q3 | [MD](downloads/2357/InvestorRelation/2025Q3_IR_Chinese.md) | [2025-11-11](downloads/2357/InvestorRelation/法說會逐字稿/華碩_2025-11-11.md) / [2025-12-08](downloads/2357/InvestorRelation/法說會逐字稿/華碩_2025-12-08.md) |
| 2025 Q2 | [MD](downloads/2357/InvestorRelation/2025Q2_IR_Chinese.md) | — |
| 2025 Q1 | [MD](downloads/2357/InvestorRelation/2025Q1_IR_Chinese.md) | — |
| 2024 Q4 | [MD](downloads/2357/InvestorRelation/2024Q4_IR_Chinese.md) | — |
| 2024 Q3 | [MD](downloads/2357/InvestorRelation/2024Q3_IR_Chinese.md) | — |
| 2024 Q2 | [MD](downloads/2357/InvestorRelation/2024Q2_IR_Chinese.md) | — |
| 2024 Q1 | [MD](downloads/2357/InvestorRelation/2024Q1_IR_Chinese.md) | — |
| 2023 Q4 | [MD](downloads/2357/InvestorRelation/2023Q4_IR_Chinese.md) | — |
| 2023 Q3 | [MD](downloads/2357/InvestorRelation/2023Q3_IR_Chinese.md) | — |
| 2023 Q2 | [MD](downloads/2357/InvestorRelation/2023Q2_IR_Chinese.md) | — |
| 2023 Q1 | [MD](downloads/2357/InvestorRelation/2023Q1_IR_Chinese.md) | — |

</details>

<details>
<summary>新聞</summary>

| 日期 | 標題 |
|------|------|
| 2026-01-29 | [華碩全力衝 AI 伺服器 2026年獨立事業群「福將」朱培蘭掌旗](downloads/2357/News/news_20260129_udn_asus_server_bg.md) |
| 2026-03-06 | [華碩 AI 伺服器戰略不攻大廠 北美四大 CSP 之外生意空間有多大？](downloads/2357/News/2026-03-06_udn_asus-ai-server-tier2-csp.md) |

</details>

---

#### 2382 廣達

<!-- END_STATUS -->

### Report Types (2025 Q3)

| Type | Count |
|------|-------|
| AI1 | 98 |
| AI2 | 11 |

### Companies Missing Recent Data

**Missing 2025 Q3** (9 companies):
2353 宏碁、6035 悠遊卡、6285 啟碁、6690 安碁資訊、6699 奇邑、6811 宏碁資訊、6850 光鼎生技、7737 凱鈿、7794 宏碁智新

**Missing 2025 Q2** (7 companies):
2353 宏碁、6285 啟碁、6690 安碁資訊、6811 宏碁資訊、6962 奕力-KY、7749 意騰-KY、7794 宏碁智新

**Missing 2025 Q1** (13 companies):
2345 智邦、2353 宏碁、2359 所羅門、2383 台光電、2405 輔信、6035 悠遊卡、6285 啟碁、6690 安碁資訊、6699 奇邑、6811 宏碁資訊、6850 光鼎生技、7737 凱鈿、7794 宏碁智新

---

## 🎯 What This Tool Does

- **Automates MOPS Downloads**: Fetches IFRSs financial reports in Chinese format from Taiwan's official MOPS system
- **Smart Report Detection**: Uses flexible targeting to find the best available reports (individual reports preferred, consolidated as fallback)
- **Organized File Management**: Downloads are systematically organized by company with consistent naming
- **Handles Real-World Complexity**: Different companies have different report types available - this tool adapts automatically
- **GitHub Actions Integration**: Automated quarterly downloads aligned with MOPS filing deadlines

## ✨ Key Features

- 📥 **Flexible Targeting System**: Prioritizes individual financial reports but falls back to consolidated reports when needed
- 📁 **Clean Organization**: Files saved in `downloads/{company_id}/` with standardized naming
- 🛡️ **Robust Error Handling**: Handles SSL issues, encoding problems, and missing reports gracefully  
- 📊 **Comprehensive Analysis**: Shows exactly what reports were found and why they were selected/rejected
- 🔄 **Two Operating Modes**: Flexible mode (default) for maximum success, strict mode for individual reports only
- 📝 **Detailed Logging**: Complete audit trail of all operations and decisions
- 🤖 **Automated Scheduling**: GitHub Actions workflow runs on MOPS filing deadlines with 5-day retry windows

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
python DownloadAll.py --year 2024 --quarter 1
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

## 🤖 Automated Downloads

### GitHub Actions Integration

The system includes automated quarterly downloads that run on a schedule aligned with Taiwan's MOPS filing deadlines.

#### MOPS Filing Schedule

Taiwan's Market Observation Post System (MOPS) requires companies to file quarterly reports by specific deadlines. Our automated system downloads reports immediately after these deadlines:

| Quarter | Period | Filing Deadline | Auto-Download Window |
|---------|--------|----------------|----------------------|
| **Q1** | Jan-Mar | **May 15** | May 15-19 (5 days) |
| **Q2** | Apr-Jun | **Aug 14** | Aug 14-18 (5 days) |
| **Q3** | Jul-Sep | **Nov 14** | Nov 14-18 (5 days) |
| **Q4** | Oct-Dec | **March 31** (next year) | March 31 - April 4 (5 days) |

#### How It Works

1. **Automatic Execution**: GitHub Actions runs on filing deadline dates at 02:00 UTC
2. **5-Day Retry Window**: Attempts download for 5 consecutive days to catch late filings
3. **Smart Skip Logic**: Only downloads missing files (won't re-download existing PDFs)
4. **Matrix Upload**: Automatically uploads status matrix to Google Sheets (if configured)
5. **Auto-Commit**: Commits all downloaded PDFs and metadata to the repository
6. **Comprehensive Logging**: Creates detailed logs and status reports for each run

#### Why Downloads Run AFTER Filing Deadlines

Reports are published **after** quarters end, so downloads are scheduled accordingly:

**Example: Q1 2025 Timeline**
```
├── Quarter Period: January 1 - March 31, 2025
├── Quarter Ends: March 31, 2025
├── Filing Deadline: May 15, 2025 ← Companies must file by this date
└── Auto-Download: May 15-19, 2025 ✅ Reports are now available!

Why the delay?
- Q1 doesn't end until March 31
- Companies need time to prepare financial statements
- Legal filing deadline is May 15 (45 days after quarter end)
- Most companies file near the deadline
- 5-day window ensures we catch all filings
```

**Example: Q4 2025 Timeline**
```
├── Quarter Period: October 1 - December 31, 2025
├── Quarter Ends: December 31, 2025
├── Filing Deadline: March 31, 2026 ← Next year!
└── Auto-Download: March 31 - April 4, 2026 ✅ Reports are now available!

Why March 2026?
- Q4 2025 is the annual report (full year)
- Companies get until March 31 of NEXT year to file
- This is 90 days after year-end for comprehensive audit
- Auto-download runs in March/April 2026 for Q4 2025 data
```

#### Manual Trigger

You can manually trigger downloads via GitHub Actions without waiting for the scheduled runs:

**Steps:**
1. Go to your repository on GitHub
2. Click the **"Actions"** tab
3. Select **"Download MOPS PDFs"** workflow from the left sidebar
4. Click **"Run workflow"** button (top right)
5. Configure parameters:
   - **Year**: Target year (e.g., 2025, 2024)
   - **Quarter**: Specific quarter (1, 2, 3, or 4)
   - **Delay**: Seconds between downloads (default: 10.0)
   - **Start from**: Optional company ID to start from (default: 2412)
   - **Skip existing files**: ✅ Recommended (only download missing files)
   - **Upload to sheets**: ✅ Enable for Google Sheets matrix view
6. Click **"Run workflow"** to start

**Use Cases for Manual Trigger:**
- Download historical data for past years
- Re-download specific quarters if needed
- Test the workflow with custom parameters
- Download immediately without waiting for scheduled run

#### Monitoring Downloads

**Check Download Status:**

- **Actions Tab**: View real-time workflow execution logs
  - See which companies are being processed
  - Track download progress and errors
  - View retry attempts (1/5, 2/5, etc.)

- **Commits**: Look for automated commit messages
  - `📥 Scheduled MOPS Download (Retry 1/5) - 2025 Q1`
  - `📥 Scheduled MOPS Download (Retry 2/5) - 2025 Q1`
  - Shows number of files downloaded

- **Google Sheets**: Matrix view (if configured)
  - Worksheet: "MOPS下載狀態"
  - Shows comprehensive download status for all companies
  - Updated automatically after each run

- **Repository Files**: Direct file inspection
  - Check `downloads/` folder for new PDFs
  - Review `logs/` for detailed execution logs
  - Check `data/reports/` for CSV matrix backups

**Example Commit Messages:**
```
📥 Scheduled MOPS Download (Retry 1/5) + 📊 Matrix Upload - 2025 Q1 (95 files from 110 companies)
📥 Scheduled MOPS Download (Retry 2/5) + 📊 Matrix Upload - 2025 Q1 (8 files from 12 companies)
📥 Scheduled MOPS Download (Retry 3/5) + 📊 Matrix Upload - 2025 Q1 (2 files from 3 companies)
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

### GitHub Actions Setup

To enable automated downloads, configure these repository secrets:

1. Go to **Settings** → **Secrets and variables** → **Actions**
2. Add the following secrets:
   - `GOOGLE_SHEETS_CREDENTIALS`: Your Google service account JSON (optional)
   - `GOOGLE_SHEET_ID`: Your Google Sheets spreadsheet ID (optional)

**Note**: Google Sheets integration is optional. The system will generate CSV backups even without Sheets credentials.

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

**GitHub Actions Not Running**:
```
Check:
1. Workflow file is in .github/workflows/Download.yaml
2. Actions are enabled in repository settings
3. Scheduled time hasn't arrived yet (check cron schedule)
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
├── .github/workflows/
│   └── Download.yaml         # GitHub Actions automation
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
- **GitHub Actions**: Optional, for automated downloads

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
- **Actions**: Monitor GitHub Actions tab for automated download status

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📝 Version History

### v2.0.0 (Current)
- ✅ Flexible targeting system with intelligent fallbacks
- ✅ Two-step download process for improved reliability  
- ✅ Comprehensive report analysis and categorization
- ✅ Enhanced error handling and logging
- ✅ Support for modern MOPS file patterns
- ✅ GitHub Actions automation with MOPS deadline alignment
- ✅ 5-day retry window for maximum success rate
- ✅ Google Sheets matrix integration

### v1.0.0
- Basic individual report downloading
- Simple file organization
- Core functionality

---

**Note**: This tool is designed to work with Taiwan's MOPS system and handles the complexities of real-world financial report availability. The flexible targeting system ensures maximum download success while providing clear explanations for any missing reports. Automated downloads run on Taiwan's official filing deadlines to ensure reports are available when downloaded.