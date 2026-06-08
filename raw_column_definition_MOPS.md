---
source: https://raw.githubusercontent.com/wenchiehlee-investment/MOPS/refs/heads/main/raw_column_definition.md
destination: https://raw.githubusercontent.com/wenchiehlee-investment/Python-Actions.GoodInfo.Analyzer/refs/heads/main/definitions/raw_column_definition_MOPS.md
---

# Raw CSV Column Definitions - MOPS Repo

---

## raw_mops_matrix.csv (Financial Report Matrix)
**No:** 60
**Source:** `data/reports/mops_matrix_*.csv`
**Purpose:** Summarize which financial reports are available for each company and quarter.
**Note:** Current implementation in `MOPS` uses a wide-format matrix which lacks the standard metadata suffix in the raw output.

### Column Definitions (Matrix Format):

| Column | Type | Description | Values |
|--------|------|-------------|--------|
| `代號` | string | Stock Code (Mapping to `stock_code`) | `2330`, `1587` |
| `名稱` | string | Company Name (Mapping to `company_name`) | `台積電`, `日月光` |
| `YYYY QN` | string | Report coverage for specific Year/Quarter | `AI1`, `AI3`, `AI1/AI3`, `-` |

### Value Mapping:
- **AI1**: Consolidated Balance Sheet (合併資產負債表)
- **AI2**: Individual Balance Sheet (個別資產負債表)
- **AI3**: Consolidated Income Statement (合併損益表)
- **AI4**: Individual Income Statement (個別損益表)
- **-**: No data available or missing.

## mops_matrix_latest.csv (Financial Report Matrix - Latest)
**No:** 60b
**Source:** `data/reports/mops_matrix_latest.csv`
**Purpose:** Stable tracking path for the matrix CSV file recording quarterly report presence. Uses same format as `raw_mops_matrix.csv`.

---

## mops_health_summary.csv (MOPS Data & Conversion Health Summary)
**No:** 62
**Source:** `data/reports/mops_health_summary.csv`
**Purpose:** Summarize download status, conversion status, and OCR needs for MOPS PDF/MD files.

### Column Definitions:

| Column | Type | Description |
|--------|------|-------------|
| `process_timestamp` | timestamp | Time when health metrics were computed (Taipei Time). |
| `download_timestamp` | timestamp | Modification time of the newest PDF file (Taipei Time). |
| `total_pdfs` | int | Total number of PDF files found under `downloads/`. |
| `total_mds` | int | Total number of converted MD files found under `downloads/`. |
| `conversion_rate_pct` | float | Overall conversion rate percentage (`total_mds / total_pdfs * 100`). |
| `ocr_needed_count` | int | Total number of MD files containing the "TODO:OCR" scanned image warning. |
| `pending_conversions` | int | Total number of PDFs waiting to be processed (always 0 after batch converter completes). |
| `failed_conversions` | int | Total number of PDFs missing matching MD files or failing to read. |
| `latest_md_time` | timestamp | Modification time of the newest MD file (Taipei Time). |
| `checked_at` | timestamp | Execution time of the health checker (same as `process_timestamp`). |

---

## [Proposed] raw_mops_financials.csv (Raw Financial Data)
**No:** 61
**Note:** Placeholder for future implementation where specific fields (Total Revenue, Net Income, etc.) are extracted.

