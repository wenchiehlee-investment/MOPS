# Raw CSV Column Definitions - MOPS Repo

## Standard Metadata Columns (Expected for All Reports)

Following the `biztrends.TW` core specification, all raw CSV files intended for the `Analyzer` layer SHOULD include these metadata columns:

| Column | Position | Type | Description | Example |
|--------|----------|------|-------------|---------|
| `stock_code` | **Column 1** | string | 4-digit Taiwan stock code | `2330`, `1587` |
| `company_name` | **Column 2** | string | Company name | `台積電`, `日月光` |
| `file_type` | **Last -5** | string | Source data type identifier | `MOPS_MATRIX` |
| `source_file` | **Last -4** | string | Original filename processed | `mops_matrix_20260401_050438.csv` |
| `download_success` | **Last -3** | boolean/null | Whether download succeeded | `True`, `False`, or `null` |
| `download_timestamp` | **Last -2** | datetime/null | When data was downloaded | `2026-04-01 05:04:38` |
| `process_timestamp` | **Last -1** | datetime/null | When downloader processed the stock | `2026-04-01 05:04:38` |
| `stage1_process_timestamp` | **Last** | datetime | When Stage 1 pipeline ran | `2026-04-01 05:04:38` |

---

## raw_mops_matrix.csv (Financial Report Matrix)
**No:** 60
**Source:** `data/reports/mops_matrix_*.csv`
**Purpose:** Summarize which financial reports are available for each company and quarter.
**Note:** Current implementation in `MOPS` uses a wide-format matrix which may lack the standard metadata suffix in the raw output.

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

---

## [Proposed] raw_mops_financials.csv (Raw Financial Data)
**No:** 61
**Note:** Placeholder for future implementation where specific fields (Total Revenue, Net Income, etc.) are extracted.

### Column Definitions:
*TBD - To be aligned with TWSE official XML/XBRL fields.*
