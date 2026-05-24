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

---

## [Proposed] raw_mops_financials.csv (Raw Financial Data)
**No:** 61
**Note:** Placeholder for future implementation where specific fields (Total Revenue, Net Income, etc.) are extracted.
