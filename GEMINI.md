# Project Overview: MOPS Data Downloader and Sheets Uploader

This project consists of a suite of Python tools designed to automate the collection, processing, and visualization of financial data from the Taiwan Market Observation Post System (MOPS). It is divided into two primary modules: `mops_downloader` and `mops_sheets_uploader`.

## `mops_downloader` Module

The `mops_downloader` module is responsible for systematically downloading financial reports (primarily PDF files) from the MOPS website.

### Key Functionality:
- **Report Retrieval**: Downloads quarterly financial reports based on specified company IDs, years, and quarters (including an option for all quarters).
- **Input Validation**: Ensures that provided company IDs, years, and quarters are valid.
- **Web Navigation**: Handles the interaction with the MOPS website to locate and fetch report pages.
- **Document Parsing**: Extracts relevant information and links from HTML content to identify target PDF reports.
- **Download Management**: Manages the actual downloading of PDF files, with options to skip already existing files.
- **File Organization**: Organizes downloaded files into a structured directory (e.g., `./downloads/<company_id>/`).
- **Error Handling**: Provides robust error handling for network issues, validation failures, and unexpected errors.

### CLI Usage Example:
```bash
python -m mops_downloader --company_id 2330 --year 2024 --quarter all
python -m mops_downloader --company_id 8272 --year 2023 --quarter 2 --output ./my_reports --only-missing-files
```

## `mops_sheets_uploader` Module

The `mops_sheets_uploader` module processes the downloaded PDF reports, builds a data matrix, and can either upload this matrix to Google Sheets or export it as a CSV file. It includes extensive features for customization and analysis of the data.

### Key Functionality:
- **PDF Scanning & Analysis**: Scans the downloaded PDF files to extract relevant data. Includes enhanced analysis for report types and coverage.
- **Stock Data Loading**: Loads stock information from a CSV file (e.g., `StockID_TWSE_TPEX.csv`) and can detect changes in the stock list.
- **Matrix Building**: Constructs a comprehensive data matrix from the scanned PDF data and stock information.
- **Google Sheets Integration**: Uploads the generated data matrix to a specified Google Sheet, with options for styling and formatting.
- **CSV Export**: Provides an option to export the data matrix to a CSV file.
- **Coverage Analysis**: Analyzes the completeness and type distribution of the downloaded reports.
- **Flexible Configuration**: Supports loading configurations from files and overriding them via command-line arguments.
- **Font & Formatting Options**: Offers various font presets and granular control over font sizes, bolding for headers and company information in Google Sheets.
- **Multiple Report Type Display**: Allows for displaying single or multiple report types found for a given quarter, with customizable separators.
- **Detailed Logging**: Provides extensive logging for tracking the processing steps and results.
- **Connection Testing**: Includes a utility to test the connection to Google Sheets.

### CLI Usage Example:
```bash
# Upload to Google Sheets with a large font preset
python -m mops_sheets_uploader --upload --font-preset large

# Export to CSV only, with specific font sizes
python -m mops_sheets_uploader --csv-only --font-size 16 --header-font-size 18

# Analyze report coverage and output to JSON
python -m mops_sheets_uploader --analyze --show-all-types --output analysis_report.json

# Test Google Sheets connection
python -m mops_sheets_uploader --test
```

### Environment Variables:
- `GOOGLE_SHEETS_CREDENTIALS`: Path to your Google Sheets service account credentials JSON file.
- `GOOGLE_SHEET_ID`: The ID of the Google Sheet to upload data to.

This `GEMINI.md` provides a high-level overview of the project's capabilities. For detailed usage and configuration, refer to the specific module documentation and CLI help messages.
