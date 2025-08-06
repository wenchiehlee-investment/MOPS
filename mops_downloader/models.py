"""Data models for MOPS downloader."""

from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime

@dataclass
class ReportMetadata:
    """Metadata for a financial report."""
    quarter: int
    filename: str
    download_url: str
    file_size: int
    upload_date: str
    company_id: str
    year: int
    report_type: str  # A12, A13, etc.

@dataclass
class ValidatedParams:
    """Validated input parameters."""
    company_id: str
    western_year: int
    roc_year: int
    quarters: List[int]  # [1,2,3,4] or specific quarters

@dataclass
class DownloadResult:
    """Result of download operation."""
    success: bool
    downloaded_files: List[str]
    missing_quarters: List[int]
    error_details: Optional[str]
    file_paths: List[str]
    total_files: int
    total_size: int
