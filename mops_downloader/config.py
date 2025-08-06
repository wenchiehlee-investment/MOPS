"""Configuration settings for MOPS downloader with expanded criteria."""

import os
from pathlib import Path

# MOPS System Configuration
BASE_URL = "https://doc.twse.com.tw/server-java/t57sb01"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

# Download Configuration
DOWNLOAD_DIR = Path("./downloads")
LOG_DIR = Path("./logs") 
MAX_RETRIES = 3
TIMEOUT = 30
RATE_LIMIT_DELAY = 1.0  # seconds between requests

# File Settings
CHUNK_SIZE = 8192  # bytes for file download
MAX_CONCURRENT_DOWNLOADS = 3

# Year Validation
MIN_YEAR = 1912  # ROC establishment
MAX_YEAR = 2030  # reasonable future limit

# Logging Configuration
LOG_FORMAT = '[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# Target Report Criteria - EXPANDED based on real MOPS data
TARGET_INDIVIDUAL_REPORTS = [
    "IFRSs個別財報",     # Original target (company 8272 style)
    "IFRSs個體財報",     # Original target (company 2330 old style) 
    "IFRSs合併財報"      # EXPANDED: Might be individual for some companies (company 2330 current)
]

# File type patterns that indicate individual reports
TARGET_FILE_PATTERNS = [
    r"A12\.pdf$",       # Pattern from company 8272
    r"A13\.pdf$",       # Pattern from company 2330 (old)
    r"AI1\.pdf$",       # Pattern from company 2330 (current) - NEW
    r"A1[0-9]\.pdf$"    # Generic individual report pattern
]

# EXCLUDED only for English and clearly non-individual types
EXCLUDED_KEYWORDS = [
    "英文版",            # English versions
    "AIA\.pdf",         # English consolidated (from real data)
    "AE2\.pdf"          # English parent-subsidiary (from real data)
]

# FLEXIBLE mode - when no strict targets found, include reasonable alternatives
FLEXIBLE_TARGETS = [
    "IFRSs合併財報",     # Consolidated reports (might be individual for some companies)
    "財務報告書"         # General financial reports
]
