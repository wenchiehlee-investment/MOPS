"""
MOPS Downloader - Taiwan Financial Reports Downloader
Downloads quarterly IFRSs individual financial reports from MOPS system.
"""

__version__ = "1.0.0"
__author__ = "MOPS Downloader Team"

from .main import MOPSDownloader
from .models import DownloadResult, ReportMetadata
from .exceptions import ValidationError, DownloadError, ParsingError

__all__ = [
    'MOPSDownloader',
    'DownloadResult', 
    'ReportMetadata',
    'ValidationError',
    'DownloadError', 
    'ParsingError'
]
