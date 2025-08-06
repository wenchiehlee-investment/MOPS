"""Custom exceptions for MOPS downloader."""

class MOPSDownloaderError(Exception):
    """Base exception for MOPS downloader."""
    pass

class ValidationError(MOPSDownloaderError):
    """Raised when input validation fails."""
    pass

class DownloadError(MOPSDownloaderError):
    """Raised when download operations fail."""
    pass

class ParsingError(MOPSDownloaderError):
    """Raised when HTML parsing fails."""
    pass

class NetworkError(MOPSDownloaderError):
    """Raised when network operations fail."""
    pass
