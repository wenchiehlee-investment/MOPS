"""Main MOPS downloader orchestration."""

from pathlib import Path
from typing import Union, List
from .models import DownloadResult, ValidatedParams
from .validators.input_validator import InputValidator
from .web.navigator import WebNavigator
from .parsers.document_parser import DocumentParser
from .downloads.download_manager import DownloadManager
from .storage.file_manager import FileManager
from .utils.logging_config import setup_logging
from .config import DOWNLOAD_DIR
import logging

class MOPSDownloader:
    """Main class for downloading MOPS financial reports."""
    
    def __init__(self, download_dir: Path = None, log_level: str = "INFO"):
        """Initialize MOPS downloader."""
        self.download_dir = Path(download_dir) if download_dir else DOWNLOAD_DIR
        self.download_dir.mkdir(exist_ok=True)
        
        # Set up logging
        self.logger = setup_logging(log_level, log_to_file=True)
        self.logger.info(f"MOPS Downloader initialized, download directory: {self.download_dir}")
        
        # Initialize components
        self.validator = InputValidator()
        self.navigator = WebNavigator()
        self.parser = DocumentParser()
        self.download_manager = DownloadManager()
        self.file_manager = FileManager(self.download_dir)
    
    def download(self, company_id: str, year: int, quarter: Union[str, int] = "all") -> DownloadResult:
        """
        Download MOPS financial reports.
        
        Args:
            company_id: Taiwan stock company ID (e.g., "2330")
            year: Western year (e.g., 2024)
            quarter: Quarter (1-4) or "all" for all quarters
            
        Returns:
            DownloadResult with success status and file information
        """
        try:
            self.logger.info(f"Starting download for company {company_id}, year {year}, quarter {quarter}")
            
            # Stage 1: Input Validation & Configuration
            self.logger.info("Stage 1: Validating inputs...")
            params = self.validator.validate_and_convert(company_id, year, quarter)
            self.logger.info(f"Validated parameters: Company {params.company_id}, "
                           f"Year {params.western_year} (ROC {params.roc_year}), "
                           f"Quarters {params.quarters}")
            
            # Stage 2: Web Navigation & Authentication
            self.logger.info("Stage 2: Fetching reports pages...")
            html_contents = self.navigator.fetch_all_quarters(params)
            
            successful_fetches = sum(1 for content in html_contents.values() if content)
            self.logger.info(f"Successfully fetched {successful_fetches}/{len(params.quarters)} quarter pages")
            
            if successful_fetches == 0:
                return DownloadResult(
                    success=False,
                    downloaded_files=[],
                    missing_quarters=params.quarters,
                    error_details="Failed to fetch any reports pages",
                    file_paths=[],
                    total_files=0,
                    total_size=0
                )
            
            # Stage 3: Document Discovery & Filtering
            self.logger.info("Stage 3: Parsing and filtering reports...")
            all_reports = []
            
            for quarter, html_content in html_contents.items():
                if html_content:
                    reports = self.parser.parse_reports(html_content, params.company_id, params.western_year)
                    all_reports.extend(reports)
                    self.logger.info(f"Found {len(reports)} target reports for Q{quarter}")
            
            if not all_reports:
                return DownloadResult(
                    success=False,
                    downloaded_files=[],
                    missing_quarters=params.quarters,
                    error_details="No individual financial reports found for specified criteria",
                    file_paths=[],
                    total_files=0,
                    total_size=0
                )
            
            self.logger.info(f"Total target reports found: {len(all_reports)}")
            
            # Stage 4: Download & File Management
            self.logger.info("Stage 4: Downloading PDF files...")
            download_results = self.download_manager.download_files(all_reports, self.download_dir)
            
            # Stage 5: Verification & Cleanup
            self.logger.info("Stage 5: Organizing files and creating metadata...")
            
            # Clean up failed downloads
            self.file_manager.cleanup_failed_downloads(params.company_id, params.western_year)
            
            # Organize files and create final result
            final_result = self.file_manager.organize_files(download_results, params.company_id, params.western_year)
            
            # Log final summary
            if final_result.success:
                self.logger.info(f"Download completed successfully: {final_result.total_files} files, "
                               f"{final_result.total_size:,} bytes")
                if final_result.missing_quarters:
                    self.logger.info(f"Missing quarters (no individual reports available): "
                                   f"Q{', Q'.join(map(str, final_result.missing_quarters))}")
            else:
                self.logger.error(f"Download failed: {final_result.error_details}")
            
            return final_result
            
        except Exception as e:
            self.logger.error(f"Unexpected error during download: {e}", exc_info=True)
            return DownloadResult(
                success=False,
                downloaded_files=[],
                missing_quarters=params.quarters if 'params' in locals() else [1, 2, 3, 4],
                error_details=f"Unexpected error: {str(e)}",
                file_paths=[],
                total_files=0,
                total_size=0
            )
