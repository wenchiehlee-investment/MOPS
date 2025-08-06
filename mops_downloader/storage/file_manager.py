"""File organization with flattened directory structure."""

import json
from pathlib import Path
from typing import List, Dict, Any
from datetime import datetime
from ..models import ReportMetadata, DownloadResult
from ..config import DOWNLOAD_DIR
import logging

class FileManager:
    """Manages file organization with flattened directory structure."""
    
    def __init__(self, base_dir: Path = None):
        self.logger = logging.getLogger("mops_downloader.storage")
        self.base_dir = Path(base_dir) if base_dir else DOWNLOAD_DIR
        self.base_dir.mkdir(exist_ok=True)
    
    def organize_files(self, download_results: List[Dict], company_id: str, year: int) -> DownloadResult:
        """Organize downloaded files with flattened structure."""
        
        # Separate successful and failed downloads
        successful_files = [r for r in download_results if r['success']]
        failed_files = [r for r in download_results if not r['success']]
        
        # Extract file information
        downloaded_files = [r['filename'] for r in successful_files]
        file_paths = [r['path'] for r in successful_files]
        total_size = sum(r.get('size', 0) for r in successful_files)
        
        # Determine missing quarters
        downloaded_quarters = set()
        for filename in downloaded_files:
            # Extract quarter from filename (e.g., 202401_2330_AI1.pdf -> Q1)
            if len(filename) >= 6 and filename[4:6].isdigit():
                quarter_num = int(filename[5])  # 5th character is quarter
                if 1 <= quarter_num <= 4:
                    downloaded_quarters.add(quarter_num)
        
        all_quarters = {1, 2, 3, 4}
        missing_quarters = sorted(list(all_quarters - downloaded_quarters))
        
        # Create metadata file in company directory (not year subdirectory)
        self._create_metadata_file(company_id, year, download_results)
        
        # Determine overall success
        success = len(successful_files) > 0
        error_details = None
        if failed_files:
            error_details = f"{len(failed_files)} files failed to download"
        
        result = DownloadResult(
            success=success,
            downloaded_files=downloaded_files,
            missing_quarters=missing_quarters,
            error_details=error_details,
            file_paths=file_paths,
            total_files=len(successful_files),
            total_size=total_size
        )
        
        self.logger.info(f"File organization complete: {result.total_files} files, "
                        f"{result.total_size:,} bytes total")
        
        return result
    
    def _create_metadata_file(self, company_id: str, year: int, download_results: List[Dict]):
        """Create metadata file in flattened company directory."""
        # Metadata goes directly in company directory
        company_dir = self.base_dir / company_id
        company_dir.mkdir(exist_ok=True)
        
        metadata_file = company_dir / "metadata.json"
        
        # Load existing metadata or create new
        if metadata_file.exists():
            try:
                with open(metadata_file, 'r', encoding='utf-8') as f:
                    metadata = json.load(f)
            except:
                metadata = {"downloads": {}}
        else:
            metadata = {"downloads": {}}
        
        # Add current session data
        session_key = f"{year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        metadata["downloads"][session_key] = {
            "timestamp": datetime.now().isoformat(),
            "company_id": company_id,
            "year": year,
            "results": download_results
        }
        
        # Update last_updated
        metadata["last_updated"] = datetime.now().isoformat()
        metadata["company_id"] = company_id
        
        # Save metadata
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)
        
        self.logger.debug(f"Updated metadata file: {metadata_file}")
    
    def get_download_history(self, company_id: str) -> Dict[str, Any]:
        """Get download history for a company."""
        metadata_file = self.base_dir / company_id / "metadata.json"
        
        if not metadata_file.exists():
            return {"downloads": {}}
        
        try:
            with open(metadata_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.logger.error(f"Failed to read metadata file: {e}")
            return {"downloads": {}}
    
    def cleanup_failed_downloads(self, company_id: str, year: int = None):
        """Clean up any partial or failed download files."""
        # With flattened structure, clean up directly in company directory
        company_dir = self.base_dir / company_id
        
        if not company_dir.exists():
            return
        
        cleaned_files = 0
        for pdf_file in company_dir.glob("*.pdf"):
            if pdf_file.stat().st_size == 0:
                pdf_file.unlink()
                cleaned_files += 1
                self.logger.info(f"Cleaned up empty file: {pdf_file.name}")
        
        if cleaned_files > 0:
            self.logger.info(f"Cleaned up {cleaned_files} empty files")
