"""
MOPS Sheets Uploader - PDF Scanner
Scans downloads directory and catalogs all PDF files with metadata extraction.
"""

import os
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from pathlib import Path
import logging

from .models import PDFFile, PDFMetadata, FutureQuarterAnalysis
from .config import MOPSConfig

logger = logging.getLogger(__name__)

class PDFScanner:
    """Scans downloads directory for MOPS PDF files"""
    
    def __init__(self, config: MOPSConfig):
        self.config = config
        self.downloads_dir = Path(config.downloads_dir)
        
    def scan_downloads_directory(self) -> Dict[str, List[PDFFile]]:
        """
        Scan downloads directory structure and return organized PDF files
        
        Expected structure:
        downloads/
        ├── 2330/
        │   ├── 202401_2330_AI1.pdf
        │   ├── 202402_2330_AI1.pdf
        │   └── metadata.json
        ├── 8272/
        │   ├── 202401_8272_A12.pdf
        │   └── metadata.json
        
        Returns:
            Dict mapping company_id to list of PDFFile objects
        """
        logger.info(f"🔍 掃描 PDF 檔案: {self.downloads_dir}")
        
        if not self.downloads_dir.exists():
            logger.error(f"Downloads directory not found: {self.downloads_dir}")
            return {}
        
        pdf_files_by_company = {}
        total_files = 0
        invalid_files = []
        future_analysis = FutureQuarterAnalysis()
        
        # Scan each company directory
        for company_dir in self.downloads_dir.iterdir():
            if not company_dir.is_dir():
                continue
                
            company_id = company_dir.name
            if not self._is_valid_company_id(company_id):
                logger.warning(f"⚠️ Invalid company directory name: {company_id}")
                continue
            
            company_pdfs = []
            
            # Scan PDF files in company directory
            for pdf_file in company_dir.glob("*.pdf"):
                try:
                    pdf_obj = self._process_pdf_file(pdf_file, company_id)
                    if pdf_obj:
                        company_pdfs.append(pdf_obj)
                        total_files += 1
                        
                        # Check for future quarters
                        if pdf_obj.is_future_quarter():
                            future_analysis.future_pdfs.append(pdf_obj)
                            months_ahead = self._calculate_months_ahead(pdf_obj)
                            if months_ahead > self.config.warn_threshold_months:
                                future_analysis.add_warning(
                                    f"PDF {pdf_obj.filename} is {months_ahead} months in the future"
                                )
                    else:
                        invalid_files.append(str(pdf_file))
                        
                except Exception as e:
                    logger.error(f"Error processing {pdf_file}: {e}")
                    invalid_files.append(str(pdf_file))
            
            if company_pdfs:
                # Sort by year and quarter
                company_pdfs.sort(key=lambda x: (x.year, x.quarter))
                pdf_files_by_company[company_id] = company_pdfs
        
        # Log summary
        companies_with_pdfs = len(pdf_files_by_company)
        discovered_quarters = self.discover_available_quarters(pdf_files_by_company)
        
        logger.info(f"   發現 {total_files} 個 PDF 檔案 ({companies_with_pdfs} 家公司)")
        if discovered_quarters:
            logger.info(f"   季度範圍: {min(discovered_quarters)} 到 {max(discovered_quarters)} ({len(discovered_quarters)} 個季度)")
        
        if invalid_files:
            logger.warning(f"⚠️ {len(invalid_files)} 個檔案無法解析: {invalid_files[:5]}...")
        
        # Handle future quarters
        if future_analysis.has_future_pdfs():
            self._log_future_quarter_warnings(future_analysis)
        
        return pdf_files_by_company
    
    def discover_available_quarters(self, pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """
        Discover all quarters that have at least one PDF file
        
        Returns:
            Sorted list of quarters (e.g., ['2025 Q4', '2025 Q3', '2025 Q1', '2024 Q4'])
        """
        quarters = set()
        
        for company_pdfs in pdf_data.values():
            for pdf in company_pdfs:
                quarters.add(pdf.quarter_key)
        
        # Sort chronologically
        def quarter_sort_key(q):
            year, quarter = q.split(' Q')
            return (int(year), int(quarter))
        
        return sorted(quarters, key=quarter_sort_key, reverse=True)
    
    def parse_pdf_filename(self, filename: str) -> Optional[PDFMetadata]:
        """
        Parse PDF filename to extract metadata with validation
        
        Pattern: YYYYQQ_COMPANYID_TYPE.pdf
        Examples:
        - 202401_2330_AI1.pdf → Year: 2024, Quarter: 1, Company: 2330, Type: AI1
        - 202403_8272_A12.pdf → Year: 2024, Quarter: 3, Company: 8272, Type: A12
        """
        return PDFMetadata.from_filename(filename)
    
    def validate_pdf_file(self, file_path: Path) -> bool:
        """Validate that file exists and is a valid PDF"""
        if not file_path.exists():
            return False
        
        if file_path.suffix.lower() != '.pdf':
            return False
        
        # Check file size (should be reasonable for a PDF)
        file_size = file_path.stat().st_size
        if file_size < 1024:  # Less than 1KB is suspicious
            logger.warning(f"Suspiciously small PDF file: {file_path} ({file_size} bytes)")
            return False
        
        # Basic PDF header check
        try:
            with open(file_path, 'rb') as f:
                header = f.read(4)
                if header != b'%PDF':
                    logger.warning(f"Invalid PDF header: {file_path}")
                    return False
        except Exception as e:
            logger.error(f"Error reading PDF file {file_path}: {e}")
            return False
        
        return True
    
    def _process_pdf_file(self, pdf_path: Path, expected_company_id: str) -> Optional[PDFFile]:
        """Process a single PDF file and create PDFFile object"""
        if not self.validate_pdf_file(pdf_path):
            return None
        
        # Parse filename
        metadata = self.parse_pdf_filename(pdf_path.name)
        if not metadata:
            logger.warning(f"Cannot parse filename: {pdf_path.name}")
            return None
        
        # Verify company ID matches directory
        if metadata.company_id != expected_company_id:
            logger.warning(f"Company ID mismatch in {pdf_path.name}: expected {expected_company_id}, got {metadata.company_id}")
            return None
        
        # Get file stats
        stat = pdf_path.stat()
        
        return PDFFile(
            company_id=metadata.company_id,
            year=metadata.year,
            quarter=metadata.quarter,
            report_type=metadata.report_type,
            filename=pdf_path.name,
            file_path=str(pdf_path),
            file_size=stat.st_size,
            modified_date=datetime.fromtimestamp(stat.st_mtime)
        )
    
    def _is_valid_company_id(self, company_id: str) -> bool:
        """Validate company ID format (4 digits)"""
        return bool(re.match(r'^\d{4}$', company_id))
    
    def _calculate_months_ahead(self, pdf_file: PDFFile) -> int:
        """Calculate how many months ahead this PDF is from current date"""
        now = datetime.now()
        current_year = now.year
        current_quarter = (now.month - 1) // 3 + 1
        
        # Convert quarters to months for easier calculation
        current_month = current_year * 12 + (current_quarter - 1) * 3
        pdf_month = pdf_file.year * 12 + (pdf_file.quarter - 1) * 3
        
        return max(0, (pdf_month - current_month) // 3)
    
    def _log_future_quarter_warnings(self, future_analysis: FutureQuarterAnalysis) -> None:
        """Log warnings about future quarter PDFs"""
        future_count = len(future_analysis.future_pdfs)
        logger.warning(f"⚠️ 發現 {future_count} 個未來季度PDF:")
        
        for pdf in future_analysis.future_pdfs[:5]:  # Show first 5
            months_ahead = self._calculate_months_ahead(pdf)
            logger.warning(f"   • {pdf.filename} ({pdf.quarter_key}, {months_ahead} 個月後)")
        
        if future_count > 5:
            logger.warning(f"   ... 還有 {future_count - 5} 個檔案")
        
        logger.warning("   這可能是檔案命名錯誤或系統時間問題")
        
        # Log detailed warnings
        for warning in future_analysis.warnings:
            logger.warning(f"   ⚠️ {warning}")

class PDFStatistics:
    """Helper class for PDF statistics and analysis"""
    
    @staticmethod
    def analyze_report_types(pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, int]:
        """Analyze distribution of report types"""
        type_counts = {}
        
        for company_pdfs in pdf_data.values():
            for pdf in company_pdfs:
                type_counts[pdf.report_type] = type_counts.get(pdf.report_type, 0) + 1
        
        return type_counts
    
    @staticmethod
    def find_companies_without_pdfs(pdf_data: Dict[str, List[PDFFile]], 
                                   all_companies: List[str]) -> List[str]:
        """Find companies that have no PDF files"""
        companies_with_pdfs = set(pdf_data.keys())
        all_companies_set = set(all_companies)
        
        return list(all_companies_set - companies_with_pdfs)
    
    @staticmethod
    def analyze_quarter_coverage(pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, List[str]]:
        """Analyze which quarters are available for each company"""
        coverage = {}
        
        for company_id, pdfs in pdf_data.items():
            quarters = [pdf.quarter_key for pdf in pdfs]
            coverage[company_id] = sorted(quarters, reverse=True)
        
        return coverage
    
    @staticmethod
    def find_missing_quarters(pdf_data: Dict[str, List[PDFFile]], 
                            expected_quarters: List[str]) -> Dict[str, List[str]]:
        """Find missing quarters for each company"""
        missing = {}
        expected_set = set(expected_quarters)
        
        for company_id, pdfs in pdf_data.items():
            available_quarters = {pdf.quarter_key for pdf in pdfs}
            missing_quarters = list(expected_set - available_quarters)
            if missing_quarters:
                missing[company_id] = sorted(missing_quarters, reverse=True)
        
        return missing