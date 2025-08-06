"""Enhanced parser with flexible criteria and user-friendly options."""

import re
from typing import List, Optional, Dict
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from ..models import ReportMetadata
from ..exceptions import ParsingError
from ..config import (TARGET_INDIVIDUAL_REPORTS, EXCLUDED_KEYWORDS, BASE_URL,
                     TARGET_FILE_PATTERNS, FLEXIBLE_TARGETS)
import logging

class DocumentParser:
    """Enhanced parser with flexible target criteria."""
    
    def __init__(self, strict_mode: bool = False):
        self.logger = logging.getLogger("mops_downloader.parser")
        self.strict_mode = strict_mode
    
    def _is_target_report(self, description: str, filename: str = "") -> tuple[bool, str]:
        """Check if report matches target criteria with reason."""
        if not description:
            return False, "Empty description"
        
        # Check for explicitly excluded content first
        for keyword in EXCLUDED_KEYWORDS:
            if keyword in description or keyword in filename:
                return False, f"Excluded by keyword: {keyword}"
        
        # Check primary targets
        for target in TARGET_INDIVIDUAL_REPORTS:
            if target in description:
                return True, f"Matched primary target: {target}"
        
        # Check file pattern targets
        for pattern in TARGET_FILE_PATTERNS:
            if re.search(pattern, filename):
                return True, f"Matched file pattern: {pattern}"
        
        # In non-strict mode, check flexible targets
        if not self.strict_mode:
            for target in FLEXIBLE_TARGETS:
                if target in description:
                    return True, f"Matched flexible target: {target}"
        
        return False, "No matching criteria"
    
    def _analyze_all_reports(self, soup) -> Dict[str, List[tuple]]:
        """Analyze all available reports with filenames."""
        all_reports = {
            'targets': [],          # Reports matching our criteria
            'consolidated': [],     # Consolidated reports
            'english': [],         # English reports  
            'other': []            # Other reports
        }
        
        # Find all table cells with report descriptions
        tables = soup.find_all('table')
        for table in tables:
            rows = table.find_all('tr')
            for row in rows[1:]:  # Skip header
                cells = row.find_all('td')
                if len(cells) < 6:
                    continue
                
                # Extract description and filename
                description = ""
                filename = ""
                
                for cell in cells:
                    text = cell.get_text(strip=True)
                    if 'IFRSs' in text or '財務報告' in text:
                        description = text
                    
                    # Look for PDF filename
                    link = cell.find('a')
                    if link:
                        href = link.get('href', '')
                        if '.pdf' in href:
                            # Extract filename from JavaScript or href
                            pdf_match = re.search(r'([^/,\"]+\.pdf)', href)
                            if pdf_match:
                                filename = pdf_match.group(1)
                
                if description and filename:
                    is_target, reason = self._is_target_report(description, filename)
                    
                    if is_target:
                        all_reports['targets'].append((description, filename, reason))
                    elif '英文版' in description:
                        all_reports['english'].append((description, filename))
                    elif '合併' in description:
                        all_reports['consolidated'].append((description, filename))
                    else:
                        all_reports['other'].append((description, filename))
        
        return all_reports
    
    def _extract_quarter_from_year_column(self, year_text: str) -> Optional[int]:
        """Extract quarter number from year column text."""
        if not year_text:
            return None
        
        quarter_patterns = {
            '第一季': 1, '第二季': 2, '第三季': 3, '第四季': 4,
            '第1季': 1, '第2季': 2, '第3季': 3, '第4季': 4
        }
        
        for pattern, quarter in quarter_patterns.items():
            if pattern in year_text:
                return quarter
        
        return None
    
    def _extract_file_info(self, file_cell) -> Optional[tuple]:
        """Extract filename and URL from file cell."""
        if not file_cell:
            return None
        
        link = file_cell.find('a')
        if not link:
            return None
        
        # Get filename from link text or href
        filename = link.get_text(strip=True)
        href = link.get('href', '')
        
        if not filename or not filename.endswith('.pdf'):
            # Try to extract from href
            if '.pdf' in href:
                pdf_match = re.search(r'([^/,\"]+\.pdf)', href)
                if pdf_match:
                    filename = pdf_match.group(1)
                else:
                    return None
            else:
                return None
        
        # Handle JavaScript download URLs
        if href.startswith('javascript:'):
            # Extract parameters from readfile2 call
            match = re.search(r'readfile2\("([^"]+)","([^"]+)","([^"]+)"\)', href)
            if match:
                kind, co_id, pdf_filename = match.groups()
                download_url = f"{BASE_URL}?step=9&kind={kind}&co_id={co_id}&filename={pdf_filename}"
            else:
                download_url = href
        else:
            # Regular href
            if href.startswith('/'):
                download_url = urljoin(BASE_URL, href)
            elif href.startswith('http'):
                download_url = href
            else:
                download_url = urljoin(BASE_URL, href)
        
        return filename, download_url
    
    def _parse_file_size(self, size_text: str) -> int:
        """Parse file size from text."""
        if not size_text:
            return 0
        
        size_clean = re.sub(r'[,\s]', '', size_text.strip())
        try:
            return int(size_clean)
        except ValueError:
            return 0
    
    def _extract_report_type(self, filename: str) -> str:
        """Extract report type code from filename."""
        match = re.search(r'_([A-Z0-9]+)\.pdf$', filename)
        return match.group(1) if match else 'UNKNOWN'
    
    def parse_reports(self, html_content: str, company_id: str, year: int) -> List[ReportMetadata]:
        """Parse HTML content with flexible criteria."""
        if not html_content:
            raise ParsingError("Empty HTML content")
        
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
        except Exception as e:
            raise ParsingError(f"Failed to parse HTML: {e}")
        
        # Analyze all available reports first
        all_reports = self._analyze_all_reports(soup)
        
        self.logger.info("📊 Report Analysis:")
        if all_reports['targets']:
            self.logger.info(f"   ✅ Target reports found: {len(all_reports['targets'])}")
            for desc, filename, reason in all_reports['targets']:
                self.logger.info(f"      • {desc} → {filename} ({reason})")
        else:
            self.logger.warning(f"   ❌ No target reports found")
        
        if all_reports['consolidated']:
            self.logger.info(f"   📋 Consolidated reports: {len(all_reports['consolidated'])}")
            for desc, filename in all_reports['consolidated'][:2]:  # Show first 2
                self.logger.info(f"      • {desc} → {filename}")
        
        if all_reports['english']:
            self.logger.info(f"   🌍 English reports: {len(all_reports['english'])}")
            for desc, filename in all_reports['english'][:2]:  # Show first 2
                self.logger.info(f"      • {desc} → {filename}")
        
        # Find the main data table and parse
        tables = soup.find_all('table')
        if not tables:
            self.logger.warning("No tables found in HTML content")
            return []
        
        reports = []
        
        for table in tables:
            rows = table.find_all('tr')
            if len(rows) < 2:
                continue
            
            header_row = rows[0]
            header_cells = header_row.find_all(['th', 'td'])
            header_text = ' '.join([cell.get_text(strip=True) for cell in header_cells])
            
            if '證券代號' not in header_text or '電子檔案' not in header_text:
                continue
            
            self.logger.info(f"Found reports table with {len(rows)-1} data rows")
            
            # Parse data rows
            for row in rows[1:]:
                cells = row.find_all('td')
                if len(cells) < 8:
                    continue
                
                try:
                    stock_code = cells[0].get_text(strip=True)
                    year_text = cells[1].get_text(strip=True)
                    
                    # Find description and file cells
                    description_text = ""
                    file_cell = None
                    size_text = ""
                    date_text = ""
                    
                    for i, cell in enumerate(cells):
                        cell_text = cell.get_text(strip=True)
                        
                        if 'IFRSs' in cell_text or '財務報告' in cell_text:
                            description_text = cell_text
                        
                        if cell.find('a') and ('.pdf' in str(cell) or 'readfile' in str(cell)):
                            file_cell = cell
                        
                        if re.match(r'^[\d,]+$', cell_text):
                            size_text = cell_text
                        
                        if '/' in cell_text and len(cell_text) > 8:
                            date_text = cell_text
                    
                    # Extract file info first to get filename
                    file_info = self._extract_file_info(file_cell)
                    if not file_info:
                        continue
                    
                    filename, download_url = file_info
                    
                    # Check if this is a target report (with filename info)
                    is_target, reason = self._is_target_report(description_text, filename)
                    
                    if not is_target:
                        self.logger.debug(f"Skipping: {description_text} → {filename} ({reason})")
                        continue
                    
                    # Extract quarter
                    quarter = self._extract_quarter_from_year_column(year_text)
                    if not quarter:
                        self.logger.warning(f"Could not extract quarter from: {year_text}")
                        continue
                    
                    # Create report metadata
                    report = ReportMetadata(
                        quarter=quarter,
                        filename=filename,
                        download_url=download_url,
                        file_size=self._parse_file_size(size_text),
                        upload_date=date_text,
                        company_id=company_id,
                        year=year,
                        report_type=self._extract_report_type(filename)
                    )
                    
                    reports.append(report)
                    self.logger.info(f"✅ Target report: Q{quarter} - {filename} ({reason})")
                    
                except Exception as e:
                    self.logger.warning(f"Error parsing row: {e}")
                    continue
        
        self.logger.info(f"Extracted {len(reports)} target reports")
        return reports
