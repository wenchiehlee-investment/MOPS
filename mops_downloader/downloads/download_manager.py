"""PDF download manager with simple, working two-step process."""

import urllib.request
import urllib.parse
import urllib.error
import ssl
import time
from pathlib import Path
from typing import List, Dict, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from ..models import ReportMetadata
from ..exceptions import DownloadError
from ..config import TIMEOUT, CHUNK_SIZE, MAX_CONCURRENT_DOWNLOADS, RATE_LIMIT_DELAY
import logging

class DownloadManager:
    """Simple, working PDF download manager for MOPS."""
    
    def __init__(self, verify_ssl: bool = False):
        self.logger = logging.getLogger("mops_downloader.download")
        
        # Create SSL context
        if verify_ssl:
            self.ssl_context = ssl.create_default_context()
        else:
            self.ssl_context = ssl.create_default_context()
            self.ssl_context.check_hostname = False
            self.ssl_context.verify_mode = ssl.CERT_NONE
            self.logger.info("SSL verification disabled for PDF downloads")
    
    def _extract_pdf_url_from_html(self, html_content: str, filename: str) -> Optional[str]:
        """Extract PDF URL from HTML using simple string matching."""
        try:
            # Decode if needed
            if isinstance(html_content, bytes):
                try:
                    html_content = html_content.decode('big5')
                except:
                    html_content = html_content.decode('utf-8', errors='ignore')
            
            self.logger.debug(f"Looking for PDF URL in HTML content...")
            
            # Simple approach: look for /pdf/ pattern in the HTML
            # Based on debug output: href='/pdf/202401_2330_AI1_20250805_120341.pdf'
            
            lines = html_content.split('\n')
            for line in lines:
                if '/pdf/' in line and '.pdf' in line:
                    # Found a line with PDF path
                    
                    # Extract between quotes - simple approach
                    if "href='" in line:
                        start = line.find("href='") + 6
                        end = line.find("'", start)
                        if start > 5 and end > start:
                            pdf_path = line[start:end]
                            if pdf_path.startswith('/pdf/') and pdf_path.endswith('.pdf'):
                                pdf_url = f"https://doc.twse.com.tw{pdf_path}"
                                self.logger.info(f"Found PDF URL: {pdf_url}")
                                return pdf_url
                    
                    elif 'href="' in line:
                        start = line.find('href="') + 6
                        end = line.find('"', start)
                        if start > 5 and end > start:
                            pdf_path = line[start:end]
                            if pdf_path.startswith('/pdf/') and pdf_path.endswith('.pdf'):
                                pdf_url = f"https://doc.twse.com.tw{pdf_path}"
                                self.logger.info(f"Found PDF URL: {pdf_url}")
                                return pdf_url
            
            # Fallback: look for any /pdf/ path
            if '/pdf/' in html_content:
                start_pos = html_content.find('/pdf/')
                # Find the end of the PDF filename
                end_pos = html_content.find('.pdf', start_pos)
                if end_pos > start_pos:
                    pdf_path = html_content[start_pos:end_pos + 4]  # +4 for '.pdf'
                    # Clean up any surrounding characters
                    if pdf_path.startswith('/pdf/') and pdf_path.endswith('.pdf'):
                        pdf_url = f"https://doc.twse.com.tw{pdf_path}"
                        self.logger.info(f"Found PDF URL (fallback): {pdf_url}")
                        return pdf_url
            
            self.logger.error("Could not find PDF URL in HTML response")
            self.logger.debug(f"HTML content (first 500 chars): {html_content[:500]}")
            return None
            
        except Exception as e:
            self.logger.error(f"Error extracting PDF URL: {e}")
            return None
    
    def _download_single_file(self, report: ReportMetadata, output_path: Path) -> Dict:
        """Download a single PDF file using two-step process."""
        try:
            self.logger.info(f"Downloading {report.filename}...")
            
            # Create output directory
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Check if file already exists and is reasonable size
            if output_path.exists():
                existing_size = output_path.stat().st_size
                if existing_size > 100000:  # More than 100KB = probably valid
                    self.logger.info(f"File already exists: {report.filename} ({existing_size:,} bytes)")
                    return {
                        'success': True,
                        'filename': report.filename,
                        'path': str(output_path),
                        'size': existing_size,
                        'skipped': True
                    }
            
            # STEP 1: Get HTML page with PDF link
            self.logger.debug(f"Step 1: Getting HTML page from {report.download_url}")
            
            request1 = urllib.request.Request(
                report.download_url,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                    'Referer': 'https://doc.twse.com.tw/server-java/t57sb01'
                }
            )
            
            with urllib.request.urlopen(request1, timeout=TIMEOUT, context=self.ssl_context) as response1:
                html_content = response1.read()
                
                # Decode
                try:
                    html_text = html_content.decode('big5')
                except:
                    html_text = html_content.decode('utf-8', errors='ignore')
                
                self.logger.debug(f"Got HTML response: {len(html_text)} characters")
            
            # STEP 2: Extract PDF URL
            pdf_url = self._extract_pdf_url_from_html(html_text, report.filename)
            
            if not pdf_url:
                raise DownloadError(f"Could not find PDF URL for {report.filename}")
            
            # STEP 3: Download actual PDF
            self.logger.debug(f"Step 2: Downloading PDF from {pdf_url}")
            
            request2 = urllib.request.Request(
                pdf_url,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                    'Accept': 'application/pdf,*/*',
                    'Referer': report.download_url
                }
            )
            
            # Download PDF
            total_size = 0
            with urllib.request.urlopen(request2, timeout=TIMEOUT, context=self.ssl_context) as response2:
                with open(output_path, 'wb') as f:
                    while True:
                        chunk = response2.read(CHUNK_SIZE)
                        if not chunk:
                            break
                        f.write(chunk)
                        total_size += len(chunk)
            
            # Verify download
            if total_size == 0:
                raise DownloadError(f"Downloaded file is empty: {report.filename}")
            
            if total_size < 1000:
                self.logger.warning(f"Downloaded file is very small ({total_size} bytes)")
            
            # Check if it's actually a PDF
            with open(output_path, 'rb') as f:
                header = f.read(4)
                if not header.startswith(b'%PDF'):
                    self.logger.warning(f"File doesn't appear to be a valid PDF")
            
            self.logger.info(f"Successfully downloaded {report.filename} ({total_size:,} bytes)")
            
            return {
                'success': True,
                'filename': report.filename,
                'path': str(output_path),
                'size': total_size,
                'skipped': False
            }
            
        except Exception as e:
            self.logger.error(f"Failed to download {report.filename}: {e}")
            
            # Clean up partial download
            if output_path.exists():
                try:
                    output_path.unlink()
                except:
                    pass
            
            return {
                'success': False,
                'filename': report.filename,
                'error': str(e),
                'skipped': False
            }
    
    def download_files(self, reports: List[ReportMetadata], base_dir: Path) -> List[Dict]:
        """Download multiple PDF files."""
        if not reports:
            self.logger.info("No files to download")
            return []
        
        self.logger.info(f"Starting download of {len(reports)} files")
        
        results = []
        
        # Simple sequential download to avoid complications
        for report in reports:
            # FLATTENED STRUCTURE: downloads/{company_id}/{filename}.pdf
            company_dir = base_dir / report.company_id
            output_path = company_dir / report.filename
            
            self.logger.debug(f"Output path: {output_path}")
            
            # Download file
            result = self._download_single_file(report, output_path)
            results.append(result)
            
            # Rate limiting
            time.sleep(RATE_LIMIT_DELAY)
        
        # Summary
        successful = sum(1 for r in results if r['success'])
        skipped = sum(1 for r in results if r.get('skipped', False))
        failed = len(results) - successful
        
        self.logger.info(f"Download complete: {successful} successful, {skipped} skipped, {failed} failed")
        
        return results
