"""Web navigation with urllib and proper Big5 encoding support."""

import urllib.request
import urllib.parse
import urllib.error
import ssl
import time
from typing import Dict, List
from ..models import ValidatedParams
from ..exceptions import NetworkError
from ..config import BASE_URL, USER_AGENT, TIMEOUT, MAX_RETRIES, RATE_LIMIT_DELAY
import logging

class WebNavigator:
    """Handles web navigation using urllib with Big5 encoding support."""
    
    def __init__(self, verify_ssl: bool = False):
        self.logger = logging.getLogger("mops_downloader.web")
        
        # Create SSL context
        if verify_ssl:
            self.ssl_context = ssl.create_default_context()
        else:
            self.ssl_context = ssl.create_default_context()
            self.ssl_context.check_hostname = False
            self.ssl_context.verify_mode = ssl.CERT_NONE
            self.logger.info("SSL verification disabled for MOPS website compatibility")
    
    def _build_url(self, company_id: str, roc_year: int, quarter: int = None) -> str:
        """Build MOPS URL with parameters."""
        params = {
            'step': '1',
            'colorchg': '1',
            'seamon': str(quarter) if quarter else '',
            'mtype': 'A',
            'co_id': company_id,
            'year': str(roc_year)
        }
        
        # Build query string using urllib
        query_string = urllib.parse.urlencode(params)
        url = f"{BASE_URL}?{query_string}"
        self.logger.debug(f"Built URL: {url}")
        return url
    
    def _decode_content(self, content_bytes: bytes) -> str:
        """Decode content with proper encoding detection."""
        # Try encodings in order of likelihood for Taiwan sites
        encodings = ['big5', 'utf-8', 'gb2312', 'latin-1']
        
        for encoding in encodings:
            try:
                content = content_bytes.decode(encoding)
                self.logger.debug(f"Successfully decoded with {encoding}")
                return content
            except UnicodeDecodeError:
                continue
        
        # Fallback to latin-1 (preserves all bytes)
        content = content_bytes.decode('latin-1')
        self.logger.warning("Using latin-1 encoding fallback")
        return content
    
    def _fetch_with_retry(self, url: str) -> str:
        """Fetch URL using urllib with retry logic."""
        last_exception = None
        
        for attempt in range(MAX_RETRIES):
            try:
                self.logger.info(f"Fetching URL (attempt {attempt + 1}/{MAX_RETRIES})")
                self.logger.debug(f"URL: {url}")
                
                # Create request with headers
                request = urllib.request.Request(
                    url,
                    headers={
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
                        'Accept-Encoding': 'gzip, deflate',
                        'Connection': 'keep-alive'
                    }
                )
                
                # Make request with SSL context
                with urllib.request.urlopen(request, timeout=TIMEOUT, context=self.ssl_context) as response:
                    content_bytes = response.read()
                    content = self._decode_content(content_bytes)
                
                # Validate response
                if not content:
                    raise urllib.error.URLError("Empty response received")
                
                if len(content) < 100:
                    self.logger.warning(f"Very short response ({len(content)} chars)")
                
                # Check for MOPS page indicators
                if '電子資料查詢作業' in content:
                    self.logger.debug("✅ Found expected MOPS page title")
                elif 'charset=big5' in content:
                    self.logger.debug("✅ Found Big5 encoded MOPS page")
                else:
                    self.logger.warning("⚠️ Response might not be a valid MOPS page")
                
                self.logger.info(f"✅ Successfully fetched {len(content):,} characters")
                return content
                
            except ssl.SSLError as e:
                last_exception = e
                self.logger.warning(f"SSL error on attempt {attempt + 1}: {e}")
                
            except urllib.error.URLError as e:
                last_exception = e
                self.logger.warning(f"URL error on attempt {attempt + 1}: {e}")
                
            except Exception as e:
                last_exception = e
                self.logger.warning(f"Unexpected error on attempt {attempt + 1}: {e}")
            
            if attempt < MAX_RETRIES - 1:
                wait_time = (2 ** attempt) * RATE_LIMIT_DELAY
                self.logger.info(f"Waiting {wait_time:.1f}s before retry...")
                time.sleep(wait_time)
        
        error_msg = f"Failed to fetch {url} after {MAX_RETRIES} attempts: {last_exception}"
        raise NetworkError(error_msg)
    
    def fetch_reports_page(self, params: ValidatedParams, quarter: int = None) -> str:
        """Fetch reports page for given parameters and quarter."""
        url = self._build_url(params.company_id, params.roc_year, quarter)
        
        self.logger.info(f"Fetching reports for company {params.company_id}, "
                        f"year {params.western_year} (ROC {params.roc_year}), "
                        f"quarter {quarter or 'all'}")
        
        # Rate limiting
        time.sleep(RATE_LIMIT_DELAY)
        
        return self._fetch_with_retry(url)
    
    def fetch_all_quarters(self, params: ValidatedParams) -> Dict[int, str]:
        """Fetch reports pages for all specified quarters."""
        results = {}
        
        for quarter in params.quarters:
            try:
                html_content = self.fetch_reports_page(params, quarter)
                results[quarter] = html_content
                self.logger.info(f"✅ Successfully fetched Q{quarter} data")
            except Exception as e:
                self.logger.error(f"❌ Failed to fetch Q{quarter}: {e}")
                results[quarter] = None
        
        return results
