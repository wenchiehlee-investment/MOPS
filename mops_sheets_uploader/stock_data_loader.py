"""
MOPS Sheets Uploader - Stock Data Loader
Loads and processes company information from StockID_TWSE_TPEX.csv
"""

import os
import pandas as pd
from typing import Dict, List, Any, Optional, Set
from pathlib import Path
import logging

from .models import StockInfo, StockListChanges
from .config import MOPSConfig

logger = logging.getLogger(__name__)

class StockDataLoader:
    """Loads and processes stock data from CSV file"""
    
    def __init__(self, config: MOPSConfig):
        self.config = config
        self.csv_path = Path(config.stock_csv_path)
        self._last_known_companies: Optional[Set[str]] = None
        
    def load_stock_csv(self) -> pd.DataFrame:
        """
        Load StockID_TWSE_TPEX.csv with proper encoding and validation
        
        Expected format:
        - Columns: 代號 (Integer), 名稱 (String)
        - Encoding: UTF-8
        - Variable number of companies (116+ as it grows)
        
        Returns:
            DataFrame with columns ['代號', '名稱']
        """
        logger.info(f"📋 載入股票清單: {self.csv_path}")
        
        if not self.csv_path.exists():
            raise FileNotFoundError(f"Stock CSV file not found: {self.csv_path}")
        
        try:
            # Try UTF-8 first, then fallback encodings
            encodings = ['utf-8', 'utf-8-sig', 'big5', 'gb2312']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(self.csv_path, encoding=encoding)
                    logger.debug(f"Successfully loaded CSV with encoding: {encoding}")
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                raise ValueError("Cannot decode CSV file with any supported encoding")
            
            # Validate and clean column names
            df = self._clean_column_names(df)
            
            # Validate required columns
            required_columns = ['代號', '名稱']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns: {missing_columns}")
            
            # Clean and validate data
            df = self._clean_stock_data(df)
            
            company_count = len(df)
            logger.info(f"   已載入 {company_count} 家公司資料")
            
            return df
            
        except Exception as e:
            logger.error(f"Error loading stock CSV: {e}")
            raise
    
    def validate_stock_data(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Validate stock data completeness and format
        
        Returns:
            Dict with validation information
        """
        validation_info = {
            'total_companies': len(df),
            'duplicate_codes': [],
            'invalid_codes': [],
            'missing_names': [],
            'validation_passed': True,
            'warnings': []
        }
        
        # Check for duplicate stock codes
        duplicates = df[df['代號'].duplicated()]
        if not duplicates.empty:
            validation_info['duplicate_codes'] = duplicates['代號'].tolist()
            validation_info['warnings'].append(f"Found {len(duplicates)} duplicate stock codes")
        
        # Check for invalid stock codes (should be 4-digit strings)
        invalid_codes = df[~df['代號'].astype(str).str.match(r'^\d{4}$')]
        if not invalid_codes.empty:
            validation_info['invalid_codes'] = invalid_codes['代號'].tolist()
            validation_info['warnings'].append(f"Found {len(invalid_codes)} invalid stock codes")
        
        # Check for missing company names
        missing_names = df[df['名稱'].isna() | (df['名稱'] == '')]
        if not missing_names.empty:
            validation_info['missing_names'] = missing_names['代號'].tolist()
            validation_info['warnings'].append(f"Found {len(missing_names)} companies with missing names")
        
        # Determine if validation passed
        validation_info['validation_passed'] = (
            len(validation_info['duplicate_codes']) == 0 and
            len(validation_info['invalid_codes']) == 0 and
            len(validation_info['missing_names']) == 0
        )
        
        # Log warnings
        for warning in validation_info['warnings']:
            logger.warning(f"⚠️ {warning}")
        
        return validation_info
    
    def detect_stock_list_changes(self, df: pd.DataFrame, 
                                 previous_companies: Optional[Set[str]] = None) -> StockListChanges:
        """
        Detect changes in stock list since last run
        
        Args:
            df: Current stock DataFrame
            previous_companies: Set of previous company codes (optional)
            
        Returns:
            StockListChanges object with detected changes
        """
        current_companies = set(df['代號'].astype(str))
        current_total = len(current_companies)
        
        if previous_companies is None:
            previous_companies = self._last_known_companies or current_companies
        
        previous_total = len(previous_companies)
        
        # Detect changes
        added_companies = list(current_companies - previous_companies)
        removed_companies = list(previous_companies - current_companies)
        
        changes = StockListChanges(
            added_companies=sorted(added_companies),
            removed_companies=sorted(removed_companies),
            total_companies=current_total,
            previous_total=previous_total
        )
        
        # Log changes if any
        if changes.has_changes:
            logger.info(f"🔄 偵測到股票清單變更:")
            if changes.added_companies:
                logger.info(f"   • 新增公司: {', '.join(changes.added_companies)}")
            if changes.removed_companies:
                logger.info(f"   • 移除公司: {', '.join(changes.removed_companies)}")
            logger.info(f"   • 總公司數: {previous_total} → {current_total} ({current_total - previous_total:+d})")
            
            # Warn if large changes detected
            total_changes = len(changes.added_companies) + len(changes.removed_companies)
            if total_changes > self.config.change_threshold_warning:
                logger.warning(f"⚠️ 大量股票清單變更: {total_changes} 家公司, 超過警告閾值 {self.config.change_threshold_warning}")
        
        # Update last known companies
        self._last_known_companies = current_companies
        
        return changes
    
    def get_company_name(self, df: pd.DataFrame, stock_id: str) -> str:
        """Get company name by stock ID"""
        try:
            mask = df['代號'].astype(str) == str(stock_id)
            matches = df[mask]
            
            if matches.empty:
                logger.warning(f"Company not found: {stock_id}")
                return f"未知公司({stock_id})"
            
            return matches.iloc[0]['名稱']
            
        except Exception as e:
            logger.error(f"Error getting company name for {stock_id}: {e}")
            return f"錯誤({stock_id})"
    
    def get_company_mapping(self, df: pd.DataFrame) -> Dict[str, str]:
        """Get mapping of stock ID to company name"""
        mapping = {}
        
        for _, row in df.iterrows():
            stock_id = str(row['代號'])
            company_name = str(row['名稱'])
            mapping[stock_id] = company_name
        
        return mapping
    
    def create_stock_info_list(self, df: pd.DataFrame) -> List[StockInfo]:
        """Create list of StockInfo objects from DataFrame"""
        stock_info_list = []
        
        for _, row in df.iterrows():
            try:
                stock_info = StockInfo(
                    stock_id=str(row['代號']),
                    company_name=str(row['名稱'])
                )
                stock_info_list.append(stock_info)
            except ValueError as e:
                logger.warning(f"Invalid stock info: {e}")
                continue
        
        return stock_info_list
    
    def _clean_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and standardize column names"""
        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Handle common column name variations
        column_mapping = {
            '股票代號': '代號',
            '公司代號': '代號',
            'code': '代號',
            'stock_id': '代號',
            '公司名稱': '名稱',
            '股票名稱': '名稱',
            'name': '名稱',
            'company_name': '名稱'
        }
        
        df = df.rename(columns=column_mapping)
        
        return df
    
    def _clean_stock_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and validate stock data"""
        # Remove rows with missing essential data
        df = df.dropna(subset=['代號'])
        
        # Convert stock code to string and pad with zeros if needed
        df['代號'] = df['代號'].astype(str).str.zfill(4)
        
        # Clean company names
        df['名稱'] = df['名稱'].astype(str).str.strip()
        df['名稱'] = df['名稱'].replace('nan', '')
        
        # Remove duplicate stock codes (keep first occurrence)
        df = df.drop_duplicates(subset=['代號'], keep='first')
        
        # Reset index
        df = df.reset_index(drop=True)
        
        return df
    
    def export_stock_changes(self, changes: StockListChanges, 
                           output_path: Optional[str] = None) -> str:
        """Export stock list changes to file"""
        import json
        from datetime import datetime
        
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"stock_changes_{timestamp}.json"
        
        change_data = {
            'timestamp': datetime.now().isoformat(),
            'added_companies': changes.added_companies,
            'removed_companies': changes.removed_companies,
            'total_companies': changes.total_companies,
            'previous_total': changes.previous_total,
            'change_summary': changes.change_summary
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(change_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"Stock changes exported to: {output_path}")
        return output_path

class StockDataValidator:
    """Helper class for advanced stock data validation"""
    
    @staticmethod
    def validate_stock_code_format(stock_codes: List[str]) -> Dict[str, List[str]]:
        """Validate stock code formats according to Taiwan market rules"""
        validation_results = {
            'valid': [],
            'invalid_length': [],
            'invalid_format': [],
            'suspicious': []
        }
        
        for code in stock_codes:
            code_str = str(code)
            
            if len(code_str) != 4:
                validation_results['invalid_length'].append(code_str)
            elif not code_str.isdigit():
                validation_results['invalid_format'].append(code_str)
            elif code_str.startswith('0000'):
                validation_results['suspicious'].append(code_str)
            else:
                validation_results['valid'].append(code_str)
        
        return validation_results
    
    @staticmethod
    def detect_encoding_issues(df: pd.DataFrame) -> List[str]:
        """Detect potential encoding issues in company names"""
        issues = []
        
        for idx, name in enumerate(df['名稱']):
            name_str = str(name)
            
            # Check for common encoding artifacts
            if '???' in name_str:
                issues.append(f"Row {idx}: Possible encoding issue in '{name_str}'")
            elif len(name_str.encode('utf-8')) != len(name_str.encode('utf-8', errors='ignore')):
                issues.append(f"Row {idx}: Encoding inconsistency in '{name_str}'")
        
        return issues