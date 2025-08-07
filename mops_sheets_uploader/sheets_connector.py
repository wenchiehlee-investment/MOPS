"""
MOPS Sheets Uploader - Google Sheets Connector
Handles Google Sheets integration, extends existing sheets_uploader functionality.
"""

import pandas as pd
import json
import os
from typing import Dict, List, Any, Optional, Tuple
import logging
from datetime import datetime

from .models import CoverageStats, StockListChanges
from .config import MOPSConfig

logger = logging.getLogger(__name__)

class MOPSSheetsConnector:
    """
    Extended sheets connector for MOPS matrix data
    
    Integrates with existing Google Sheets infrastructure and reuses:
    - Google Sheets authentication (GOOGLE_SHEETS_CREDENTIALS)
    - Connection management
    - Error handling and retry logic
    - CSV fallback functionality
    """
    
    def __init__(self, config: MOPSConfig):
        self.config = config
        self.worksheet_name = config.worksheet_name
        self.sheet_id = config.google_sheet_id
        self.credentials = config.google_credentials
        self._worksheet = None
        self._client = None
        
    def setup_connection(self) -> bool:
        """Setup Google Sheets connection using existing credentials"""
        try:
            logger.info("🔧 開始設定 Google Sheets 連線...")
            
            # Check if gspread is available
            try:
                import gspread
                from google.oauth2.service_account import Credentials
                logger.info("✅ Google 函式庫載入成功")
            except ImportError as e:
                logger.error(f"❌ Google 函式庫載入失敗: {e}")
                logger.error("請執行: pip install gspread google-auth")
                return False
            
            # Load credentials from environment
            if not self.credentials:
                logger.error("❌ GOOGLE_SHEETS_CREDENTIALS 環境變數為空")
                return False
            
            logger.info(f"📋 憑證長度: {len(self.credentials)} 字元")
            
            # Parse credentials JSON
            try:
                creds_dict = json.loads(self.credentials)
                logger.info("✅ 憑證 JSON 解析成功")
                logger.info(f"📧 服務帳戶: {creds_dict.get('client_email', 'N/A')}")
                logger.info(f"🏗️ 專案 ID: {creds_dict.get('project_id', 'N/A')}")
            except json.JSONDecodeError as e:
                logger.error(f"❌ 憑證 JSON 解析失敗: {e}")
                logger.error("請檢查 GOOGLE_SHEETS_CREDENTIALS 格式")
                return False
            
            # Setup authentication
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            try:
                credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
                logger.info("✅ Google 憑證物件建立成功")
            except Exception as e:
                logger.error(f"❌ Google 憑證物件建立失敗: {e}")
                return False
            
            try:
                self._client = gspread.authorize(credentials)
                logger.info("✅ Google Sheets 客戶端授權成功")
            except Exception as e:
                logger.error(f"❌ Google Sheets 客戶端授權失敗: {e}")
                logger.error("可能原因: 網路連線問題或憑證權限不足")
                return False
            
            # Test spreadsheet access
            if not self.sheet_id:
                logger.error("❌ GOOGLE_SHEET_ID 未設定")
                return False
            
            logger.info(f"📊 測試試算表存取: {self.sheet_id}")
            
            try:
                spreadsheet = self._client.open_by_key(self.sheet_id)
                logger.info(f"✅ 試算表開啟成功: {spreadsheet.title}")
            except gspread.SpreadsheetNotFound:
                logger.error(f"❌ 找不到試算表 ID: {self.sheet_id}")
                logger.error("請檢查:")
                logger.error("  1. 試算表 ID 是否正確")
                logger.error("  2. 服務帳戶是否有存取權限")
                return False
            except gspread.APIError as e:
                logger.error(f"❌ Google Sheets API 錯誤: {e}")
                logger.error("可能原因: API 配額超限或權限不足")
                return False
            except Exception as e:
                logger.error(f"❌ 試算表存取失敗: {e}")
                return False
            
            # Create or get MOPS worksheet
            try:
                self._worksheet = self._get_or_create_worksheet(spreadsheet)
                logger.info(f"✅ 工作表準備完成: {self.worksheet_name}")
            except Exception as e:
                logger.error(f"❌ 工作表設定失敗: {e}")
                return False
            
            logger.info("🎉 Google Sheets 連線設定完成")
            return True
            
        except ImportError:
            logger.error("❌ gspread 函式庫未安裝")
            logger.error("請執行: pip install gspread google-auth")
            return False
        except Exception as e:
            logger.error(f"❌ Google Sheets 連線失敗: {e}")
            import traceback
            logger.error(f"詳細錯誤: {traceback.format_exc()}")
            return False
    
    def upload_matrix(self, matrix_df: pd.DataFrame, 
                     coverage_stats: Optional[CoverageStats] = None,
                     stock_changes: Optional[StockListChanges] = None) -> bool:
        """
        Upload matrix data to Google Sheets with formatting
        
        Process:
        1. Setup connection if needed
        2. Auto-resize worksheet for current matrix dimensions
        3. Upload matrix data with proper formatting
        4. Apply MOPS-specific styling
        5. Create summary if enabled
        """
        logger.info("🚀 上傳到 Google Sheets...")
        
        if not self._worksheet and not self.setup_connection():
            logger.error("Cannot establish Google Sheets connection")
            return False
        
        try:
            # Auto-resize worksheet for current matrix
            required_rows = len(matrix_df) + 1  # +1 for header
            required_cols = len(matrix_df.columns)
            self.auto_resize_worksheet(required_rows, required_cols)
            
            # Prepare data for upload
            upload_data = self._prepare_upload_data(matrix_df)
            
            # Clear existing content
            self._worksheet.clear()
            
            # Upload data in batch
            self._worksheet.update('A1', upload_data, value_input_option='USER_ENTERED')
            
            # Apply formatting
            self.format_matrix_worksheet(
                matrix_df, 
                coverage_stats,
                stock_changes
            )
            
            # Create summary sheet if enabled
            if self.config.include_summary_sheet and coverage_stats:
                self._create_summary_sheet(coverage_stats, stock_changes)
            
            logger.info(f"✅ 主矩陣上傳成功 ({len(matrix_df)} 公司 × {len(matrix_df.columns)-2} 季度)")
            
            return True
            
        except Exception as e:
            logger.error(f"Matrix upload failed: {e}")
            return False
    
    def format_matrix_worksheet(self, matrix_df: pd.DataFrame,
                               coverage_stats: Optional[CoverageStats] = None,
                               stock_changes: Optional[StockListChanges] = None) -> None:
        """
        Apply MOPS-specific formatting for matrix display
        
        Features:
        - Header row styling (company info vs quarters)
        - Future quarter column highlighting (orange background)
        - New company row highlighting (light blue background)
        - Status symbol color coding (✅🟡⚠️❌)
        - Auto-resize columns for readability
        - Freeze first two columns (代號, 名稱)
        """
        try:
            # Header formatting
            header_range = f"A1:{self._get_column_letter(len(matrix_df.columns))}1"
            self._worksheet.format(header_range, {
                "backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2},
                "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}, "bold": True},
                "horizontalAlignment": "CENTER"
            })
            
            # Company info columns (代號, 名稱) - different styling
            company_cols_range = f"A1:B{len(matrix_df)+1}"
            self._worksheet.format(company_cols_range, {
                "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                "textFormat": {"bold": True}
            })
            
            # Future quarter highlighting
            if coverage_stats and coverage_stats.future_quarters:
                self._highlight_future_quarters(matrix_df, coverage_stats.future_quarters)
            
            # New company highlighting
            if stock_changes and stock_changes.added_companies:
                self._highlight_new_companies(matrix_df, stock_changes.added_companies)
            
            # Status symbol color coding
            self._apply_status_color_coding(matrix_df)
            
            # Freeze first two columns
            self._worksheet.freeze(rows=1, cols=2)
            
            # Auto-resize columns
            if self.config.auto_resize_columns:
                self._auto_resize_columns(matrix_df)
            
            logger.info("✅ 矩陣格式設定完成")
            
        except Exception as e:
            logger.warning(f"Formatting failed (matrix still uploaded): {e}")
    
    def auto_resize_worksheet(self, required_rows: int, required_cols: int) -> None:
        """
        Automatically resize worksheet to fit matrix dimensions
        
        Args:
            required_rows: Current company count + header (e.g., 118 + 1 = 119)
            required_cols: Base columns + quarter columns (e.g., 2 + 12 = 14)
        """
        try:
            current_rows = self._worksheet.row_count
            current_cols = self._worksheet.col_count
            
            # Add buffer for safety
            target_rows = max(required_rows + 5, current_rows)
            target_cols = max(required_cols + 2, current_cols)
            
            if current_rows < target_rows or current_cols < target_cols:
                self._worksheet.resize(rows=target_rows, cols=target_cols)
                logger.info(f"📊 工作表已調整大小: {current_rows}×{current_cols} → {target_rows}×{target_cols}")
            
        except Exception as e:
            logger.warning(f"Worksheet resize failed: {e}")
    
    def export_csv_backup(self, matrix_df: pd.DataFrame, 
                         coverage_stats: Optional[CoverageStats] = None) -> str:
        """Export matrix to CSV as backup"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = self.config.csv_filename_pattern.format(timestamp=timestamp)
        
        # Ensure output directory exists
        os.makedirs(self.config.output_dir, exist_ok=True)
        csv_path = os.path.join(self.config.output_dir, filename)
        
        # Export main matrix
        matrix_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        
        # Export summary if available
        if coverage_stats and self.config.include_future_analysis:
            summary_path = csv_path.replace('.csv', '_summary.csv')
            self._export_summary_csv(coverage_stats, summary_path)
        
        logger.info(f"💾 CSV 備份檔案: {csv_path}")
        return csv_path
    
    def test_connection(self) -> bool:
        """Test Google Sheets connection"""
        try:
            success = self.setup_connection()
            if success:
                logger.info("✅ Google Sheets 連線測試成功")
            else:
                logger.error("❌ Google Sheets 連線測試失敗")
            return success
        except Exception as e:
            logger.error(f"Connection test failed: {e}")
            return False
    
    def _get_or_create_worksheet(self, spreadsheet) -> Any:
        """Get existing worksheet or create new one"""
        try:
            # Try to get existing worksheet
            worksheet = spreadsheet.worksheet(self.worksheet_name)
            logger.info(f"Found existing worksheet: {self.worksheet_name}")
            return worksheet
        except:
            # Create new worksheet
            worksheet = spreadsheet.add_worksheet(
                title=self.worksheet_name,
                rows=200,  # Initial size
                cols=20
            )
            logger.info(f"Created new worksheet: {self.worksheet_name}")
            return worksheet
    
    def _prepare_upload_data(self, matrix_df: pd.DataFrame) -> List[List[str]]:
        """Prepare matrix data for Google Sheets upload"""
        # Convert DataFrame to list of lists
        data = []
        
        # Add header row
        data.append(list(matrix_df.columns))
        
        # Add data rows
        for _, row in matrix_df.iterrows():
            data.append([str(cell) for cell in row])
        
        return data
    
    def _highlight_future_quarters(self, matrix_df: pd.DataFrame, 
                                  future_quarters: List[str]) -> None:
        """Highlight future quarter columns"""
        try:
            for quarter in future_quarters:
                if quarter in matrix_df.columns:
                    col_index = matrix_df.columns.get_loc(quarter) + 1  # +1 for 1-based indexing
                    col_letter = self._get_column_letter(col_index)
                    
                    # Highlight entire column
                    range_name = f"{col_letter}:{col_letter}"
                    self._worksheet.format(range_name, {
                        "backgroundColor": {"red": 1, "green": 0.8, "blue": 0.4}  # Orange
                    })
        except Exception as e:
            logger.warning(f"Future quarter highlighting failed: {e}")
    
    def _highlight_new_companies(self, matrix_df: pd.DataFrame, 
                                new_companies: List[str]) -> None:
        """Highlight new company rows"""
        try:
            for company_id in new_companies:
                company_mask = matrix_df['代號'].astype(str) == company_id
                row_indices = matrix_df[company_mask].index.tolist()
                
                for row_idx in row_indices:
                    row_num = row_idx + 2  # +2 for 1-based indexing and header row
                    range_name = f"A{row_num}:{self._get_column_letter(len(matrix_df.columns))}{row_num}"
                    
                    self._worksheet.format(range_name, {
                        "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1}  # Light blue
                    })
        except Exception as e:
            logger.warning(f"New company highlighting failed: {e}")
    
    def _apply_status_color_coding(self, matrix_df: pd.DataFrame) -> None:
        """Apply color coding for status symbols"""
        try:
            # Define colors for different status symbols
            status_colors = {
                '✅': {"red": 0.8, "green": 1, "blue": 0.8},      # Light green
                '🟡': {"red": 1, "green": 1, "blue": 0.8},        # Light yellow
                '⚠️': {"red": 1, "green": 0.9, "blue": 0.8},      # Light orange
                '❌': {"red": 1, "green": 0.8, "blue": 0.8},       # Light red
            }
            
            # Apply colors (simplified approach - would need more complex logic for full implementation)
            quarter_cols = [col for col in matrix_df.columns if 'Q' in col]
            for col in quarter_cols:
                col_index = matrix_df.columns.get_loc(col) + 1
                col_letter = self._get_column_letter(col_index)
                
                # Apply conditional formatting for each status type
                # Note: This is a simplified version - full implementation would need cell-by-cell formatting
                
        except Exception as e:
            logger.warning(f"Status color coding failed: {e}")
    
    def _auto_resize_columns(self, matrix_df: pd.DataFrame) -> None:
        """Auto-resize columns for better readability"""
        try:
            # Set column widths based on content type
            column_widths = {}
            
            for i, col in enumerate(matrix_df.columns):
                col_letter = self._get_column_letter(i + 1)
                
                if col == '代號':
                    width = 80
                elif col == '名稱':
                    width = 150
                elif 'Q' in col:
                    width = 100
                else:
                    width = 120
                
                column_widths[col_letter] = width
            
            # Apply column widths
            for col_letter, width in column_widths.items():
                self._worksheet.columns_auto_resize(start_column_index=ord(col_letter) - ord('A'), 
                                                  end_column_index=ord(col_letter) - ord('A'))
                
        except Exception as e:
            logger.warning(f"Column auto-resize failed: {e}")
    
    def _create_summary_sheet(self, coverage_stats: CoverageStats,
                             stock_changes: Optional[StockListChanges] = None) -> None:
        """Create summary statistics sheet"""
        try:
            # This would create a separate worksheet with summary statistics
            # Implementation details depend on specific requirements
            logger.info("📊 Summary sheet creation not yet implemented")
        except Exception as e:
            logger.warning(f"Summary sheet creation failed: {e}")
    
    def _export_summary_csv(self, coverage_stats: CoverageStats, summary_path: str) -> None:
        """Export summary statistics to CSV"""
        summary_data = {
            '統計項目': [
                '總公司數',
                '有PDF公司數',
                '無PDF公司數',
                '總季度數',
                '總報告數',
                '涵蓋率'
            ],
            '數值': [
                coverage_stats.total_companies,
                coverage_stats.companies_with_pdfs,
                coverage_stats.companies_without_pdfs,
                coverage_stats.total_quarters,
                coverage_stats.total_actual_reports,
                f"{coverage_stats.coverage_percentage:.1f}%"
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_csv(summary_path, index=False, encoding='utf-8-sig')
    
    def _get_column_letter(self, col_num: int) -> str:
        """Convert column number to letter (1 -> A, 2 -> B, etc.)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result

class SheetsUploadManager:
    """Manager for handling upload with fallback strategies"""
    
    def __init__(self, config: MOPSConfig):
        self.config = config
        self.connector = MOPSSheetsConnector(config)
    
    def upload_with_fallback(self, matrix_df: pd.DataFrame,
                           coverage_stats: Optional[CoverageStats] = None,
                           stock_changes: Optional[StockListChanges] = None) -> Dict[str, Any]:
        """
        Upload with automatic fallback strategies
        
        Returns:
            Dict with upload results and paths
        """
        result = {
            'sheets_success': False,
            'csv_exported': False,
            'sheets_url': None,
            'csv_path': None,
            'error': None
        }
        
        # Try Google Sheets upload if credentials are available
        if self.config.google_sheet_id and self.config.google_credentials:
            logger.info("🚀 嘗試上傳到 Google Sheets...")
            try:
                sheets_success = self.connector.upload_matrix(matrix_df, coverage_stats, stock_changes)
                result['sheets_success'] = sheets_success
                
                if sheets_success:
                    result['sheets_url'] = f"https://docs.google.com/spreadsheets/d/{self.config.google_sheet_id}"
                    logger.info(f"✅ Google Sheets 上傳成功")
                    logger.info(f"🔗 Google Sheets URL: {result['sheets_url']}")
                else:
                    logger.warning("⚠️ Google Sheets 上傳失敗，將使用 CSV 備份")
                
            except Exception as e:
                logger.warning(f"⚠️ Google Sheets 上傳失敗: {e}")
                result['error'] = str(e)
        else:
            logger.info("ℹ️ Google Sheets 憑證未設定，跳過上傳")
        
        # Create CSV backup if enabled OR if Sheets failed
        if self.config.csv_backup or not result['sheets_success']:
            logger.info("💾 建立 CSV 備份...")
            try:
                csv_path = self.connector.export_csv_backup(matrix_df, coverage_stats)
                result['csv_exported'] = True
                result['csv_path'] = csv_path
                logger.info(f"✅ CSV 備份建立成功: {csv_path}")
            except Exception as e:
                logger.error(f"❌ CSV 備份失敗: {e}")
                if not result['error']:
                    result['error'] = f"CSV export failed: {e}"
        
        return result