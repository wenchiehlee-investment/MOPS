"""
Complete Enhanced Sheets Connector for MOPS Sheets Uploader v1.1.1
Replace your existing sheets_connector.py with this complete version
"""

import pandas as pd
import json
import os
from typing import Dict, List, Any, Optional, Tuple
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

class MOPSSheetsConnector:
    """Enhanced Google Sheets connector for MOPS matrix data with v1.1.1 font support"""
    
    def __init__(self, config):
        self.config = config
        self.worksheet_name = config.worksheet_name
        self.sheet_id = config.google_sheet_id
        self.credentials = config.google_credentials
        self._worksheet = None
        self._client = None
        self._spreadsheet = None
        
        # Get enhanced font configuration
        self.font_config = self._get_enhanced_font_config()
        
    def _get_enhanced_font_config(self) -> Dict[str, Any]:
        """Get enhanced font configuration from config"""
        if hasattr(self.config, 'get_enhanced_font_config'):
            return self.config.get_enhanced_font_config()
        elif hasattr(self.config, 'get_font_config'):
            return self.config.get_font_config()
        else:
            # Fallback configuration
            return {
                'font_size': getattr(self.config, 'font_size', 14),
                'header_font_size': getattr(self.config, 'header_font_size', 14),
                'bold_headers': getattr(self.config, 'bold_headers', True),
                'bold_company_info': getattr(self.config, 'bold_company_info', True),
                'preset_equivalent': 'custom'
            }
    
    def setup_connection(self) -> bool:
        """Enhanced setup Google Sheets connection with better error handling"""
        try:
            logger.info("🔧 開始設定 Google Sheets 連線 (v1.1.1)...")
            
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
                return False
            
            # Test spreadsheet access
            if not self.sheet_id:
                logger.error("❌ GOOGLE_SHEET_ID 未設定")
                return False
            
            logger.info(f"📊 測試試算表存取: {self.sheet_id}")
            
            try:
                self._spreadsheet = self._client.open_by_key(self.sheet_id)
                logger.info(f"✅ 試算表開啟成功: {self._spreadsheet.title}")
            except Exception as e:
                logger.error(f"❌ 試算表存取失敗: {e}")
                logger.error(f"   請檢查試算表 ID: {self.sheet_id}")
                logger.error(f"   請確認服務帳戶有存取權限")
                return False
            
            # Create or get MOPS worksheet
            try:
                self._worksheet = self._get_or_create_worksheet()
                logger.info(f"✅ 工作表準備完成: {self.worksheet_name}")
            except Exception as e:
                logger.error(f"❌ 工作表設定失敗: {e}")
                return False
            
            # Log font configuration that will be used
            logger.info(f"🔤 字體設定: {self.font_config['font_size']}pt (標題: {self.font_config['header_font_size']}pt)")
            logger.info(f"   預設模式: {self.font_config.get('preset_equivalent', 'custom')}")
            
            logger.info("🎉 Google Sheets 連線設定完成")
            return True
            
        except Exception as e:
            logger.error(f"❌ Google Sheets 連線失敗: {e}")
            return False
    
    def upload_matrix(self, matrix_df: pd.DataFrame, 
                     coverage_stats=None,
                     stock_changes=None) -> bool:
        """Enhanced matrix upload with v1.1.1 font formatting"""
        logger.info("🚀 上傳到 Google Sheets (v1.1.1 增強版)...")
        
        if not self._worksheet and not self.setup_connection():
            logger.error("Cannot establish Google Sheets connection")
            return False
        
        try:
            # Auto-resize worksheet for current matrix
            required_rows = len(matrix_df) + 1  # +1 for header
            required_cols = len(matrix_df.columns)
            self.auto_resize_worksheet(required_rows, required_cols)
            
            # Prepare enhanced data for upload
            upload_data = self._prepare_enhanced_upload_data(matrix_df)
            
            # Clear existing content
            logger.info("🧹 清除現有內容...")
            self._worksheet.clear()
            
            # Upload data in batch
            logger.info(f"📤 批次上傳資料 ({len(upload_data)} 列 × {len(upload_data[0])} 欄)...")
            self._worksheet.update('A1', upload_data, value_input_option='USER_ENTERED')
            
            # Apply enhanced v1.1.1 formatting
            logger.info("🎨 套用 v1.1.1 增強格式...")
            self.format_enhanced_matrix_worksheet(
                matrix_df, coverage_stats, stock_changes
            )
            
            logger.info(f"✅ v1.1.1 增強矩陣上傳成功")
            logger.info(f"   公司數: {len(matrix_df)}, 季度數: {len(matrix_df.columns)-2}")
            logger.info(f"   字體: {self.font_config['font_size']}pt/{self.font_config['header_font_size']}pt")
            logger.info(f"   多重類型支援: {'啟用' if getattr(self.config, 'show_all_report_types', True) else '停用'}")
            
            return True
            
        except Exception as e:
            logger.error(f"Enhanced matrix upload failed: {e}")
            return False
    
    def format_enhanced_matrix_worksheet(self, matrix_df: pd.DataFrame,
                                       coverage_stats=None,
                                       stock_changes=None) -> None:
        """Apply enhanced MOPS-specific formatting with v1.1.1 features"""
        try:
            logger.info("🎨 套用增強型格式設定...")
            
            # Step 1: Apply base font size to entire worksheet
            entire_range = f"A1:{self._get_column_letter(len(matrix_df.columns))}{len(matrix_df)+1}"
            logger.info(f"   設定全域字體: {self.font_config['font_size']}pt")
            self._worksheet.format(entire_range, {
                "textFormat": {
                    "fontSize": self.font_config['font_size'],
                    "fontFamily": "Arial"
                }
            })
            
            # Step 2: Enhanced header formatting
            header_range = f"A1:{self._get_column_letter(len(matrix_df.columns))}1"
            header_format = {
                "backgroundColor": {"red": 0.2, "green": 0.8, "blue": 0.8},
                "textFormat": {
                    "foregroundColor": {"red": 1, "green": 1, "blue": 1}, 
                    "fontSize": self.font_config['header_font_size'],
                    "fontFamily": "Arial"
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            }
            
            if self.font_config['bold_headers']:
                header_format["textFormat"]["bold"] = True
                
            logger.info(f"   設定標題格式: {self.font_config['header_font_size']}pt, 粗體={self.font_config['bold_headers']}")
            self._worksheet.format(header_range, header_format)
            
            # Step 3: Enhanced company info columns formatting
            company_cols_range = f"A2:B{len(matrix_df)+1}"
            company_format = {
                "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                "textFormat": {
                    "fontSize": self.font_config['font_size'],
                    "fontFamily": "Arial"
                },
                "horizontalAlignment": "LEFT",
                "verticalAlignment": "MIDDLE"
            }
            
            if self.font_config['bold_company_info']:
                company_format["textFormat"]["bold"] = True
                
            logger.info(f"   設定公司資訊格式: 粗體={self.font_config['bold_company_info']}")
            self._worksheet.format(company_cols_range, company_format)
            
            # Step 4: v1.1.1 Enhanced quarter data formatting
            if len(matrix_df.columns) > 2:
                quarter_cols_range = f"C1:{self._get_column_letter(len(matrix_df.columns))}{len(matrix_df)+1}"
                quarter_format = {
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "textFormat": {
                        "fontSize": self.font_config['font_size'],
                        "fontFamily": "Arial"
                    }
                }
                self._worksheet.format(quarter_cols_range, quarter_format)
            
            # Step 5: v1.1.1 Special formatting for multiple types
            if hasattr(self.config, 'show_all_report_types') and self.config.show_all_report_types:
                self._apply_multiple_type_formatting(matrix_df)
            
            # Step 6: Future quarter highlighting (if applicable)
            if hasattr(self.config, 'highlight_future_quarters') and self.config.highlight_future_quarters:
                self._apply_future_quarter_formatting(matrix_df)
            
            # Step 7: New company highlighting (if applicable)
            if stock_changes and stock_changes.added_companies:
                self._apply_new_company_formatting(matrix_df, stock_changes.added_companies)
            
            # Step 8: Freeze panes and final touches
            self._worksheet.freeze(rows=1, cols=2)
            
            # Step 9: Auto-resize columns for optimal readability
            self._auto_resize_columns(matrix_df)
            
            logger.info(f"✅ v1.1.1 增強格式設定完成")
            logger.info(f"   字體配置: {self.font_config.get('preset_equivalent', 'custom')} 預設")
            
        except Exception as e:
            logger.warning(f"Enhanced formatting failed (matrix still uploaded): {e}")
    
    def _apply_multiple_type_formatting(self, matrix_df: pd.DataFrame) -> None:
        """Apply special formatting for cells with multiple report types (v1.1.1)"""
        try:
            logger.info("   🔄 套用多重類型特殊格式...")
            
            separator = getattr(self.config, 'report_type_separator', '/')
            
            for row_idx, row in matrix_df.iterrows():
                for col_idx, col_name in enumerate(matrix_df.columns[2:], start=3):  # Skip 代號, 名稱
                    cell_value = str(row[col_name])
                    
                    if cell_value != '-' and separator in cell_value:
                        # This cell has multiple types
                        cell_range = f"{self._get_column_letter(col_idx)}{row_idx + 2}"  # +2 for header and 0-based index
                        
                        multiple_type_format = {
                            "backgroundColor": {"red": 0.9, "green": 0.8, "blue": 1.0},  # Light purple
                            "textFormat": {
                                "fontSize": self.font_config['font_size'],
                                "fontFamily": "Arial",
                                "bold": True
                            }
                        }
                        
                        self._worksheet.format(cell_range, multiple_type_format)
            
        except Exception as e:
            logger.warning(f"Multiple type formatting failed: {e}")
    
    def _apply_future_quarter_formatting(self, matrix_df: pd.DataFrame) -> None:
        """Apply formatting for future quarter columns"""
        try:
            logger.info("   ⏭️ 套用未來季度格式...")
            
            from datetime import datetime
            current_year = datetime.now().year
            current_quarter = (datetime.now().month - 1) // 3 + 1
            
            for col_idx, col_name in enumerate(matrix_df.columns[2:], start=3):
                if 'Q' in col_name:
                    try:
                        year_str, quarter_str = col_name.split(' Q')
                        year, quarter = int(year_str), int(quarter_str)
                        
                        if year > current_year or (year == current_year and quarter > current_quarter):
                            # This is a future quarter
                            col_range = f"{self._get_column_letter(col_idx)}1:{self._get_column_letter(col_idx)}{len(matrix_df)+1}"
                            
                            future_format = {
                                "backgroundColor": {"red": 1.0, "green": 0.9, "blue": 0.8},  # Light orange
                            }
                            
                            self._worksheet.format(col_range, future_format)
                    except:
                        continue
                        
        except Exception as e:
            logger.warning(f"Future quarter formatting failed: {e}")
    
    def _apply_new_company_formatting(self, matrix_df: pd.DataFrame, added_companies: List[str]) -> None:
        """Apply formatting for newly added companies"""
        try:
            logger.info(f"   🆕 套用新公司格式 ({len(added_companies)} 家)...")
            
            for row_idx, row in matrix_df.iterrows():
                company_id = str(row['代號'])
                if company_id in added_companies:
                    row_range = f"A{row_idx + 2}:{self._get_column_letter(len(matrix_df.columns))}{row_idx + 2}"
                    
                    new_company_format = {
                        "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0},  # Light blue
                    }
                    
                    self._worksheet.format(row_range, new_company_format)
                    
        except Exception as e:
            logger.warning(f"New company formatting failed: {e}")
    
    def _auto_resize_columns(self, matrix_df: pd.DataFrame) -> None:
        """Auto-resize columns for optimal readability"""
        try:
            logger.info("   📏 自動調整欄寬...")
            
            # Get the number of columns
            num_cols = len(matrix_df.columns)
            
            # Auto-resize columns based on content
            requests = []
            
            for i in range(num_cols):
                if i == 0:  # 代號 column
                    width = 80
                elif i == 1:  # 名稱 column  
                    width = 150
                else:  # Quarter columns
                    width = 100
                
                requests.append({
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self._worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": i,
                            "endIndex": i + 1
                        },
                        "properties": {
                            "pixelSize": width
                        },
                        "fields": "pixelSize"
                    }
                })
            
            if requests:
                self._spreadsheet.batch_update({"requests": requests})
                
        except Exception as e:
            logger.warning(f"Column auto-resize failed: {e}")
    
    def auto_resize_worksheet(self, required_rows: int, required_cols: int) -> None:
        """Enhanced worksheet resizing with better logic"""
        try:
            current_rows = self._worksheet.row_count
            current_cols = self._worksheet.col_count
            
            # Add buffer for safety
            target_rows = max(required_rows + 10, current_rows)
            target_cols = max(required_cols + 5, current_cols)
            
            if current_rows < target_rows or current_cols < target_cols:
                logger.info(f"📊 調整工作表大小: {current_rows}×{current_cols} → {target_rows}×{target_cols}")
                self._worksheet.resize(rows=target_rows, cols=target_cols)
                
                # Log the adjustment
                logger.info(f"   可容納: {target_rows-1} 家公司, {target_cols-2} 個季度")
            
        except Exception as e:
            logger.warning(f"Worksheet resize failed: {e}")
    
    def export_csv_backup(self, matrix_df: pd.DataFrame, coverage_stats=None) -> str:
        """Enhanced CSV export with v1.1.1 metadata"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Use configurable filename pattern
        filename_pattern = getattr(self.config, 'csv_filename_pattern', 'mops_matrix_{timestamp}.csv')
        filename = filename_pattern.format(timestamp=timestamp)
        
        # Ensure output directory exists
        os.makedirs(self.config.output_dir, exist_ok=True)
        csv_path = os.path.join(self.config.output_dir, filename)
        
        # Export main matrix with enhanced metadata
        matrix_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        
        # Create metadata file
        metadata_path = csv_path.replace('.csv', '_metadata.json')
        metadata = {
            'version': '1.1.1',
            'export_time': datetime.now().isoformat(),
            'font_configuration': self.font_config,
            'matrix_dimensions': {
                'rows': len(matrix_df),
                'columns': len(matrix_df.columns),
                'companies': len(matrix_df),
                'quarters': len(matrix_df.columns) - 2
            },
            'multiple_types_enabled': getattr(self.config, 'show_all_report_types', True),
            'type_separator': getattr(self.config, 'report_type_separator', '/'),
        }
        
        if coverage_stats:
            metadata['coverage_stats'] = {
                'coverage_percentage': coverage_stats.coverage_percentage,
                'total_companies': coverage_stats.total_companies,
                'companies_with_pdfs': coverage_stats.companies_with_pdfs,
                'total_reports': coverage_stats.total_actual_reports
            }
            
            if hasattr(coverage_stats, 'cells_with_multiple_types'):
                metadata['coverage_stats']['multiple_types'] = {
                    'count': coverage_stats.cells_with_multiple_types,
                    'percentage': coverage_stats.multiple_types_percentage
                }
        
        with open(metadata_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
        
        logger.info(f"💾 v1.1.1 增強 CSV 備份: {csv_path}")
        logger.info(f"📋 元資料檔案: {metadata_path}")

        # Auto-update README.md with new download status
        self._update_readme_status(csv_path, matrix_df)

        return csv_path

    def _update_readme_status(self, csv_path: str, matrix_df: pd.DataFrame) -> None:
        """Update README.md Current Download Status section from the generated CSV"""
        import re
        from pathlib import Path
        from collections import Counter

        # Find README.md relative to output_dir
        readme_path = None
        for d in [
            Path(self.config.output_dir).parent,
            Path(self.config.output_dir).parent.parent,
            Path('.'),
        ]:
            candidate = d / 'README.md'
            if candidate.exists():
                readme_path = candidate
                break

        if not readme_path:
            logger.warning("⚠️ README.md 找不到，跳過狀態更新")
            return

        # Extract date from csv filename: mops_matrix_YYYYMMDD_HHMMSS.csv
        date_match = re.search(r'(\d{8})', Path(csv_path).stem)
        if date_match:
            d = date_match.group(1)
            last_updated = f"{d[:4]}-{d[4:6]}-{d[6:8]}"
        else:
            last_updated = datetime.now().strftime('%Y-%m-%d')

        # Quarter columns only (exclude 代號, 名稱, and summary cols)
        quarter_cols = [c for c in matrix_df.columns if ' Q' in c]
        total = len(matrix_df)

        filing_deadlines = {
            'Q1': 'May 15', 'Q2': 'Aug 14', 'Q3': 'Nov 14', 'Q4': 'Mar 31 (next year)'
        }

        # Build coverage table
        table_lines = [
            "| Quarter | Downloaded | Coverage | Notes |",
            "|---------|-----------|----------|-------|",
        ]
        for col in quarter_cols:
            count = int((matrix_df[col] != '-').sum())
            if count == 0:
                pct = "—"
                q_key = col.split(' ')[1]
                note = f"Filing deadline: {filing_deadlines.get(q_key, '')}"
            else:
                pct = f"{count / total * 100:.0f}%"
                note = ""
            table_lines.append(f"| {col} | {count} / {total} | {pct} | {note} |")

        # Report types for most recent non-empty quarter
        type_section = ""
        for col in quarter_cols:
            filled = matrix_df.loc[matrix_df[col] != '-', col]
            if not filled.empty:
                types = Counter(filled)
                rows = "\n".join(f"| {t} | {c} |" for t, c in sorted(types.items()))
                type_section = (
                    f"### Report Types ({col})\n\n"
                    f"| Type | Count |\n|------|-------|\n{rows}\n"
                )
                break

        # Companies missing recent data (show quarters with partial coverage)
        missing_parts = []
        for col in quarter_cols:
            count = int((matrix_df[col] != '-').sum())
            if 0 < count < total:
                missing = matrix_df.loc[matrix_df[col] == '-', ['代號', '名稱']]
                names = "\u3001".join(f"{r['代號']} {r['名稱']}" for _, r in missing.iterrows())
                missing_parts.append(f"**Missing {col}** ({len(missing)} companies):\n{names}")
                if len(missing_parts) >= 3:
                    break

        missing_text = "\n\n".join(missing_parts) if missing_parts else "_No recent missing data_"

        # Assemble the new section
        new_section = (
            "## 📊 Current Download Status\n\n"
            f"> **Last Updated**: {last_updated} | **Source**: `{Path(csv_path).name}`\n\n"
            f"**{total} companies tracked**\n\n"
            + "\n".join(table_lines) + "\n\n"
            + (type_section + "\n" if type_section else "")
            + f"### Companies Missing Recent Data\n\n{missing_text}\n\n---\n"
        )

        # Replace existing section in README.md
        readme_text = readme_path.read_text(encoding='utf-8')
        updated = re.sub(
            r'## 📊 Current Download Status.*?---\n',
            new_section,
            readme_text,
            flags=re.DOTALL,
        )

        if updated == readme_text:
            logger.warning("⚠️ README.md 未找到 '📊 Current Download Status' 區塊，跳過更新")
            return

        readme_path.write_text(updated, encoding='utf-8')
        logger.info(f"✅ README.md 下載狀態已更新: {readme_path}")
    
    def test_connection(self) -> bool:
        """Enhanced connection test with detailed feedback"""
        try:
            success = self.setup_connection()
            if success:
                logger.info("✅ v1.1.1 Google Sheets 連線測試成功")
                logger.info(f"   試算表: {self._spreadsheet.title}")
                logger.info(f"   工作表: {self.worksheet_name}")
                logger.info(f"   字體配置: {self.font_config['preset_equivalent']} ({self.font_config['font_size']}pt)")
            else:
                logger.error("❌ v1.1.1 Google Sheets 連線測試失敗")
            return success
        except Exception as e:
            logger.error(f"Connection test failed: {e}")
            return False
    
    def _get_or_create_worksheet(self):
        """Enhanced worksheet creation/retrieval with better error handling"""
        try:
            # Try to get existing worksheet
            worksheet = self._spreadsheet.worksheet(self.worksheet_name)
            logger.info(f"📋 找到現有工作表: {self.worksheet_name}")
            return worksheet
        except:
            try:
                # Create new worksheet
                worksheet = self._spreadsheet.add_worksheet(
                    title=self.worksheet_name,
                    rows=300,  # Initial size - will be auto-resized
                    cols=30
                )
                logger.info(f"📋 建立新工作表: {self.worksheet_name}")
                return worksheet
            except Exception as e:
                logger.error(f"Failed to create worksheet: {e}")
                raise
    
    def _prepare_enhanced_upload_data(self, matrix_df: pd.DataFrame) -> List[List[str]]:
        """Enhanced data preparation for Google Sheets upload"""
        data = []
        
        # Add header row
        data.append(list(matrix_df.columns))
        
        # Add data rows with enhanced processing
        for _, row in matrix_df.iterrows():
            row_data = []
            for col_name in matrix_df.columns:
                cell_value = row[col_name]
                
                # Enhanced cell value processing for v1.1.1
                if col_name in ['代號', '名稱']:
                    # Company info - keep as is
                    row_data.append(str(cell_value))
                else:
                    # Quarter data - handle multiple types
                    if pd.isna(cell_value) or cell_value == '':
                        row_data.append('-')
                    else:
                        row_data.append(str(cell_value))
            
            data.append(row_data)
        
        return data
    
    def _get_column_letter(self, col_num: int) -> str:
        """Convert column number to letter (1 -> A, 2 -> B, etc.)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result

class SheetsUploadManager:
    """Enhanced manager for handling upload with fallback strategies (v1.1.1)"""
    
    def __init__(self, config):
        self.config = config
        self.connector = MOPSSheetsConnector(config)
    
    def upload_with_fallback(self, matrix_df: pd.DataFrame,
                           coverage_stats=None,
                           stock_changes=None) -> Dict[str, Any]:
        """Enhanced upload with automatic fallback strategies"""
        result = {
            'sheets_success': False,
            'csv_exported': False,
            'sheets_url': None,
            'csv_path': None,
            'error': None,
            'v1_1_1_features': {
                'font_config': self.connector.font_config,
                'multiple_types_enabled': getattr(self.config, 'show_all_report_types', True)
            }
        }
        
        # Try Google Sheets upload if credentials are available
        if self.config.google_sheet_id and self.config.google_credentials:
            logger.info("🚀 嘗試 v1.1.1 增強上傳到 Google Sheets...")
            try:
                sheets_success = self.connector.upload_matrix(matrix_df, coverage_stats, stock_changes)
                result['sheets_success'] = sheets_success
                
                if sheets_success:
                    result['sheets_url'] = f"https://docs.google.com/spreadsheets/d/{self.config.google_sheet_id}"
                    logger.info(f"✅ v1.1.1 Google Sheets 上傳成功")
                    logger.info(f"🔗 Google Sheets URL: {result['sheets_url']}")
                    logger.info(f"📋 工作表: {self.config.worksheet_name}")
                    logger.info(f"🔤 字體: {self.connector.font_config['preset_equivalent']} ({self.connector.font_config['font_size']}pt)")
                else:
                    logger.warning("⚠️ v1.1.1 Google Sheets 上傳失敗，將使用 CSV 備份")
                
            except Exception as e:
                logger.warning(f"⚠️ v1.1.1 Google Sheets 上傳失敗: {e}")
                result['error'] = str(e)
        else:
            logger.info("ℹ️ Google Sheets 憑證未設定，跳過上傳")
        
        # Create enhanced CSV backup if enabled OR if Sheets failed
        if self.config.csv_backup or not result['sheets_success']:
            logger.info("💾 建立 v1.1.1 增強 CSV 備份...")
            try:
                csv_path = self.connector.export_csv_backup(matrix_df, coverage_stats)
                result['csv_exported'] = True
                result['csv_path'] = csv_path
                logger.info(f"✅ v1.1.1 增強 CSV 備份建立成功: {csv_path}")
            except Exception as e:
                logger.error(f"❌ v1.1.1 CSV 備份失敗: {e}")
                if not result['error']:
                    result['error'] = f"CSV export failed: {e}"
        
        return result