"""
MOPS Sheets Uploader - Matrix Builder
Constructs the main data matrix combining stock data and PDF information.
"""

import pandas as pd
from typing import Dict, List, Optional, Tuple, Set
from datetime import datetime
import logging

from .models import PDFFile, MatrixCell, CoverageStats, StockListChanges
from .config import MOPSConfig

logger = logging.getLogger(__name__)

class MatrixBuilder:
    """Builds the main data matrix for Google Sheets upload"""
    
    def __init__(self, config: MOPSConfig):
        self.config = config
        
    def build_base_matrix(self, stock_df: pd.DataFrame, 
                         pdf_data: Dict[str, List[PDFFile]]) -> pd.DataFrame:
        """
        Create base matrix with companies and dynamically discovered quarter columns
        
        Process:
        1. Start with N companies from StockID_TWSE_TPEX.csv (auto-detects count)
        2. Discover quarters from actual PDF files
        3. Add expected quarters within max_years range  
        4. Create matrix with dynamic row and column count
        
        Args:
            stock_df: DataFrame with companies (代號, 名稱)
            pdf_data: Dictionary mapping company_id to PDFFile list
            
        Returns:
            DataFrame with shape (N, 2 + dynamic_quarters) where:
            - N rows: current number of companies in CSV
            - Columns: 代號, 名稱, [quarters in reverse chronological order]
        """
        logger.info("🏗️ 建構矩陣資料...")
        
        # Get discovered quarters from PDF data
        discovered_quarters = self._discover_quarters_from_pdfs(pdf_data)
        
        # Generate complete quarter columns based on config
        quarter_columns = self.config.get_quarter_columns(discovered_quarters)
        
        # Create base matrix structure
        matrix_data = {
            '代號': stock_df['代號'].astype(str),
            '名稱': stock_df['名稱'].astype(str)
        }
        
        # Initialize quarter columns with empty values
        for quarter in quarter_columns:
            matrix_data[quarter] = ['-'] * len(stock_df)
        
        matrix_df = pd.DataFrame(matrix_data)
        
        logger.info(f"   矩陣大小: {len(matrix_df)} × {len(matrix_df.columns)} (公司 × 欄位)")
        logger.info(f"   動態季度欄位: {', '.join(quarter_columns[:5])}{'...' if len(quarter_columns) > 5 else ''}")
        
        return matrix_df
    
    def populate_pdf_status(self, matrix_df: pd.DataFrame, 
                           pdf_data: Dict[str, List[PDFFile]]) -> pd.DataFrame:
        """
        Fill matrix with PDF availability status for all companies
        
        Handles:
        - Existing companies with PDFs (normal processing)
        - New companies without PDFs (shows "-" for all quarters)  
        - Companies with partial PDF coverage
        - Multiple report types per quarter (priority rules)
        """
        logger.info("📊 填入 PDF 狀態資料...")
        
        # Create matrix cells for each company-quarter combination
        matrix_cells = self._create_matrix_cells(matrix_df, pdf_data)
        
        # Fill matrix with PDF status
        filled_matrix = matrix_df.copy()
        
        for idx, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            
            # Skip base columns (代號, 名稱)
            for col in matrix_df.columns[2:]:
                quarter_key = col
                cell_key = f"{company_id}_{quarter_key}"
                
                if cell_key in matrix_cells:
                    cell = matrix_cells[cell_key]
                    filled_matrix.loc[idx, col] = cell.get_display_value()
                else:
                    filled_matrix.loc[idx, col] = '-'
        
        # Log summary statistics
        self._log_population_summary(filled_matrix, pdf_data)
        
        return filled_matrix
    
    def identify_new_companies(self, matrix_df: pd.DataFrame, 
                              pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """
        Identify companies in stock list but not in PDF data
        
        Returns:
            List of company codes that have no PDFs downloaded yet
        """
        all_companies = set(matrix_df['代號'].astype(str))
        companies_with_pdfs = set(pdf_data.keys())
        
        new_companies = list(all_companies - companies_with_pdfs)
        
        if new_companies:
            logger.info(f"🆕 發現 {len(new_companies)} 家公司無 PDF 檔案: {', '.join(new_companies[:5])}{'...' if len(new_companies) > 5 else ''}")
        
        return sorted(new_companies)
    
    def apply_priority_rules(self, matrix_df: pd.DataFrame, 
                           pdf_data: Dict[str, List[PDFFile]]) -> pd.DataFrame:
        """
        Apply report type priority rules for conflicting files
        
        Priority order (from models.py):
        1. A12, A13 (Individual reports)
        2. AI1, A1L (Consolidated reports)  
        3. A10, A11 (Generic reports)
        9. AIA, AE2 (English reports - excluded)
        """
        logger.info("🎯 套用報告優先級規則...")
        
        conflicts_resolved = 0
        
        for company_id, pdfs in pdf_data.items():
            # Group PDFs by quarter
            quarter_groups = {}
            for pdf in pdfs:
                quarter_key = pdf.quarter_key
                if quarter_key not in quarter_groups:
                    quarter_groups[quarter_key] = []
                quarter_groups[quarter_key].append(pdf)
            
            # Resolve conflicts for quarters with multiple PDFs
            for quarter_key, quarter_pdfs in quarter_groups.items():
                if len(quarter_pdfs) > 1:
                    # Sort by priority (lower number = higher priority)
                    quarter_pdfs.sort(key=lambda x: x.priority)
                    best_pdf = quarter_pdfs[0]
                    
                    # Update matrix with best report
                    company_mask = matrix_df['代號'].astype(str) == company_id
                    if quarter_key in matrix_df.columns:
                        matrix_df.loc[company_mask, quarter_key] = best_pdf.report_type
                        conflicts_resolved += 1
                        
                        logger.debug(f"Resolved conflict for {company_id} {quarter_key}: chose {best_pdf.report_type} over {[p.report_type for p in quarter_pdfs[1:]]}")
        
        if conflicts_resolved > 0:
            logger.info(f"   解決 {conflicts_resolved} 個報告類型衝突")
        
        return matrix_df
    
    def add_summary_columns(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """Add summary columns like total reports, coverage percentage"""
        if not self.config.include_summary_sheet:
            return matrix_df
        
        logger.info("📈 新增統計欄位...")
        
        # Calculate summary statistics for each company
        quarter_columns = [col for col in matrix_df.columns if col not in ['代號', '名稱']]
        total_quarters = len(quarter_columns)
        
        summary_data = []
        
        for _, row in matrix_df.iterrows():
            company_id = row['代號']
            company_name = row['名稱']
            
            # Count non-empty quarters
            non_empty_quarters = sum(1 for col in quarter_columns if row[col] != '-')
            coverage_percentage = (non_empty_quarters / total_quarters * 100) if total_quarters > 0 else 0
            
            # Count by report type
            report_types = [row[col] for col in quarter_columns if row[col] != '-']
            individual_count = sum(1 for rt in report_types if rt in ['A12', 'A13'])
            consolidated_count = sum(1 for rt in report_types if rt in ['AI1', 'A1L'])
            
            summary_data.append({
                '代號': company_id,
                '名稱': company_name,
                '總報告數': non_empty_quarters,
                '涵蓋率': f"{coverage_percentage:.1f}%",
                '個別報告': individual_count,
                '合併報告': consolidated_count
            })
        
        # Add summary columns to main matrix
        matrix_with_summary = matrix_df.copy()
        
        for col_name in ['總報告數', '涵蓋率']:
            matrix_with_summary[col_name] = [item[col_name] for item in summary_data]
        
        return matrix_with_summary
    
    def generate_coverage_stats(self, matrix_df: pd.DataFrame, 
                               pdf_data: Dict[str, List[PDFFile]]) -> CoverageStats:
        """Generate comprehensive coverage statistics"""
        logger.info("📊 產生涵蓋率統計...")
        
        # Basic counts
        total_companies = len(matrix_df)
        companies_with_pdfs = len(pdf_data)
        
        # Quarter analysis
        quarter_columns = [col for col in matrix_df.columns if col not in ['代號', '名稱', '總報告數', '涵蓋率']]
        total_quarters = len(quarter_columns)
        
        # Count actual reports
        total_possible_reports = total_companies * total_quarters
        total_actual_reports = sum(len(pdfs) for pdfs in pdf_data.values())
        
        # Calculate coverage percentage
        coverage_percentage = (total_actual_reports / total_possible_reports * 100) if total_possible_reports > 0 else 0
        
        # Analyze report types
        report_type_distribution = {}
        for pdfs in pdf_data.values():
            for pdf in pdfs:
                report_type = pdf.report_type
                report_type_distribution[report_type] = report_type_distribution.get(report_type, 0) + 1
        
        # Find missing quarters by company
        missing_quarters_by_company = {}
        for _, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            missing_quarters = []
            
            for col in quarter_columns:
                if row[col] == '-':
                    missing_quarters.append(col)
            
            if missing_quarters:
                missing_quarters_by_company[company_id] = missing_quarters
        
        # Identify future quarters
        future_quarters = self._identify_future_quarters(quarter_columns)
        
        stats = CoverageStats(
            total_companies=total_companies,
            companies_with_pdfs=companies_with_pdfs,
            total_quarters=total_quarters,
            total_possible_reports=total_possible_reports,
            total_actual_reports=total_actual_reports,
            coverage_percentage=coverage_percentage,
            report_type_distribution=report_type_distribution,
            missing_quarters_by_company=missing_quarters_by_company,
            future_quarters=future_quarters
        )
        
        # Log key statistics
        logger.info(f"   總公司數: {total_companies}")
        logger.info(f"   有PDF公司: {companies_with_pdfs}")
        logger.info(f"   涵蓋率: {coverage_percentage:.1f}%")
        logger.info(f"   報告類型分布: {dict(list(report_type_distribution.items())[:3])}")
        
        return stats
    
    def _discover_quarters_from_pdfs(self, pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """Discover all quarters from PDF data"""
        quarters = set()
        
        for pdfs in pdf_data.values():
            for pdf in pdfs:
                quarters.add(pdf.quarter_key)
        
        return sorted(quarters, reverse=True)
    
    def _create_matrix_cells(self, matrix_df: pd.DataFrame, 
                           pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, MatrixCell]:
        """Create matrix cells for each company-quarter combination"""
        matrix_cells = {}
        
        for company_id, pdfs in pdf_data.items():
            for pdf in pdfs:
                cell_key = f"{company_id}_{pdf.quarter_key}"
                
                if cell_key not in matrix_cells:
                    matrix_cells[cell_key] = MatrixCell()
                
                matrix_cells[cell_key].add_pdf(pdf)
        
        return matrix_cells
    
    def _identify_future_quarters(self, quarter_columns: List[str]) -> List[str]:
        """Identify quarters that are in the future"""
        now = datetime.now()
        current_year = now.year
        current_quarter = (now.month - 1) // 3 + 1
        
        future_quarters = []
        
        for quarter_str in quarter_columns:
            try:
                year, quarter = quarter_str.split(' Q')
                year, quarter = int(year), int(quarter)
                
                if year > current_year or (year == current_year and quarter > current_quarter):
                    future_quarters.append(quarter_str)
            except (ValueError, IndexError):
                continue
        
        return future_quarters
    
    def _log_population_summary(self, matrix_df: pd.DataFrame, 
                               pdf_data: Dict[str, List[PDFFile]]) -> None:
        """Log summary of matrix population"""
        total_cells = len(matrix_df) * (len(matrix_df.columns) - 2)  # Exclude 代號, 名稱
        filled_cells = 0
        
        for _, row in matrix_df.iterrows():
            for col in matrix_df.columns[2:]:  # Skip 代號, 名稱
                if row[col] != '-':
                    filled_cells += 1
        
        fill_percentage = (filled_cells / total_cells * 100) if total_cells > 0 else 0
        
        logger.info(f"   填入狀態: {filled_cells}/{total_cells} 個儲存格 ({fill_percentage:.1f}%)")

class MatrixOptimizer:
    """Helper class for matrix optimization and analysis"""
    
    @staticmethod
    def optimize_column_order(matrix_df: pd.DataFrame, 
                            prioritize_recent: bool = True) -> pd.DataFrame:
        """Optimize column order for better visualization"""
        base_columns = ['代號', '名稱']
        quarter_columns = [col for col in matrix_df.columns if col not in base_columns and 'Q' in col]
        summary_columns = [col for col in matrix_df.columns if col not in base_columns + quarter_columns]
        
        # Sort quarter columns
        def quarter_sort_key(q):
            try:
                year, quarter = q.split(' Q')
                return (int(year), int(quarter))
            except:
                return (0, 0)
        
        sorted_quarters = sorted(quarter_columns, key=quarter_sort_key, reverse=prioritize_recent)
        
        # Rebuild column order
        new_columns = base_columns + sorted_quarters + summary_columns
        
        return matrix_df[new_columns]
    
    @staticmethod
    def compress_matrix_for_display(matrix_df: pd.DataFrame, 
                                  max_quarters: int = 12) -> pd.DataFrame:
        """Compress matrix for display by limiting quarters shown"""
        base_columns = ['代號', '名稱']
        quarter_columns = [col for col in matrix_df.columns if col not in base_columns and 'Q' in col]
        summary_columns = [col for col in matrix_df.columns if col not in base_columns + quarter_columns]
        
        # Take only the most recent quarters
        recent_quarters = quarter_columns[:max_quarters]
        
        # Rebuild with compressed columns
        display_columns = base_columns + recent_quarters + summary_columns
        
        return matrix_df[display_columns]
    
    @staticmethod
    def calculate_data_quality_score(matrix_df: pd.DataFrame) -> float:
        """Calculate overall data quality score for the matrix"""
        total_data_cells = len(matrix_df) * (len(matrix_df.columns) - 2)
        filled_cells = 0
        
        for _, row in matrix_df.iterrows():
            for col in matrix_df.columns[2:]:
                if row[col] != '-':
                    filled_cells += 1
        
        return (filled_cells / total_data_cells) if total_data_cells > 0 else 0.0