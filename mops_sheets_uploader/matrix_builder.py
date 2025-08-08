"""
MOPS Sheets Uploader - Matrix Builder (Enhanced for Multiple Report Types)
Constructs the main data matrix combining stock data and PDF information.
"""

import pandas as pd
from typing import Dict, List, Optional, Tuple, Set
from datetime import datetime
import logging

from .models import (
    PDFFile, MatrixCell, CoverageStats, StockListChanges,
    analyze_report_type_combinations, get_report_type_category_stats,
    REPORT_TYPE_CATEGORIES
)
from .config import MOPSConfig

logger = logging.getLogger(__name__)

class MatrixBuilder:
    """Builds the main data matrix for Google Sheets upload with enhanced multiple report type support"""
    
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
        logger.info("🏗️ 建構增強型矩陣資料...")
        
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
        logger.info(f"   📊 支援多重報告類型顯示: {'是' if self.config.show_all_report_types else '否'}")
        
        return matrix_df
    
    def populate_pdf_status(self, matrix_df: pd.DataFrame, 
                           pdf_data: Dict[str, List[PDFFile]]) -> pd.DataFrame:
        """
        Fill matrix with PDF availability status for all companies
        Enhanced to support multiple report types per cell
        
        Handles:
        - Multiple report types per quarter (A12/A13/AI1)
        - Existing companies with PDFs (normal processing)
        - New companies without PDFs (shows "-" for all quarters)  
        - Companies with partial PDF coverage
        - Configurable display options (all types vs best only)
        """
        logger.info("📊 填入增強型 PDF 狀態資料...")
        
        # Create enhanced matrix cells for each company-quarter combination
        matrix_cells = self._create_enhanced_matrix_cells(matrix_df, pdf_data)
        
        # Fill matrix with PDF status using enhanced display logic
        filled_matrix = matrix_df.copy()
        
        cells_with_multiple_types = 0
        total_filled_cells = 0
        
        for idx, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            
            # Skip base columns (代號, 名稱)
            for col in matrix_df.columns[2:]:
                quarter_key = col
                cell_key = f"{company_id}_{quarter_key}"
                
                if cell_key in matrix_cells:
                    cell = matrix_cells[cell_key]
                    
                    # Get display value based on configuration
                    display_value = cell.get_display_value(
                        show_all_types=getattr(self.config, 'show_all_report_types', True),
                        separator=getattr(self.config, 'report_type_separator', '/'),
                        max_types=getattr(self.config, 'max_display_types', 5)
                    )
                    
                    filled_matrix.loc[idx, col] = display_value
                    total_filled_cells += 1
                    
                    # Count cells with multiple types for statistics
                    if len(cell.report_types) > 1:
                        cells_with_multiple_types += 1
                        
                else:
                    filled_matrix.loc[idx, col] = '-'
        
        # Log enhanced summary statistics
        self._log_enhanced_population_summary(
            filled_matrix, pdf_data, matrix_cells, 
            cells_with_multiple_types, total_filled_cells
        )
        
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
    
    def apply_enhanced_categorization(self, matrix_df: pd.DataFrame, 
                                    pdf_data: Dict[str, List[PDFFile]]) -> pd.DataFrame:
        """
        Apply enhanced categorization instead of priority-based conflict resolution
        This replaces the old apply_priority_rules method
        
        Instead of resolving conflicts, this method:
        1. Groups report types by category
        2. Applies intelligent display logic
        3. Maintains all information while keeping display clean
        """
        logger.info("🎯 套用增強型報告分類邏輯...")
        
        enhanced_matrix = matrix_df.copy()
        categorization_applied = 0
        
        for company_id, pdfs in pdf_data.items():
            # Group PDFs by quarter
            quarter_groups = {}
            for pdf in pdfs:
                quarter_key = pdf.quarter_key
                if quarter_key not in quarter_groups:
                    quarter_groups[quarter_key] = []
                quarter_groups[quarter_key].append(pdf)
            
            # Apply categorization for quarters with multiple PDFs
            for quarter_key, quarter_pdfs in quarter_groups.items():
                if len(quarter_pdfs) > 1:
                    # Create categorized display
                    categorized_display = self._create_categorized_display(quarter_pdfs)
                    
                    # Update matrix with categorized display
                    company_mask = enhanced_matrix['代號'].astype(str) == company_id
                    if quarter_key in enhanced_matrix.columns:
                        enhanced_matrix.loc[company_mask, quarter_key] = categorized_display
                        categorization_applied += 1
                        
                        logger.debug(f"Enhanced display for {company_id} {quarter_key}: {categorized_display}")
        
        if categorization_applied > 0:
            logger.info(f"   套用增強分類: {categorization_applied} 個多重報告類型季度")
        
        return enhanced_matrix
    
    def add_summary_columns(self, matrix_df: pd.DataFrame) -> pd.DataFrame:
        """Add enhanced summary columns including multiple report type statistics"""
        if not self.config.include_summary_sheet:
            return matrix_df
        
        logger.info("📈 新增增強型統計欄位...")
        
        # Calculate enhanced summary statistics for each company
        quarter_columns = [col for col in matrix_df.columns if col not in ['代號', '名稱']]
        total_quarters = len(quarter_columns)
        
        summary_data = []
        
        for _, row in matrix_df.iterrows():
            company_id = row['代號']
            company_name = row['名稱']
            
            # Count non-empty quarters
            non_empty_quarters = sum(1 for col in quarter_columns if row[col] != '-')
            coverage_percentage = (non_empty_quarters / total_quarters * 100) if total_quarters > 0 else 0
            
            # Analyze report types (enhanced)
            all_report_values = [row[col] for col in quarter_columns if row[col] != '-']
            
            # Count by enhanced categories
            individual_count = sum(1 for val in all_report_values 
                                 if self._contains_individual_reports(val))
            consolidated_count = sum(1 for val in all_report_values 
                                   if self._contains_consolidated_reports(val))
            multiple_types_count = sum(1 for val in all_report_values 
                                     if '/' in val)  # Contains multiple types
            
            # Calculate diversity score (how many different report type patterns)
            unique_patterns = len(set(all_report_values))
            diversity_score = (unique_patterns / non_empty_quarters) if non_empty_quarters > 0 else 0
            
            summary_data.append({
                '代號': company_id,
                '名稱': company_name,
                '總報告數': non_empty_quarters,
                '涵蓋率': f"{coverage_percentage:.1f}%",
                '個別報告': individual_count,
                '合併報告': consolidated_count,
                '多重類型': multiple_types_count,
                '類型多樣性': f"{diversity_score:.2f}"
            })
        
        # Add enhanced summary columns to main matrix
        matrix_with_summary = matrix_df.copy()
        
        for col_name in ['總報告數', '涵蓋率', '個別報告', '合併報告', '多重類型', '類型多樣性']:
            matrix_with_summary[col_name] = [item[col_name] for item in summary_data]
        
        return matrix_with_summary
    
    def generate_enhanced_coverage_stats(self, matrix_df: pd.DataFrame, 
                                       pdf_data: Dict[str, List[PDFFile]]) -> CoverageStats:
        """Generate comprehensive coverage statistics with enhanced multiple type analysis"""
        logger.info("📊 產生增強型涵蓋率統計...")
        
        # Basic counts
        total_companies = len(matrix_df)
        companies_with_pdfs = len(pdf_data)
        
        # Quarter analysis
        quarter_columns = [col for col in matrix_df.columns if col not in ['代號', '名稱', '總報告數', '涵蓋率', '個別報告', '合併報告', '多重類型', '類型多樣性']]
        total_quarters = len(quarter_columns)
        
        # Count actual reports
        total_possible_reports = total_companies * total_quarters
        total_actual_reports = sum(len(pdfs) for pdfs in pdf_data.values())
        
        # Calculate coverage percentage
        coverage_percentage = (total_actual_reports / total_possible_reports * 100) if total_possible_reports > 0 else 0
        
        # Enhanced analysis: report type distribution
        report_type_distribution = {}
        for pdfs in pdf_data.values():
            for pdf in pdfs:
                report_type = pdf.report_type
                report_type_distribution[report_type] = report_type_distribution.get(report_type, 0) + 1
        
        # Enhanced analysis: create matrix cells for combination analysis
        matrix_cells = self._create_enhanced_matrix_cells(matrix_df, pdf_data)
        
        # Count cells with multiple types
        cells_with_multiple_types = sum(1 for cell in matrix_cells.values() 
                                       if cell.has_pdf and len(cell.report_types) > 1)
        
        # Analyze report type combinations
        report_type_combinations = analyze_report_type_combinations(matrix_cells)
        
        # Analyze category distribution
        category_stats = get_report_type_category_stats(matrix_cells)
        
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
        
        # Create enhanced coverage stats
        stats = CoverageStats(
            total_companies=total_companies,
            companies_with_pdfs=companies_with_pdfs,
            total_quarters=total_quarters,
            total_possible_reports=total_possible_reports,
            total_actual_reports=total_actual_reports,
            coverage_percentage=coverage_percentage,
            report_type_distribution=report_type_distribution,
            missing_quarters_by_company=missing_quarters_by_company,
            future_quarters=future_quarters,
            # Enhanced fields
            cells_with_multiple_types=cells_with_multiple_types,
            report_type_combinations=report_type_combinations,
            category_distribution=category_stats
        )
        
        # Log enhanced statistics
        self._log_enhanced_coverage_insights(stats)
        
        return stats
    
    # Enhanced helper methods
    
    def _create_enhanced_matrix_cells(self, matrix_df: pd.DataFrame, 
                                    pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, MatrixCell]:
        """Create enhanced matrix cells that support multiple report types"""
        matrix_cells = {}
        
        for company_id, pdfs in pdf_data.items():
            for pdf in pdfs:
                cell_key = f"{company_id}_{pdf.quarter_key}"
                
                if cell_key not in matrix_cells:
                    matrix_cells[cell_key] = MatrixCell()
                
                matrix_cells[cell_key].add_pdf(pdf)
        
        return matrix_cells
    
    def _create_categorized_display(self, pdfs: List[PDFFile]) -> str:
        """
        Create intelligent categorized display for multiple PDFs
        
        Logic:
        1. Group by category (individual, consolidated, etc.)
        2. Show highest priority from each category
        3. Use intelligent ordering and separators
        """
        categories = {
            'individual': [],
            'consolidated': [],
            'generic': [],
            'other': []
        }
        
        # Categorize PDFs
        for pdf in pdfs:
            if pdf.report_type in REPORT_TYPE_CATEGORIES['individual']:
                categories['individual'].append(pdf.report_type)
            elif pdf.report_type in REPORT_TYPE_CATEGORIES['consolidated']:
                categories['consolidated'].append(pdf.report_type)
            elif pdf.report_type in REPORT_TYPE_CATEGORIES['generic']:
                categories['generic'].append(pdf.report_type)
            else:
                categories['other'].append(pdf.report_type)
        
        # Build display string by priority
        display_parts = []
        
        # Individual reports first (highest priority)
        if categories['individual']:
            individual_types = sorted(set(categories['individual']))
            display_parts.append('/'.join(individual_types))
        
        # Consolidated reports second
        if categories['consolidated']:
            consolidated_types = sorted(set(categories['consolidated']))
            display_parts.append('/'.join(consolidated_types))
        
        # Generic and other reports last
        for category in ['generic', 'other']:
            if categories[category]:
                category_types = sorted(set(categories[category]))
                display_parts.append('/'.join(category_types))
        
        # Use different separator for categories vs types within categories
        category_separator = getattr(self.config, 'category_separator', '/')
        return category_separator.join(display_parts)
    
    def _contains_individual_reports(self, value: str) -> bool:
        """Check if value contains individual report types"""
        individual_types = REPORT_TYPE_CATEGORIES['individual']
        return any(rtype in value for rtype in individual_types)
    
    def _contains_consolidated_reports(self, value: str) -> bool:
        """Check if value contains consolidated report types"""
        consolidated_types = REPORT_TYPE_CATEGORIES['consolidated']
        return any(rtype in value for rtype in consolidated_types)
    
    def _discover_quarters_from_pdfs(self, pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """Discover all quarters from PDF data"""
        quarters = set()
        
        for pdfs in pdf_data.values():
            for pdf in pdfs:
                quarters.add(pdf.quarter_key)
        
        return sorted(quarters, reverse=True)
    
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
    
    def _log_enhanced_population_summary(self, matrix_df: pd.DataFrame, 
                                       pdf_data: Dict[str, List[PDFFile]],
                                       matrix_cells: Dict[str, MatrixCell],
                                       cells_with_multiple_types: int,
                                       total_filled_cells: int) -> None:
        """Log enhanced summary of matrix population"""
        total_cells = len(matrix_df) * (len(matrix_df.columns) - 2)  # Exclude 代號, 名稱
        fill_percentage = (total_filled_cells / total_cells * 100) if total_cells > 0 else 0
        multiple_types_percentage = (cells_with_multiple_types / total_filled_cells * 100) if total_filled_cells > 0 else 0
        
        logger.info(f"   填入狀態: {total_filled_cells}/{total_cells} 個儲存格 ({fill_percentage:.1f}%)")
        logger.info(f"   多重類型: {cells_with_multiple_types}/{total_filled_cells} 個儲存格 ({multiple_types_percentage:.1f}%)")
        
        # Log most common combinations
        if cells_with_multiple_types > 0:
            combinations = analyze_report_type_combinations(matrix_cells)
            top_combinations = sorted(combinations.items(), key=lambda x: x[1], reverse=True)[:3]
            logger.info(f"   常見組合: {', '.join([f'{combo} ({count}次)' for combo, count in top_combinations])}")
    
    def _log_enhanced_coverage_insights(self, stats: CoverageStats) -> None:
        """Log enhanced coverage insights"""
        logger.info(f"📊 增強型涵蓋率分析結果:")
        logger.info(f"   • 總公司數: {stats.total_companies}")
        logger.info(f"   • 有PDF公司: {stats.companies_with_pdfs} ({stats.companies_with_pdfs/stats.total_companies:.1%})")
        logger.info(f"   • 整體涵蓋率: {stats.coverage_percentage:.1f}%")
        logger.info(f"   • 多重類型比例: {stats.multiple_types_percentage:.1f}%")
        logger.info(f"   • 報告類型: {len(stats.report_type_distribution)} 種")
        
        if stats.report_type_combinations:
            top_combo = max(stats.report_type_combinations.items(), key=lambda x: x[1])
            logger.info(f"   • 最常見組合: {top_combo[0]} ({top_combo[1]} 次)")
        
        if stats.future_quarters:
            logger.info(f"   ⚠️ 未來季度: {len(stats.future_quarters)} 個")

# Compatibility alias for the old method name
MatrixBuilder.generate_coverage_stats = MatrixBuilder.generate_enhanced_coverage_stats
MatrixBuilder.apply_priority_rules = MatrixBuilder.apply_enhanced_categorization

class MatrixOptimizer:
    """Helper class for matrix optimization and analysis with enhanced multiple type support"""
    
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
    def calculate_enhanced_data_quality_score(matrix_df: pd.DataFrame, 
                                            coverage_stats: CoverageStats) -> float:
        """Calculate enhanced overall data quality score for the matrix"""
        # Base coverage score (0-4 points)
        coverage_score = (coverage_stats.coverage_percentage / 100) * 4
        
        # Multiple types diversity bonus (0-2 points)
        diversity_score = min((coverage_stats.multiple_types_percentage / 50) * 2, 2)
        
        # Category balance score (0-2 points) - bonus for having both individual and consolidated
        category_balance = 0
        if coverage_stats.category_distribution.get('individual_only', 0) > 0 and \
           coverage_stats.category_distribution.get('consolidated_only', 0) > 0:
            category_balance = 1
        if coverage_stats.category_distribution.get('mixed_types', 0) > 0:
            category_balance = 2
        
        # Completeness score (0-2 points)
        companies_with_data_ratio = coverage_stats.companies_with_pdfs / coverage_stats.total_companies
        completeness_score = companies_with_data_ratio * 2
        
        total_score = coverage_score + diversity_score + category_balance + completeness_score
        return round(min(total_score, 10), 1)