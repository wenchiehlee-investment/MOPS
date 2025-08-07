"""
MOPS Sheets Uploader - Report Analyzer
Analyzes PDF files and generates insights and recommendations.
"""

import pandas as pd
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime, timedelta
from collections import defaultdict
import logging

from .models import PDFFile, CoverageStats, StockListChanges, FutureQuarterAnalysis
from .config import MOPSConfig

logger = logging.getLogger(__name__)

class ReportAnalyzer:
    """Analyzes PDF files and generates comprehensive insights"""
    
    def __init__(self, config: MOPSConfig):
        self.config = config
    
    def analyze_coverage(self, matrix_df: pd.DataFrame, 
                        pdf_data: Dict[str, List[PDFFile]]) -> CoverageStats:
        """
        Analyze report coverage statistics
        
        Returns comprehensive coverage analysis including:
        - Overall coverage percentages
        - Company-wise coverage analysis
        - Quarter-wise availability patterns
        - Report type distribution
        """
        logger.info("📊 分析報告涵蓋率...")
        
        # Basic metrics
        total_companies = len(matrix_df)
        companies_with_pdfs = len(pdf_data)
        
        # Quarter analysis
        quarter_columns = [col for col in matrix_df.columns if 'Q' in col and col not in ['代號', '名稱']]
        total_quarters = len(quarter_columns)
        
        # Calculate total possible vs actual reports
        total_possible_reports = total_companies * total_quarters
        total_actual_reports = sum(len(pdfs) for pdfs in pdf_data.values())
        
        # Coverage percentage
        coverage_percentage = (total_actual_reports / total_possible_reports * 100) if total_possible_reports > 0 else 0
        
        # Report type distribution analysis
        report_type_distribution = self._analyze_report_types(pdf_data)
        
        # Missing quarters analysis
        missing_quarters_by_company = self._find_missing_quarters(matrix_df, pdf_data, quarter_columns)
        
        # Future quarters identification
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
        
        # Log key insights
        self._log_coverage_insights(stats)
        
        return stats
    
    def identify_missing_reports(self, matrix_df: pd.DataFrame, 
                               pdf_data: Dict[str, List[PDFFile]]) -> pd.DataFrame:
        """
        Generate detailed list of missing reports by priority
        
        Returns DataFrame with missing report analysis:
        - Company details
        - Missing quarters
        - Priority recommendations
        - Expected report types
        """
        logger.info("🔍 識別缺失報告...")
        
        missing_reports = []
        quarter_columns = [col for col in matrix_df.columns if 'Q' in col and col not in ['代號', '名稱']]
        
        for _, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            company_name = row['名稱']
            
            # Check each quarter for missing reports
            missing_quarters = []
            for quarter in quarter_columns:
                if row[quarter] == '-':
                    missing_quarters.append(quarter)
            
            if missing_quarters:
                # Determine priority based on company's existing report pattern
                priority = self._calculate_download_priority(company_id, pdf_data, missing_quarters)
                expected_type = self._predict_report_type(company_id, pdf_data)
                
                missing_reports.append({
                    '代號': company_id,
                    '名稱': company_name,
                    '缺失季度數': len(missing_quarters),
                    '缺失季度': ', '.join(missing_quarters[:3]) + ('...' if len(missing_quarters) > 3 else ''),
                    '優先級': priority,
                    '預期報告類型': expected_type,
                    '最近季度': missing_quarters[0] if missing_quarters else 'N/A'
                })
        
        missing_df = pd.DataFrame(missing_reports)
        
        if not missing_df.empty:
            # Sort by priority (high to low)
            priority_order = {'高': 3, '中': 2, '低': 1}
            missing_df['優先級數值'] = missing_df['優先級'].map(priority_order)
            missing_df = missing_df.sort_values(['優先級數值', '缺失季度數'], ascending=[False, False])
            missing_df = missing_df.drop('優先級數值', axis=1)
            
            logger.info(f"   發現 {len(missing_df)} 家公司有缺失報告")
            logger.info(f"   高優先級: {len(missing_df[missing_df['優先級'] == '高'])} 家")
            logger.info(f"   中優先級: {len(missing_df[missing_df['優先級'] == '中'])} 家")
        
        return missing_df
    
    def generate_download_suggestions(self, matrix_df: pd.DataFrame,
                                    pdf_data: Dict[str, List[PDFFile]],
                                    stock_changes: Optional[StockListChanges] = None) -> List[str]:
        """
        Generate intelligent download suggestions based on analysis
        
        Returns list of actionable download recommendations
        """
        logger.info("💡 產生下載建議...")
        
        suggestions = []
        
        # 1. New companies without any PDFs
        if stock_changes and stock_changes.added_companies:
            suggestions.append("🆕 新增公司下載建議:")
            for company_id in stock_changes.added_companies[:5]:  # Top 5
                company_name = self._get_company_name(matrix_df, company_id)
                suggestions.append(f"   • {company_id} ({company_name}): 下載所有可用季度")
        
        # 2. Companies with partial coverage - focus on recent quarters
        recent_quarters = self._get_recent_quarters(2)  # Last 2 quarters
        high_priority_missing = self._find_high_priority_missing(matrix_df, pdf_data, recent_quarters)
        
        if high_priority_missing:
            suggestions.append("🎯 高優先級缺失報告:")
            for item in high_priority_missing[:10]:  # Top 10
                suggestions.append(f"   • {item['company_id']} ({item['company_name']}): {item['quarter']} - {item['reason']}")
        
        # 3. Companies with only consolidated reports - suggest individual reports
        individual_candidates = self._find_individual_report_candidates(pdf_data)
        
        if individual_candidates:
            suggestions.append("📋 建議下載個別財報 (目前僅有合併報告):")
            for company_id in individual_candidates[:5]:
                company_name = self._get_company_name(matrix_df, company_id)
                suggestions.append(f"   • {company_id} ({company_name}): 優先下載 A12/A13 個別財報")
        
        # 4. Future quarter analysis and warnings
        future_analysis = self._analyze_future_quarters(pdf_data)
        if future_analysis.has_future_pdfs():
            suggestions.append("⚠️ 未來季度檔案警告:")
            for warning in future_analysis.warnings[:3]:
                suggestions.append(f"   • {warning}")
        
        # 5. Bulk download recommendations
        bulk_suggestions = self._generate_bulk_download_suggestions(matrix_df, pdf_data)
        if bulk_suggestions:
            suggestions.extend(bulk_suggestions)
        
        logger.info(f"   產生 {len([s for s in suggestions if s.startswith('   •')])} 項具體建議")
        
        return suggestions
    
    def analyze_temporal_patterns(self, pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, Any]:
        """
        Analyze temporal patterns in PDF availability
        
        Returns analysis of:
        - Quarterly release patterns
        - Company reporting consistency
        - Seasonal availability trends
        - Late reporting identification
        """
        logger.info("📅 分析時間模式...")
        
        # Group PDFs by quarter for temporal analysis
        quarter_patterns = defaultdict(list)
        company_consistency = {}
        
        for company_id, pdfs in pdf_data.items():
            # Analyze consistency for each company
            quarters = [pdf.quarter_key for pdf in pdfs]
            quarter_patterns[company_id] = sorted(quarters)
            
            # Calculate consistency score (how regularly they report)
            consistency_score = self._calculate_consistency_score(quarters)
            company_consistency[company_id] = consistency_score
        
        # Identify seasonal patterns
        seasonal_analysis = self._analyze_seasonal_patterns(pdf_data)
        
        # Find companies with irregular reporting
        irregular_reporters = [
            company_id for company_id, score in company_consistency.items() 
            if score < 0.5  # Less than 50% consistency
        ]
        
        # Analyze release timing relative to quarter end
        release_timing = self._analyze_release_timing(pdf_data)
        
        analysis = {
            'company_consistency': company_consistency,
            'irregular_reporters': irregular_reporters,
            'seasonal_patterns': seasonal_analysis,
            'release_timing': release_timing,
            'total_companies_analyzed': len(pdf_data),
            'average_consistency': sum(company_consistency.values()) / len(company_consistency) if company_consistency else 0
        }
        
        logger.info(f"   分析 {len(pdf_data)} 家公司的時間模式")
        logger.info(f"   平均一致性: {analysis['average_consistency']:.1%}")
        logger.info(f"   不規律報告公司: {len(irregular_reporters)} 家")
        
        return analysis
    
    def generate_comprehensive_report(self, matrix_df: pd.DataFrame,
                                    pdf_data: Dict[str, List[PDFFile]], 
                                    coverage_stats: CoverageStats,
                                    stock_changes: Optional[StockListChanges] = None) -> Dict[str, Any]:
        """Generate comprehensive analysis report"""
        logger.info("📋 產生綜合分析報告...")
        
        # Collect all analyses
        missing_reports = self.identify_missing_reports(matrix_df, pdf_data)
        download_suggestions = self.generate_download_suggestions(matrix_df, pdf_data, stock_changes)
        temporal_patterns = self.analyze_temporal_patterns(pdf_data)
        
        # Quality assessment
        quality_score = self._calculate_overall_quality_score(coverage_stats, temporal_patterns)
        
        # Generate summary insights
        key_insights = self._generate_key_insights(coverage_stats, missing_reports, temporal_patterns)
        
        report = {
            'summary': {
                'generation_time': datetime.now().isoformat(),
                'total_companies': coverage_stats.total_companies,
                'coverage_percentage': coverage_stats.coverage_percentage,
                'quality_score': quality_score,
                'key_insights': key_insights
            },
            'coverage_analysis': coverage_stats,
            'missing_reports': missing_reports.to_dict('records') if not missing_reports.empty else [],
            'download_suggestions': download_suggestions,
            'temporal_patterns': temporal_patterns,
            'stock_changes': stock_changes,
            'report_metadata': {
                'analyzer_version': '1.0.0',
                'config_used': {
                    'max_years': self.config.max_years,
                    'include_future_quarters': self.config.include_future_quarters,
                    'preferred_types': self.config.preferred_types
                }
            }
        }
        
        logger.info(f"✅ 綜合報告產生完成 (品質分數: {quality_score:.1f}/10)")
        
        return report
    
    # Helper methods
    
    def _analyze_report_types(self, pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, int]:
        """Analyze distribution of report types"""
        type_counts = defaultdict(int)
        
        for pdfs in pdf_data.values():
            for pdf in pdfs:
                type_counts[pdf.report_type] += 1
        
        return dict(type_counts)
    
    def _find_missing_quarters(self, matrix_df: pd.DataFrame, 
                              pdf_data: Dict[str, List[PDFFile]], 
                              quarter_columns: List[str]) -> Dict[str, List[str]]:
        """Find missing quarters for each company"""
        missing_by_company = {}
        
        for _, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            missing_quarters = []
            
            for quarter in quarter_columns:
                if row[quarter] == '-':
                    missing_quarters.append(quarter)
            
            if missing_quarters:
                missing_by_company[company_id] = missing_quarters
        
        return missing_by_company
    
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
    
    def _calculate_download_priority(self, company_id: str, 
                                   pdf_data: Dict[str, List[PDFFile]], 
                                   missing_quarters: List[str]) -> str:
        """Calculate download priority for a company"""
        if company_id not in pdf_data:
            return "高"  # New company, high priority
        
        # Check if missing recent quarters
        recent_quarters = self._get_recent_quarters(2)
        missing_recent = any(q in missing_quarters for q in recent_quarters)
        
        if missing_recent:
            return "高"
        elif len(missing_quarters) > 3:
            return "中"
        else:
            return "低"
    
    def _predict_report_type(self, company_id: str, 
                           pdf_data: Dict[str, List[PDFFile]]) -> str:
        """Predict expected report type for a company"""
        if company_id not in pdf_data:
            return "A12/A13"  # Default to individual
        
        # Analyze existing reports
        existing_types = [pdf.report_type for pdf in pdf_data[company_id]]
        
        if any(rt in ['A12', 'A13'] for rt in existing_types):
            return "A12/A13"
        elif any(rt in ['AI1', 'A1L'] for rt in existing_types):
            return "AI1/A1L"
        else:
            return "未知"
    
    def _get_recent_quarters(self, count: int = 2) -> List[str]:
        """Get list of recent quarters"""
        now = datetime.now()
        quarters = []
        
        for i in range(count):
            quarter = (now.month - 1) // 3 + 1 - i
            year = now.year
            
            if quarter <= 0:
                quarter += 4
                year -= 1
            
            quarters.append(f"{year} Q{quarter}")
        
        return quarters
    
    def _find_high_priority_missing(self, matrix_df: pd.DataFrame,
                                   pdf_data: Dict[str, List[PDFFile]],
                                   recent_quarters: List[str]) -> List[Dict[str, str]]:
        """Find high priority missing reports"""
        high_priority = []
        
        for _, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            company_name = row['名稱']
            
            for quarter in recent_quarters:
                if quarter in matrix_df.columns and row[quarter] == '-':
                    high_priority.append({
                        'company_id': company_id,
                        'company_name': company_name,
                        'quarter': quarter,
                        'reason': '缺失最近季度報告'
                    })
        
        return high_priority
    
    def _find_individual_report_candidates(self, pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """Find companies that only have consolidated reports"""
        candidates = []
        
        for company_id, pdfs in pdf_data.items():
            report_types = [pdf.report_type for pdf in pdfs]
            
            # Has consolidated but no individual reports
            has_consolidated = any(rt in ['AI1', 'A1L'] for rt in report_types)
            has_individual = any(rt in ['A12', 'A13'] for rt in report_types)
            
            if has_consolidated and not has_individual:
                candidates.append(company_id)
        
        return candidates
    
    def _analyze_future_quarters(self, pdf_data: Dict[str, List[PDFFile]]) -> FutureQuarterAnalysis:
        """Analyze PDFs with future quarter dates"""
        analysis = FutureQuarterAnalysis()
        
        for pdfs in pdf_data.values():
            for pdf in pdfs:
                if pdf.is_future_quarter():
                    analysis.future_pdfs.append(pdf)
                    
                    # Check if suspiciously far in future
                    months_ahead = self._calculate_months_ahead(pdf)
                    if months_ahead > self.config.warn_threshold_months:
                        analysis.suspicious_files.append(pdf.filename)
                        analysis.add_warning(f"{pdf.filename} 超前 {months_ahead} 個月")
        
        return analysis
    
    def _generate_bulk_download_suggestions(self, matrix_df: pd.DataFrame,
                                          pdf_data: Dict[str, List[PDFFile]]) -> List[str]:
        """Generate bulk download suggestions"""
        suggestions = []
        
        # Find companies with no PDFs at all
        companies_no_pdfs = []
        for _, row in matrix_df.iterrows():
            company_id = str(row['代號'])
            if company_id not in pdf_data:
                companies_no_pdfs.append(company_id)
        
        if len(companies_no_pdfs) > 5:
            suggestions.append("🔄 批量下載建議:")
            suggestions.append(f"   • {len(companies_no_pdfs)} 家公司完全無PDF，建議批量下載")
            suggestions.append(f"   • 優先下載最近2個季度，再補充歷史資料")
        
        return suggestions
    
    def _calculate_consistency_score(self, quarters: List[str]) -> float:
        """Calculate reporting consistency score for a company"""
        if len(quarters) < 2:
            return 0.0
        
        # Simple consistency: ratio of actual to expected quarters
        # This is a simplified calculation - could be more sophisticated
        sorted_quarters = sorted(quarters)
        expected_quarters = self._get_expected_quarters_between(sorted_quarters[0], sorted_quarters[-1])
        
        return len(quarters) / len(expected_quarters) if expected_quarters else 0.0
    
    def _get_expected_quarters_between(self, start_quarter: str, end_quarter: str) -> List[str]:
        """Get list of expected quarters between start and end"""
        # Simplified implementation
        try:
            start_year, start_q = start_quarter.split(' Q')
            end_year, end_q = end_quarter.split(' Q')
            
            start_year, start_q = int(start_year), int(start_q)
            end_year, end_q = int(end_year), int(end_q)
            
            quarters = []
            year, quarter = start_year, start_q
            
            while year < end_year or (year == end_year and quarter <= end_q):
                quarters.append(f"{year} Q{quarter}")
                quarter += 1
                if quarter > 4:
                    quarter = 1
                    year += 1
            
            return quarters
        except:
            return []
    
    def _analyze_seasonal_patterns(self, pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, Any]:
        """Analyze seasonal patterns in reporting"""
        quarter_counts = {1: 0, 2: 0, 3: 0, 4: 0}
        
        for pdfs in pdf_data.values():
            for pdf in pdfs:
                quarter_counts[pdf.quarter] += 1
        
        total_reports = sum(quarter_counts.values())
        quarter_percentages = {q: (count / total_reports * 100) if total_reports > 0 else 0 
                             for q, count in quarter_counts.items()}
        
        return {
            'quarter_distribution': quarter_counts,
            'quarter_percentages': quarter_percentages,
            'most_common_quarter': max(quarter_counts, key=quarter_counts.get),
            'least_common_quarter': min(quarter_counts, key=quarter_counts.get)
        }
    
    def _analyze_release_timing(self, pdf_data: Dict[str, List[PDFFile]]) -> Dict[str, Any]:
        """Analyze release timing relative to quarter end"""
        # This would analyze file modification dates relative to quarter end dates
        # Simplified implementation for now
        return {
            'average_delay_days': 45,  # Placeholder
            'early_reporters': [],
            'late_reporters': []
        }
    
    def _calculate_overall_quality_score(self, coverage_stats: CoverageStats,
                                       temporal_patterns: Dict[str, Any]) -> float:
        """Calculate overall data quality score (0-10)"""
        # Coverage component (0-5 points)
        coverage_score = (coverage_stats.coverage_percentage / 100) * 5
        
        # Consistency component (0-3 points)
        avg_consistency = temporal_patterns.get('average_consistency', 0)
        consistency_score = avg_consistency * 3
        
        # Completeness component (0-2 points)
        companies_with_data_ratio = coverage_stats.companies_with_pdfs / coverage_stats.total_companies
        completeness_score = companies_with_data_ratio * 2
        
        total_score = coverage_score + consistency_score + completeness_score
        return round(min(total_score, 10), 1)
    
    def _generate_key_insights(self, coverage_stats: CoverageStats,
                             missing_reports: pd.DataFrame,
                             temporal_patterns: Dict[str, Any]) -> List[str]:
        """Generate key insights summary"""
        insights = []
        
        # Coverage insight
        if coverage_stats.coverage_percentage > 80:
            insights.append(f"整體涵蓋率良好 ({coverage_stats.coverage_percentage:.1f}%)")
        elif coverage_stats.coverage_percentage > 50:
            insights.append(f"涵蓋率中等 ({coverage_stats.coverage_percentage:.1f}%)，有改善空間")
        else:
            insights.append(f"涵蓋率偏低 ({coverage_stats.coverage_percentage:.1f}%)，需要大量補充")
        
        # Missing reports insight
        if not missing_reports.empty:
            high_priority_count = len(missing_reports[missing_reports['優先級'] == '高'])
            if high_priority_count > 0:
                insights.append(f"{high_priority_count} 家公司需要優先下載報告")
        
        # Report type insight
        individual_count = coverage_stats.report_type_distribution.get('A12', 0) + \
                          coverage_stats.report_type_distribution.get('A13', 0)
        consolidated_count = coverage_stats.report_type_distribution.get('AI1', 0) + \
                           coverage_stats.report_type_distribution.get('A1L', 0)
        
        if individual_count > consolidated_count:
            insights.append(f"個別財報比例較高 ({individual_count}/{individual_count + consolidated_count})")
        else:
            insights.append(f"合併財報比例較高，建議補充個別財報")
        
        return insights
    
    def _calculate_months_ahead(self, pdf_file: PDFFile) -> int:
        """Calculate how many months ahead this PDF is"""
        now = datetime.now()
        current_year = now.year
        current_quarter = (now.month - 1) // 3 + 1
        
        current_month = current_year * 12 + (current_quarter - 1) * 3
        pdf_month = pdf_file.year * 12 + (pdf_file.quarter - 1) * 3
        
        return max(0, (pdf_month - current_month) // 3)
    
    def _get_company_name(self, matrix_df: pd.DataFrame, company_id: str) -> str:
        """Get company name from matrix"""
        try:
            mask = matrix_df['代號'].astype(str) == company_id
            matches = matrix_df[mask]
            return matches.iloc[0]['名稱'] if not matches.empty else f"未知({company_id})"
        except:
            return f"未知({company_id})"
    
    def _log_coverage_insights(self, stats: CoverageStats) -> None:
        """Log key coverage insights"""
        logger.info(f"📊 涵蓋率分析結果:")
        logger.info(f"   • 總公司數: {stats.total_companies}")
        logger.info(f"   • 有PDF公司: {stats.companies_with_pdfs} ({stats.companies_with_pdfs/stats.total_companies:.1%})")
        logger.info(f"   • 整體涵蓋率: {stats.coverage_percentage:.1f}%")
        logger.info(f"   • 報告類型: {len(stats.report_type_distribution)} 種")
        
        if stats.future_quarters:
            logger.info(f"   ⚠️ 未來季度: {len(stats.future_quarters)} 個")