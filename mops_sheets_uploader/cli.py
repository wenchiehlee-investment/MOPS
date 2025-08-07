"""
MOPS Sheets Uploader - Command Line Interface
Provides command-line access to MOPS Sheets Uploader functionality.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Optional
import logging

from .main import MOPSSheetsUploader, QuickStart
from .config import MOPSConfig, load_config
from .models import ProcessingResult

def create_parser() -> argparse.ArgumentParser:
    """Create command line argument parser"""
    parser = argparse.ArgumentParser(
        description="MOPS Sheets Uploader - Upload PDF matrix to Google Sheets",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic usage with existing Google Sheet
  python -m mops_sheets_uploader --sheet-id YOUR_SHEET_ID

  # CSV-only export
  python -m mops_sheets_uploader --csv-only --output-dir ./reports

  # Analysis only (no upload)
  python -m mops_sheets_uploader --analyze-only --output report.json

  # Custom directories and settings
  python -m mops_sheets_uploader \\
    --downloads-dir ./data/downloads \\
    --stock-csv ./data/stocks.csv \\
    --max-years 5 \\
    --worksheet-name "財報矩陣"

  # Test connection only
  python -m mops_sheets_uploader --test-connection
        """
    )
    
    # Input/Output options
    io_group = parser.add_argument_group('Input/Output Options')
    io_group.add_argument(
        '--downloads-dir',
        default='downloads',
        help='Downloads directory path (default: downloads)'
    )
    io_group.add_argument(
        '--stock-csv',
        default='StockID_TWSE_TPEX.csv',
        help='Stock CSV file path (default: StockID_TWSE_TPEX.csv)'
    )
    io_group.add_argument(
        '--output-dir',
        default='data/reports',
        help='Output directory for CSV files (default: data/reports)'
    )
    
    # Google Sheets options
    sheets_group = parser.add_argument_group('Google Sheets Options')
    sheets_group.add_argument(
        '--sheet-id',
        help='Google Sheets ID for upload'
    )
    sheets_group.add_argument(
        '--worksheet-name',
        default='MOPS下載狀態',
        help='Worksheet name (default: MOPS下載狀態)'
    )
    sheets_group.add_argument(
        '--test-connection',
        action='store_true',
        help='Test Google Sheets connection and exit'
    )
    
    # Matrix options
    matrix_group = parser.add_argument_group('Matrix Options')
    matrix_group.add_argument(
        '--max-years',
        type=int,
        default=3,
        help='Maximum years of quarters to include (default: 3)'
    )
    matrix_group.add_argument(
        '--include-future',
        action='store_true',
        default=True,
        help='Include future quarters in matrix (default: True)'
    )
    matrix_group.add_argument(
        '--no-future',
        action='store_true',
        help='Exclude future quarters from matrix'
    )
    
    # Export options
    export_group = parser.add_argument_group('Export Options')
    export_group.add_argument(
        '--csv-only',
        action='store_true',
        help='Export to CSV only (skip Google Sheets)'
    )
    export_group.add_argument(
        '--csv-backup',
        action='store_true',
        default=True,
        help='Create CSV backup even when uploading to Sheets (default: True)'
    )
    export_group.add_argument(
        '--no-csv-backup',
        action='store_true',
        help='Disable CSV backup'
    )
    
    # Analysis options
    analysis_group = parser.add_argument_group('Analysis Options')
    analysis_group.add_argument(
        '--analyze-only',
        action='store_true',
        help='Generate analysis report only (no upload)'
    )
    analysis_group.add_argument(
        '--no-suggestions',
        action='store_true',
        help='Disable download suggestions'
    )
    analysis_group.add_argument(
        '--output',
        help='Output file for analysis report (JSON format)'
    )
    
    # Logging options
    log_group = parser.add_argument_group('Logging Options')
    log_group.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='Logging level (default: INFO)'
    )
    log_group.add_argument(
        '--quiet',
        action='store_true',
        help='Quiet mode (minimal output)'
    )
    log_group.add_argument(
        '--verbose',
        action='store_true',
        help='Verbose mode (detailed output)'
    )
    
    # Configuration options
    config_group = parser.add_argument_group('Configuration Options')
    config_group.add_argument(
        '--config',
        help='Configuration file path (YAML or JSON)'
    )
    config_group.add_argument(
        '--save-config',
        help='Save current configuration to file'
    )
    
    return parser

def setup_logging(args) -> None:
    """Setup logging based on command line arguments"""
    if args.quiet:
        level = logging.WARNING
    elif args.verbose:
        level = logging.DEBUG
    else:
        level = getattr(logging, args.log_level)
    
    # Setup basic logging
    logging.basicConfig(
        level=level,
        format='%(message)s' if args.quiet else '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    )
    
    # Reduce noise from external libraries
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    logging.getLogger('requests').setLevel(logging.WARNING)
    logging.getLogger('gspread').setLevel(logging.WARNING)

def load_configuration(args) -> MOPSConfig:
    """Load configuration from arguments and files"""
    # Start with config file if provided
    if args.config:
        config = MOPSConfig.from_file(args.config)
    else:
        config = load_config()
    
    # Override with command line arguments
    if args.downloads_dir != 'downloads':
        config.downloads_dir = args.downloads_dir
    
    if args.stock_csv != 'StockID_TWSE_TPEX.csv':
        config.stock_csv_path = args.stock_csv
    
    if args.output_dir != 'data/reports':
        config.output_dir = args.output_dir
    
    if args.sheet_id:
        config.google_sheet_id = args.sheet_id
    
    if args.worksheet_name != 'MOPS下載狀態':
        config.worksheet_name = args.worksheet_name
    
    if args.max_years != 3:
        config.max_years = args.max_years
    
    # Handle boolean flags
    if args.no_future:
        config.include_future_quarters = False
    elif args.include_future:
        config.include_future_quarters = True
    
    if args.csv_only:
        config.csv_backup = True
        config.google_sheet_id = None  # Disable Sheets upload
    
    if args.no_csv_backup:
        config.csv_backup = False
    elif args.csv_backup:
        config.csv_backup = True
    
    if args.no_suggestions:
        config.auto_suggest_downloads = False
    
    return config

def handle_test_connection(config: MOPSConfig) -> int:
    """Handle connection test"""
    print("🔍 測試 Google Sheets 連線...")
    
    uploader = MOPSSheetsUploader(config=config)
    
    if uploader.test_connection():
        print("✅ Google Sheets 連線成功")
        if config.google_sheet_id:
            print(f"📊 目標試算表: https://docs.google.com/spreadsheets/d/{config.google_sheet_id}")
        return 0
    else:
        print("❌ Google Sheets 連線失敗")
        print("💡 請檢查:")
        print("   • GOOGLE_SHEETS_CREDENTIALS 環境變數")
        print("   • GOOGLE_SHEET_ID 環境變數")
        print("   • 服務帳戶權限設定")
        return 1

def handle_analyze_only(config: MOPSConfig, output_file: Optional[str]) -> int:
    """Handle analysis-only mode"""
    print("📊 執行分析模式...")
    
    uploader = MOPSSheetsUploader(config=config)
    
    try:
        report = uploader.generate_report(include_analysis=True)
        
        if output_file:
            # Save to file
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2, default=str)
            print(f"📄 分析報告已儲存: {output_file}")
        else:
            # Print summary to console
            summary = report.get('summary', {})
            print("📊 分析摘要:")
            print(f"   • 總公司數: {summary.get('total_companies', 'N/A')}")
            print(f"   • 涵蓋率: {summary.get('coverage_percentage', 'N/A'):.1f}%")
            print(f"   • 品質分數: {summary.get('quality_score', 'N/A')}/10")
            
            if 'key_insights' in summary:
                print("💡 關鍵洞察:")
                for insight in summary['key_insights']:
                    print(f"   • {insight}")
        
        return 0
        
    except Exception as e:
        print(f"❌ 分析失敗: {e}")
        return 1

def handle_full_process(config: MOPSConfig) -> int:
    """Handle full processing pipeline"""
    uploader = MOPSSheetsUploader(config=config)
    
    try:
        result = uploader.run()
        
        if result.success:
            print("\n🎉 處理完成!")
            
            if result.matrix_uploaded and result.sheets_url:
                print(f"🔗 Google Sheets: {result.sheets_url}")
            
            if result.csv_exported and result.csv_path:
                print(f"💾 CSV 檔案: {result.csv_path}")
            
            return 0
        else:
            print(f"\n❌ 處理失敗: {result.error_message}")
            return 1
            
    except Exception as e:
        print(f"❌ 執行錯誤: {e}")
        return 1

def save_configuration(config: MOPSConfig, config_path: str) -> None:
    """Save configuration to file"""
    import yaml
    
    config_dict = {
        'downloads_dir': config.downloads_dir,
        'stock_csv_path': config.stock_csv_path,
        'output_dir': config.output_dir,
        'max_years': config.max_years,
        'worksheet_name': config.worksheet_name,
        'csv_backup': config.csv_backup,
        'include_future_quarters': config.include_future_quarters,
        'auto_suggest_downloads': config.auto_suggest_downloads,
        'preferred_types': config.preferred_types,
        'acceptable_types': config.acceptable_types,
        'excluded_types': config.excluded_types
    }
    
    with open(config_path, 'w', encoding='utf-8') as f:
        if config_path.endswith('.json'):
            json.dump(config_dict, f, ensure_ascii=False, indent=2)
        else:
            yaml.dump(config_dict, f, allow_unicode=True, default_flow_style=False)
    
    print(f"💾 設定已儲存: {config_path}")

def main() -> int:
    """Main CLI entry point"""
    parser = create_parser()
    args = parser.parse_args()
    
    # Setup logging first
    setup_logging(args)
    
    try:
        # Load configuration
        config = load_configuration(args)
        
        # Save configuration if requested
        if args.save_config:
            save_configuration(config, args.save_config)
            return 0
        
        # Handle different modes
        if args.test_connection:
            return handle_test_connection(config)
        elif args.analyze_only:
            return handle_analyze_only(config, args.output)
        else:
            return handle_full_process(config)
            
    except KeyboardInterrupt:
        print("\n⏹️ 使用者中斷操作")
        return 130
    except Exception as e:
        print(f"❌ 程式錯誤: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

def quick_commands():
    """Quick command shortcuts"""
    import sys
    
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == 'quick-csv':
            # Quick CSV export
            print("🚀 快速 CSV 匯出...")
            result = QuickStart.csv_only_export()
            if result.success:
                print(f"✅ 完成! 檔案: {result.csv_path}")
                return 0
            else:
                print(f"❌ 失敗: {result.error_message}")
                return 1
        
        elif command == 'quick-analyze':
            # Quick analysis
            print("📊 快速分析...")
            report = QuickStart.analyze_only()
            summary = report.get('summary', {})
            print(f"📈 涵蓋率: {summary.get('coverage_percentage', 'N/A'):.1f}%")
            print(f"🏢 公司數: {summary.get('total_companies', 'N/A')}")
            return 0
    
    # Default to full CLI
    return main()

if __name__ == '__main__':
    sys.exit(main())