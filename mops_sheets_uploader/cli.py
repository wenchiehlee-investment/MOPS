"""
Complete Enhanced CLI for MOPS Sheets Uploader v1.1.1
Replace your existing cli.py with this complete version
"""

import argparse
import json
import sys
import os
from pathlib import Path
from typing import Optional
import logging

from .main import MOPSSheetsUploader, QuickStart
from .config import MOPSConfig, load_config, create_font_config_preset
from .models import ProcessingResult

def create_parser() -> argparse.ArgumentParser:
    """Create enhanced command line argument parser with v1.1.1 font support"""
    parser = argparse.ArgumentParser(
        description="MOPS Sheets Uploader v1.1.1 - Upload PDF matrix to Google Sheets with font support",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
v1.1.1 Enhanced Examples:

  # Basic upload with professional font
  python -m mops_sheets_uploader --upload --font-preset large

  # Custom font configuration
  python -m mops_sheets_uploader --upload \\
    --font-size 16 --header-font-size 18 --bold-headers --bold-company-info

  # Multiple report types with large font for presentation
  python -m mops_sheets_uploader --upload \\
    --font-preset extra_large --show-all-types --type-separator "/"

  # Analysis with multiple type support
  python -m mops_sheets_uploader --analyze \\
    --show-all-types --output analysis_v1.1.1.json

  # CSV-only export with specific font for consistency
  python -m mops_sheets_uploader --csv-only --font-preset large

  # Test connection
  python -m mops_sheets_uploader --test

Font Presets Available:
  - small: 10pt (compact)
  - normal: 12pt (standard)  
  - large: 14pt (professional, recommended)
  - extra_large: 16pt (enhanced readability)
  - huge: 20pt (presentation mode)
        """
    )
    
    # Action arguments (mutually exclusive)
    action_group = parser.add_mutually_exclusive_group(required=True)
    action_group.add_argument('--upload', action='store_true', help='上傳到 Google Sheets')
    action_group.add_argument('--csv-only', action='store_true', help='僅匯出 CSV 檔案')
    action_group.add_argument('--analyze', action='store_true', help='僅分析涵蓋率')
    action_group.add_argument('--test', action='store_true', help='測試 Google Sheets 連線')
    
    # Input/Output options
    io_group = parser.add_argument_group('Input/Output Options')
    io_group.add_argument('--downloads-dir', default='downloads', help='Downloads 目錄路徑')
    io_group.add_argument('--stock-csv', default='StockID_TWSE_TPEX.csv', help='股票清單 CSV 檔案')
    io_group.add_argument('--output-dir', default='data/reports', help='CSV 輸出目錄')
    
    # Google Sheets options
    sheets_group = parser.add_argument_group('Google Sheets Options')
    sheets_group.add_argument('--sheet-id', help='Google Sheets ID for upload')
    sheets_group.add_argument('--worksheet-name', default='MOPS下載狀態', help='工作表名稱')
    
    # v1.1.1 Enhanced Font and Formatting Options
    font_group = parser.add_argument_group('v1.1.1 Font and Formatting Options')
    font_group.add_argument('--font-preset', 
                           choices=['small', 'normal', 'large', 'extra_large', 'huge'],
                           help='字體預設組合 (會覆蓋其他字體設定)')
    font_group.add_argument('--font-size', type=int, 
                           help='字體大小 pt (8-72, 預設: 14)')
    font_group.add_argument('--header-font-size', type=int, 
                           help='標題字體大小 pt (8-72, 預設: 與 --font-size 相同)')
    font_group.add_argument('--bold-headers', action='store_true',
                           help='標題使用粗體 (預設: 是)')
    font_group.add_argument('--no-bold-headers', action='store_true',
                           help='標題不使用粗體')
    font_group.add_argument('--bold-company-info', action='store_true',
                           help='公司資訊欄位使用粗體 (預設: 是)')
    font_group.add_argument('--no-bold-company-info', action='store_true',
                           help='公司資訊欄位不使用粗體')
    
    # v1.1.1 Enhanced Multiple Report Type Options
    type_group = parser.add_argument_group('v1.1.1 Multiple Report Type Display Options')
    type_group.add_argument('--show-all-types', action='store_true',
                           help='顯示所有報告類型 (如: A12/A13/AI1) (預設: 是)')
    type_group.add_argument('--show-best-only', action='store_true',
                           help='僅顯示最佳報告類型 (傳統模式)')
    type_group.add_argument('--type-separator', default='/',
                           help='報告類型分隔符號 (預設: /)')
    type_group.add_argument('--max-display-types', type=int, default=5,
                           help='最大顯示類型數量 (預設: 5)')
    type_group.add_argument('--categorized-display', action='store_true',
                           help='使用分類顯示模式 (個別→合併)')
    
    # Matrix options
    matrix_group = parser.add_argument_group('Matrix Options')
    matrix_group.add_argument('--max-years', type=int, default=3, help='最大年數範圍')
    matrix_group.add_argument('--include-future', action='store_true', default=True, help='包含未來季度')
    matrix_group.add_argument('--no-future', action='store_true', help='排除未來季度')
    
    # Export options
    export_group = parser.add_argument_group('Export Options')
    export_group.add_argument('--csv-backup', action='store_true', default=True, help='即使上傳 Sheets 也建立 CSV 備份')
    export_group.add_argument('--no-csv-backup', action='store_true', help='停用 CSV 備份')
    
    # Analysis options
    analysis_group = parser.add_argument_group('Analysis Options')
    analysis_group.add_argument('--no-suggestions', action='store_true', help='停用下載建議')
    analysis_group.add_argument('--output', help='分析報告輸出檔案 (JSON 格式)')
    analysis_group.add_argument('--include-type-analysis', action='store_true', default=True,
                               help='包含多重類型分析 (預設: 是)')
    
    # Logging options
    log_group = parser.add_argument_group('Logging Options')
    log_group.add_argument('--log-level', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'], default='INFO', help='記錄層級')
    log_group.add_argument('--quiet', action='store_true', help='安靜模式 (最少輸出)')
    log_group.add_argument('--verbose', action='store_true', help='詳細模式 (詳細輸出)')
    
    # Configuration options
    config_group = parser.add_argument_group('Configuration Options')
    config_group.add_argument('--config', help='設定檔路徑 (YAML 或 JSON)')
    config_group.add_argument('--save-config', help='儲存目前設定到檔案')
    config_group.add_argument('--print-presets', action='store_true', help='列印可用的字體預設')
    
    return parser

def load_configuration(args) -> MOPSConfig:
    """Load configuration from arguments and files with v1.1.1 font support"""
    # Start with config file if provided
    if args.config:
        config = MOPSConfig.from_file(args.config)
    else:
        config = load_config()
    
    # v1.1.1: Apply font preset first (if specified)
    if args.font_preset:
        config.apply_font_preset(args.font_preset)
        if not args.quiet:
            preset_info = create_font_config_preset(args.font_preset)
            print(f"📝 套用字體預設: {args.font_preset} ({preset_info['font_size']}pt)")
    
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
    
    # v1.1.1: Apply font configuration from command line
    if args.font_size:
        config.font_size = args.font_size
    
    if args.header_font_size:
        config.header_font_size = args.header_font_size
    elif args.font_size and not args.font_preset:
        # Default header font size to content font size if not specified
        config.header_font_size = args.font_size
    
    if args.no_bold_headers:
        config.bold_headers = False
    elif args.bold_headers:
        config.bold_headers = True
    
    if args.no_bold_company_info:
        config.bold_company_info = False
    elif args.bold_company_info:
        config.bold_company_info = True
    
    # v1.1.1: Multiple report type settings
    if args.show_best_only:
        config.show_all_report_types = False
    elif args.show_all_types:
        config.show_all_report_types = True
    
    if args.type_separator != '/':
        config.report_type_separator = args.type_separator
    
    if args.max_display_types != 5:
        config.max_display_types = args.max_display_types
    
    if args.categorized_display:
        config.use_categorized_display = True
    
    # Handle boolean flags
    if args.no_future:
        config.include_future_quarters = False
    elif args.include_future:
        config.include_future_quarters = True
    
    if args.csv_only:
        config.csv_backup = True
        config.google_sheet_id = None
    
    if args.no_csv_backup:
        config.csv_backup = False
    elif args.csv_backup:
        config.csv_backup = True
    
    if args.no_suggestions:
        config.auto_suggest_downloads = False
    
    return config

def validate_font_args(args):
    """Validate font-related arguments for v1.1.1"""
    errors = []
    
    if args.font_size and (args.font_size < 8 or args.font_size > 72):
        errors.append("Font size must be between 8 and 72 pt")
    
    if args.header_font_size and (args.header_font_size < 8 or args.header_font_size > 72):
        errors.append("Header font size must be between 8 and 72 pt")
    
    if args.max_display_types and (args.max_display_types < 1 or args.max_display_types > 20):
        errors.append("Max display types must be between 1 and 20")
    
    if args.bold_headers and args.no_bold_headers:
        errors.append("Cannot specify both --bold-headers and --no-bold-headers")
    
    if args.bold_company_info and args.no_bold_company_info:
        errors.append("Cannot specify both --bold-company-info and --no-bold-company-info")
    
    if args.show_all_types and args.show_best_only:
        errors.append("Cannot specify both --show-all-types and --show-best-only")
    
    return errors

def print_font_presets():
    """Print available font presets with details"""
    print("🎨 Available Font Presets (v1.1.1):")
    print("=" * 50)
    
    from .models import create_font_config_preset
    
    presets = ['small', 'normal', 'large', 'extra_large', 'huge']
    for preset_name in presets:
        preset = create_font_config_preset(preset_name)
        print(f"📝 {preset_name:12} : {preset['font_size']:2}pt / {preset['header_font_size']:2}pt - {preset['description']}")
    
    print()
    print("Usage examples:")
    print("  --font-preset large              # 14pt professional (recommended)")
    print("  --font-preset extra_large        # 16pt enhanced readability")
    print("  --font-preset huge               # 20pt presentation mode")
    print()

def handle_upload(args):
    """Handle upload to Google Sheets with v1.1.1 enhancements"""
    try:
        config = load_configuration(args)
        
        uploader = MOPSSheetsUploader(config=config)
        
        result = uploader.run()
        
        if result.success:
            if not args.quiet:
                print("✅ v1.1.1 上傳成功!")
                if result.sheets_url:
                    print(f"🔗 Google Sheets: {result.sheets_url}")
                if result.csv_path:
                    print(f"💾 CSV 檔案: {result.csv_path}")
                print(f"📊 涵蓋率: {result.coverage_stats.coverage_percentage:.1f}%")
                
                # v1.1.1 Enhanced output
                if result.font_config_used:
                    font_config = result.font_config_used
                    print(f"🔤 字體設定: {font_config.get('font_size', 14)}pt (標題: {font_config.get('header_font_size', 14)}pt)")
                    print(f"   預設模式: {font_config.get('preset_equivalent', 'custom')}")
                
                if result.multiple_types_found > 0:
                    print(f"🔄 多重類型: {result.multiple_types_found} 個季度")
                
                if result.type_combinations_analyzed:
                    top_combo = max(result.type_combinations_analyzed.items(), key=lambda x: x[1], default=None)
                    if top_combo:
                        print(f"📊 最常見組合: {top_combo[0]} ({top_combo[1]} 次)")
            
            return 0
        else:
            if not args.quiet:
                print(f"❌ 上傳失敗: {result.error_message}")
            return 1
            
    except Exception as e:
        if not args.quiet:
            print(f"❌ 執行錯誤: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
        return 1

def handle_csv_only(args):
    """Handle CSV-only export with v1.1.1 enhancements"""
    try:
        config = load_configuration(args)
        config.csv_backup = True
        config.google_sheet_id = None  # Disable Sheets upload
        
        uploader = MOPSSheetsUploader(config=config)
        
        result = uploader.run()
        
        if result.success:
            if not args.quiet:
                print("✅ v1.1.1 CSV 匯出成功!")
                print(f"📁 檔案位置: {result.csv_path}")
                if result.coverage_stats:
                    print(f"📊 涵蓋率: {result.coverage_stats.coverage_percentage:.1f}%")
                
                # v1.1.1 Enhanced output
                if result.multiple_types_found > 0:
                    print(f"🔄 多重類型: {result.multiple_types_found} 個季度")
                
                if result.font_config_used:
                    font_config = result.font_config_used
                    print(f"🔤 字體配置: {font_config.get('preset_equivalent', 'custom')} ({font_config.get('font_size', 14)}pt)")
            
            return 0
        else:
            if not args.quiet:
                print(f"❌ CSV 匯出失敗: {result.error_message}")
            return 1
            
    except Exception as e:
        if not args.quiet:
            print(f"❌ 匯出錯誤: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
        return 1

def handle_analyze(args):
    """Handle analysis-only mode with v1.1.1 enhancements"""
    try:
        config = load_configuration(args)
        
        uploader = MOPSSheetsUploader(config=config)
        
        report = uploader.generate_enhanced_report(include_analysis=True)
        
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2, default=str)
            if not args.quiet:
                print(f"📄 v1.1.1 分析報告已儲存: {args.output}")
        
        summary = report.get('summary', {})
        coverage = report.get('coverage', {})
        
        if not args.quiet:
            print(f"📈 v1.1.1 增強分析結果:")
            print(f"   • 總公司數: {summary.get('total_companies', 'N/A')}")
            print(f"   • 涵蓋率: {coverage.get('coverage_percentage', 0):.1f}%")
            print(f"   • 處理時間: {summary.get('processing_time', 0):.1f} 秒")
            
            # v1.1.1 Enhanced analysis output
            if summary.get('multiple_types_found', 0) > 0:
                print(f"   🔄 多重類型: {summary['multiple_types_found']} 個季度")
            
            if coverage.get('type_combinations'):
                combinations = coverage['type_combinations']
                top_combo = max(combinations.items(), key=lambda x: x[1])
                print(f"   📊 最常見組合: {top_combo[0]} ({top_combo[1]} 次)")
            
            if coverage.get('multiple_types'):
                mt_info = coverage['multiple_types']
                print(f"   🔄 多重類型比例: {mt_info.get('percentage', 0):.1f}%")
            
            # Font configuration used
            font_config = report.get('configuration', {}).get('font_config', {})
            if font_config:
                print(f"   🔤 字體配置: {font_config.get('preset_equivalent', 'custom')} ({font_config.get('font_size', 14)}pt)")
        else:
            # Quiet mode - minimal output
            print(f"Companies: {summary.get('total_companies', 'N/A')}")
            print(f"Coverage: {coverage.get('coverage_percentage', 0):.1f}%")
            print(f"Multiple Types: {summary.get('multiple_types_found', 0)}")
        
        return 0
        
    except Exception as e:
        if not args.quiet:
            print(f"❌ 分析錯誤: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
        return 1

def handle_test(args):
    """Handle connection test"""
    try:
        config = load_configuration(args)
        
        uploader = MOPSSheetsUploader(config=config)
        
        if uploader.test_connection():
            if not args.quiet:
                print("✅ v1.1.1 連線測試成功")
                if config.google_sheet_id:
                    print(f"📊 試算表: https://docs.google.com/spreadsheets/d/{config.google_sheet_id}")
                    print(f"📋 工作表名稱: {config.worksheet_name}")
                
                # Show font configuration that would be used
                font_config = config.get_enhanced_font_config()
                print(f"🔤 字體設定: {font_config['font_size']}pt/{font_config['header_font_size']}pt ({font_config['preset_equivalent']})")
            return 0
        else:
            if not args.quiet:
                print("❌ 連線測試失敗")
                print("   請檢查 GOOGLE_SHEETS_CREDENTIALS 和 GOOGLE_SHEET_ID 環境變數")
            return 1
            
    except Exception as e:
        if not args.quiet:
            print(f"❌ 連線測試錯誤: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
        return 1

def save_configuration(config: MOPSConfig, config_path: str) -> None:
    """Save configuration to file with v1.1.1 font settings"""
    config_dict = config.to_dict()
    
    # Add v1.1.1 specific metadata
    config_dict['_metadata'] = {
        'version': '1.1.1',
        'created_by': 'MOPS Sheets Uploader CLI',
        'creation_time': str(datetime.now()),
        'font_preset_equivalent': config.get_enhanced_font_config().get('preset_equivalent', 'custom')
    }
    
    file_ext = Path(config_path).suffix.lower()
    
    if file_ext == '.json':
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=2)
    else:
        # Default to YAML
        try:
            import yaml
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config_dict, f, allow_unicode=True, default_flow_style=False)
        except ImportError:
            # Fallback to JSON if YAML not available
            config_path = config_path.replace('.yaml', '.json').replace('.yml', '.json')
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, ensure_ascii=False, indent=2)
    
    print(f"💾 v1.1.1 設定已儲存: {config_path}")

def setup_logging(args) -> None:
    """Setup logging based on command line arguments"""
    if args.quiet:
        log_level = logging.WARNING
    elif args.verbose:
        log_level = logging.DEBUG
    else:
        log_level = getattr(logging, args.log_level.upper(), logging.INFO)
    
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - [CLI v1.1.1] %(message)s',
        handlers=[logging.StreamHandler()]
    )

def main() -> int:
    """Main CLI entry point with v1.1.1 enhancements"""
    parser = create_parser()
    args = parser.parse_args()
    
    # Handle special actions first
    if hasattr(args, 'print_presets') and args.print_presets:
        print_font_presets()
        return 0
    
    # Validate font arguments
    font_errors = validate_font_args(args)
    if font_errors:
        for error in font_errors:
            print(f"❌ {error}")
        return 1
    
    # Setup logging
    setup_logging(args)
    
    try:
        # Load and validate configuration
        config = load_configuration(args)
        
        # Save configuration if requested
        if args.save_config:
            from datetime import datetime
            save_configuration(config, args.save_config)
            return 0
        
        # Environment validation
        if not args.quiet:
            # Check environment
            has_credentials = bool(os.getenv('GOOGLE_SHEETS_CREDENTIALS'))
            has_sheet_id = bool(os.getenv('GOOGLE_SHEET_ID'))
            
            print(f"🔍 v1.1.1 Environment check:")
            print(f"   Google Credentials: {'✅' if has_credentials else '❌'}")
            print(f"   Sheet ID: {'✅' if has_sheet_id else '❌'}")
            
            # Show font configuration
            font_config = config.get_enhanced_font_config()
            print(f"   Font Config: {font_config['font_size']}pt/{font_config['header_font_size']}pt ({font_config['preset_equivalent']})")
            print(f"   Multiple Types: {'✅' if config.show_all_report_types else '❌'}")
        
        # Handle different actions
        if args.test:
            return handle_test(args)
        elif args.analyze:
            return handle_analyze(args)
        elif args.csv_only:
            return handle_csv_only(args)
        elif args.upload:
            # Check requirements for upload
            if not os.getenv('GOOGLE_SHEETS_CREDENTIALS') or not os.getenv('GOOGLE_SHEET_ID'):
                if not args.quiet:
                    print("❌ 缺少 Google Sheets 憑證，無法上傳")
                    print("   請設定 GOOGLE_SHEETS_CREDENTIALS 和 GOOGLE_SHEET_ID 環境變數")
                return 1
            return handle_upload(args)
        
    except KeyboardInterrupt:
        print("\n⏹️ 使用者中斷操作")
        return 130
    except Exception as e:
        print(f"❌ v1.1.1 程式錯誤: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

if __name__ == '__main__':
    sys.exit(main())