#!/usr/bin/env python3
"""
MOPS Sheets Uploader - Fixed Version for Scripts Folder
Upload MOPS PDF status to Google Sheets using command line parameters.
Works from scripts/ folder with proper path resolution.
"""

import os
import sys
import argparse
from pathlib import Path

def load_environment():
    """Load environment variables from .env file with path resolution"""
    try:
        from dotenv import load_dotenv
        
        # Try to find .env file in multiple locations
        script_dir = Path(__file__).parent
        env_locations = [
            script_dir / '.env',          # Same directory as script
            script_dir.parent / '.env',   # Parent directory (for scripts/)
            Path('.env'),                 # Current working directory
            Path.cwd() / '.env'          # Explicit current working directory
        ]
        
        env_loaded = False
        for env_file in env_locations:
            if env_file.exists():
                load_dotenv(env_file)
                env_loaded = True
                print(f"✅ Loaded .env from: {env_file}")
                break
        
        if not env_loaded:
            load_dotenv()  # Try default location
            
    except ImportError:
        # Fallback: manually load .env file
        script_dir = Path(__file__).parent
        env_locations = [
            script_dir / '.env',
            script_dir.parent / '.env', 
            Path('.env'),
            Path.cwd() / '.env'
        ]
        
        for env_file in env_locations:
            if env_file.exists():
                print(f"✅ Manually loading .env from: {env_file}")
                with open(env_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#') and '=' in line:
                            key, value = line.split('=', 1)
                            os.environ[key.strip()] = value.strip()
                break

def setup_package_path():
    """Setup package path for import"""
    script_dir = Path(__file__).parent
    
    # Try different possible package locations
    possible_package_dirs = [
        script_dir / "mops_sheets_uploader",  # Same directory
        script_dir.parent / "mops_sheets_uploader",  # Parent directory (for scripts/)
        Path("mops_sheets_uploader"),  # Current working directory
        Path.cwd() / "mops_sheets_uploader"  # Explicit current working directory
    ]
    
    package_dir = None
    for pkg_dir in possible_package_dirs:
        if pkg_dir.exists() and (pkg_dir / "__init__.py").exists():
            package_dir = pkg_dir
            print(f"✅ Found package at: {package_dir}")
            break
    
    if not package_dir:
        print("❌ mops_sheets_uploader 套件目錄不存在")
        print("   已搜尋位置:")
        for pkg_dir in possible_package_dirs:
            print(f"     • {pkg_dir}")
        return None
    
    # Add package directory to Python path if needed
    package_parent = str(package_dir.parent.absolute())
    if package_parent not in sys.path:
        sys.path.insert(0, package_parent)
        print(f"✅ Added to Python path: {package_parent}")
    
    return package_dir

# Load environment at module level
load_environment()

def create_parser():
    """Create command line argument parser"""
    parser = argparse.ArgumentParser(
        description="MOPS Sheets Uploader - Upload PDF matrix to Google Sheets",
        epilog="""
使用範例:
  python scripts/sheets_uploader.py --upload         # 上傳到 Google Sheets
  python scripts/sheets_uploader.py --csv-only       # 僅匯出 CSV
  python scripts/sheets_uploader.py --analyze        # 僅分析
  python scripts/sheets_uploader.py --test           # 測試連線
        """
    )
    
    # Action arguments (mutually exclusive)
    action_group = parser.add_mutually_exclusive_group(required=True)
    action_group.add_argument('--upload', action='store_true', help='上傳到 Google Sheets')
    action_group.add_argument('--csv-only', action='store_true', help='僅匯出 CSV 檔案')
    action_group.add_argument('--analyze', action='store_true', help='僅分析涵蓋率')
    action_group.add_argument('--test', action='store_true', help='測試 Google Sheets 連線')
    
    # Directory options
    parser.add_argument('--downloads-dir', default='downloads', help='Downloads 目錄路徑')
    parser.add_argument('--stock-csv', default='StockID_TWSE_TPEX.csv', help='股票清單 CSV 檔案')
    parser.add_argument('--output-dir', default='data/reports', help='CSV 輸出目錄')
    
    # Google Sheets options
    parser.add_argument('--sheet-id', help='Google Sheets ID')
    parser.add_argument('--worksheet-name', default='MOPS下載狀態', help='工作表名稱')
    
    # Analysis options
    parser.add_argument('--max-years', type=int, default=3, help='最大年數範圍')
    parser.add_argument('--no-suggestions', action='store_true', help='停用下載建議')
    parser.add_argument('--output-report', help='分析報告輸出檔案 (JSON)')
    
    # Display options
    parser.add_argument('--quiet', '-q', action='store_true', help='安靜模式')
    parser.add_argument('--verbose', '-v', action='store_true', help='詳細模式')
    
    return parser

def handle_upload(args):
    """Handle upload to Google Sheets"""
    try:
        from mops_sheets_uploader import MOPSSheetsUploader
        
        config_overrides = {}
        if args.sheet_id:
            config_overrides['google_sheet_id'] = args.sheet_id
        if args.worksheet_name != 'MOPS下載狀態':
            config_overrides['worksheet_name'] = args.worksheet_name
        if args.max_years != 3:
            config_overrides['max_years'] = args.max_years
        if args.no_suggestions:
            config_overrides['auto_suggest_downloads'] = False
        
        uploader = MOPSSheetsUploader(
            downloads_dir=args.downloads_dir,
            stock_csv_path=args.stock_csv,
            **config_overrides
        )
        uploader.config.output_dir = args.output_dir
        
        result = uploader.run()
        
        if result.success:
            if not args.quiet:
                print("✅ 上傳成功!")
                if result.sheets_url:
                    print(f"🔗 Google Sheets: {result.sheets_url}")
                if result.csv_path:
                    print(f"💾 CSV 檔案: {result.csv_path}")
                print(f"📊 涵蓋率: {result.coverage_stats.coverage_percentage:.1f}%")
            return 0
        else:
            if not args.quiet:
                print(f"❌ 上傳失敗: {result.error_message}")
            return 1
            
    except Exception as e:
        if not args.quiet:
            print(f"❌ 執行錯誤: {e}")
        return 1

def handle_csv_only(args):
    """Handle CSV-only export"""
    try:
        from mops_sheets_uploader import export_to_csv
        
        result = export_to_csv(
            output_dir=args.output_dir,
            downloads_dir=args.downloads_dir,
            stock_csv_path=args.stock_csv,
            max_years=args.max_years
        )
        
        if result.success:
            if not args.quiet:
                print(f"✅ CSV 匯出成功!")
                print(f"📁 檔案位置: {result.csv_path}")
                if result.coverage_stats:
                    print(f"📊 涵蓋率: {result.coverage_stats.coverage_percentage:.1f}%")
            return 0
        else:
            if not args.quiet:
                print(f"❌ CSV 匯出失敗: {result.error_message}")
            return 1
            
    except Exception as e:
        if not args.quiet:
            print(f"❌ 匯出錯誤: {e}")
        return 1

def handle_analyze(args):
    """Handle analysis only"""
    try:
        from mops_sheets_uploader import analyze_coverage
        
        report = analyze_coverage(
            downloads_dir=args.downloads_dir,
            stock_csv_path=args.stock_csv,
            include_suggestions=not args.no_suggestions,
            max_years=args.max_years
        )
        
        if args.output_report:
            import json
            with open(args.output_report, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2, default=str)
            if not args.quiet:
                print(f"📄 分析報告已儲存: {args.output_report}")
        
        summary = report.get('summary', {})
        if not args.quiet:
            print(f"📈 分析結果:")
            print(f"   • 總公司數: {summary.get('total_companies', 'N/A')}")
            print(f"   • 有PDF公司: {summary.get('companies_with_pdfs', 'N/A')}")
            print(f"   • 涵蓋率: {summary.get('coverage_percentage', 0):.1f}%")
            print(f"   • 品質分數: {summary.get('quality_score', 'N/A')}/10")
        else:
            print(f"Companies: {summary.get('total_companies', 'N/A')}")
            print(f"Coverage: {summary.get('coverage_percentage', 0):.1f}%")
            print(f"Quality: {summary.get('quality_score', 'N/A')}/10")
        
        return 0
        
    except Exception as e:
        if not args.quiet:
            print(f"❌ 分析錯誤: {e}")
        return 1

def handle_test(args):
    """Handle connection test"""
    try:
        from mops_sheets_uploader import test_google_sheets_connection
        
        sheet_id = args.sheet_id or os.getenv('GOOGLE_SHEET_ID')
        
        if test_google_sheets_connection(sheet_id):
            if not args.quiet:
                print("✅ 連線測試成功")
                if sheet_id:
                    print(f"📊 試算表: https://docs.google.com/spreadsheets/d/{sheet_id}")
            return 0
        else:
            if not args.quiet:
                print("❌ 連線測試失敗")
            return 1
            
    except Exception as e:
        if not args.quiet:
            print(f"❌ 連線測試錯誤: {e}")
        return 1

def main():
    """Main function"""
    parser = create_parser()
    args = parser.parse_args()
    
    # Setup package path
    package_dir = setup_package_path()
    if not package_dir:
        return 1
    
    try:
        # Import package
        from mops_sheets_uploader import MOPSSheetsUploader
        
        # Check environment variables
        has_credentials = bool(os.getenv('GOOGLE_SHEETS_CREDENTIALS'))
        has_sheet_id = bool(os.getenv('GOOGLE_SHEET_ID'))
        
        if not args.quiet:
            print(f"🔍 Environment check:")
            print(f"   Credentials: {'✅' if has_credentials else '❌'}")
            print(f"   Sheet ID: {'✅' if has_sheet_id else '❌'}")
        
        # Execute the requested action
        if args.upload:
            if not has_credentials or not has_sheet_id:
                if not args.quiet:
                    print("❌ 缺少 Google Sheets 憑證，無法上傳")
                return 1
            return handle_upload(args)
            
        elif args.csv_only:
            return handle_csv_only(args)
            
        elif args.analyze:
            return handle_analyze(args)
            
        elif args.test:
            if not has_credentials:
                if not args.quiet:
                    print("❌ 缺少 Google Sheets 憑證，無法測試連線")
                return 1
            return handle_test(args)
        
    except ImportError as e:
        if not args.quiet:
            print(f"❌ 套件匯入失敗: {e}")
            print(f"   套件位置: {package_dir}")
        return 1
    
    except Exception as e:
        if not args.quiet:
            print(f"❌ 程式執行錯誤: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())