"""Command-line interface for MOPS downloader."""

import sys
import argparse
from pathlib import Path
from .main import MOPSDownloader
from .exceptions import ValidationError, DownloadError, NetworkError
from .__init__ import __version__

def create_parser():
    """Create command-line argument parser."""
    parser = argparse.ArgumentParser(
        description="Download Taiwan MOPS quarterly financial reports",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --company_id 2330 --year 2024
  %(prog)s --company_id 8272 --year 2023 --quarter 2
  %(prog)s --company_id 2382 --year 2023 --quarter all --output ./reports
        """
    )
    
    parser.add_argument(
        '--company_id', 
        required=True,
        help='Taiwan stock company ID (4 digits, e.g., 2330, 8272)'
    )
    
    parser.add_argument(
        '--year', 
        type=int,
        required=True,
        help='Reporting year in Western format (e.g., 2024, 2023)'
    )
    
    parser.add_argument(
        '--quarter',
        default='all',
        help='Quarter to download: 1, 2, 3, 4, or "all" (default: all)'
    )
    
    parser.add_argument(
        '--output', '-o',
        type=Path,
        help='Output directory for downloads (default: ./downloads)'
    )
    
    parser.add_argument(
        '--log_level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='Logging level (default: INFO)'
    )
    
    parser.add_argument(
        '--version', '-v',
        action='version',
        version=f'%(prog)s {__version__}'
    )
    
    return parser

def format_result_summary(result):
    """Format download result for display."""
    lines = []
    
    if result.success:
        lines.append("✅ Download completed successfully!")
        lines.append(f"   Files downloaded: {result.total_files}")
        lines.append(f"   Total size: {result.total_size:,} bytes")
        
        if result.downloaded_files:
            lines.append(f"   Downloaded files:")
            for filename in result.downloaded_files:
                lines.append(f"     • {filename}")
        
        if result.missing_quarters:
            lines.append(f"   Missing quarters: Q{', Q'.join(map(str, result.missing_quarters))}")
            lines.append(f"     (Individual reports not available for these quarters)")
    else:
        lines.append("❌ Download failed!")
        if result.error_details:
            lines.append(f"   Error: {result.error_details}")
        
        if result.downloaded_files:
            lines.append(f"   Partial success - {len(result.downloaded_files)} files downloaded")
    
    return '\n'.join(lines)

def main():
    """Main CLI entry point."""
    parser = create_parser()
    args = parser.parse_args()
    
    try:
        # Initialize downloader
        downloader = MOPSDownloader(
            download_dir=args.output,
            log_level=args.log_level
        )
        
        # Convert quarter argument
        quarter = args.quarter
        if quarter.lower() == 'all':
            quarter = 'all'
        else:
            try:
                quarter = int(quarter)
            except ValueError:
                print(f"Error: Invalid quarter '{args.quarter}'. Must be 1, 2, 3, 4, or 'all'")
                return 1
        
        # Perform download
        print(f"Downloading reports for company {args.company_id}, year {args.year}, quarter {quarter}...")
        result = downloader.download(args.company_id, args.year, quarter)
        
        # Display results
        print("\n" + "="*60)
        print(format_result_summary(result))
        print("="*60)
        
        # Exit with appropriate code
        return 0 if result.success else 1
        
    except ValidationError as e:
        print(f"Validation Error: {e}")
        return 1
    except (DownloadError, NetworkError) as e:
        print(f"Download Error: {e}")
        return 1
    except KeyboardInterrupt:
        print("\nDownload interrupted by user")
        return 1
    except Exception as e:
        print(f"Unexpected Error: {e}")
        if args.log_level == 'DEBUG':
            import traceback
            traceback.print_exc()
        return 1

if __name__ == '__main__':
    sys.exit(main())
