#!/usr/bin/env python3
"""
DownloadAll.py - batch orchestrator using skill-mops-financialreport-pdf-md

This version adds the ability to skip existing files to avoid redundant downloads.

Usage:
    python DownloadAll.py --year 2025 --quarter 1 --only-missing-files
"""

import argparse
import csv
import subprocess
import sys
import os
import time
import glob
from typing import List, Tuple, Optional


def find_skill_runner() -> str:
    """Locate skill-mops-financialreport-pdf-md runner from the MOPS repo."""
    repo = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(os.path.dirname(repo), 'skills', 'common', 'skill-mops-financialreport-pdf-md', 'scripts', 'run_mops_financialreport_pdf_md.py'),
        os.path.join(repo, 'skills', 'common', 'skill-mops-financialreport-pdf-md', 'scripts', 'run_mops_financialreport_pdf_md.py'),
        os.path.join(repo, 'skills', 'skill-mops-financialreport-pdf-md', 'scripts', 'run_mops_financialreport_pdf_md.py'),
    ]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    raise FileNotFoundError('Cannot find skill-mops-financialreport-pdf-md runner. Expected ../skills/common/skill-mops-financialreport-pdf-md/.')


def read_company_ids(csv_file: str) -> List[Tuple[int, str]]:
    """Read company IDs and names from CSV file."""
    companies = []
    
    try:
        with open(csv_file, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            
            for row in reader:
                company_id_key = '代號' if '代號' in row else 'company_id'
                company_name_key = '名稱' if '名稱' in row else 'company_name'
                
                if company_id_key in row and company_name_key in row:
                    company_id = int(row[company_id_key])
                    company_name = row[company_name_key].strip()
                    companies.append((company_id, company_name))
                    
    except FileNotFoundError:
        print(f"Error: CSV file '{csv_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        sys.exit(1)
    
    return companies


def check_existing_file(company_id: int, year: int, quarter: Optional[int], min_size: int = 102400) -> bool:
    """
    Check if PDF file already exists and meets minimum size requirement.
    
    Args:
        company_id: Company ID
        year: Year
        quarter: Quarter (None for all quarters)
        min_size: Minimum file size in bytes (default: 100KB = 102400 bytes)
        
    Returns:
        True if existing valid file found, False otherwise
    """
    downloads_dir = "downloads"
    company_dir = os.path.join(downloads_dir, str(company_id))
    
    if not os.path.exists(company_dir):
        return False
    
    # Check for existing PDF files. Skip mode is only complete when the
    # same-stem Markdown sidecar also exists; otherwise the skill must run to
    # create or repair the .md file.
    if quarter is not None:
        quarter_str = f"{year}{quarter:02d}"
        pattern = os.path.join(company_dir, f"{quarter_str}_{company_id}_*.pdf")
    else:
        pattern = os.path.join(company_dir, f"{year}*_{company_id}_*.pdf")

    existing_files = glob.glob(pattern)

    if not existing_files:
        return False

    for file_path in existing_files:
        md_path = os.path.splitext(file_path)[0] + '.md'
        try:
            if os.path.getsize(file_path) >= min_size and os.path.exists(md_path) and os.path.getsize(md_path) > 0:
                return True
        except OSError:
            continue

    return False


def should_skip_download(company_id: int, company_name: str, year: int, 
                        quarter: Optional[int], only_missing: bool, min_size: int = 102400) -> bool:
    """
    Determine if download should be skipped based on existing files.
    
    Returns:
        True if download should be skipped, False otherwise
    """
    if not only_missing:
        return False
    
    if check_existing_file(company_id, year, quarter, min_size):
        print(f"  --> SKIP: {company_id} ({company_name}) - Existing PDF+MD found (>={min_size//1024}KB)")
        return True
    
    return False


def download_company_pdf_simple(company_id: int, company_name: str, year: int, 
                               quarter: Optional[int] = None, delay: float = 1.0,
                               only_missing: bool = False) -> bool:
    """
    Download and convert reports for a single company through skill-mops-financialreport-pdf-md.
    """
    # Check if we should skip this download
    if should_skip_download(company_id, company_name, year, quarter, only_missing):
        return True  # Count as success since file already exists
    
    # Build the command through the maintained skill. The skill owns the public
    # workflow: MOPS download plus same-stem Markdown sidecar generation.
    cmd = [
        sys.executable, find_skill_runner(),
        str(company_id),
        str(year),
        str(quarter) if quarter is not None else 'all'
    ]
    if only_missing:
        cmd.append('--only-missing-files')
    
    print(f"Downloading {company_id} ({company_name}) - Year: {year}" + 
          (f", Quarter: {quarter}" if quarter else ""))
    
    try:
        # Set environment for Unicode support
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        
        # Execute the command without capturing output (to avoid encoding issues)
        # Let mops_downloader print directly to console
        result = subprocess.run(
            cmd,
            timeout=300,  # 5 minute timeout
            env=env
        )
        
        # Check return code to determine success
        if result.returncode == 0:
            print(f"  --> SUCCESS: {company_id} ({company_name})")
            return True
        else:
            print(f"  --> FAILED: {company_id} ({company_name}) - Return code: {result.returncode}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"  --> TIMEOUT: {company_id} ({company_name})")
        return False
    except Exception as e:
        print(f"  --> ERROR: {company_id} ({company_name}) - {e}")
        return False
    finally:
        # Add delay between requests
        if delay > 0:
            print(f"  Waiting {delay}s before next download...")
            time.sleep(delay)
        print("-" * 50)


def main():
    parser = argparse.ArgumentParser(
        description='Download MOPS PDFs and Markdown sidecars for all companies via skill-mops-financialreport-pdf-md',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Download all Q1 2025 reports
  python DownloadAll.py --year 2025 --quarter 1

  # Download all 2024 reports, skip existing files > 100KB
  python DownloadAll.py --year 2024 --only-missing-files

  # Download starting from company 8045, with 10 second delays
  python DownloadAll.py --year 2025 --quarter 1 --start-from 8045 --delay 10.0
  
  # Test run without downloading
  python DownloadAll.py --year 2025 --quarter 1 --dry-run
        """
    )
    
    parser.add_argument(
        '--year', '-y',
        type=int,
        required=True,
        help='Year to download (e.g., 2025)'
    )
    
    parser.add_argument(
        '--quarter', '-q',
        type=int,
        choices=[1, 2, 3, 4],
        help='Quarter to download (1-4, optional). If not specified, downloads all quarters.'
    )
    
    parser.add_argument(
        '--csv-file', '-f',
        type=str,
        default='StockID_TWSE_TPEX.csv',
        help='CSV file containing company IDs (default: StockID_TWSE_TPEX.csv)'
    )
    
    parser.add_argument(
        '--delay', '-d',
        type=float,
        default=1.0,
        help='Delay between downloads in seconds (default: 1.0)'
    )
    
    parser.add_argument(
        '--start-from',
        type=int,
        help='Start downloading from specific company ID (useful for resuming)'
    )
    
    parser.add_argument(
        '--only-missing-files',
        action='store_true',
        help='Skip companies that already have a valid PDF and same-stem Markdown sidecar'
    )
    
    parser.add_argument(
        '--min-file-size',
        type=int,
        default=102400,
        help='Minimum file size in bytes to consider existing file valid (default: 102400 = 100KB)'
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Show what would be downloaded without actually downloading'
    )
    
    args = parser.parse_args()
    
    # Read companies from CSV
    print(f"Reading companies from {args.csv_file}...")
    companies = read_company_ids(args.csv_file)
    
    if not companies:
        print("No companies found in CSV file.")
        sys.exit(1)
    
    print(f"Found {len(companies)} companies.")
    
    # Filter companies if start_from is specified
    if args.start_from:
        companies = [(cid, name) for cid, name in companies if cid >= args.start_from]
        print(f"Starting from company ID {args.start_from}. {len(companies)} companies remaining.")
    
    if args.dry_run:
        print("\nDRY RUN - Commands that would be executed:")
        for company_id, company_name in companies:
            if args.only_missing_files:
                if should_skip_download(company_id, company_name, args.year, args.quarter, True, args.min_file_size):
                    print(f"  SKIP: {company_id} ({company_name}) - Existing PDF+MD found")
                    continue
            
            quarter = args.quarter if args.quarter else 'all'
            cmd_str = f"python ../skills/common/skill-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py {company_id} {args.year} {quarter}"
            if args.only_missing_files:
                cmd_str += " --only-missing-files"
            print(f"  {cmd_str}  # {company_name}")
        return
    
    # Count existing files if skip mode is enabled
    if args.only_missing_files:
        existing_count = 0
        for company_id, company_name in companies:
            if check_existing_file(company_id, args.year, args.quarter, args.min_file_size):
                existing_count += 1
        
        print(f"Skip mode enabled: {existing_count} companies already have valid PDF+MD sidecars (>={args.min_file_size//1024}KB)")
        print(f"Will attempt to download {len(companies) - existing_count} companies")
    
    # Start downloading
    print(f"\nStarting skill-driven download/convert for year {args.year}" + 
          (f", quarter {args.quarter}" if args.quarter else " (all quarters)"))
    print(f"Delay between requests: {args.delay}s")
    print(f"Skip existing files: {args.only_missing_files}")
    if args.only_missing_files:
        print(f"Minimum file size: {args.min_file_size//1024}KB")
    print("=" * 50)
    
    successful = 0
    failed = 0
    skipped = 0
    
    for i, (company_id, company_name) in enumerate(companies, 1):
        print(f"\nProgress: {i}/{len(companies)}")
        
        # Check if we should skip (for statistics)
        if args.only_missing_files and should_skip_download(company_id, company_name, args.year, args.quarter, True, args.min_file_size):
            skipped += 1
            continue
        
        success = download_company_pdf_simple(
            company_id=company_id,
            company_name=company_name,
            year=args.year,
            quarter=args.quarter,
            delay=args.delay,
            only_missing=args.only_missing_files
        )
        
        if success:
            successful += 1
        else:
            failed += 1
    
    # Summary
    print("\n" + "=" * 50)
    print("DOWNLOAD/CONVERSION SUMMARY")
    print("=" * 50)
    print(f"Total companies: {len(companies)}")
    print(f"Successful skill runs: {successful}")
    print(f"Failed skill runs: {failed}")
    if args.only_missing_files:
        print(f"Skipped (existing files): {skipped}")
        print(f"Attempted downloads: {len(companies) - skipped}")
        if len(companies) - skipped > 0:
            print(f"Success rate (attempted): {(successful/(len(companies) - skipped)*100):.1f}%")
    print(f"Overall success rate: {(successful/len(companies)*100):.1f}%")
    
    # Create downloads directory summary
    downloads_dir = "downloads"
    if os.path.exists(downloads_dir):
        company_dirs = [d for d in os.listdir(downloads_dir) if os.path.isdir(os.path.join(downloads_dir, d))]
        total_pdfs = 0
        for company_dir in company_dirs:
            pdf_files = glob.glob(os.path.join(downloads_dir, company_dir, "*.pdf"))
            total_pdfs += len(pdf_files)
        
        print(f"\nDownloads directory: {len(company_dirs)} company folders, {total_pdfs} total PDF files")


if __name__ == "__main__":
    main()