#!/usr/bin/env python3
"""
DownloadAll_Simple.py - Simple version that avoids Unicode encoding issues

This version doesn't capture output from mops_downloader to avoid Windows encoding problems.
Instead, it checks for successful downloads by looking at the return code.

Usage:
    python DownloadAll_Simple.py --year 2025 --quarter 1
"""

import argparse
import csv
import subprocess
import sys
import os
import time
from typing import List, Tuple, Optional


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


def download_company_pdf_simple(company_id: int, company_name: str, year: int, 
                               quarter: Optional[int] = None, delay: float = 1.0) -> bool:
    """
    Download PDF for a single company using mops_downloader (simple version).
    This version doesn't capture output to avoid encoding issues.
    """
    # Build the command
    cmd = [
        sys.executable, '-m', 'mops_downloader.cli',
        '--company_id', str(company_id),
        '--year', str(year)
    ]
    
    # Add quarter if specified
    if quarter is not None:
        cmd.extend(['--quarter', str(quarter)])
    
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
        description='Download MOPS PDFs for all companies (Simple version - no output capture)',
        formatter_class=argparse.RawDescriptionHelpFormatter
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
        help='Quarter to download (1-4, optional)'
    )
    
    parser.add_argument(
        '--csv-file', '-f',
        type=str,
        default='StockID_TWSE_TPEX.csv',
        help='CSV file containing company IDs'
    )
    
    parser.add_argument(
        '--delay', '-d',
        type=float,
        default=1.0,
        help='Delay between downloads in seconds'
    )
    
    parser.add_argument(
        '--start-from',
        type=int,
        help='Start downloading from specific company ID'
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
            cmd_str = f"python -m mops_downloader.cli --company_id {company_id} --year {args.year}"
            if args.quarter:
                cmd_str += f" --quarter {args.quarter}"
            print(f"  {cmd_str}  # {company_name}")
        return
    
    # Start downloading
    print(f"\nStarting download for year {args.year}" + 
          (f", quarter {args.quarter}" if args.quarter else ""))
    print(f"Delay between requests: {args.delay}s")
    print("=" * 50)
    
    successful = 0
    failed = 0
    
    for i, (company_id, company_name) in enumerate(companies, 1):
        print(f"\nProgress: {i}/{len(companies)}")
        
        success = download_company_pdf_simple(
            company_id=company_id,
            company_name=company_name,
            year=args.year,
            quarter=args.quarter,
            delay=args.delay
        )
        
        if success:
            successful += 1
        else:
            failed += 1
    
    # Summary
    print("\n" + "=" * 50)
    print("DOWNLOAD SUMMARY")
    print("=" * 50)
    print(f"Total companies: {len(companies)}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print(f"Success rate: {(successful/len(companies)*100):.1f}%")


if __name__ == "__main__":
    main()