#!/usr/bin/env python3
import csv
import os
import shutil
from datetime import datetime, timezone, timedelta
from pathlib import Path

# Paths
REPO_ROOT = Path(__file__).resolve().parent.parent
DOWNLOADS = REPO_ROOT / "downloads"
REPORTS_DIR = REPO_ROOT / "data" / "reports"
HEALTH_SUMMARY_CSV = REPORTS_DIR / "mops_health_summary.csv"
MATRIX_LATEST_CSV = REPORTS_DIR / "mops_matrix_latest.csv"
# Written by batch_convert.py: PDF stems that errored out on the most recent
# conversion attempt. Used to tell "genuinely failed" apart from "not yet
# converted" (pending) below.
BATCH_CONVERT_FAILURES_LOG = REPORTS_DIR / "batch_convert_failures.log"

TAIPEI_TZ = timezone(timedelta(hours=8))

def format_timestamp(mtime):
    dt = datetime.fromtimestamp(mtime, tz=timezone.utc).astimezone(TAIPEI_TZ)
    return dt.isoformat()

def main():
    print("=== Generating MOPS Data Health Summary ===")
    
    if not REPORTS_DIR.exists():
        os.makedirs(REPORTS_DIR, exist_ok=True)

    # 1. Scan downloads for PDF and MD files
    pdf_files = []
    md_files = []
    
    if DOWNLOADS.exists():
        pdf_files = sorted(DOWNLOADS.rglob("*.pdf"))
        md_files = sorted(DOWNLOADS.rglob("*.md"))
    
    total_pdfs = len(pdf_files)
    total_mds = len(md_files)
    
    print(f"Total PDFs found: {total_pdfs}")
    print(f"Total MDs found: {total_mds}")

    # 2. Check for OCR Needed, genuinely failed, and pending (not yet attempted) conversions
    ocr_needed_count = 0
    failed_conversions = 0
    pending_conversions = 0

    # PDF stems that batch_convert.py actually attempted and failed on its
    # most recent run. Anything else with no matching .md is simply pending
    # (nobody has tried to convert it yet).
    known_failed_stems = set()
    if BATCH_CONVERT_FAILURES_LOG.exists():
        try:
            known_failed_stems = {
                line.strip() for line in
                BATCH_CONVERT_FAILURES_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
                if line.strip()
            }
        except Exception as e:
            print(f"Warning: Failed to read {BATCH_CONVERT_FAILURES_LOG.name}: {e}")

    existing_mds = {}
    for md in md_files:
        existing_mds[md.stem] = md

    for pdf in pdf_files:
        stem = pdf.stem
        if stem in existing_mds:
            md_path = existing_mds[stem]
            try:
                content = md_path.read_text(encoding="utf-8", errors="replace")
                if "TODO:OCR" in content:
                    ocr_needed_count += 1
            except Exception as e:
                print(f"Warning: Failed to read {md_path.name}: {e}")
                failed_conversions += 1
        elif stem in known_failed_stems:
            failed_conversions += 1
        else:
            # No matching MD file and no record of a failed attempt -> pending
            pending_conversions += 1

    # Calculate conversion rate (completed MD without OCR or failure / total_pdfs)
    successful_mds = total_mds - ocr_needed_count - failed_conversions
    conversion_rate_pct = round((max(0, successful_mds) / total_pdfs * 100), 2) if total_pdfs > 0 else 0.0

    print(f"OCR Needed count: {ocr_needed_count}")
    print(f"Pending conversions: {pending_conversions}")
    print(f"Failed conversions: {failed_conversions}")
    print(f"Overall conversion rate: {conversion_rate_pct}%")

    # 3. Determine timestamps from file mtimes
    latest_pdf_time = ""
    latest_md_time = ""
    
    if pdf_files:
        latest_pdf_mtime = max(os.path.getmtime(p) for p in pdf_files)
        latest_pdf_time = format_timestamp(latest_pdf_mtime)
    
    if md_files:
        latest_md_mtime = max(os.path.getmtime(p) for p in md_files)
        latest_md_time = format_timestamp(latest_md_mtime)
        
    now = datetime.now(TAIPEI_TZ)
    checked_at = now.isoformat()
    
    # If no files, default download_timestamp to checked_at
    process_timestamp = checked_at
    download_timestamp = latest_pdf_time if latest_pdf_time else checked_at

    # 4. Copy the latest matrix CSV to data/reports/mops_matrix_latest.csv
    csvs = sorted(REPORTS_DIR.glob("mops_matrix_*.csv"))
    # Filter out mops_matrix_latest.csv to avoid self-reference
    csvs = [c for c in csvs if c.name != "mops_matrix_latest.csv"]
    
    if csvs:
        latest_csv = csvs[-1]
        print(f"Processing and copying latest matrix CSV: {latest_csv.name} -> mops_matrix_latest.csv")
        try:
            with open(latest_csv, "r", encoding="utf-8-sig") as f_in:
                reader = csv.DictReader(f_in)
                fieldnames = reader.fieldnames if reader.fieldnames else []
                new_fieldnames = list(fieldnames)
                if "process_timestamp" not in new_fieldnames:
                    new_fieldnames.append("process_timestamp")
                rows = list(reader)
                for r in rows:
                    r["process_timestamp"] = process_timestamp
            with open(MATRIX_LATEST_CSV, "w", encoding="utf-8", newline="") as f_out:
                writer = csv.DictWriter(f_out, fieldnames=new_fieldnames)
                writer.writeheader()
                writer.writerows(rows)
        except Exception as e:
            print(f"Error processing matrix CSV: {e}")
            shutil.copy2(latest_csv, MATRIX_LATEST_CSV)
    else:
        print("Warning: No mops_matrix_*.csv found to copy to mops_matrix_latest.csv")

    # 5. Write the summary row to data/reports/mops_health_summary.csv
    summary_data = {
        "process_timestamp": process_timestamp,
        "download_timestamp": download_timestamp,
        "total_pdfs": total_pdfs,
        "total_mds": total_mds,
        "conversion_rate_pct": conversion_rate_pct,
        "ocr_needed_count": ocr_needed_count,
        "pending_conversions": pending_conversions,
        "failed_conversions": failed_conversions,
        "latest_md_time": latest_md_time if latest_md_time else "N/A",
        "mops_financials_extracted_count": 0,
        "ready_to_use_rate_pct": 0.0,
        "checked_at": checked_at
    }

    fieldnames = [
        "process_timestamp",
        "download_timestamp",
        "total_pdfs",
        "total_mds",
        "conversion_rate_pct",
        "ocr_needed_count",
        "pending_conversions",
        "failed_conversions",
        "latest_md_time",
        "mops_financials_extracted_count",
        "ready_to_use_rate_pct",
        "checked_at"
    ]

    with open(HEALTH_SUMMARY_CSV, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerow(summary_data)
        
    print(f"Successfully generated health summary at {HEALTH_SUMMARY_CSV.name}")

if __name__ == "__main__":
    main()
