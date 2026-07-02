#!/usr/bin/env python3
"""
batch_convert.py -- Parallel PDF-to-Markdown batch converter for MOPS downloads/.

Finds all PDFs in downloads/ that don't already have a matching .md file,
converts each using pdf_to_md.py:convert_pdf_to_md() directly,
runs up to 4 workers via concurrent.futures.ProcessPoolExecutor.
"""

import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import argparse
import subprocess
import time
import concurrent.futures
from pathlib import Path

SCRIPT_DIR      = Path(__file__).resolve().parent
REPO_ROOT       = SCRIPT_DIR.parent
DOWNLOADS       = REPO_ROOT / "downloads"
# Read by generate_mops_health.py to tell a genuine conversion failure apart
# from a PDF that simply hasn't been converted yet (pending).
FAILURES_LOG    = REPO_ROOT / "data" / "reports" / "batch_convert_failures.log"
# Lowered from 4: OCR-eligible pages call the (single) mac-mini-ocr API, which
# can take minutes per page -- too many concurrent workers would queue up and
# risk each other's requests timing out.
MAX_WORKERS = 2
# Raised from 180s: a scanned page OCR call alone may take up to 900s
# (see scripts/ocr_client.py), on top of normal PyMuPDF conversion time.
CONVERT_TIMEOUT = 1200
PDF_TO_MD   = SCRIPT_DIR / "pdf_to_md.py"


def _convert_one(pdf_path_str: str, no_ocr: bool = False) -> tuple:
    """Worker thread: convert a single PDF via subprocess."""
    try:
        cmd = [sys.executable, str(PDF_TO_MD), "--pdf", pdf_path_str]
        if no_ocr:
            cmd.append("--no-ocr")
        result = subprocess.run(
            cmd,
            capture_output=True, text=True, timeout=CONVERT_TIMEOUT,
            encoding="utf-8", errors="replace"
        )
        ok = result.returncode == 0
        err = result.stderr if not ok else ""
        return (pdf_path_str, ok, err)
    except subprocess.TimeoutExpired:
        return (pdf_path_str, False, f"Timeout after {CONVERT_TIMEOUT}s")
    except Exception as e:
        return (pdf_path_str, False, str(e))


def main():
    parser = argparse.ArgumentParser(
        description="Batch convert MOPS PDFs to Markdown")
    parser.add_argument('--no-ocr', action='store_true',
                        help="Disable mac-mini-ocr transcription for scanned pages")
    parser.add_argument('--retry-ocr', action='store_true',
                        help="Also reconvert existing .md files that still contain "
                             "a TODO:OCR placeholder (requires the matching .pdf)")
    parser.add_argument('--limit', type=int, default=None,
                        help="Only process the first N pending files (useful for "
                             "a quick sample run before committing to the full batch)")
    args = parser.parse_args()

    if not DOWNLOADS.exists():
        print(f"ERROR: downloads/ not found at {DOWNLOADS}", file=sys.stderr)
        sys.exit(1)

    all_pdfs = sorted(DOWNLOADS.rglob("*.pdf"))
    all_mds = sorted(DOWNLOADS.rglob("*.md"))
    existing_stems = {p.stem for p in all_mds}
    pending = [p for p in all_pdfs if p.stem not in existing_stems]

    if args.retry_ocr:
        ocr_stems = set()
        for md in all_mds:
            try:
                if "TODO:OCR" in md.read_text(encoding="utf-8", errors="replace"):
                    ocr_stems.add(md.stem)
            except Exception:
                continue
        retry = [p for p in all_pdfs if p.stem in ocr_stems]
        pending = pending + retry
        print(f"Retry OCR      : {len(retry)} PDF(s) with TODO:OCR placeholders")

    if args.limit is not None:
        pending = pending[:args.limit]

    n = len(pending)
    print(f"Downloads dir  : {DOWNLOADS}")
    print(f"Total PDFs     : {len(all_pdfs)}")
    print(f"Existing MDs   : {len(existing_stems)}")
    print(f"To convert     : {n}")
    if n == 0:
        print("Nothing to do -- all PDFs already have a matching .md file.")
        return
    print(f"Workers        : {MAX_WORKERS}")
    print(f"OCR enabled    : {not args.no_ocr}")
    print("-" * 70, flush=True)

    succeeded, failed = 0, 0
    failures = []
    done = 0
    t_start = time.time()

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_convert_one, str(p), args.no_ocr): p for p in pending}
        for future in concurrent.futures.as_completed(futures):
            done += 1
            pdf_path_str, ok, err = future.result()
            rel = Path(pdf_path_str).relative_to(DOWNLOADS)
            elapsed = time.time() - t_start
            avg = elapsed / done
            eta = avg * (n - done)
            if ok:
                succeeded += 1
                status = "OK  "
            else:
                failed += 1
                failures.append((pdf_path_str, err))
                status = "FAIL"
            print(
                f"[{done:4d}/{n}] {status}  {str(rel):<55}  "
                f"elapsed={elapsed:6.0f}s  eta={eta:6.0f}s",
                flush=True,
            )

    print("-" * 70)
    print(f"Conversion complete.  Succeeded: {succeeded}  Failed: {failed}")
    if failures:
        print("\nFailed files:")
        for path_str, err in failures:
            rel = Path(path_str).relative_to(DOWNLOADS)
            print(f"  {rel}")
            if err:
                for line in err.strip().splitlines()[:2]:
                    print(f"    {line}")

    # Record which PDFs actually errored out this run (as opposed to PDFs
    # nobody has tried to convert yet), so generate_mops_health.py can
    # distinguish "failed" from "pending" in its summary. Only replace the
    # entries for stems attempted this run -- a partial/--limit run must not
    # erase failure records for files it didn't touch.
    attempted_stems = {p.stem for p in pending}
    failed_stems = {Path(path_str).stem for path_str, _err in failures}
    prior_stems = set()
    if FAILURES_LOG.exists():
        try:
            prior_stems = {
                line.strip() for line in
                FAILURES_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
                if line.strip()
            }
        except Exception:
            pass
    merged_stems = (prior_stems - attempted_stems) | failed_stems

    FAILURES_LOG.parent.mkdir(parents=True, exist_ok=True)
    with open(FAILURES_LOG, "w", encoding="utf-8") as f:
        for stem in sorted(merged_stems):
            f.write(stem + "\n")


if __name__ == "__main__":
    main()
