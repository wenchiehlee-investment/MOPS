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

import subprocess
import time
import concurrent.futures
from pathlib import Path

SCRIPT_DIR  = Path(__file__).resolve().parent
REPO_ROOT   = SCRIPT_DIR.parent
DOWNLOADS   = REPO_ROOT / "downloads"
MAX_WORKERS = 4
PDF_TO_MD   = SCRIPT_DIR / "pdf_to_md.py"


def _convert_one(pdf_path_str: str) -> tuple:
    """Worker thread: convert a single PDF via subprocess."""
    try:
        result = subprocess.run(
            [sys.executable, str(PDF_TO_MD), "--pdf", pdf_path_str],
            capture_output=True, text=True, timeout=180,
            encoding="utf-8", errors="replace"
        )
        ok = result.returncode == 0
        err = result.stderr if not ok else ""
        return (pdf_path_str, ok, err)
    except subprocess.TimeoutExpired:
        return (pdf_path_str, False, "Timeout after 180s")
    except Exception as e:
        return (pdf_path_str, False, str(e))


def main():
    if not DOWNLOADS.exists():
        print(f"ERROR: downloads/ not found at {DOWNLOADS}", file=sys.stderr)
        sys.exit(1)

    all_pdfs = sorted(DOWNLOADS.rglob("*.pdf"))
    existing_stems = {p.stem for p in DOWNLOADS.rglob("*.md")}
    pending = [p for p in all_pdfs if p.stem not in existing_stems]

    n = len(pending)
    print(f"Downloads dir  : {DOWNLOADS}")
    print(f"Total PDFs     : {len(all_pdfs)}")
    print(f"Existing MDs   : {len(existing_stems)}")
    print(f"To convert     : {n}")
    if n == 0:
        print("Nothing to do -- all PDFs already have a matching .md file.")
        return
    print(f"Workers        : {MAX_WORKERS}")
    print("-" * 70, flush=True)

    succeeded, failed = 0, 0
    failures = []
    done = 0
    t_start = time.time()

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_convert_one, str(p)): p for p in pending}
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


if __name__ == "__main__":
    main()
