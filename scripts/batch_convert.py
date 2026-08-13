#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deprecated compatibility wrapper for batch PDF-to-Markdown conversion.

The maintained path is skill-company-mops-financialreport-pdf-md, which downloads MOPS
financial-report PDFs and creates same-stem Markdown sidecars through
skill-mac-mini-ocr.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _find_repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def _find_skill_runner(repo: Path) -> Path:
    candidates = [
        repo.parent / "skills" / "common" / "skill-company-mops-financialreport-pdf-md" / "scripts" / "run_mops_financialreport_pdf_md.py",
        repo / "skills" / "common" / "skill-company-mops-financialreport-pdf-md" / "scripts" / "run_mops_financialreport_pdf_md.py",
        repo / "skills" / "skill-company-mops-financialreport-pdf-md" / "scripts" / "run_mops_financialreport_pdf_md.py",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise SystemExit("Cannot find skill-company-mops-financialreport-pdf-md runner. Expected ../skills/common/skill-company-mops-financialreport-pdf-md/.")


def _pending_targets(downloads: Path) -> list[tuple[str, str, str]]:
    targets = set()
    for pdf in downloads.rglob("*.pdf"):
        if pdf.with_suffix(".md").is_file():
            continue
        parts = pdf.stem.split("_")
        if len(parts) < 3 or len(parts[0]) < 6:
            continue
        period, company_id = parts[0], parts[1]
        targets.add((company_id, period[:4], str(int(period[5:6]))))
    return sorted(targets)


def main() -> int:
    parser = argparse.ArgumentParser(description="Deprecated wrapper; use skill-company-mops-financialreport-pdf-md.")
    parser.add_argument("--force-convert", action="store_true")
    parser.add_argument("--no-refine", action="store_true")
    args = parser.parse_args()

    repo = _find_repo_root()
    runner = _find_skill_runner(repo)
    targets = _pending_targets(repo / "downloads")
    print("DEPRECATED: scripts/batch_convert.py is retired. Delegating to skill-company-mops-financialreport-pdf-md.", file=sys.stderr)
    print(f"Targets needing Markdown: {len(targets)}")

    exit_code = 0
    for company_id, year, quarter in targets:
        cmd = [sys.executable, str(runner), company_id, year, quarter, "--skip-download"]
        if args.force_convert:
            cmd.append("--force-convert")
        if args.no_refine:
            cmd.append("--no-refine")
        result = subprocess.run(cmd, cwd=repo)
        exit_code = exit_code or result.returncode
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
