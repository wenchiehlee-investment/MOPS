#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deprecated compatibility wrapper for MOPS PDF-to-Markdown conversion.

Use ../skills/common/skill-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py
for MOPS financial-report PDF download and Markdown sidecar generation.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _find_repo_root() -> Path:
    candidates = [Path.cwd(), *Path.cwd().parents]
    for candidate in candidates:
        if (candidate / "mops_downloader").is_dir() and (candidate / "downloads").exists():
            return candidate.resolve()
    return Path(__file__).resolve().parents[1]


def _find_skill_runner(repo: Path) -> Path:
    candidates = [
        repo.parent / "skills" / "common" / "skill-mops-financialreport-pdf-md" / "scripts" / "run_mops_financialreport_pdf_md.py",
        repo / "skills" / "common" / "skill-mops-financialreport-pdf-md" / "scripts" / "run_mops_financialreport_pdf_md.py",
        repo / "skills" / "skill-mops-financialreport-pdf-md" / "scripts" / "run_mops_financialreport_pdf_md.py",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise SystemExit("Cannot find skill-mops-financialreport-pdf-md runner. Expected ../skills/common/skill-mops-financialreport-pdf-md/.")


def _infer_target(pdf: Path) -> tuple[str, str, str]:
    parts = pdf.stem.split("_")
    if len(parts) < 3 or len(parts[0]) < 6:
        raise SystemExit(f"Cannot infer company/year/quarter from {pdf}. Expected YYYYQQ_company_report.pdf")
    period, company_id = parts[0], parts[1]
    return company_id, period[:4], str(int(period[5:6]))


def main() -> int:
    parser = argparse.ArgumentParser(description="Deprecated wrapper; use skill-mops-financialreport-pdf-md.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--pdf", "-p")
    group.add_argument("--dir", "-d")
    parser.add_argument("--force-convert", action="store_true")
    parser.add_argument("--no-refine", action="store_true")
    args = parser.parse_args()

    repo = _find_repo_root()
    runner = _find_skill_runner(repo)
    print("DEPRECATED: scripts/pdf_to_md.py is retired. Delegating to skill-mops-financialreport-pdf-md.", file=sys.stderr)

    targets: list[tuple[str, str, str]] = []
    if args.pdf:
        targets.append(_infer_target(Path(args.pdf)))
    else:
        by_target = set()
        for pdf in Path(args.dir).glob("*.pdf"):
            by_target.add(_infer_target(pdf))
        targets.extend(sorted(by_target))

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
