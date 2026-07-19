"""Deprecated public console entry point routed through the MOPS PDF/MD skill."""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _find_repo_root() -> Path:
    candidates = [Path.cwd(), *Path.cwd().parents, Path(__file__).resolve().parents[1]]
    for candidate in candidates:
        if (candidate / "mops_downloader").is_dir():
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
    raise SystemExit("Cannot find skill-mops-financialreport-pdf-md runner. Expected ../skills/common/skill-mops-financialreport-pdf-md/ or skills/common/...")


def main() -> int:
    parser = argparse.ArgumentParser(description="Deprecated wrapper; use skill-mops-financialreport-pdf-md.")
    parser.add_argument("--company_id", required=True)
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--quarter", default="all")
    parser.add_argument("--only-missing-files", action="store_true")
    parser.add_argument("--log_level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    parser.add_argument("--output", "-o", help="Deprecated; the skill writes to downloads/<company_id>.")
    args, unknown = parser.parse_known_args()

    if unknown:
        print(f"DEPRECATED: ignoring unsupported legacy options: {' '.join(unknown)}", file=sys.stderr)
    if args.output:
        print("DEPRECATED: --output is ignored by the skill wrapper; using downloads/<company_id>.", file=sys.stderr)

    repo = _find_repo_root()
    runner = _find_skill_runner(repo)
    cmd = [sys.executable, str(runner), args.company_id, str(args.year), str(args.quarter), "--log-level", args.log_level]
    if args.only_missing_files:
        cmd.append("--only-missing-files")

    print("DEPRECATED: mops-downloader console script is retired. Delegating to skill-mops-financialreport-pdf-md.", file=sys.stderr)
    return subprocess.run(cmd, cwd=repo).returncode


if __name__ == "__main__":
    raise SystemExit(main())
