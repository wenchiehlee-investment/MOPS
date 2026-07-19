# Repository Guidelines

## Project Structure & Module Organization
- Primary automation lives in `.github/workflows/Download.yaml`.
- Supporting runtime code:
  - `DownloadAll.py`: batch orchestrator used by the workflow; delegates to `skill-mops-financialreport-pdf-md`.
  - `../skills/common/skill-mops-financialreport-pdf-md/`: maintained public workflow for MOPS PDF download plus Markdown sidecar generation.
  - `mops_downloader/`: internal downloader package used by the skill.
  - `scripts/sheets_uploader.py` + `mops_sheets_uploader/`: matrix generation and Google Sheets upload.
- Generated artifacts:
  - `downloads/<company_id>/*.pdf` and same-stem `*.md` sidecars
  - `logs/*.log`
  - `data/reports/mops_matrix_*.csv`

## Build, Test, and Development Commands
- `python -m venv venv && source venv/bin/activate`: local dev environment.
- `pip install -r requirements.txt && pip install -e .`: match CI dependency setup.
- `python ../skills/common/skill-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py --help`: quick skill runner sanity check.
- `python ../skills/common/skill-mops-financialreport-pdf-md/scripts/run_mops_financialreport_pdf_md.py 2382 2025 all --only-missing-files`: single-company download plus Markdown conversion.
- `python DownloadAll.py --year 2025 --quarter 1 --only-missing-files`: reproduce batch path through the skill.
- `python scripts/sheets_uploader.py --csv-only`: validate matrix export without credentials.
- `python scripts/sheets_uploader.py --upload`: test upload path when `.env` or secrets are configured.

## Coding Style & Naming Conventions
- Python: 4-space indentation, `snake_case` for functions/variables, `PascalCase` for classes.
- Workflow YAML: keep 2-space indentation and preserve explicit step names for GitHub Actions logs.
- Keep shell snippets in workflow steps POSIX-compatible and quote variable expansions.

## Download.yaml Change Rules
- Update schedule and quarter mapping consistently across all duplicated logic blocks (download target, summary, commit message, and status report).
- Preserve manual inputs: `year`, `quarter`, `delay`, `start_from`, `only_missing_files`, `upload_to_sheets`.
- Keep fallback behavior intact: if bulk download fails, workflow must still attempt limited per-company download and CSV backup.
- Do not remove `actions/upload-artifact` or status-report generation; they are operational diagnostics.

## Testing Guidelines
- For workflow edits, run local command equivalents above and verify outputs in `downloads/`, `logs/`, and `data/reports/`.
- Include at least one smoke run for download and one for matrix generation before PR.

## Commit & Pull Request Guidelines
- Prefer concise commit prefixes (`feat:`, `fix:`, `chore:`) for manual changes.
- PRs touching `Download.yaml` should include: changed trigger/schedule behavior, rollback plan, and sample command/output used for validation.
- If behavior affects release/tagging or auto-commit steps, call it out explicitly in PR description.

## Security & Configuration Tips
- Never commit secrets. Use GitHub Secrets: `GOOGLE_SHEETS_CREDENTIALS`, `GOOGLE_SHEET_ID`.
- Treat generated PDFs/CSVs as artifacts; review large diffs before merging.

## Retired Entry Points
- `scripts/pdf_to_md.py` and `scripts/batch_convert.py` are compatibility wrappers only. Do not add new conversion logic there; extend `skill-mops-financialreport-pdf-md` or `skill-mac-mini-ocr`.
- Avoid calling `scripts/mops_downloader.py` as the public workflow. Direct PDF acquisition remains in `mops_downloader/` for the skill to use internally.
