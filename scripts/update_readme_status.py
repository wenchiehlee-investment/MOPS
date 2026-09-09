"""Auto-update README.md Current Download Status section.

Scans:
  1. data/reports/mops_matrix_*.csv → 季財報 概況

Usage:
  python scripts/update_readme_status.py
"""

import csv
import re
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
README    = REPO_ROOT / "README.md"
DOWNLOADS = REPO_ROOT / "downloads"

MARKER_BEGIN = "<!-- BEGIN_STATUS -->"
MARKER_END   = "<!-- END_STATUS -->"

DEADLINE_NOTES = {
    "2026 Q1": "Filing deadline: May 15",
    "2025 Q4": "Filing deadline: Mar 31 (next year)",
}


# ── CSV parsing ──────────────────────────────────────────────────────────────

def latest_csv() -> Path | None:
    csvs = sorted((REPO_ROOT / "data" / "reports").glob("mops_matrix_*.csv"))
    return csvs[-1] if csvs else None


def parse_matrix(csv_path: Path):
    with open(csv_path, encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    quarters = [k for k in rows[0] if k not in ("代號", "名稱")]
    total    = len(rows)
    stats    = {q: sum(1 for r in rows if r[q] and r[q] != "-") for q in quarters}
    names    = {r["代號"]: r["名稱"] for r in rows}
    return rows, quarters, stats, total, names


def _quarter_filename_prefix(quarter: str, code: str) -> str | None:
    m = re.match(r"^(\d{4}) Q([1-4])$", quarter)
    if not m:
        return None
    return f"{m.group(1)}{int(m.group(2)):02d}_{code}_"


def _pdf_link(code: str, quarter: str, report_type: str) -> str:
    prefix = _quarter_filename_prefix(quarter, code)
    if not prefix:
        return report_type

    pdf_name = f"{prefix}{report_type}.pdf"
    pdf_path = DOWNLOADS / code / pdf_name
    if pdf_path.exists():
        return f"[{report_type}](downloads/{code}/{pdf_name})"

    matches = sorted((DOWNLOADS / code).glob(f"{prefix}{report_type}*.pdf"))
    if matches:
        rel = matches[0].relative_to(REPO_ROOT).as_posix()
        return f"[{report_type}]({rel})"

    return report_type


def _pdf_cell(code: str, quarter: str, value: str) -> str:
    if not value or value == "-":
        return "-"
    return " / ".join(_pdf_link(code, quarter, t.strip()) for t in value.split("/") if t.strip())


def generate_mops_pdfs_table(rows: list[dict[str, str]], quarters: list[str], csv_path: Path) -> list[str]:
    lines = [
        "## Current MOPS PDFs",
        "",
        f"> **Source**: `{csv_path.name}`",
        "",
        "| 代號 | 名稱 | " + " | ".join(quarters) + " |",
        "|------|------|" + "|".join(["------"] * len(quarters)) + "|",
    ]
    for row in rows:
        code = row["代號"]
        cells = [_pdf_cell(code, q, row[q]) for q in quarters]
        lines.append(f"| {code} | {row['名稱']} | " + " | ".join(cells) + " |")
    return lines


# ── section generator ────────────────────────────────────────────────────────

def generate_section(csv_path: Path) -> str:
    rows, quarters, stats, total, _names = parse_matrix(csv_path)
    today    = datetime.now().strftime("%Y-%m-%d")
    lines    = generate_mops_pdfs_table(rows, quarters, csv_path)

    # ── 季財報 概況 ──
    lines += [
        "",
        "---",
        "",
        "## 📊 Current Download Status",
        "",
        f"> **Last Updated**: {today} | **Source**: `{csv_path.name}`",
        "",
        f"**{total} companies tracked**",
        "",
        "### 季財報 概況",
        "",
        "| Quarter | 季財報 | Coverage | Notes |",
        "|---------|--------|----------|-------|",
    ]
    for q in quarters:
        cnt      = stats[q]
        coverage = f"{cnt / total:.0%}" if cnt else "—"
        note     = DEADLINE_NOTES.get(q, "")
        lines.append(f"| {q} | {cnt} / {total} | {coverage} | {note} |")

    return "\n".join(lines)


# ── README update ────────────────────────────────────────────────────────────

def update_readme(section: str) -> bool:
    content = README.read_text(encoding="utf-8")
    begin   = content.find(MARKER_BEGIN)
    end     = content.find(MARKER_END)
    if begin == -1 or end == -1:
        print("ERROR: markers not found in README.md")
        return False
    new_content = (
        content[:begin]
        + MARKER_BEGIN + "\n\n"
        + section + "\n\n"
        + MARKER_END
        + content[end + len(MARKER_END):]
    )
    README.write_text(new_content, encoding="utf-8", newline="\n")
    return True


def main():
    csv_path = latest_csv()
    if not csv_path:
        print("No mops_matrix CSV found in data/reports/")
        return
    print(f"Using CSV: {csv_path.name}")
    section = generate_section(csv_path)
    if update_readme(section):
        print("README.md updated successfully")
    else:
        print("Failed — add markers to README.md first")


if __name__ == "__main__":
    main()
