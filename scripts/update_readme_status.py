"""Auto-update README.md Current Download Status section.

Scans:
  1. data/reports/mops_matrix_*.csv                    → 季財報 概況
  2. downloads/*.pdf + StockID_TWSE_TPEX.csv            → 本季申報焦點 (filing-deadline banner)

Deadline dates come from skill-mops-fetch/scripts/filing_deadlines.py, the
single source of truth shared with generate_mops_health.py and the CI
workflows, so this file never re-derives month/day deadline windows itself.

Usage:
  python scripts/update_readme_status.py
"""

import csv
import re
import sys
from datetime import datetime
from pathlib import Path

REPO_ROOT     = Path(__file__).parent.parent
README        = REPO_ROOT / "README.md"
DOWNLOADS     = REPO_ROOT / "downloads"
WATCHLIST_CSV = REPO_ROOT / "StockID_TWSE_TPEX.csv"

MARKER_BEGIN = "<!-- BEGIN_STATUS -->"
MARKER_END   = "<!-- END_STATUS -->"


def _find_skill_scripts_dir() -> Path:
    candidates = [
        REPO_ROOT.parent / "skills" / "common" / "skill-mops-fetch" / "scripts",
        REPO_ROOT / "skills" / "common" / "skill-mops-fetch" / "scripts",
        REPO_ROOT / "skills" / "skill-mops-fetch" / "scripts",
    ]
    for candidate in candidates:
        if (candidate / "filing_deadlines.py").is_file():
            return candidate
    raise SystemExit(
        "Cannot find skill-mops-fetch filing_deadlines.py. "
        "Expected ../skills/common/skill-mops-fetch/scripts/ or skills/common/..."
    )


sys.path.insert(0, str(_find_skill_scripts_dir()))
import filing_deadlines as fd  # noqa: E402


# ── CSV parsing ──────────────────────────────────────────────────────────────

def latest_csv() -> Path | None:
    csvs = sorted((REPO_ROOT / "data" / "reports").glob("mops_matrix_*.csv"))
    return csvs[-1] if csvs else None


def parse_matrix(csv_path: Path):
    with open(csv_path, encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    quarters = [k for k in rows[0] if k not in ("代號", "名稱", "process_timestamp")]
    total    = len(rows)
    stats    = {q: sum(1 for r in rows if r[q] and r[q] != "-") for q in quarters}
    names    = {r["代號"]: r["名稱"] for r in rows}
    return rows, quarters, stats, total, names


def _parse_quarter(quarter: str) -> tuple[int, int] | None:
    m = re.match(r"^(\d{4}) Q([1-4])$", quarter)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def _quarter_filename_prefix(quarter: str, code: str) -> str | None:
    parsed = _parse_quarter(quarter)
    if not parsed:
        return None
    year, q = parsed
    return f"{year}{q:02d}_{code}_"


# ── live downloads/ scan (source of truth for both the deadline banner and
#    the 季財報 概況 coverage stats below, so the two never disagree even if
#    data/reports/mops_matrix_*.csv hasn't been regenerated recently) ─────────

def scan_filed_by_quarter(downloads_dir: Path) -> dict[str, set[str]]:
    """{"<year> Q<quarter>": {company_id, ...}} built from downloads/*.pdf filenames.

    Matches on the "<YYYYQQ>_<company_id>_" filename prefix used by every MOPS
    financial-report PDF, regardless of report-type code (AI1/AI2/AI3/AE2/AI4/
    AIA/...), so a new report-type code never silently drops out of the count.
    This also naturally excludes non-MOPS PDFs (e.g. IR presentations under
    downloads/<id>/InvestorRelation/), which use a different filename shape.
    """
    filed: dict[str, set[str]] = {}
    if not downloads_dir.exists():
        return filed
    for pdf in downloads_dir.rglob("*.pdf"):
        parts = pdf.stem.split("_")
        if len(parts) < 3 or len(parts[0]) != 6 or not parts[0].isdigit():
            continue
        year, q = parts[0][:4], parts[0][4:6]
        try:
            label = f"{int(year)} Q{int(q)}"
        except ValueError:
            continue
        filed.setdefault(label, set()).add(parts[1])
    return filed


def load_watchlist(watchlist_csv: Path) -> list[tuple[str, str]]:
    watchlist: list[tuple[str, str]] = []
    if watchlist_csv.exists():
        with open(watchlist_csv, encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                code = (row.get("代號") or "").strip()
                if code:
                    watchlist.append((code, (row.get("名稱") or "").strip()))
    return watchlist


# ── filing-deadline banner ───────────────────────────────────────────────────

def generate_filing_focus_section(filed_by_quarter: dict[str, set[str]], watchlist: list[tuple[str, str]]) -> list[str]:
    focus_qd = fd.current_focus_quarter()
    days = fd.days_since_deadline(focus_qd)
    filed_ids = filed_by_quarter.get(focus_qd.label, set())
    overdue = [(code, name) for code, name in watchlist if code not in filed_ids]
    total = len(watchlist)
    filed = total - len(overdue)
    coverage = f"{filed / total:.0%}" if total else "—"

    lines = [
        f"### 📅 本季申報焦點：{focus_qd.label}",
        "",
        f"- 一般上市公司申報期限：**{focus_qd.deadline.isoformat()}**"
        f"（{'已逾期 ' + str(days) + ' 天' if days >= 0 else '尚有 ' + str(-days) + ' 天'}）",
        f"- 金控/銀行申報期限：{focus_qd.holding_deadline.isoformat()}",
        f"- 完成度：{filed} / {total}（{coverage}）｜ 尚缺 {len(overdue)} 家",
    ]

    if overdue:
        lines += [
            "",
            "<details>",
            f"<summary>尚未申報 {focus_qd.label} 財報的公司（{len(overdue)}）</summary>",
            "",
            "| 代號 | 名稱 |",
            "|------|------|",
        ]
        for code, name in overdue:
            lines.append(f"| {code} | {name} |")
        lines += ["", "</details>"]

    lines += [
        "",
        "> ⚠️ 金控/銀行申報期限較晚，上列名單可能包含少數仍在合法期限內的金控/銀行公司。",
    ]
    return lines


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
    rows, quarters, _stats, matrix_total, _names = parse_matrix(csv_path)
    today    = datetime.now().strftime("%Y-%m-%d")
    lines    = generate_mops_pdfs_table(rows, quarters, csv_path)

    # Live scan: source of truth for both the deadline banner and the 季財報
    # 概況 coverage below, so they always agree with each other regardless of
    # how stale data/reports/mops_matrix_*.csv (used only for the per-cell
    # AI1/AI3 links in the table above) happens to be.
    filed_by_quarter = scan_filed_by_quarter(DOWNLOADS)
    watchlist = load_watchlist(WATCHLIST_CSV)
    total = len(watchlist) or matrix_total

    # ── 季財報 概況 ──
    lines += [
        "",
        "---",
        "",
        "## 📊 Current Download Status",
        "",
        f"> **Last Updated**: {today} | **Source**: `downloads/` (live scan) + `{csv_path.name}`",
        "",
        f"**{total} companies tracked**",
        "",
    ]
    lines += generate_filing_focus_section(filed_by_quarter, watchlist)
    lines += [
        "",
        "---",
        "",
        "### 季財報 概況",
        "",
        "| Quarter | 季財報 | Coverage | Notes |",
        "|---------|--------|----------|-------|",
    ]
    for q in quarters:
        cnt      = len(filed_by_quarter.get(q, set()))
        coverage = f"{cnt / total:.0%}" if cnt and total else "—"
        parsed   = _parse_quarter(q)
        note     = fd.deadline_note(*parsed) if parsed else ""
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
