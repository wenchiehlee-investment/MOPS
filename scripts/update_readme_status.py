"""Auto-update README.md Current Download Status section.

Scans:
  1. data/reports/mops_matrix_*.csv       → 季財報 概況
  2. downloads/*/InvestorRelation/*.md    → 法說會 PDF/MD (per quarter)
  3. downloads/*/InvestorRelation/法說會逐字稿/*.md → 逐字稿
  4. downloads/*/News/*.md                → 新聞 (title from first # line)

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

IR_RE = re.compile(r"^(\d{4})Q(\d)_IR_Chinese\.md$")

# Which calendar months correspond to which results quarter
# (meetings are held after the quarter ends)
_MONTH_TO_RESULT_Q = {
    11: 3, 12: 3,   # Nov/Dec → Q3 results
     2: 4,  3: 4,   # Feb/Mar → Q4 results
     5: 1,  6: 1,   # May/Jun → Q1 results
     8: 2,  9: 2,   # Aug/Sep → Q2 results
}

DEADLINE_NOTES = {
    "2026 Q1": "Filing deadline: May 15",
    "2025 Q4": "Filing deadline: Mar 31 (next year)",
}


# ── helpers ─────────────────────────────────────────────────────────────────

def _quarter_key(q: str):
    m = re.match(r"(\d{4}) Q(\d)", q)
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def _date_to_result_quarter(date_str: str) -> str | None:
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d")
        rq = _MONTH_TO_RESULT_Q.get(d.month)
        if rq is None:
            return None
        year = d.year if d.month >= 2 else d.year - 1
        return f"{year} Q{rq}"
    except ValueError:
        return None


def _extract_date(filename: str) -> str:
    m = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
    if m:
        return m.group(1)
    m = re.search(r"(\d{8})", filename)
    if m:
        raw = m.group(1)
        return f"{raw[:4]}-{raw[4:6]}-{raw[6:]}"
    return "unknown"


def _read_title(path: Path) -> str:
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line.startswith("# "):
                    return line[2:].strip()
    except Exception:
        pass
    return path.stem


# ── scanners ────────────────────────────────────────────────────────────────

def scan_ir(code: str) -> dict[str, str]:
    """Quarter → relative MD path."""
    ir_dir = DOWNLOADS / code / "InvestorRelation"
    result = {}
    if not ir_dir.exists():
        return result
    for f in ir_dir.glob("*_IR_Chinese.md"):
        m = IR_RE.match(f.name)
        if m:
            q = f"{m.group(1)} Q{m.group(2)}"
            result[q] = f"downloads/{code}/InvestorRelation/{f.name}"
    return result


def scan_transcripts(code: str) -> list[tuple[str, str]]:
    """List of (date, relative path)."""
    ts_dir = DOWNLOADS / code / "InvestorRelation" / "法說會逐字稿"
    result = []
    if not ts_dir.exists():
        return result
    for f in sorted(ts_dir.glob("*.md")):
        date = _extract_date(f.name)
        result.append((date, f"downloads/{code}/InvestorRelation/法說會逐字稿/{f.name}"))
    return result


def scan_news(code: str) -> list[tuple[str, str, str]]:
    """List of (date, title, relative path), newest first."""
    news_dir = DOWNLOADS / code / "News"
    result = []
    if not news_dir.exists():
        return result
    for f in sorted(news_dir.glob("*.md"), reverse=True):
        date  = _extract_date(f.name)
        title = _read_title(f)
        result.append((date, title, f"downloads/{code}/News/{f.name}"))
    return result


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
    return quarters, stats, total, names


# ── section generator ────────────────────────────────────────────────────────

def generate_section(csv_path: Path) -> str:
    quarters, stats, total, names = parse_matrix(csv_path)
    today    = datetime.now().strftime("%Y-%m-%d")
    lines    = []

    # ── 季財報 概況 ──
    lines += [
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

    # ── 法說會 & 新聞 資料庫 ──
    company_codes = sorted(p.name for p in DOWNLOADS.iterdir() if p.is_dir())
    companies = []
    for code in company_codes:
        ir          = scan_ir(code)
        transcripts = scan_transcripts(code)
        news        = scan_news(code)
        if ir or transcripts or news:
            companies.append((code, names.get(code, code), ir, transcripts, news))

    if not companies:
        return "\n".join(lines)

    lines += [
        "",
        "---",
        "",
        "### 📂 法說會 & 新聞 資料庫",
        "",
        "#### 完成度概況",
        "",
        "> 各公司法說會簡報、逐字稿、新聞收錄數量。點擊公司名稱展開詳細連結。",
        "",
        "| 公司 | 法說會 PDF/MD | 逐字稿 | 新聞 |",
        "|------|:------------:|:------:|:----:|",
    ]

    for code, name, ir, transcripts, news in companies:
        anchor    = f"#{code}-{name}".lower().replace(" ", "-")
        ir_cnt    = f"{len(ir)} 季" if ir else "—"
        ts_cnt    = str(len(transcripts)) if transcripts else "—"
        news_cnt  = str(len(news)) if news else "—"
        lines.append(f"| [{code} {name}]({anchor}) | {ir_cnt} | {ts_cnt} | {news_cnt} |")

    lines.append("")
    lines.append(f"**覆蓋率**：{len(companies)} / {total} companies")

    # per-company details
    for code, name, ir, transcripts, news in companies:
        lines += ["", "---", "", f"#### {code} {name}"]

        if ir:
            # group transcripts by result quarter
            ts_by_q: dict[str, list] = {}
            for date, path in transcripts:
                rq = _date_to_result_quarter(date)
                if rq:
                    ts_by_q.setdefault(rq, []).append((date, path))

            sorted_quarters = sorted(ir, key=_quarter_key, reverse=True)
            lines += [
                "",
                "<details>",
                "<summary>法說會（季度）</summary>",
                "",
                "| Quarter | 法說會 PDF/MD | 逐字稿 |",
                "|---------|:------------:|:------:|",
            ]
            for q in sorted_quarters:
                ir_link = f"[MD]({ir[q]})"
                ts      = ts_by_q.get(q, [])
                ts_md   = " / ".join(f"[{d}]({p})" for d, p in ts) if ts else "—"
                lines.append(f"| {q} | {ir_link} | {ts_md} |")
            lines += ["", "</details>"]

        if news:
            lines += [
                "",
                "<details>",
                "<summary>新聞</summary>",
                "",
                "| 日期 | 標題 |",
                "|------|------|",
            ]
            for date, title, path in news:
                lines.append(f"| {date} | [{title}]({path}) |")
            lines += ["", "</details>"]

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
    README.write_text(new_content, encoding="utf-8")
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
