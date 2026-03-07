"""
Convert ASUS IR presentation PDFs to Markdown.

Usage:
    python scripts/ir_pdf_to_md.py downloads/2357/InvestorRelation/2025Q1_IR_Chinese.pdf
    python scripts/ir_pdf_to_md.py downloads/2357/InvestorRelation/   # convert all PDFs in dir
    python scripts/ir_pdf_to_md.py --force downloads/2357/InvestorRelation/  # overwrite existing
"""

import sys
import io
import re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

try:
    import pymupdf
except ImportError:
    print("ERROR: pymupdf not installed. Run: pip install pymupdf")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
SKIP_TEXTS = {"CONFIDENTIAL AND PROPRIETARY INFORMATION"}

# A block is "table-like" if it spans this many points wide and has >= 2 \n-separated values
TABLE_BLOCK_MIN_WIDTH = 400
TABLE_MIN_COLS = 2

# At least one non-first cell must match this to confirm it's a data table
_NUMERIC_RE = re.compile(
    r'^[\(\-]?[\d,]+[\.\d]*[\)\%]?$'   # numbers / (negative) / percentages
    r'|^\-\d'                            # negative number starting with -
    r'|^[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4}$'  # date: "Sep 30, 2025"
    r'|^\d+%$'                           # percentage
)

PAGE_TITLE_HINTS = [
    ("損益", "損益表"),
    ("業外損益", "業外損益"),
    ("資產負債表", "資產負債表"),
    ("營收組合", "營收組合"),
    ("營運展望", "營運展望"),
    ("System Business Group", "System Business Group"),
    ("Open Platform Business Group", "Open Platform Business Group"),
    ("AIoT Business Group", "AIoT Business Group"),
    ("議程", "議程"),
    ("聲明", "聲明"),
    ("問與答", "問與答"),
    ("財務結果", "章節標題"),
    ("策略與展望", "章節標題"),
]


# ---------------------------------------------------------------------------
# Block parsing
# ---------------------------------------------------------------------------
def get_blocks(page):
    """Return all text blocks sorted by (y0, x0)."""
    blocks = []
    for b in page.get_text("blocks"):
        x0, y0, x1, y1, text, _, btype = b
        if btype != 0:
            continue
        text = text.strip()
        if not text:
            continue
        # Skip watermark / confidential footer / standalone page numbers
        if "CONFIDENTIAL AND PROPRIETARY INFORMATION" in text:
            continue
        if re.fullmatch(r"\d{1,3}", text):
            continue
        # Skip unaudited footnote
        if text.startswith("(unaudited"):
            continue
        # Skip source footnote
        if text.startswith("(Source:"):
            continue
        blocks.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1, "text": text})
    return sorted(blocks, key=lambda b: (round(b["y0"] / 5) * 5, b["x0"]))


def is_table_block(block):
    """Wide block where at least one non-header cell is numeric = financial table row."""
    width = block["x1"] - block["x0"]
    if width < TABLE_BLOCK_MIN_WIDTH:
        return False
    values = [v.strip() for v in block["text"].split("\n") if v.strip()]
    if len(values) < TABLE_MIN_COLS:
        return False
    # At least one non-first value must be numeric/date to qualify as a table row
    return any(_NUMERIC_RE.match(v) for v in values[1:])


def parse_table_values(block):
    return [v.strip() for v in block["text"].split("\n") if v.strip()]


# ---------------------------------------------------------------------------
# Table grouping
# ---------------------------------------------------------------------------
def clean_block_lines(text):
    """Split block text into lines, stripping trailing standalone page numbers."""
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    while lines and re.fullmatch(r"\d{1,3}", lines[-1]):
        lines.pop()
    return lines


def extract_tables_and_text(blocks):
    """
    Returns list of items:
      {"type": "table", "rows": [[col, col, ...], ...]}
      {"type": "text",  "text": str, "is_bullet": bool}
    Peek-ahead: if a non-numeric wide block is immediately followed by table blocks
    with the same column count, treat it as the table header row.
    """
    items = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if is_table_block(b):
            # Collect consecutive table blocks
            table_rows = []
            while i < len(blocks) and is_table_block(blocks[i]):
                table_rows.append(parse_table_values(blocks[i]))
                i += 1
            col_count = max(set(len(r) for r in table_rows),
                            key=lambda n: sum(1 for r in table_rows if len(r) == n))
            padded = [r + [""] * (col_count - len(r)) if len(r) <= col_count
                      else r[:col_count] for r in table_rows]
            items.append({"type": "table", "rows": padded})
        else:
            lines = clean_block_lines(b["text"])
            if not lines:
                i += 1
                continue

            # Peek: is the next block a table with matching column count?
            next_j = i + 1
            while next_j < len(blocks) and not is_table_block(blocks[next_j]):
                next_j += 1

            used_as_header = False
            if (next_j < len(blocks)
                    and next_j == i + 1          # immediately adjacent
                    and len(lines) >= TABLE_MIN_COLS
                    and not any(_NUMERIC_RE.match(v) for v in lines)):
                # Check column count matches
                first_data = parse_table_values(blocks[next_j])
                if len(lines) == len(first_data):
                    # Collect the table rows
                    table_rows = []
                    k = next_j
                    while k < len(blocks) and is_table_block(blocks[k]):
                        table_rows.append(parse_table_values(blocks[k]))
                        k += 1
                    col_count = len(lines)
                    padded = [r + [""] * (col_count - len(r)) if len(r) <= col_count
                              else r[:col_count] for r in table_rows]
                    items.append({"type": "table", "rows": [lines] + padded})
                    i = k
                    used_as_header = True

            if not used_as_header:
                for ln in lines:
                    is_bullet = ln.startswith("•") or ln.startswith("·")
                    if is_bullet:
                        ln = "- " + ln.lstrip("•·").strip()
                    items.append({"type": "text", "text": ln, "is_bullet": is_bullet})
                i += 1
    return items


# ---------------------------------------------------------------------------
# Markdown rendering
# ---------------------------------------------------------------------------
def render_markdown(items):
    md = []
    for item in items:
        if item["type"] == "table":
            rows = item["rows"]
            if not rows:
                continue
            header = rows[0]
            md.append("| " + " | ".join(header) + " |")
            md.append("|" + "|".join(" --- " for _ in header) + "|")
            for row in rows[1:]:
                md.append("| " + " | ".join(row) + " |")
            md.append("")
        else:
            md.append(item["text"])
    return "\n".join(md)


# ---------------------------------------------------------------------------
# Page title inference
# ---------------------------------------------------------------------------
def infer_page_title(raw_text):
    for hint, label in PAGE_TITLE_HINTS:
        if hint in raw_text:
            return label
    return ""


# ---------------------------------------------------------------------------
# Main conversion
# ---------------------------------------------------------------------------
def convert_pdf(pdf_path: Path) -> Path:
    doc = pymupdf.open(str(pdf_path))
    stem = pdf_path.stem  # e.g. 2025Q3_IR_Chinese
    quarter = stem.replace("_IR_Chinese", "")  # e.g. 2025Q3

    q_map = {"Q1": "第1季", "Q2": "第2季", "Q3": "第3季", "Q4": "第4季"}
    year = quarter[:4]
    q = quarter[4:]
    q_label = q_map.get(q, q)
    title = f"華碩電腦 {year}年{q_label}投資人說明會"
    url = (f"https://www.asus.com/event/Investor/Content/attachment/"
           f"{quarter}%20IR(Chinese).pdf")

    md_lines = [
        f"# {title}",
        "",
        "**來源**：ASUS Investor Relations",
        "**類型**：Unaudited Brand Consolidated Financials",
        f"**URL**：{url}",
        "",
        "---",
        "",
    ]

    for page_num, page in enumerate(doc, 1):
        raw_text = page.get_text()
        page_title = infer_page_title(raw_text)

        header = f"## Page {page_num}"
        if page_title:
            header += f" — {page_title}"
        md_lines.append(header)
        md_lines.append("")

        blocks = get_blocks(page)
        items = extract_tables_and_text(blocks)
        content = render_markdown(items)
        if content.strip():
            md_lines.append(content)
        md_lines.append("")
        md_lines.append("---")
        md_lines.append("")

    out_path = pdf_path.with_suffix(".md")
    out_path.write_text("\n".join(md_lines), encoding="utf-8")
    print(f"✅ {out_path.name}")
    return out_path


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    args = sys.argv[1:]
    force = "--force" in args
    args = [a for a in args if not a.startswith("--")]

    if not args:
        print(__doc__)
        sys.exit(1)

    target = Path(args[0])

    if target.is_dir():
        pdfs = sorted(target.glob("*_IR_Chinese.pdf"))
        if not pdfs:
            print(f"No *_IR_Chinese.pdf files found in {target}")
            sys.exit(1)
        for pdf in pdfs:
            md_path = pdf.with_suffix(".md")
            if md_path.exists() and not force:
                print(f"⏭  {pdf.name} → already exists, skipping")
                continue
            convert_pdf(pdf)
    elif target.suffix == ".pdf":
        convert_pdf(target)
    else:
        print(f"ERROR: {target} is not a PDF or directory")
        sys.exit(1)


if __name__ == "__main__":
    main()
