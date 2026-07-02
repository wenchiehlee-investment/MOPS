#!/usr/bin/env python3
"""
MOPS PDF to Markdown Converter - "AI-TOC Linked" Version
Specialized for Taiwan MOPS Financial Reports (財務報告).
First Pass: Extracts Table of Contents (TOC).
Second Pass: Uses PyMuPDF find_tables() + text blocks to structure output.
"""

import os
import argparse
import sys
import re
import tempfile
from pathlib import Path
from typing import Optional, List, Tuple, Dict
import logging
import traceback

try:
    import fitz  # PyMuPDF
except ImportError:
    print("❌ Error: PyMuPDF is not installed. Please install it using:")
    print("   pip install pymupdf")
    sys.exit(1)

try:
    from ocr_client import transcribe_document_to_markdown
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Patterns
PAGE_NUM_PATTERN = re.compile(r'^~\s*\d+\s*~$')
CONT_PATTERN = re.compile(r'^\(續\s*[次上]\s*頁\)$')
# Cross-page table continuation markers
CONT_NEXT_RE = re.compile(r'[（(](?:接次頁|續次頁)[）)]')
CONT_PREV_RE = re.compile(r'[（(](?:承前頁|續上頁)[）)]')
# Matches "五、 合併資產負債表" on one line (title only, page number on next line)
TOC_TITLE_PATTERN = re.compile(
    r'^(一|二|三|四|五|六|七|八|九|十)、\s*(.+)$'
)
# Matches "11 ~ 12" or "13" (page number line)
TOC_PAGE_PATTERN = re.compile(r'^(\d+)\s*(?:[~～]\s*(\d+))?$')
# Also match single-line format: "五、 合併資產負債表 11 ~ 12"
TOC_ENTRY_PATTERN = re.compile(
    r'^(一|二|三|四|五|六|七|八|九|十)、\s*(.+)\s+(\d+)\s*(?:[~～]\s*(\d+))?$'
)
# Minimum non-empty columns to consider a text-strategy detection a real table
MIN_TABLE_COLS = 3


def extract_toc(doc) -> Tuple[Dict[int, str], Dict[int, int]]:
    """Pass 1: Scans early pages to build a page-to-section map.

    Handles both formats:
    - Single line: "五、 合併資產負債表 11 ~ 12"
    - Two lines:   "五、 合併資產負債表" then "11 ~ 12"

    Returns:
        toc_map: {page_num_0indexed: section_title}
        toc_ranges: {start_page_0indexed: end_page_0indexed}
    """
    toc_map = {}
    toc_ranges = {}

    for p_idx in range(1, min(5, len(doc))):
        page = doc.load_page(p_idx)
        text = page.get_text("text")
        lines = [l.strip() for l in text.splitlines()]

        pending_title = None  # (prefix, title)
        for line in lines:
            if not line:
                continue

            # Try single-line format first
            match = TOC_ENTRY_PATTERN.match(line)
            if match:
                prefix, title, start_pg, end_pg = match.groups()
                start_pg = int(start_pg)
                end_pg = int(end_pg) if end_pg else start_pg
                toc_map[start_pg - 1] = f"{prefix}、{title.strip()}"
                toc_ranges[start_pg - 1] = end_pg - 1
                logger.info(f"📍 TOC Found: {prefix}、{title.strip()} → Page {start_pg}")
                pending_title = None
                continue

            # Try two-line format: title line
            title_match = TOC_TITLE_PATTERN.match(line)
            if title_match:
                pending_title = (title_match.group(1), title_match.group(2).strip())
                continue

            # Try two-line format: page number line after a title
            if pending_title:
                page_match = TOC_PAGE_PATTERN.match(line)
                if page_match:
                    prefix, title = pending_title
                    start_pg = int(page_match.group(1))
                    end_pg_str = page_match.group(2)
                    end_pg = int(end_pg_str) if end_pg_str else start_pg
                    toc_map[start_pg - 1] = f"{prefix}、{title}"
                    toc_ranges[start_pg - 1] = end_pg - 1
                    logger.info(f"📍 TOC Found: {prefix}、{title} → Page {start_pg}")
                pending_title = None

    return toc_map, toc_ranges


def _is_noise_line(text: str) -> bool:
    """Check if a text line is noise (page number, continuation marker)."""
    text = text.strip()
    if not text:
        return True
    if PAGE_NUM_PATTERN.match(text):
        return True
    if CONT_PATTERN.match(text):
        return True
    return False


def _clean_cell(cell: Optional[str]) -> str:
    """Clean a table cell value."""
    if cell is None:
        return ""
    cell = cell.strip()
    cell = re.sub(r'~\s*\d+\s*~', '', cell)
    cell = re.sub(r'\(續\s*[次上]\s*頁\)', '', cell)
    lines = cell.split('\n')
    lines = [' '.join(l.split()) for l in lines]
    cell = '\n'.join(l for l in lines if l)
    return cell


def _block_in_table_area(block_bbox, table_rects) -> bool:
    """Check if a text block's center overlaps with any table area."""
    bx0, by0, bx1, by1 = block_bbox
    b_mid_x = (bx0 + bx1) / 2
    b_mid_y = (by0 + by1) / 2
    for trect in table_rects:
        if (trect[0] - 5 <= b_mid_x <= trect[2] + 5 and
                trect[1] - 5 <= b_mid_y <= trect[3] + 5):
            return True
    return False


def _count_effective_cols(rows: List[List[str]]) -> int:
    """Count how many columns have at least some non-empty data."""
    if not rows:
        return 0
    max_cols = max(len(r) for r in rows)
    count = 0
    for col_idx in range(max_cols):
        for row in rows:
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                count += 1
                break
    return count


def _merge_continuation_rows(rows: List[List[str]]) -> List[List[str]]:
    """Merge rows where only the label column has text (continuation of previous row's label).

    For financial tables, continuation rows have:
    - Empty first column (account code)
    - Text in second column (continuation of account name)
    - All other columns empty
    """
    if len(rows) < 2:
        return rows

    merged = [list(rows[0])]
    for row in rows[1:]:
        if not row:
            continue
        # Check if this is a continuation row:
        # - First col (code) is empty
        # - Second col (name) has text
        # - All remaining cols are empty
        is_continuation = False
        if len(row) >= 2 and not row[0].strip() and row[1].strip():
            remaining = row[2:] if len(row) > 2 else []
            if all(not c.strip() for c in remaining):
                is_continuation = True

        if is_continuation and merged:
            # Append the continuation text to the previous row's name column
            prev = merged[-1]
            if len(prev) >= 2:
                prev[1] = (prev[1] + row[1]).strip()
            continue

        merged.append(list(row))

    return merged


def _postprocess_table(raw_rows: List[List[str]]) -> List[List[str]]:
    """Post-process table: clean cells, remove empty rows/cols, merge continuations."""
    if not raw_rows:
        return []

    # Clean all cells
    rows = [[_clean_cell(c) for c in row] for row in raw_rows]

    # Normalize row lengths
    max_cols = max(len(r) for r in rows)
    rows = [r + [""] * (max_cols - len(r)) for r in rows]

    # Remove fully empty rows
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return []

    # Merge continuation rows (account name spanning multiple lines)
    rows = _merge_continuation_rows(rows)

    # Remove fully empty columns
    non_empty_col_indices = []
    for col_idx in range(max_cols):
        if any(row[col_idx].strip() for row in rows if col_idx < len(row)):
            non_empty_col_indices.append(col_idx)

    if not non_empty_col_indices:
        return []

    rows = [[row[i] for i in non_empty_col_indices] for row in rows]

    return rows


def format_grid_table(rows: List[List[str]]) -> List[str]:
    """Format processed table rows as Markdown table."""
    if not rows:
        return []

    max_cols = max(len(r) for r in rows)
    if max_cols < 2:
        return [r[0] for r in rows if r[0].strip()] + [""]

    # Normalize
    rows = [r + [""] * (max_cols - len(r)) for r in rows]

    # Replace newlines and escape pipes
    for i in range(len(rows)):
        for j in range(len(rows[i])):
            rows[i][j] = rows[i][j].replace('\n', ' ').replace('|', '\\|')

    md = []
    md.append("| " + " | ".join(rows[0]) + " |")
    md.append("| " + " | ".join(["---"] * max_cols) + " |")
    for i in range(1, len(rows)):
        md.append("| " + " | ".join(rows[i]) + " |")
    md.append("")
    return md


def _rows_similar(row1: List[str], row2: List[str], threshold: float = 0.6) -> bool:
    """Check if two table rows have similar content (for repeated header detection)."""
    if not row1 and not row2:
        return True
    if not row1 or not row2:
        return False
    s1 = ''.join(c.strip() for c in row1)
    s2 = ''.join(c.strip() for c in row2)
    if not s1 and not s2:
        return True
    if not s1 or not s2:
        return False
    shorter, longer = (s1, s2) if len(s1) <= len(s2) else (s2, s1)
    return len(shorter) / len(longer) >= threshold


_HEADER_ROW_MAX_CHARS = 120  # rows longer than this are data rows, not headers


def _skip_continuation_header(orig_rows: List[List[str]],
                               cont_rows: List[List[str]]) -> List[List[str]]:
    """Skip repeated header rows at the top of a continuation table.

    Financial tables repeat column headers on each continuation page.
    Matches the continuation's leading rows against the original table's
    first rows; skips all that are sufficiently similar.

    Safeguard: only rows whose total content is short (≤ _HEADER_ROW_MAX_CHARS)
    are treated as candidate headers.  Long rows are data rows and are never
    skipped, even if their character-length ratio happens to be ≥ threshold.
    """
    if not cont_rows:
        return []
    skip = 0
    for i in range(min(4, len(orig_rows), len(cont_rows))):
        cont_content = ''.join(c.strip() for c in cont_rows[i])
        # If the continuation row is long it is a data row, never a header
        if len(cont_content) > _HEADER_ROW_MAX_CHARS:
            break
        if _rows_similar(orig_rows[i], cont_rows[i]):
            skip = i + 1
        else:
            break
    return cont_rows[skip:]


def _merge_cross_page_tables(all_content: list) -> list:
    """Merge tables split across PDF pages.

    Scans the sorted all_content list for the pattern:
        table  →  [text(接次頁)]  →  page  →  [text(承前頁)]  →  table
    and merges the two table fragments, stripping repeated headers from
    the continuation table.  Any leftover continuation-marker text items
    are removed.  Handles chains (table spanning 3+ pages) by re-checking
    each merged table.
    """
    items = list(all_content)
    result = []
    skip_set: set = set()
    _page_num_re = re.compile(r'^[-~]+\s*\d+\s*[-~]+$')

    i = 0
    while i < len(items):
        if i in skip_set:
            i += 1
            continue

        key, ctype, data = items[i]

        if ctype != "table":
            result.append(items[i])
            i += 1
            continue

        # --- look ahead for cross-page continuation ---
        j = i + 1
        found_cont_next = False
        found_page = False
        found_cont_prev = False
        cont_table_idx = None
        candidate_skips: List[int] = []

        eff = 0  # effective (non-skip) steps
        while j < len(items) and eff < 12:
            if j in skip_set:
                j += 1
                continue
            jkey, jtype, jdata = items[j]

            # Free-pass: page-number footers before the page break don't consume
            # effective steps (important after multi-page chain merges)
            if jtype == "text":
                text_str = str(jdata).strip()
                if not found_page and not found_cont_next:
                    if _page_num_re.match(text_str):
                        j += 1
                        continue  # free pass — no eff consumed
                    # Check for continuation markers before counting eff
                    if CONT_NEXT_RE.search(text_str):
                        eff += 1
                        found_cont_next = True
                        candidate_skips.append(j)
                        j += 1
                        continue
                    # Any other real text before the page break = not cross-page
                    break
                elif found_page:
                    if CONT_PREV_RE.search(text_str):
                        eff += 1
                        found_cont_prev = True
                    # Any text between the page marker and the continuation table
                    # (title text, date, sub-title, etc.) belongs to the continuation
                    # page's header and must be added to skip_set to avoid breaking
                    # the structural look-ahead on subsequent chain iterations.
                    candidate_skips.append(j)
                    j += 1
                    continue
                # Other text after page break — count it
            eff += 1

            if jtype == "page":
                found_page = True
                candidate_skips.append(j)
            elif jtype == "table":
                if found_page:
                    cont_table_idx = j
                break
            elif jtype == "section":
                break  # section boundary — never merge across it

            j += 1

        should_merge = (
            cont_table_idx is not None
            and found_page
            and (found_cont_next or found_cont_prev)
        )

        if should_merge:
            orig_rows = data
            cont_rows = items[cont_table_idx][2]
            tail_rows = _skip_continuation_header(orig_rows, cont_rows)

            # Only merge if the continuation actually contributes data rows;
            # an empty tail means the entire continuation was repeated headers
            # (or a false positive) — skip silently.
            if not tail_rows:
                result.append(items[i])
                i += 1
            else:
                merged_rows = orig_rows + tail_rows
                items[i] = (key, "table", merged_rows)

                for idx in candidate_skips:
                    skip_set.add(idx)
                skip_set.add(cont_table_idx)

                logger.info(
                    f"🔗 Merged cross-page table at item {i}: "
                    f"{len(orig_rows)}+{len(tail_rows)} rows"
                )
                # Re-check merged table for further continuations (3+ page tables)
                # Safety cap to prevent runaway chaining
                if len(merged_rows) > 5000:
                    result.append(items[i])
                    i += 1
        else:
            # --- Secondary: structural merge (no markers) ---
            # Pattern: table → (page-num text?) → page → table
            # Condition: continuation table's first row ≈ original's first row
            #            (repeated header) AND same column count.
            j2 = i + 1
            found_page2 = False
            struct_table_idx = None
            struct_skips: List[int] = []
            eff_steps = 0  # count only non-skip items toward window

            while j2 < len(items) and eff_steps < 10:
                if j2 in skip_set:
                    j2 += 1
                    continue
                jkey2, jtype2, jdata2 = items[j2]

                # Free-pass: page-number footers and short noise don't consume
                # effective steps (important when many merged pages leave behind
                # their page-number texts between the merged table and the next page)
                if jtype2 == "text":
                    text_str2 = str(jdata2).strip()
                    if not found_page2:
                        if _page_num_re.match(text_str2) or len(text_str2) <= 6:
                            j2 += 1
                            continue  # free pass — no eff_steps consumed
                        break  # real text before page break — not a continuation

                eff_steps += 1

                if jtype2 == "page":
                    found_page2 = True
                    struct_skips.append(j2)
                elif jtype2 == "table":
                    if found_page2:
                        struct_table_idx = j2
                    break
                elif jtype2 == "section":
                    break
                j2 += 1

            struct_merge = False
            if struct_table_idx is not None and found_page2 and data:
                orig_rows = data
                cont_rows = items[struct_table_idx][2]
                if cont_rows:
                    # Same column count (within ±1) AND similar first (header) row
                    orig_cols = len(orig_rows[0]) if orig_rows else 0
                    cont_cols = len(cont_rows[0]) if cont_rows else 0
                    same_cols = abs(orig_cols - cont_cols) <= 1
                    header_match = _rows_similar(orig_rows[0], cont_rows[0],
                                                 threshold=0.7)
                    if same_cols and header_match:
                        tail_rows = _skip_continuation_header(orig_rows, cont_rows)
                        if tail_rows:
                            merged_rows = orig_rows + tail_rows
                            items[i] = (key, "table", merged_rows)
                            for idx in struct_skips:
                                skip_set.add(idx)
                            skip_set.add(struct_table_idx)
                            logger.info(
                                f"🔗 Structural merge at item {i}: "
                                f"{len(orig_rows)}+{len(tail_rows)} rows "
                                f"(cols {orig_cols}≈{cont_cols})"
                            )
                            struct_merge = True
                            if len(merged_rows) > 5000:
                                result.append(items[i])
                                i += 1

            if not struct_merge:
                result.append(items[i])
                i += 1

    return result


def _find_page_tables(page, force_text: bool = False) -> List:
    """Find tables on a page using best available strategy.

    Args:
        page: PyMuPDF page object.
        force_text: If True, always try text strategy (for known table pages
                    from TOC like balance sheet, income statement).

    1. Try default (line-based) strategy first.
    2. If no tables found (or force_text), also try text-based strategy.
    3. Use whichever strategy found better coverage.
    4. Filter out false positives from text strategy.
    """
    line_tables = page.find_tables()
    text_tables = page.find_tables(strategy='text')

    # Collect valid text-strategy tables
    valid_text_tables = []
    for t in text_tables.tables:
        raw = t.extract()
        if not raw:
            continue
        effective_cols = _count_effective_cols(raw)
        if effective_cols >= MIN_TABLE_COLS:
            valid_text_tables.append(t)

    # If force_text or no line tables, prefer text tables
    if force_text and valid_text_tables:
        return valid_text_tables

    if line_tables.tables:
        # Compare coverage: if text strategy covers significantly more area, use it
        if valid_text_tables:
            line_area = sum(
                (t.bbox[2] - t.bbox[0]) * (t.bbox[3] - t.bbox[1])
                for t in line_tables.tables
            )
            text_area = sum(
                (t.bbox[2] - t.bbox[0]) * (t.bbox[3] - t.bbox[1])
                for t in valid_text_tables
            )
            if text_area > line_area * 2:
                return valid_text_tables
        return list(line_tables.tables)

    return valid_text_tables


_OCR_RESULTS_MARKER = "===============save results:==============="


def _clean_ocr_output(raw: str) -> str:
    """The mac-mini-ocr API response includes raw <|det|> layout-detection
    tags followed by a clean plain-text rendering after a marker line. Keep
    only the clean section; fall back to stripping the tags if the marker
    is absent.
    """
    if _OCR_RESULTS_MARKER in raw:
        raw = raw.split(_OCR_RESULTS_MARKER, 1)[1]
    else:
        raw = re.sub(r'<\|det\|>.*?<\|/det\|>', '', raw)
    raw = raw.replace('<PAGE>', '').strip()
    return raw


def _ocr_page(doc, pdf_path: Path, page_num: int, dpi: int = 200) -> Optional[str]:
    """Extract a single scanned page to a temp single-page PDF and transcribe
    it via the mac-mini-ocr API. Returns the Markdown text, or None on failure
    (caller falls back to the TODO:OCR placeholder).
    """
    tmp_path = None
    try:
        single_doc = fitz.open()
        single_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
        fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        single_doc.save(tmp_path)
        single_doc.close()

        markdown = transcribe_document_to_markdown(tmp_path, dpi=dpi)
        markdown = _clean_ocr_output(markdown) if markdown else None
        return markdown if markdown else None
    except Exception as e:
        logger.warning(f"⚠️ OCR failed for {pdf_path.name} page {page_num + 1}: {e}")
        return None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def convert_pdf_to_md(pdf_path: Path, output_path: Optional[Path] = None,
                       use_ocr: bool = True) -> bool:
    if not pdf_path.exists():
        return False
    if not output_path:
        output_path = pdf_path.with_suffix('.md')

    logger.info(f"📄 Converting (AI-TOC Mode): {pdf_path} -> {output_path}")

    try:
        doc = fitz.open(str(pdf_path))

        # --- Pass 1: Build Section Map ---
        toc_map, toc_ranges = extract_toc(doc)

        # Build set of pages that belong to financial statement sections
        # These pages should use text-based table detection
        FINANCIAL_KEYWORDS = ('資產負債表', '損益表', '權益變動表', '現金流量表',
                              '明細表')
        table_pages = set()
        for start_pg, title in toc_map.items():
            if any(kw in title for kw in FINANCIAL_KEYWORDS):
                end_pg = toc_ranges.get(start_pg, start_pg)
                for pg in range(start_pg, end_pg + 1):
                    table_pages.add(pg)

        # content_items: (sort_key, type, data)
        all_content = []

        # --- Pass 2: Extract Content ---
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            page_offset = page_num * 100000  # ensure page ordering

            # Insert page separator
            all_content.append((page_offset, "page", page_num + 1))

            # Insert section headers from TOC
            if page_num in toc_map:
                section_title = toc_map[page_num]
                all_content.append((page_offset, "section", section_title))

            # Detect scanned (image-only) pages
            text = page.get_text("text").strip()
            images = page.get_images()
            if not text and images:
                ocr_markdown = None
                if use_ocr and OCR_AVAILABLE:
                    ocr_markdown = _ocr_page(doc, pdf_path, page_num)
                if ocr_markdown:
                    all_content.append((page_offset + 1, "text",
                                        f"<!-- OCR transcribed via mac-mini-ocr -->\n\n{ocr_markdown}"))
                else:
                    # TODO:OCR - this page is a scanned image, needs OCR to extract text
                    all_content.append((page_offset + 1, "text",
                                        f"> **TODO:OCR** - Page {page_num + 1} 為掃描圖片，需要 OCR 處理"))
                continue

            # Find tables
            tables = _find_page_tables(page, force_text=page_num in table_pages)
            table_bboxes = [t.bbox for t in tables]

            for table in tables:
                y_pos = page_offset + table.bbox[1]
                raw = table.extract()
                if raw:
                    processed = _postprocess_table(raw)
                    if processed:
                        all_content.append((y_pos, "table", processed))

            # Get text blocks NOT inside table areas
            blocks = page.get_text("dict",
                                   flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]

            for block in blocks:
                if block["type"] != 0:
                    continue

                block_bbox = block["bbox"]
                if _block_in_table_area(block_bbox, table_bboxes):
                    continue

                y_pos = page_offset + block_bbox[1]

                lines = []
                for line in block["lines"]:
                    line_text = ""
                    for span in line["spans"]:
                        line_text += span["text"]
                    line_text = line_text.strip()
                    if line_text and not _is_noise_line(line_text):
                        lines.append(line_text)

                if lines:
                    text = "\n".join(lines)
                    all_content.append((y_pos, "text", text))

        # Sort by position
        all_content.sort(key=lambda x: x[0])

        # Merge tables that were split across PDF pages
        all_content = _merge_cross_page_tables(all_content)

        # --- Final Assembly ---
        md_final = [f"# MOPS Report: {pdf_path.name}",
                    f"- **Source**: {pdf_path}", ""]

        if toc_map:
            md_final.append("## Quick Navigation")
            for pg, title in sorted(toc_map.items()):
                link = title.replace(" ", "").replace("、", "")
                md_final.append(f"- [{title}](#{link})")
            md_final.append("\n---")

        for _, content_type, data in all_content:
            if content_type == "page":
                md_final.append(f"\n---\n<!-- Page {data} -->\n")
            elif content_type == "section":
                md_final.append(f"\n## {data}\n")
            elif content_type == "table":
                md_final.extend(format_grid_table(data))
            elif content_type == "text":
                # Skip any remaining cross-page continuation markers
                if CONT_NEXT_RE.search(data) or CONT_PREV_RE.search(data):
                    continue
                md_final.append(data)
                md_final.append("")

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(md_final))
        doc.close()
        logger.info(f"✅ Successfully converted with TOC: {output_path}")
        return True
    except Exception as e:
        logger.error(f"❌ Failed to convert {pdf_path}: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Convert MOPS PDF to TOC-Linked Markdown")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--pdf', '-p', type=str)
    group.add_argument('--dir', '-d', type=str)
    parser.add_argument('--no-ocr', action='store_true',
                        help="Disable mac-mini-ocr transcription for scanned pages "
                             "(fall back to TODO:OCR placeholder)")
    args = parser.parse_args()
    use_ocr = not args.no_ocr
    if args.pdf:
        convert_pdf_to_md(Path(args.pdf), use_ocr=use_ocr)
    elif args.dir:
        for f in Path(args.dir).glob("*.pdf"):
            convert_pdf_to_md(f, use_ocr=use_ocr)


if __name__ == "__main__":
    main()
