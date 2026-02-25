import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
import re, sys, statistics
from pathlib import Path

DOWNLOADS_DIR = Path(__file__).resolve().parent.parent / "downloads"
PAGE_BREAK_RE = re.compile(r"<!--\s*Page\s+(\d+)\s*-->", re.IGNORECASE)
CONTEXT_LINES = 4


def is_filler(line):
    s = line.strip()
    if not s:
        return True
    if re.match(r"^-+$", s):
        return True
    if re.match(r"^-+\s*\d+\s*-+$", s):
        return True
    if len(s) <= 12 and s[0] in "(\uff08":
        return True
    return False


def count_pipes(line):
    return line.count("|")


def is_table_row(line):
    return line.strip().startswith("|")


def is_separator_row(line):
    stripped = line.strip()
    if not stripped.startswith("|"):
        return False
    inner = stripped.strip("|").replace(" ", "")
    return bool(inner) and all(c in "-:|" for c in inner)


def find_table_before(lines, pb_idx, window=15):
    for j in range(pb_idx - 1, max(-1, pb_idx - window - 1), -1):
        if is_table_row(lines[j]):
            return j
        if not is_filler(lines[j]):
            return None
    return None


def find_table_after(lines, pb_idx, window=15):
    for j in range(pb_idx + 1, min(len(lines), pb_idx + window + 1)):
        if is_table_row(lines[j]):
            return j
        if not is_filler(lines[j]):
            return None
    return None


def analyze_cols(lines, start, end):
    rows = [(i, count_pipes(lines[i])) for i in range(start, end)
            if is_table_row(lines[i]) and not is_separator_row(lines[i])]
    if len(rows) < 3:
        return []
    counts = [c for _, c in rows]
    med = statistics.median(counts)
    if med == 0:
        return []
    return [{"lineno": ln, "pipes": p, "median": med}
            for ln, p in rows if p / med < 0.6 or p / med > 1.5]


def find_blocks(lines):
    blocks, in_block, bs = [], False, 0
    for i, line in enumerate(lines):
        is_pb = bool(PAGE_BREAK_RE.match(line.strip()))
        if is_table_row(line) and not is_pb:
            if not in_block:
                in_block, bs = True, i
        else:
            if in_block:
                blocks.append((bs, i))
                in_block = False
    if in_block:
        blocks.append((bs, len(lines)))
    return blocks

def scan_file(md_path):
    try:
        lines = md_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception as e:
        return {"cross_page": [], "inconsistent_cols": [], "error": str(e)}
    cp, ic = [], []
    for i, line in enumerate(lines):
        m = PAGE_BREAK_RE.match(line.strip())
        if not m:
            continue
        pg = int(m.group(1))
        lb = find_table_before(lines, i)
        fa = find_table_after(lines, i)
        if lb is not None and fa is not None:
            ctx_start = max(0, lb - CONTEXT_LINES + 1)
            cp.append({
                "page_num": pg, "page_break_line": i,
                "last_before": lb, "first_after": fa,
                "ctx_start": ctx_start,
                "pre":  lines[ctx_start: i + 1],
                "post": lines[i + 1: min(len(lines), fa + CONTEXT_LINES)],
            })
    for bs, be in find_blocks(lines):
        for anom in analyze_cols(lines, bs, be):
            ln = anom["lineno"]
            ic.append({
                "lineno": ln, "pipes": anom["pipes"], "median": anom["median"],
                "content": lines[ln],
                "pre":  lines[max(0, ln - CONTEXT_LINES): ln],
                "post": lines[ln + 1: min(len(lines), ln + CONTEXT_LINES + 1)],
            })
    return {"cross_page": cp, "inconsistent_cols": ic}

def report(all_findings):
    tcp = sum(len(v["cross_page"]) for v in all_findings.values())
    tic = sum(len(v["inconsistent_cols"]) for v in all_findings.values())
    bad = sum(1 for v in all_findings.values()
              if v["cross_page"] or v["inconsistent_cols"])
    sep = "=" * 80
    print(sep)
    print("MOPS Markdown  --  Incomplete / Cross-Page Table Scanner")
    print(sep)
    print("Scanned dir    : " + str(DOWNLOADS_DIR))
    print("Files w/issues : " + str(bad))
    print("Cross-pg breaks: " + str(tcp))
    print("Inconsist cols : " + str(tic))
    print(sep)
    for path_str, res in sorted(all_findings.items()):
        cp = res["cross_page"]
        ic = res["inconsistent_cols"]
        if not cp and not ic:
            continue
        rel = Path(path_str).relative_to(DOWNLOADS_DIR)
        print("\n" + "-" * 80)
        print("FILE: " + str(rel))
        print("-" * 80)
        if cp:
            print("\n  [CROSS-PAGE TABLES]  (" + str(len(cp)) + " found)")
            for f in cp:
                pg, pb = f["page_num"], f["page_break_line"]
                lb, fa = f["last_before"], f["first_after"]
                pre, post, cs = f["pre"], f["post"], f["ctx_start"]
                print("\n    >> Table split across Page " + str(pg) +
                      "  (page-break at line " + str(pb + 1) + ")")
                print("       Last table row before break : line " + str(lb + 1))
                print("       First table row after break : line " + str(fa + 1))
                print("    -- context before break --")
                for off, c in enumerate(pre):
                    ln = cs + off
                    tag = ""
                    if ln == lb:
                        tag = "  <-- last table row"
                    if ln == pb:
                        tag = "  <-- PAGE BREAK"
                    print("    " + str(ln + 1).rjust(5) + ": " + c + tag)
                print("    -- context after break --")
                for off, c in enumerate(post):
                    ln = pb + 1 + off
                    tag = "  <-- first table row" if ln == fa else ""
                    print("    " + str(ln + 1).rjust(5) + ": " + c + tag)
        if ic:
            print("\n  [INCONSISTENT COLUMN COUNTS]  (" + str(len(ic)) + " found)")
            for f in ic:
                ln, pre, post = f["lineno"], f["pre"], f["post"]
                print("\n    >> Line " + str(ln + 1) + ": " + str(f["pipes"]) +
                      " pipes  (block median: " + str(int(f["median"])) + ")")
                print("    -- context --")
                for off, c in enumerate(pre):
                    print("    " + str(ln - len(pre) + off + 1).rjust(5) + ": " + c)
                print("    " + str(ln + 1).rjust(5) + ": " +
                      f["content"].rstrip() + "  <-- ANOMALY")
                for off, c in enumerate(post):
                    print("    " + str(ln + off + 2).rjust(5) + ": " + c)
    print("\n" + "=" * 80)
    print("Scan complete.")
    print("  Cross-page table breaks : " + str(tcp))
    print("  Inconsistent col rows   : " + str(tic))
    print("=" * 80)


def main():
    if not DOWNLOADS_DIR.exists():
        print("ERROR: " + str(DOWNLOADS_DIR), file=sys.stderr)
        sys.exit(1)
    md_files = sorted(DOWNLOADS_DIR.rglob("*.md"))
    print("Found " + str(len(md_files)) + " .md files. Scanning...", flush=True)
    all_findings = {str(p): scan_file(p) for p in md_files}
    report(all_findings)


if __name__ == "__main__":
    main()
