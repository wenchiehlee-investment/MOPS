#!/usr/bin/env python3
"""
Identify companies that have likely already filed their quarterly financial
report with MOPS ahead of the general filing deadline (e.g. TSMC files
around its own earnings call, weeks before the regulatory deadline that
Download.yaml's seasonal cron windows are built around).

Reads the earnings calendar synced from the InvestorConference repo
(data/InvestorConference/raw_event_upcoming_earnings.csv) and cross-checks
"財報公告" / "台股" rows whose announcement date has already passed against
what's already been downloaded, to build a short candidate list for an
early, targeted download run.
"""

import argparse
import csv
import platform
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List, NamedTuple, Optional

if platform.system() == 'Windows':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

REPO_ROOT = Path(__file__).resolve().parent.parent
DOWNLOADS = REPO_ROOT / "downloads"
CALENDAR_CSV = REPO_ROOT / "data" / "InvestorConference" / "raw_event_upcoming_earnings.csv"
CANDIDATES_CSV = REPO_ROOT / "data" / "reports" / "early_filer_candidates.csv"

EVENT_RE = re.compile(r'\((\d{4})\)\s*(\d{4})\s*Q(\d)\s*財報')


class Candidate(NamedTuple):
    company_id: str
    year: int
    quarter: int
    announce_date: date
    event_name: str


def _parse_row(row: dict) -> Optional[Candidate]:
    if row.get('類別') != '財報公告' or row.get('子類別') != '台股':
        return None

    match = EVENT_RE.search(row.get('事件名稱', ''))
    if not match:
        return None
    company_id, year_str, quarter_str = match.groups()

    date_str = row.get('開始日期', '').strip()
    try:
        announce_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    except ValueError:
        return None

    return Candidate(company_id, int(year_str), int(quarter_str),
                     announce_date, row.get('事件名稱', ''))


def _already_downloaded(company_id: str, year: int, quarter: int) -> bool:
    prefix = f"{year}{quarter:02d}_{company_id}_"
    company_dir = DOWNLOADS / company_id
    if not company_dir.exists():
        return False
    return any(company_dir.glob(f"{prefix}*.pdf"))


def find_candidates(lookback_days: int = 14, today: Optional[date] = None) -> List[Candidate]:
    if today is None:
        today = date.today()
    window_start = today - timedelta(days=lookback_days)

    if not CALENDAR_CSV.exists():
        print(f"⚠️ Earnings calendar not found at {CALENDAR_CSV} "
              f"(has the InvestorConference sync run yet?)", file=sys.stderr)
        return []

    candidates = []
    with open(CALENDAR_CSV, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            candidate = _parse_row(row)
            if not candidate:
                continue
            if not (window_start <= candidate.announce_date <= today):
                continue
            if _already_downloaded(candidate.company_id, candidate.year, candidate.quarter):
                continue
            candidates.append(candidate)

    # De-duplicate (company/year/quarter can appear more than once in the calendar)
    seen = set()
    unique = []
    for c in candidates:
        key = (c.company_id, c.year, c.quarter)
        if key in seen:
            continue
        seen.add(key)
        unique.append(c)

    return sorted(unique, key=lambda c: c.announce_date)


def main():
    parser = argparse.ArgumentParser(
        description="Find companies whose quarterly report announcement date "
                    "has passed but hasn't been downloaded yet")
    parser.add_argument('--lookback-days', type=int, default=14,
                        help="How many days back to consider an announcement "
                             "date 'recent enough' to check (default: 14)")
    args = parser.parse_args()

    candidates = find_candidates(lookback_days=args.lookback_days)

    CANDIDATES_CSV.parent.mkdir(parents=True, exist_ok=True)
    with open(CANDIDATES_CSV, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['company_id', 'year', 'quarter', 'announce_date', 'event_name'])
        for c in candidates:
            writer.writerow([c.company_id, c.year, c.quarter, c.announce_date.isoformat(), c.event_name])

    print(f"Found {len(candidates)} early-filer candidate(s):")
    for c in candidates:
        print(f"  {c.company_id}  {c.year} Q{c.quarter}  announced {c.announce_date}  ({c.event_name})")
    print(f"Written to {CANDIDATES_CSV}")


if __name__ == "__main__":
    main()
