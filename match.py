#!/usr/bin/env python3
"""
Compare a check-list of SharePoint URLs against a directory of "sensitive data"
xlsx files. For every URL in the check-list that also appears in a sensitive
file, group the match by department (the path segment after /sites/) and write
one output xlsx per department.

Match basis : exact full URL (whitespace-trimmed, case-insensitive).
Grouping    : department only (category is recorded as a column).

Usage:
    python match.py \
        --sensitive-dir ./sensitive \
        --checklist     ./checklist.xlsx \
        --out-dir       ./output
"""

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

# Matches https://<tenant>.sharepoint.com/... up to the first whitespace.
URL_RE = re.compile(r"https?://[^\s\"'<>]*sharepoint\.com[^\s\"'<>]*", re.IGNORECASE)
# Department = first path segment after /sites/ (or /teams/ for MS Teams sites).
DEPT_RE = re.compile(r"/(?:sites|teams)/([^/?#]+)", re.IGNORECASE)


def normalize(url: str) -> str:
    """Normalization used only for matching, never for display/output."""
    return url.strip().rstrip("/").lower()


def department_of(url: str) -> str:
    m = DEPT_RE.search(url)
    return m.group(1) if m else "_no_department"


def safe_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]", "_", name) or "_unnamed"


def iter_urls(xlsx_path: Path):
    """Yield every SharePoint URL found in any cell value or hyperlink target."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=False)
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    for source in (cell.value, getattr(cell.hyperlink, "target", None)):
                        if not source:
                            continue
                        for found in URL_RE.findall(str(source)):
                            yield found
    finally:
        wb.close()


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--sensitive-dir", required=True, type=Path,
                   help="Directory of sensitive xlsx files (secrets.xlsx, ...)")
    p.add_argument("--checklist", required=True, type=Path,
                   help="Single xlsx with the URLs to confirm")
    p.add_argument("--out-dir", required=True, type=Path,
                   help="Where per-department xlsx files are written")
    p.add_argument("--glob", default="*.xlsx",
                   help="Filename pattern for sensitive files (default: *.xlsx)")
    args = p.parse_args()

    if not args.sensitive_dir.is_dir():
        p.error(f"--sensitive-dir not a directory: {args.sensitive_dir}")
    if not args.checklist.is_file():
        p.error(f"--checklist not found: {args.checklist}")

    # 1. Build the check-list lookup: normalized URL -> original URL (first seen).
    checklist = {}
    for url in iter_urls(args.checklist):
        checklist.setdefault(normalize(url), url)
    print(f"check-list: {len(checklist)} unique SharePoint URLs "
          f"from {args.checklist.name}")
    if not checklist:
        print("Nothing to match against; exiting.", file=sys.stderr)
        return 1

    sensitive_files = sorted(f for f in args.sensitive_dir.glob(args.glob)
                             if f.is_file() and not f.name.startswith("~$"))
    if not sensitive_files:
        print(f"No files matching {args.glob} in {args.sensitive_dir}",
              file=sys.stderr)
        return 1

    # 2. Scan each sensitive file; collect matches grouped by department.
    #    dept -> set of (display_url, category, source_file)  [set = dedup]
    by_dept = defaultdict(set)
    for sf in sensitive_files:
        category = sf.stem               # secrets, addresses, creditcard, ...
        matched_here = 0
        for url in iter_urls(sf):
            key = normalize(url)
            if key in checklist:
                display = checklist[key]  # use the check-list's form of the URL
                by_dept[department_of(display)].add((display, category, sf.name))
                matched_here += 1
        print(f"  {sf.name:<30} category={category:<15} matches={matched_here}")

    if not by_dept:
        print("\nNo matches found. No output files written.")
        return 0

    # 3. Write one workbook per department.
    args.out_dir.mkdir(parents=True, exist_ok=True)
    print()
    grand_total = 0
    for dept in sorted(by_dept):
        rows = sorted(by_dept[dept], key=lambda r: (r[1], r[0]))
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = safe_filename(dept)[:31] or "Sheet1"
        ws.append(["SharePoint URL", "Sensitive Category", "Source File"])
        for url, category, src in rows:
            ws.append([url, category, src])
        ws.freeze_panes = "A2"
        out_path = args.out_dir / f"{safe_filename(dept)}.xlsx"
        wb.save(out_path)
        grand_total += len(rows)
        print(f"  {out_path.name:<30} {len(rows)} matched URL(s)")

    print(f"\nDone: {grand_total} match row(s) across "
          f"{len(by_dept)} department file(s) in {args.out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
