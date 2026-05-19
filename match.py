#!/usr/bin/env python3
"""
Compare a check-list of SharePoint URLs against a directory of "sensitive data"
xlsx files. For every URL in the check-list that also appears in a sensitive
file, group the match by department (the path segment after /sites/) and write
one output xlsx per department.

Match basis : exact full URL (whitespace-trimmed, case-insensitive).
Grouping    : department only (category is recorded as a column).

Column names (the columns that hold the SharePoint URL) default to the
constants below and can be overridden per run with --checklist-column /
--sensitive-column. Header matching is case-insensitive and whitespace-trimmed.

Usage:
    python match.py \
        --sensitive-dir ./sensitive \
        --checklist     ./checklist.xlsx \
        --out-dir       ./output
"""

import argparse
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

# csv default field-size limit is too small for long SharePoint URLs in
# pathological rows; raise it.
csv.field_size_limit(10 * 1024 * 1024)

# ---------------------------------------------------------------------------
# Column names that hold the SharePoint URL. Change here, or override per run
# with --checklist-column / --sensitive-column.
# ---------------------------------------------------------------------------
CHECKLIST_COLUMN = "ObjectId"   # the file listing the URLs to confirm
SENSITIVE_COLUMN = "FileUrl"    # the scanned secrets/addresses/... files
# ---------------------------------------------------------------------------

# Pulls the URL out of a cell that may also contain surrounding text.
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


def _extract_url(raw):
    """Normalize a raw cell/field value to a URL string (or None).
    If the text wraps a URL in other text, extract just the URL."""
    if raw is None:
        return None
    text = str(raw).strip()
    if not text:
        return None
    m = URL_RE.search(text)
    return m.group(0) if m else text


def _cell_url(cell):
    """openpyxl cell -> URL: value, falling back to a hyperlink target."""
    raw = cell.value
    if raw is None or str(raw).strip() == "":
        raw = getattr(cell.hyperlink, "target", None)
    return _extract_url(raw)


def _iter_xlsx(path: Path, want: str, seen_headers: set):
    """Yield URLs from the named column across all worksheets of an xlsx."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            rows = ws.iter_rows()
            try:
                header = next(rows)
            except StopIteration:
                continue  # empty sheet
            col_idx = None
            for i, hcell in enumerate(header):
                hv = hcell.value
                if hv is None:
                    continue
                seen_headers.add(str(hv).strip())
                if str(hv).strip().lower() == want:
                    col_idx = i
                    break
            if col_idx is None:
                continue
            for row in rows:
                if col_idx < len(row):
                    url = _cell_url(row[col_idx])
                    if url:
                        yield url
    finally:
        wb.close()


def _iter_csv(path: Path, want: str, seen_headers: set):
    """Yield URLs from the named column of a CSV. Delimiter is sniffed
    (comma/semicolon/tab/pipe); a BOM is tolerated."""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        sample = fh.read(8192)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel  # default to comma
        reader = csv.reader(fh, dialect)
        try:
            header = next(reader)
        except StopIteration:
            return  # empty file
        col_idx = None
        for i, h in enumerate(header):
            if h is None:
                continue
            seen_headers.add(str(h).strip())
            if str(h).strip().lower() == want:
                col_idx = i
                break
        if col_idx is None:
            return
        for row in reader:
            if col_idx < len(row):
                url = _extract_url(row[col_idx])
                if url:
                    yield url


def iter_urls(path: Path, column: str):
    """Yield every URL found in the named column of an xlsx or csv file.

    Header matching is case-insensitive and whitespace-trimmed. Raises
    ValueError if the column is not present (xlsx: in any worksheet).
    """
    want = column.strip().lower()
    seen_headers: set = set()
    suffix = path.suffix.lower()
    if suffix == ".csv":
        reader = _iter_csv(path, want, seen_headers)
    elif suffix in (".xlsx", ".xlsm"):
        reader = _iter_xlsx(path, want, seen_headers)
    else:
        raise ValueError(
            f"{path.name}: unsupported extension '{path.suffix}' "
            f"(expected .csv, .xlsx, or .xlsm)"
        )

    yielded = False
    for url in reader:
        yielded = True
        yield url

    if not yielded and want not in {h.strip().lower() for h in seen_headers}:
        raise ValueError(
            f"{path.name}: no column named '{column}'. "
            f"Headers seen: {sorted(seen_headers) or '(none)'}"
        )


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--sensitive-dir", required=True, type=Path,
                   help="Directory of sensitive xlsx files (secrets.xlsx, ...)")
    p.add_argument("--checklist", required=True, type=Path,
                   help="Single xlsx with the URLs to confirm")
    p.add_argument("--out-dir", required=True, type=Path,
                   help="Where per-department xlsx files are written")
    p.add_argument("--checklist-column", default=CHECKLIST_COLUMN,
                   help=f"URL column in the check-list (default: {CHECKLIST_COLUMN})")
    p.add_argument("--sensitive-column", default=SENSITIVE_COLUMN,
                   help=f"URL column in the sensitive files (default: {SENSITIVE_COLUMN})")
    p.add_argument("--glob", default="*.xlsx",
                   help="Filename pattern for sensitive files (default: *.xlsx)")
    args = p.parse_args()

    if not args.sensitive_dir.is_dir():
        p.error(f"--sensitive-dir not a directory: {args.sensitive_dir}")
    if not args.checklist.is_file():
        p.error(f"--checklist not found: {args.checklist}")

    # 1. Build the check-list lookup: normalized URL -> original URL (first seen).
    try:
        checklist = {}
        for url in iter_urls(args.checklist, args.checklist_column):
            checklist.setdefault(normalize(url), url)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1
    print(f"check-list: {len(checklist)} unique URLs from "
          f"{args.checklist.name} (column '{args.checklist_column}')")
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
    print(f"scanning {len(sensitive_files)} sensitive file(s) "
          f"(column '{args.sensitive_column}'):")
    for sf in sensitive_files:
        category = sf.stem               # secrets, addresses, creditcard, ...
        matched_here = 0
        try:
            for url in iter_urls(sf, args.sensitive_column):
                key = normalize(url)
                if key in checklist:
                    display = checklist[key]  # use the check-list's form
                    by_dept[department_of(display)].add(
                        (display, category, sf.name))
                    matched_here += 1
            note = f"matches={matched_here}"
        except ValueError as e:
            note = f"SKIPPED ({e})"
        print(f"  {sf.name:<30} category={category:<15} {note}")

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
