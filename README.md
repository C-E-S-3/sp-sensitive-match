# sp-sensitive-match

Compare a check-list of SharePoint URLs against a directory of "sensitive data"
spreadsheets and produce, per department, the URLs that match.

Each xlsx in the sensitive directory (e.g. `secrets.xlsx`, `addresses.xlsx`,
`creditcard.xlsx`) lists SharePoint file URLs flagged for that kind of sensitive
content. The check-list is a single xlsx of SharePoint URLs you want to confirm.
For every check-list URL that also appears in a sensitive file, the match is
grouped by **department** — the path segment after `/sites/` (or `/teams/`) — and
written to one output xlsx per department.

- **Match basis:** exact full URL, whitespace-trimmed and case-insensitive.
- **Grouping:** department only; the sensitive category and source file are
  recorded as columns, so a URL flagged by multiple sensitive files yields one
  row per category.
- **Input formats:** `.xlsx`/`.xlsm` and `.csv` are both supported, for the
  check-list and the sensitive files independently (dispatched by extension).
  CSV delimiter is auto-detected (comma/semicolon/tab/pipe) and a BOM is
  tolerated. To scan CSV sensitive files pass `--glob "*.csv"`. Output
  per-department files are always `.xlsx`.
- **Extraction:** the URL is read from a named column. The check-list column
  defaults to `ObjectId`; the sensitive-file column defaults to `FileUrl`.
  Change the `CHECKLIST_COLUMN` / `SENSITIVE_COLUMN` constants at the top of
  `match.py`, or override per run with `--checklist-column` /
  `--sensitive-column`. Header matching is case-insensitive and
  whitespace-trimmed; all worksheets in a workbook are searched for the column.

## Setup

```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
```

## Usage

```bash
./.venv/bin/python match.py \
    --sensitive-dir /path/to/sensitive-xlsx-dir \
    --checklist     /path/to/checklist.xlsx \
    --out-dir       /path/to/output
```

Override the URL columns if your files differ from the defaults:

```bash
./.venv/bin/python match.py ... \
    --checklist-column ObjectId \
    --sensitive-column FileUrl
```

Optional `--glob` overrides the sensitive-file pattern (default `*.xlsx`).
A sensitive file missing the column is skipped (with the headers it has
listed); the check-list missing its column is a hard error.

Output: `<out-dir>/<department>.xlsx`, columns `SharePoint URL`,
`Sensitive Category`, `Source File`. URLs with no `/sites/` or `/teams/`
segment go to `_no_department.xlsx`.

## Notes

- Matching lowercases and trims URLs (SharePoint treats URLs
  case-insensitively); the output preserves the check-list's original casing.
- Temporary Excel lock files (`~$*.xlsx`) are skipped.
- If your tenant uses a vanity domain instead of `*.sharepoint.com`, widen the
  `URL_RE` regex in `match.py`.
