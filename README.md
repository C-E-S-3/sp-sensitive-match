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
- **Extraction:** every cell on every sheet (and hyperlink targets) is scanned
  with a SharePoint URL regex, so column layout / headers don't matter.

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

Optional `--glob` overrides the sensitive-file pattern (default `*.xlsx`).

Output: `<out-dir>/<department>.xlsx`, columns `SharePoint URL`,
`Sensitive Category`, `Source File`. URLs with no `/sites/` or `/teams/`
segment go to `_no_department.xlsx`.

## Notes

- Matching lowercases and trims URLs (SharePoint treats URLs
  case-insensitively); the output preserves the check-list's original casing.
- Temporary Excel lock files (`~$*.xlsx`) are skipped.
- If your tenant uses a vanity domain instead of `*.sharepoint.com`, widen the
  `URL_RE` regex in `match.py`.
