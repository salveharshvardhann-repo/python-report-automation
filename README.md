# Python Report Automation

A small but real automation pipeline: merge monthly sales exports, clean them,
and produce a formatted Excel report with **one command** instead of an afternoon
of copy-paste.

**Author:** Harshvardhan Salve · [LinkedIn](https://linkedin.com/in/harshvardhan-salve-7601s)

---

## The problem

A recurring report was assembled manually from multiple spreadsheets every month:
copy each export, fix inconsistent columns, remove duplicates, re-format amounts,
rebuild the summary. Slow, and every manual step was a chance to ship a wrong number.

## The approach

Script the whole chain so it's **repeatable and auditable**:

```
data/raw/*.csv  →  validate  →  clean  →  dedupe  →  output/report.xlsx (+ console summary)
```

Every transformation is a named function with a comment — easy to review, easy to trust.

## Repository structure

```
python-report-automation/
├── automate_report.py    # the pipeline
├── requirements.txt
├── data/
│   └── raw/
│       ├── sales_2026_01.csv   # sample exports (same shape as the real ones)
│       └── sales_2026_02.csv
└── README.md
```

## How to run

```bash
pip install -r requirements.txt
python automate_report.py
```

Output: `output/report.xlsx` with two sheets — `details` (clean, merged rows) and
`summary` (revenue by region) — plus a summary printed to the console.

## What the script handles

| Messy input | What the script does |
|---|---|
| Column names in mixed case with stray spaces | Standardized to `snake_case` |
| Amounts like `₹45,000` or `45,000.0` as text | Parsed to numeric |
| Dates in `DD-MM-YYYY` / `YYYY/MM/DD` | Parsed to proper dates, reported as ISO |
| Duplicate rows across monthly exports | Dropped, count reported |
| Missing values in optional fields | Filled with sensible defaults, flagged |

## Sample console output

```
Report automation
=================
Rows read       : 13
Duplicates removed: 1
Rows after clean: 12
Total revenue   : ₹ 558,000
Regions         : 4
Wrote output/report.xlsx
```

## Impact

This pattern replaced a manual monthly assembly process — the kind of automation that
saves **10+ hours a month** once it replaces a real recurring report.

## Skills demonstrated

- pandas data cleaning (schema validation, type parsing, deduplication)
- Defensive scripting: every fix is counted and reported, never silent
- Producing stakeholder-ready Excel output (formatted, multi-sheet)

## Note

The CSVs here are representative samples with the same *shape* as real exports I
automated; values are illustrative.
