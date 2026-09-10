"""automate_report.py
Merge monthly sales exports into one clean, formatted Excel report.

Usage:
    python automate_report.py

Reads every CSV in data/raw/, cleans and merges them, and writes
output/report.xlsx (sheets: "details", "summary") plus a console summary.
"""

from pathlib import Path

import pandas as pd

RAW_DIR = Path("data/raw")
OUT_DIR = Path("output")
OUT_FILE = OUT_DIR / "report.xlsx"

EXPECTED_COLUMNS = {"date", "region", "product", "revenue", "units"}
DATE_FORMATS = ("%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d")


def parse_amount(value) -> float:
    """'₹45,000' / '45,000.0' / 45000 -> 45000.0"""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = "".join(ch for ch in str(value) if ch.isdigit() or ch == ".")
    return float(cleaned) if cleaned else 0.0


def parse_date(value) -> pd.Timestamp | None:
    """Try the known export formats; return NaT if none match."""
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.to_datetime(text, format=fmt)
        except ValueError:
            continue
    return pd.NaT


def load_raw() -> pd.DataFrame:
    """Read and concatenate every raw CSV."""
    files = sorted(RAW_DIR.glob("*.csv"))
    if not files:
        raise SystemExit(f"No CSV files found in {RAW_DIR}")
    frames = []
    for f in files:
        frame = pd.read_csv(f)
        # Standardize column names per file BEFORE concat, so exports that
        # differ only in header spacing (" Region " vs "Region") merge into
        # one column instead of becoming two half-empty ones.
        frame.columns = frame.columns.str.strip().str.lower().str.replace(" ", "_")
        frames.append(frame)
    return pd.concat(frames, ignore_index=True)


def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Standardize schema, parse types, dedupe. Returns (df, report)."""
    report = {"rows_read": len(df)}

    # 1. Standardize column names: "Revenue " -> "revenue"
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    missing = EXPECTED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing expected columns: {sorted(missing)}")

    # 2. Parse messy types
    df["revenue"] = df["revenue"].map(parse_amount)
    df["date"] = df["date"].map(parse_date)
    df["units"] = pd.to_numeric(df["units"], errors="coerce").fillna(0).astype(int)

    # 3. Fill + flag gaps in optional fields
    df["region"] = df["region"].fillna("Unassigned")
    df["product"] = df["product"].fillna("Unknown")

    # 4. Drop exact duplicates (exports often overlap)
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    report["duplicates_removed"] = before - len(df)

    # 5. Sort for readability
    df = df.sort_values(["date", "region"]).reset_index(drop=True)
    report["rows_clean"] = len(df)
    return df, report


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Revenue by region — the number leadership asks for first."""
    return (
        df.groupby("region", as_index=False)
        .agg(total_revenue=("revenue", "sum"), total_units=("units", "sum"), rows=("revenue", "size"))
        .sort_values("total_revenue", ascending=False)
    )


def write_report(df: pd.DataFrame, summary: pd.DataFrame) -> None:
    OUT_DIR.mkdir(exist_ok=True)
    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="details", index=False)
        summary.to_excel(writer, sheet_name="summary", index=False)
    print(f"Wrote {OUT_FILE}")


def main() -> None:
    print("Report automation")
    print("=================")
    df, report = clean(load_raw())
    summary = build_summary(df)

    print(f"Rows read       : {report['rows_read']}")
    print(f"Duplicates removed: {report['duplicates_removed']}")
    print(f"Rows after clean: {report['rows_clean']}")
    print(f"Total revenue   : ₹ {df['revenue'].sum():,.0f}")
    print(f"Regions         : {df['region'].nunique()}")
    write_report(df, summary)


if __name__ == "__main__":
    main()
