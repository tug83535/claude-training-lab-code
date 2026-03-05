"""
KBT Universal Tools — Date Format Unifier
Works on ANY CSV/Excel file — no project-specific setup required.

Reads a file, detects date columns, and converts all dates to a single
user-specified format (default: YYYY-MM-DD). Handles common date formats:
  - MM/DD/YYYY, DD/MM/YYYY, YYYY-MM-DD, MM-DD-YYYY
  - Jan 15, 2026 / 15-Jan-2026 / January 15 2026
  - Excel serial dates (e.g., 46054)

Usage:
    python date_format_unifier.py input.csv --output output.csv --format "%Y-%m-%d"
    python date_format_unifier.py input.xlsx --sheet "Sheet1" --format "%m/%d/%Y"
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd


# Common date formats to try when parsing
DATE_FORMATS = [
    "%Y-%m-%d",       # 2026-03-05
    "%m/%d/%Y",       # 03/05/2026
    "%d/%m/%Y",       # 05/03/2026
    "%m-%d-%Y",       # 03-05-2026
    "%d-%m-%Y",       # 05-03-2026
    "%Y/%m/%d",       # 2026/03/05
    "%b %d, %Y",      # Mar 05, 2026
    "%d-%b-%Y",       # 05-Mar-2026
    "%d %b %Y",       # 05 Mar 2026
    "%B %d, %Y",      # March 05, 2026
    "%B %d %Y",       # March 05 2026
    "%Y%m%d",         # 20260305
    "%m/%d/%y",       # 03/05/26
    "%d/%m/%y",       # 05/03/26
]

# Excel epoch for serial date conversion
EXCEL_EPOCH = datetime(1899, 12, 30)


def is_excel_serial(value) -> bool:
    """Check if a value looks like an Excel serial date number."""
    try:
        num = float(value)
        return 1 < num < 200000  # Reasonable date range
    except (ValueError, TypeError):
        return False


def parse_date(value, day_first: bool = False) -> datetime | None:
    """Try to parse a value as a date using multiple format strategies."""
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None

    # Already a datetime
    if isinstance(value, datetime):
        return value

    # Pandas Timestamp
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    # Excel serial date
    if is_excel_serial(value):
        try:
            return EXCEL_EPOCH + timedelta(days=float(value))
        except (ValueError, OverflowError):
            pass

    # String parsing
    text = str(value).strip()
    if len(text) == 0:
        return None

    # Try pandas first (it's smart about formats)
    try:
        result = pd.to_datetime(text, dayfirst=day_first)
        return result.to_pydatetime()
    except (ValueError, TypeError):
        pass

    # Try explicit formats
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    return None


def detect_date_columns(df: pd.DataFrame, sample_size: int = 20,
                        threshold: float = 0.6) -> list[str]:
    """Detect which columns likely contain dates by sampling values."""
    date_cols = []
    for col in df.columns:
        non_null = df[col].dropna()
        if len(non_null) == 0:
            continue

        sample = non_null.head(sample_size)
        date_count = sum(1 for v in sample if parse_date(v) is not None)
        ratio = date_count / len(sample)

        if ratio >= threshold:
            date_cols.append(col)

    return date_cols


def unify_dates(df: pd.DataFrame, columns: list[str],
                target_format: str, day_first: bool = False) -> tuple[pd.DataFrame, dict]:
    """Convert date columns to the target format. Returns (df, stats)."""
    stats = {}
    for col in columns:
        converted = 0
        failed = 0
        for idx in df.index:
            val = df.at[idx, col]
            if val is None or (isinstance(val, str) and val.strip() == ""):
                continue

            parsed = parse_date(val, day_first=day_first)
            if parsed is not None:
                df.at[idx, col] = parsed.strftime(target_format)
                converted += 1
            else:
                failed += 1

        stats[col] = {"converted": converted, "failed": failed}

    return df, stats


def main():
    parser = argparse.ArgumentParser(
        description="Unify date formats across a CSV or Excel file."
    )
    parser.add_argument("input", help="Input file path (CSV or Excel)")
    parser.add_argument("--output", "-o", help="Output file path (defaults to input_unified.ext)")
    parser.add_argument("--format", "-f", default="%Y-%m-%d",
                        help="Target date format (default: %%Y-%%m-%%d)")
    parser.add_argument("--sheet", "-s", default=0,
                        help="Sheet name or index for Excel files (default: first sheet)")
    parser.add_argument("--columns", "-c", nargs="*",
                        help="Specific columns to convert (auto-detects if not specified)")
    parser.add_argument("--day-first", action="store_true",
                        help="Parse ambiguous dates as DD/MM (European format)")
    parser.add_argument("--preview", action="store_true",
                        help="Preview changes without writing output")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    # Read input
    print(f"Reading: {input_path}")
    if input_path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        df = pd.read_excel(input_path, sheet_name=args.sheet, dtype=str)
    elif input_path.suffix.lower() == ".csv":
        df = pd.read_csv(input_path, dtype=str)
    else:
        print(f"ERROR: Unsupported file type: {input_path.suffix}")
        sys.exit(1)

    print(f"  Loaded {len(df)} rows, {len(df.columns)} columns")

    # Detect or use specified columns
    if args.columns:
        date_cols = [c for c in args.columns if c in df.columns]
        missing = [c for c in args.columns if c not in df.columns]
        if missing:
            print(f"  WARNING: Columns not found: {', '.join(missing)}")
    else:
        date_cols = detect_date_columns(df, day_first=args.day_first)
        print(f"  Auto-detected {len(date_cols)} date column(s): {', '.join(date_cols)}")

    if not date_cols:
        print("  No date columns found. Nothing to convert.")
        sys.exit(0)

    # Convert
    df, stats = unify_dates(df, date_cols, args.format, day_first=args.day_first)

    # Report
    print(f"\nDate Unification Report (target: {args.format}):")
    print("-" * 50)
    total_converted = 0
    total_failed = 0
    for col, s in stats.items():
        status = "OK" if s["failed"] == 0 else "WARNINGS"
        print(f"  {col}: {s['converted']} converted, {s['failed']} failed [{status}]")
        total_converted += s["converted"]
        total_failed += s["failed"]
    print("-" * 50)
    print(f"  Total: {total_converted} converted, {total_failed} failed")

    if args.preview:
        print("\n[PREVIEW MODE — no file written]")
        print(df.head(10).to_string())
        return

    # Write output
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_stem(input_path.stem + "_unified")

    if output_path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        df.to_excel(output_path, index=False)
    else:
        df.to_csv(output_path, index=False)

    print(f"\nOutput written to: {output_path}")


if __name__ == "__main__":
    main()
