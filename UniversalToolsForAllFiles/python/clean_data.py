"""
KBT Universal Tools — Universal Data Cleaner
============================================================
Cleans any Excel file in one command:
  - Removes completely empty rows and columns
  - Trims leading/trailing spaces from all text cells
  - Converts text-stored numbers to numeric
  - Standardizes date formats to YYYY-MM-DD
  - Removes duplicate rows
  - Reports a before/after summary

Usage:
    python clean_data.py "C:\\path\\to\\your_file.xlsx"
    python clean_data.py "C:\\path\\to\\your_file.xlsx" --sheet "Sheet1"
    python clean_data.py "C:\\path\\to\\your_file.xlsx" --no-dedupe

Output: Saves a cleaned copy as "your_file_CLEANED.xlsx" in the same folder
"""

import sys
import os
import argparse
from datetime import datetime

try:
    import pandas as pd
    import openpyxl
except ImportError:
    print("ERROR: Required packages not installed.")
    print("Run: pip install pandas openpyxl")
    sys.exit(1)


def clean_excel_file(file_path: str, sheet_name: str = None, dedupe: bool = True) -> None:
    print(f"\n{'='*55}")
    print("  KBT Universal Data Cleaner")
    print(f"{'='*55}")
    print(f"  File:  {os.path.basename(file_path)}")
    print(f"  Start: {datetime.now().strftime('%m/%d/%Y %I:%M %p')}")
    print(f"{'='*55}\n")

    if not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)

    try:
        xl = pd.ExcelFile(file_path)
        sheets_to_process = [sheet_name] if sheet_name else xl.sheet_names
        print(f"Sheets found: {', '.join(xl.sheet_names)}")
        print(f"Processing:   {', '.join(sheets_to_process)}\n")
    except Exception as e:
        print(f"ERROR reading file: {e}")
        sys.exit(1)

    results = {}

    for sht in sheets_to_process:
        print(f"--- Sheet: {sht} ---")
        try:
            df = pd.read_excel(file_path, sheet_name=sht, header=0)
        except Exception as e:
            print(f"  Could not read sheet '{sht}': {e}")
            continue

        original_rows = len(df)
        original_cols = len(df.columns)
        print(f"  Original size: {original_rows} rows x {original_cols} columns")

        # 1. Drop completely empty rows and columns
        df.dropna(how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        empty_rows_removed = original_rows - len(df)
        print(f"  Empty rows removed:       {empty_rows_removed}")

        # 2. Strip leading/trailing spaces from string columns
        space_fixes = 0
        for col in df.select_dtypes(include='object').columns:
            before = df[col].copy()
            df[col] = df[col].str.strip()
            space_fixes += (before != df[col]).sum()
        print(f"  Cells trimmed:            {space_fixes}")

        # 3. Convert text-stored numbers
        num_fixes = 0
        for col in df.select_dtypes(include='object').columns:
            converted = pd.to_numeric(df[col], errors='coerce')
            mask = converted.notna() & df[col].notna()
            df.loc[mask, col] = converted[mask]
            num_fixes += mask.sum()
        print(f"  Text-to-number fixes:     {num_fixes}")

        # 4. Standardize date columns
        date_fixes = 0
        for col in df.select_dtypes(include='object').columns:
            try:
                converted_dates = pd.to_datetime(df[col], infer_datetime_format=True, errors='coerce')
                valid_mask = converted_dates.notna() & (converted_dates.notna() != df[col].isna())
                if valid_mask.sum() > len(df) * 0.5:  # Only convert if >50% look like dates
                    df[col] = converted_dates
                    date_fixes += valid_mask.sum()
            except Exception:
                pass
        print(f"  Date columns standardized: {date_fixes}")

        # 5. Remove duplicate rows
        dupes_removed = 0
        if dedupe:
            before_count = len(df)
            df.drop_duplicates(inplace=True)
            dupes_removed = before_count - len(df)
            print(f"  Duplicate rows removed:   {dupes_removed}")

        results[sht] = df
        print(f"  Final size:  {len(df)} rows x {len(df.columns)} columns")
        print()

    # Save cleaned file
    base, ext = os.path.splitext(file_path)
    output_path = f"{base}_CLEANED.xlsx"

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sht, df in results.items():
                df.to_excel(writer, sheet_name=sht, index=False)
                # Auto-fit columns
                ws = writer.sheets[sht]
                for col in ws.columns:
                    max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
                    ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
    except Exception as e:
        print(f"ERROR saving file: {e}")
        sys.exit(1)

    print(f"{'='*55}")
    print(f"  DONE! Cleaned file saved to:")
    print(f"  {output_path}")
    print(f"{'='*55}\n")


def main():
    parser = argparse.ArgumentParser(
        description='KBT Universal Data Cleaner — cleans any Excel file')
    parser.add_argument('file', help='Path to the Excel file to clean')
    parser.add_argument('--sheet', help='Specific sheet name to process (default: all sheets)')
    parser.add_argument('--no-dedupe', action='store_true',
                        help='Skip duplicate row removal')

    args = parser.parse_args()
    clean_excel_file(args.file, sheet_name=args.sheet, dedupe=not args.no_dedupe)


if __name__ == '__main__':
    main()
