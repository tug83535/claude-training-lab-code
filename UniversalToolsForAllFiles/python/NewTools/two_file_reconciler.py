"""
KBT Universal Tools — Two-File Reconciler
Works on ANY CSV/Excel files — no project-specific setup required.

Compares two files row-by-row on a key column and reports:
  - Rows only in File A (missing from B)
  - Rows only in File B (missing from A)
  - Rows in both but with value differences (with column-level detail)

Usage:
    python two_file_reconciler.py fileA.csv fileB.csv --key "Invoice ID" --output recon_report.xlsx
    python two_file_reconciler.py gl_export.xlsx budget.csv --key "Account" --tolerance 0.01
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


def read_input(file_path: Path, sheet: str | int = 0) -> pd.DataFrame:
    """Read a CSV or Excel file into a DataFrame."""
    if file_path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        return pd.read_excel(file_path, sheet_name=sheet)
    elif file_path.suffix.lower() == ".csv":
        return pd.read_csv(file_path)
    else:
        raise ValueError(f"Unsupported file type: {file_path.suffix}")


def reconcile(df_a: pd.DataFrame, df_b: pd.DataFrame,
              key_col: str, tolerance: float = 0.0,
              ignore_cols: list[str] | None = None) -> dict:
    """
    Reconcile two DataFrames on a key column.

    Returns a dict with:
        'only_a': DataFrame — rows only in A
        'only_b': DataFrame — rows only in B
        'differences': list of dicts — column-level differences
        'matched': int — rows that match exactly
        'stats': summary statistics
    """
    if key_col not in df_a.columns:
        raise ValueError(f"Key column '{key_col}' not found in File A. Columns: {list(df_a.columns)}")
    if key_col not in df_b.columns:
        raise ValueError(f"Key column '{key_col}' not found in File B. Columns: {list(df_b.columns)}")

    ignore = set(ignore_cols or [])

    # Normalize keys to string for matching
    keys_a = set(df_a[key_col].astype(str).unique())
    keys_b = set(df_b[key_col].astype(str).unique())

    only_a_keys = keys_a - keys_b
    only_b_keys = keys_b - keys_a
    common_keys = keys_a & keys_b

    # Rows only in A
    only_a = df_a[df_a[key_col].astype(str).isin(only_a_keys)].copy()
    only_b = df_b[df_b[key_col].astype(str).isin(only_b_keys)].copy()

    # Compare common rows
    common_cols = [c for c in df_a.columns if c in df_b.columns and c != key_col and c not in ignore]
    differences = []
    matched = 0

    # Index B by key for fast lookup
    b_indexed = df_b.set_index(df_b[key_col].astype(str))

    for _, row_a in df_a[df_a[key_col].astype(str).isin(common_keys)].iterrows():
        key_val = str(row_a[key_col])

        # Get matching row(s) from B — take first match
        b_matches = b_indexed.loc[[key_val]]
        if len(b_matches) == 0:
            continue
        row_b = b_matches.iloc[0]

        row_diffs = []
        for col in common_cols:
            val_a = row_a[col]
            val_b = row_b[col]

            # Handle NaN
            if pd.isna(val_a) and pd.isna(val_b):
                continue

            if pd.isna(val_a) or pd.isna(val_b):
                row_diffs.append({
                    "column": col,
                    "file_a": val_a if not pd.isna(val_a) else "(blank)",
                    "file_b": val_b if not pd.isna(val_b) else "(blank)",
                })
                continue

            # Numeric comparison with tolerance
            try:
                num_a = float(val_a)
                num_b = float(val_b)
                if abs(num_a - num_b) > tolerance:
                    row_diffs.append({
                        "column": col,
                        "file_a": val_a,
                        "file_b": val_b,
                        "difference": num_a - num_b,
                    })
                continue
            except (ValueError, TypeError):
                pass

            # String comparison
            if str(val_a).strip() != str(val_b).strip():
                row_diffs.append({
                    "column": col,
                    "file_a": val_a,
                    "file_b": val_b,
                })

        if row_diffs:
            differences.append({
                "key": key_val,
                "diffs": row_diffs,
            })
        else:
            matched += 1

    stats = {
        "total_a": len(df_a),
        "total_b": len(df_b),
        "unique_keys_a": len(keys_a),
        "unique_keys_b": len(keys_b),
        "only_in_a": len(only_a_keys),
        "only_in_b": len(only_b_keys),
        "common_keys": len(common_keys),
        "matched": matched,
        "with_differences": len(differences),
        "columns_compared": len(common_cols),
    }

    return {
        "only_a": only_a,
        "only_b": only_b,
        "differences": differences,
        "matched": matched,
        "stats": stats,
    }


def write_report(result: dict, output_path: Path,
                 name_a: str, name_b: str):
    """Write reconciliation results to an Excel file with multiple sheets."""
    stats = result["stats"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Summary sheet
        summary_data = {
            "Metric": [
                "File A", "File B",
                "Total rows in A", "Total rows in B",
                "Unique keys in A", "Unique keys in B",
                "Keys only in A", "Keys only in B",
                "Common keys", "Exact matches", "With differences",
                "Columns compared",
            ],
            "Value": [
                name_a, name_b,
                stats["total_a"], stats["total_b"],
                stats["unique_keys_a"], stats["unique_keys_b"],
                stats["only_in_a"], stats["only_in_b"],
                stats["common_keys"], stats["matched"], stats["with_differences"],
                stats["columns_compared"],
            ],
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="Summary", index=False)

        # Only in A
        if len(result["only_a"]) > 0:
            result["only_a"].to_excel(writer, sheet_name="Only in File A", index=False)

        # Only in B
        if len(result["only_b"]) > 0:
            result["only_b"].to_excel(writer, sheet_name="Only in File B", index=False)

        # Differences detail
        if result["differences"]:
            diff_rows = []
            for item in result["differences"]:
                for d in item["diffs"]:
                    row = {
                        "Key": item["key"],
                        "Column": d["column"],
                        "File A Value": d["file_a"],
                        "File B Value": d["file_b"],
                    }
                    if "difference" in d:
                        row["Numeric Difference"] = d["difference"]
                    diff_rows.append(row)
            pd.DataFrame(diff_rows).to_excel(writer, sheet_name="Differences", index=False)


def print_report(result: dict, name_a: str, name_b: str):
    """Print reconciliation summary to console."""
    stats = result["stats"]
    print("\n" + "=" * 60)
    print("  Two-File Reconciliation Report")
    print("=" * 60)
    print(f"  File A: {name_a} ({stats['total_a']:,} rows, {stats['unique_keys_a']:,} unique keys)")
    print(f"  File B: {name_b} ({stats['total_b']:,} rows, {stats['unique_keys_b']:,} unique keys)")
    print(f"  Columns compared: {stats['columns_compared']}")
    print()
    print(f"  Keys only in A:     {stats['only_in_a']:>8,}")
    print(f"  Keys only in B:     {stats['only_in_b']:>8,}")
    print(f"  Common keys:        {stats['common_keys']:>8,}")
    print(f"    Exact matches:    {stats['matched']:>8,}")
    print(f"    With differences: {stats['with_differences']:>8,}")
    print()

    total_issues = stats["only_in_a"] + stats["only_in_b"] + stats["with_differences"]
    if total_issues == 0:
        print("  RESULT: FULL MATCH — files are identical on all common keys")
    else:
        print(f"  RESULT: {total_issues} issue(s) found — review report for details")

    print("=" * 60)


def main():
    parser = argparse.ArgumentParser(
        description="Reconcile two CSV/Excel files and report differences."
    )
    parser.add_argument("file_a", help="First file (File A)")
    parser.add_argument("file_b", help="Second file (File B)")
    parser.add_argument("--key", "-k", required=True,
                        help="Key column name to match rows on")
    parser.add_argument("--output", "-o",
                        help="Output report file (Excel). Defaults to recon_report.xlsx")
    parser.add_argument("--tolerance", "-t", type=float, default=0.0,
                        help="Numeric tolerance for value comparison (default: 0 = exact)")
    parser.add_argument("--sheet-a", default=0,
                        help="Sheet name/index for File A if Excel (default: first)")
    parser.add_argument("--sheet-b", default=0,
                        help="Sheet name/index for File B if Excel (default: first)")
    parser.add_argument("--ignore", nargs="*", default=[],
                        help="Column names to ignore in comparison")
    parser.add_argument("--preview", action="store_true",
                        help="Print report only — don't write file")
    args = parser.parse_args()

    path_a = Path(args.file_a)
    path_b = Path(args.file_b)

    if not path_a.exists():
        print(f"ERROR: File not found: {path_a}")
        sys.exit(1)
    if not path_b.exists():
        print(f"ERROR: File not found: {path_b}")
        sys.exit(1)

    print(f"Reading File A: {path_a}")
    df_a = read_input(path_a, sheet=args.sheet_a)
    print(f"  {len(df_a):,} rows, {len(df_a.columns)} columns")

    print(f"Reading File B: {path_b}")
    df_b = read_input(path_b, sheet=args.sheet_b)
    print(f"  {len(df_b):,} rows, {len(df_b.columns)} columns")

    print(f"\nReconciling on key column: '{args.key}'")
    if args.tolerance > 0:
        print(f"  Numeric tolerance: {args.tolerance}")

    result = reconcile(df_a, df_b, key_col=args.key,
                       tolerance=args.tolerance,
                       ignore_cols=args.ignore)

    print_report(result, path_a.name, path_b.name)

    if not args.preview:
        output_path = Path(args.output) if args.output else Path("recon_report.xlsx")
        write_report(result, output_path, path_a.name, path_b.name)
        print(f"\nDetailed report written to: {output_path}")


if __name__ == "__main__":
    main()
