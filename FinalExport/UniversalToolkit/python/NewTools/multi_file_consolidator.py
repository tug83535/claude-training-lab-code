"""
KBT Universal Tools — Multi-File Consolidator
Works on ANY CSV/Excel files — no project-specific setup required.

Combines multiple CSV or Excel files into a single consolidated file.
Handles column mismatches, different formats, and adds a source file column.

Usage:
    python multi_file_consolidator.py file1.csv file2.csv file3.csv --output combined.xlsx
    python multi_file_consolidator.py *.csv --output all_data.csv
    python multi_file_consolidator.py folder/ --output combined.xlsx --recursive
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls", ".xlsm"}


def find_files(paths: list[str], recursive: bool = False) -> list[Path]:
    """Resolve input paths to a list of supported files."""
    files = []
    for p in paths:
        path = Path(p)
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(path)
        elif path.is_dir():
            pattern = "**/*" if recursive else "*"
            for ext in SUPPORTED_EXTENSIONS:
                files.extend(sorted(path.glob(f"{pattern}{ext}")))
        else:
            # Try as glob pattern
            matches = sorted(Path(".").glob(p))
            files.extend(f for f in matches
                         if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS)

    # Deduplicate while preserving order
    seen = set()
    unique = []
    for f in files:
        resolved = f.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(f)

    return unique


def read_file(file_path: Path, sheet: str | int = 0) -> list[tuple[pd.DataFrame, str]]:
    """Read a file and return list of (DataFrame, source_label) tuples."""
    results = []

    if file_path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        xls = pd.ExcelFile(file_path)
        if isinstance(sheet, str) and sheet == "all":
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                label = f"{file_path.name} | {sheet_name}"
                results.append((df, label))
        else:
            df = pd.read_excel(file_path, sheet_name=sheet)
            results.append((df, file_path.name))
    elif file_path.suffix.lower() == ".csv":
        df = pd.read_csv(file_path)
        results.append((df, file_path.name))

    return results


def consolidate(dataframes: list[tuple[pd.DataFrame, str]],
                add_source: bool = True,
                match_mode: str = "union") -> tuple[pd.DataFrame, dict]:
    """
    Consolidate multiple DataFrames into one.

    match_mode:
        'union' — include all columns from all files (fill missing with NaN)
        'intersection' — only include columns common to ALL files
    """
    if not dataframes:
        return pd.DataFrame(), {}

    stats = {
        "files": len(dataframes),
        "total_rows": 0,
        "all_columns": set(),
        "common_columns": None,
        "per_file": [],
    }

    frames = []
    for df, source in dataframes:
        if add_source:
            df = df.copy()
            df.insert(0, "_Source_File", source)

        stats["total_rows"] += len(df)
        cols = set(df.columns) - {"_Source_File"}
        stats["all_columns"] |= cols
        if stats["common_columns"] is None:
            stats["common_columns"] = cols.copy()
        else:
            stats["common_columns"] &= cols

        stats["per_file"].append({
            "source": source,
            "rows": len(df),
            "columns": len(df.columns) - (1 if add_source else 0),
        })

        frames.append(df)

    if match_mode == "intersection" and stats["common_columns"]:
        keep_cols = sorted(stats["common_columns"])
        if add_source:
            keep_cols = ["_Source_File"] + keep_cols
        frames = [f[keep_cols] for f in frames]

    combined = pd.concat(frames, ignore_index=True, sort=False)
    return combined, stats


def print_report(stats: dict, match_mode: str):
    """Print consolidation summary."""
    print("\n" + "=" * 60)
    print("  Multi-File Consolidation Report")
    print("=" * 60)
    print(f"  Files consolidated: {stats['files']}")
    print(f"  Total rows: {stats['total_rows']:,}")
    print(f"  Column mode: {match_mode}")
    print(f"  All columns (union): {len(stats['all_columns'])}")
    print(f"  Common columns (intersection): {len(stats['common_columns'] or set())}")

    print("\n  Per-file breakdown:")
    print(f"  {'Source':<40} {'Rows':>8} {'Cols':>6}")
    print(f"  {'-'*40} {'-'*8} {'-'*6}")
    for pf in stats["per_file"]:
        print(f"  {pf['source']:<40} {pf['rows']:>8,} {pf['columns']:>6}")

    # Column mismatch warning
    if stats["common_columns"] and len(stats["common_columns"]) < len(stats["all_columns"]):
        extra = stats["all_columns"] - stats["common_columns"]
        if len(extra) <= 10:
            print(f"\n  WARNING: These columns are not in all files: {', '.join(sorted(extra))}")
        else:
            print(f"\n  WARNING: {len(extra)} columns are not present in all files")

    print("=" * 60)


def main():
    parser = argparse.ArgumentParser(
        description="Consolidate multiple CSV/Excel files into one."
    )
    parser.add_argument("files", nargs="+",
                        help="Input files, directories, or glob patterns")
    parser.add_argument("--output", "-o", required=True,
                        help="Output file path (CSV or Excel)")
    parser.add_argument("--sheet", "-s", default="all",
                        help="Sheet to read from Excel files (default: all)")
    parser.add_argument("--mode", "-m", choices=["union", "intersection"],
                        default="union",
                        help="Column matching: union (all cols) or intersection (common only)")
    parser.add_argument("--no-source", action="store_true",
                        help="Don't add _Source_File column")
    parser.add_argument("--recursive", "-r", action="store_true",
                        help="Search directories recursively")
    parser.add_argument("--preview", action="store_true",
                        help="Preview only — don't write output")
    args = parser.parse_args()

    # Find all input files
    input_files = find_files(args.files, recursive=args.recursive)
    if not input_files:
        print("ERROR: No supported files found.")
        sys.exit(1)

    print(f"Found {len(input_files)} file(s) to consolidate:")
    for f in input_files:
        print(f"  - {f}")

    # Read all files
    all_data = []
    for file_path in input_files:
        try:
            results = read_file(file_path, sheet=args.sheet)
            all_data.extend(results)
            for _, label in results:
                pass  # Already printed in read_file
        except Exception as e:
            print(f"  ERROR reading {file_path}: {e}")

    if not all_data:
        print("ERROR: No data loaded from any file.")
        sys.exit(1)

    # Consolidate
    combined, stats = consolidate(all_data, add_source=not args.no_source,
                                   match_mode=args.mode)
    print_report(stats, args.mode)

    if args.preview:
        print("\n[PREVIEW MODE — no file written]")
        print(combined.head(10).to_string())
        return

    # Write output
    output_path = Path(args.output)
    if output_path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        combined.to_excel(output_path, index=False)
    else:
        combined.to_csv(output_path, index=False)

    print(f"\nOutput written to: {output_path}")
    print(f"  {len(combined):,} rows, {len(combined.columns)} columns")


if __name__ == "__main__":
    main()
