"""
KBT Universal Tools — SQL Query Tool
Works on ANY CSV/Excel file — no project-specific setup required.

Lets you run SQL queries against CSV or Excel files using SQLite.
Each file/sheet becomes a table you can SELECT, JOIN, filter, and aggregate.

Usage:
    python sql_query_tool.py data.csv --query "SELECT * FROM data WHERE Amount > 1000"
    python sql_query_tool.py sales.xlsx orders.csv --query "SELECT s.Product, o.Qty FROM sales s JOIN orders o ON s.ID = o.ID"
    python sql_query_tool.py data.csv --interactive
"""

import argparse
import sqlite3
import sys
from pathlib import Path

import pandas as pd


def load_file_to_sqlite(conn: sqlite3.Connection, file_path: Path,
                        sheet: str | int = 0) -> list[str]:
    """Load a CSV or Excel file into SQLite. Returns list of table names created."""
    tables_created = []

    if file_path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        # Load all sheets (or specific one)
        xls = pd.ExcelFile(file_path)
        if isinstance(sheet, str) and sheet != "all":
            sheets = [sheet]
        elif sheet == "all":
            sheets = xls.sheet_names
        else:
            sheets = [xls.sheet_names[int(sheet)]] if isinstance(sheet, int) else [xls.sheet_names[0]]

        for sheet_name in sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # Sanitize table name: replace spaces/special chars
            table_name = sanitize_name(sheet_name)
            df.to_sql(table_name, conn, if_exists="replace", index=False)
            tables_created.append(table_name)
            print(f"  Loaded sheet '{sheet_name}' as table '{table_name}' ({len(df)} rows, {len(df.columns)} cols)")

    elif file_path.suffix.lower() == ".csv":
        df = pd.read_csv(file_path)
        table_name = sanitize_name(file_path.stem)
        df.to_sql(table_name, conn, if_exists="replace", index=False)
        tables_created.append(table_name)
        print(f"  Loaded '{file_path.name}' as table '{table_name}' ({len(df)} rows, {len(df.columns)} cols)")

    else:
        print(f"  WARNING: Skipping unsupported file: {file_path}")

    return tables_created


def sanitize_name(name: str) -> str:
    """Make a string safe for use as a SQLite table name."""
    safe = "".join(c if c.isalnum() or c == "_" else "_" for c in name)
    # Ensure it doesn't start with a digit
    if safe and safe[0].isdigit():
        safe = "t_" + safe
    return safe


def show_schema(conn: sqlite3.Connection, tables: list[str]):
    """Print column names and types for each table."""
    cursor = conn.cursor()
    for table in tables:
        cursor.execute(f"PRAGMA table_info({table})")
        cols = cursor.fetchall()
        print(f"\n  Table: {table}")
        print(f"  {'Column':<30} {'Type':<15}")
        print(f"  {'-'*30} {'-'*15}")
        for col in cols:
            print(f"  {col[1]:<30} {col[2]:<15}")


def run_query(conn: sqlite3.Connection, query: str,
              output_path: Path | None = None) -> pd.DataFrame | None:
    """Execute a SQL query and return/display/export results."""
    try:
        df = pd.read_sql_query(query, conn)
    except Exception as e:
        print(f"  SQL ERROR: {e}")
        return None

    print(f"\n  Query returned {len(df)} row(s), {len(df.columns)} column(s)")

    if output_path:
        if output_path.suffix.lower() == ".csv":
            df.to_csv(output_path, index=False)
        else:
            df.to_excel(output_path, index=False)
        print(f"  Results exported to: {output_path}")
    else:
        # Display results (limit to 50 rows for console)
        if len(df) <= 50:
            print(df.to_string(index=False))
        else:
            print(df.head(50).to_string(index=False))
            print(f"\n  ... showing first 50 of {len(df)} rows")

    return df


def interactive_mode(conn: sqlite3.Connection, tables: list[str]):
    """Run an interactive SQL shell."""
    print("\n" + "=" * 60)
    print("  SQL Query Tool — Interactive Mode")
    print("  Type SQL queries, 'schema' to see tables, 'quit' to exit")
    print("=" * 60)
    show_schema(conn, tables)

    while True:
        try:
            query = input("\nSQL> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.")
            break

        if not query:
            continue
        if query.lower() in ("quit", "exit", "q"):
            break
        if query.lower() == "schema":
            show_schema(conn, tables)
            continue
        if query.lower().startswith("export "):
            # export filename.csv SELECT ...
            parts = query.split(" ", 2)
            if len(parts) >= 3:
                run_query(conn, parts[2], Path(parts[1]))
            else:
                print("  Usage: export output.csv SELECT ...")
            continue

        run_query(conn, query)


def main():
    parser = argparse.ArgumentParser(
        description="Run SQL queries against CSV and Excel files."
    )
    parser.add_argument("files", nargs="+", help="Input file(s) — CSV or Excel")
    parser.add_argument("--query", "-q", help="SQL query to execute")
    parser.add_argument("--output", "-o", help="Export results to file (CSV or Excel)")
    parser.add_argument("--sheet", "-s", default="all",
                        help="Sheet name or 'all' for Excel files (default: all)")
    parser.add_argument("--interactive", "-i", action="store_true",
                        help="Enter interactive SQL mode")
    args = parser.parse_args()

    # Create in-memory SQLite database
    conn = sqlite3.connect(":memory:")
    all_tables = []

    print("Loading files into SQL engine...")
    for file_str in args.files:
        file_path = Path(file_str)
        if not file_path.exists():
            print(f"  WARNING: File not found: {file_path}")
            continue
        tables = load_file_to_sqlite(conn, file_path, sheet=args.sheet)
        all_tables.extend(tables)

    if not all_tables:
        print("ERROR: No data loaded.")
        conn.close()
        sys.exit(1)

    print(f"\n{len(all_tables)} table(s) ready: {', '.join(all_tables)}")

    if args.interactive:
        interactive_mode(conn, all_tables)
    elif args.query:
        output_path = Path(args.output) if args.output else None
        run_query(conn, args.query, output_path)
    else:
        # Default: show schema
        show_schema(conn, all_tables)
        print("\nUse --query 'SELECT ...' or --interactive to query data.")

    conn.close()


if __name__ == "__main__":
    main()
