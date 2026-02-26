#!/usr/bin/env python3
"""
pnl_snapshot.py — Point-in-Time P&L Snapshots
================================================

PURPOSE: Save timestamped snapshots of key P&L metrics to a SQLite history
         table. Enables tracking how the P&L evolves during close cycles.

USAGE:
    python pnl_snapshot.py save                              # Save current snapshot
    python pnl_snapshot.py save --label "Pre-adjustment"     # With label
    python pnl_snapshot.py list                              # List all snapshots
    python pnl_snapshot.py compare --a 1 --b 3               # Compare snapshot 1 vs 3
    python pnl_snapshot.py export --id 3                     # Export snapshot to Excel

    from pnl_snapshot import SnapshotManager
    mgr = SnapshotManager("KeystoneBenefitTech_PL_Model.xlsx")
    mgr.save_snapshot(label="February Pre-Close")
"""

import os
import sys
import sqlite3
import argparse
import json
from typing import Dict, List, Optional

import pandas as pd
import numpy as np

try:
    from pnl_config import *
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pnl_config import *


SNAPSHOT_DB = "pnl_snapshots.db"


class SnapshotManager(PnLBase):
    """Manages point-in-time P&L snapshots."""

    def __init__(self, file_path: str = None, db_path: str = None, verbose: bool = True):
        super().__init__(verbose)
        self.file_path = file_path or SOURCE_FILE
        self.db_path = db_path or SNAPSHOT_DB
        self.gl = None
        self._init_db()

    def _init_db(self):
        """Create snapshot tables if they don't exist."""
        conn = sqlite3.connect(self.db_path)
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS snapshots (
                snapshot_id     INTEGER PRIMARY KEY AUTOINCREMENT,
                label           TEXT NOT NULL,
                source_file     TEXT NOT NULL,
                created_at      TEXT NOT NULL DEFAULT (datetime('now')),
                fiscal_year     TEXT NOT NULL,
                total_gl_rows   INTEGER NOT NULL DEFAULT 0,
                metrics_json    TEXT NOT NULL DEFAULT '{}'
            );

            CREATE TABLE IF NOT EXISTS snapshot_details (
                detail_id       INTEGER PRIMARY KEY AUTOINCREMENT,
                snapshot_id     INTEGER NOT NULL REFERENCES snapshots(snapshot_id),
                dimension_type  TEXT NOT NULL,   -- 'product', 'department', 'category', 'total'
                dimension_value TEXT NOT NULL,
                month           INTEGER,
                total_spend     REAL DEFAULT 0,
                abs_spend       REAL DEFAULT 0,
                txn_count       INTEGER DEFAULT 0,
                vendor_count    INTEGER DEFAULT 0
            );

            CREATE INDEX IF NOT EXISTS idx_snap_detail ON snapshot_details(snapshot_id);
        """)
        conn.close()

    def save_snapshot(self, label: str = None) -> int:
        """Save a snapshot of current P&L state."""
        self.gl = self._load_gl(self.file_path)
        gl = self.gl

        if not label:
            label = f"Snapshot {PnLBase.timestamp()}"

        self._section(f"SAVING SNAPSHOT: {label}")

        # Compute high-level metrics
        metrics = {
            "total_net_spend": float(gl["Amount"].sum()),
            "total_abs_spend": float(gl["Abs_Amount"].sum()),
            "total_transactions": int(len(gl)),
            "unique_vendors": int(gl["Vendor"].nunique()),
            "months_covered": sorted(gl["Month"].dropna().unique().astype(int).tolist()),
            "products": {p: float(gl[gl["Product"] == p]["Abs_Amount"].sum()) for p in PRODUCTS},
            "departments": {d: float(gl[gl["Department"] == d]["Abs_Amount"].sum()) for d in DEPARTMENTS},
        }

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Insert snapshot header
        cursor.execute("""
            INSERT INTO snapshots (label, source_file, fiscal_year, total_gl_rows, metrics_json)
            VALUES (?, ?, ?, ?, ?)
        """, (label, self.file_path, FY_LABEL, len(gl), json.dumps(metrics)))
        snap_id = cursor.lastrowid

        # Insert detail rows: by product × month
        detail_rows = []
        for prod in PRODUCTS:
            for month in sorted(gl["Month"].dropna().unique()):
                mask = (gl["Product"] == prod) & (gl["Month"] == month)
                subset = gl[mask]
                if len(subset) > 0:
                    detail_rows.append((
                        snap_id, "product", prod, int(month),
                        float(subset["Amount"].sum()),
                        float(subset["Abs_Amount"].sum()),
                        int(len(subset)),
                        int(subset["Vendor"].nunique()),
                    ))

        # By department × month
        for dept in DEPARTMENTS:
            for month in sorted(gl["Month"].dropna().unique()):
                mask = (gl["Department"] == dept) & (gl["Month"] == month)
                subset = gl[mask]
                if len(subset) > 0:
                    detail_rows.append((
                        snap_id, "department", dept, int(month),
                        float(subset["Amount"].sum()),
                        float(subset["Abs_Amount"].sum()),
                        int(len(subset)),
                        int(subset["Vendor"].nunique()),
                    ))

        # Total by month
        for month in sorted(gl["Month"].dropna().unique()):
            subset = gl[gl["Month"] == month]
            detail_rows.append((
                snap_id, "total", "All", int(month),
                float(subset["Amount"].sum()),
                float(subset["Abs_Amount"].sum()),
                int(len(subset)),
                int(subset["Vendor"].nunique()),
            ))

        cursor.executemany("""
            INSERT INTO snapshot_details
            (snapshot_id, dimension_type, dimension_value, month,
             total_spend, abs_spend, txn_count, vendor_count)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, detail_rows)

        conn.commit()
        conn.close()

        self._print(f"Snapshot #{snap_id} saved: {label}", "OK")
        self._print(f"  GL rows: {len(gl):,}")
        self._print(f"  Detail rows: {len(detail_rows):,}")
        self._print(f"  Total spend: {format_currency(metrics['total_abs_spend'])}")

        return snap_id

    def list_snapshots(self) -> pd.DataFrame:
        """List all saved snapshots."""
        conn = sqlite3.connect(self.db_path)
        df = pd.read_sql_query("""
            SELECT snapshot_id, label, created_at, fiscal_year, total_gl_rows,
                   (SELECT COUNT(*) FROM snapshot_details WHERE snapshot_id = s.snapshot_id) AS detail_rows
            FROM snapshots s
            ORDER BY snapshot_id DESC
        """, conn)
        conn.close()

        self._section("SAVED SNAPSHOTS")
        if len(df) == 0:
            self._print("No snapshots saved yet. Use 'save' to create one.")
        else:
            for _, row in df.iterrows():
                self._print(f"  #{row['snapshot_id']:3d}  {row['created_at']}  "
                            f"{row['label']:30s}  ({row['total_gl_rows']:,} GL rows)")
        return df

    def compare_snapshots(self, snap_a: int, snap_b: int) -> pd.DataFrame:
        """Compare two snapshots and show deltas."""
        conn = sqlite3.connect(self.db_path)

        # Get snapshot labels
        labels = pd.read_sql_query(
            "SELECT snapshot_id, label, created_at FROM snapshots WHERE snapshot_id IN (?, ?)",
            conn, params=(snap_a, snap_b)
        )

        # Get details for both
        details_a = pd.read_sql_query(
            "SELECT * FROM snapshot_details WHERE snapshot_id = ?",
            conn, params=(snap_a,)
        )
        details_b = pd.read_sql_query(
            "SELECT * FROM snapshot_details WHERE snapshot_id = ?",
            conn, params=(snap_b,)
        )
        conn.close()

        if len(details_a) == 0 or len(details_b) == 0:
            self._print("One or both snapshots not found", "ERROR")
            return pd.DataFrame()

        self._section(f"COMPARING SNAPSHOT #{snap_a} vs #{snap_b}")

        # Merge on dimension keys
        merge_cols = ["dimension_type", "dimension_value", "month"]
        comparison = details_a[merge_cols + ["abs_spend", "txn_count"]].merge(
            details_b[merge_cols + ["abs_spend", "txn_count"]],
            on=merge_cols, how="outer", suffixes=("_a", "_b")
        ).fillna(0)

        comparison["spend_delta"] = comparison["abs_spend_b"] - comparison["abs_spend_a"]
        comparison["spend_pct_change"] = comparison.apply(
            lambda r: r["spend_delta"] / r["abs_spend_a"] if r["abs_spend_a"] != 0 else 0, axis=1
        )

        # Print summary by product (totals across months)
        products = comparison[comparison["dimension_type"] == "product"]
        if len(products) > 0:
            prod_summary = products.groupby("dimension_value").agg(
                Spend_A=("abs_spend_a", "sum"),
                Spend_B=("abs_spend_b", "sum"),
                Delta=("spend_delta", "sum"),
            ).reset_index()
            prod_summary["Pct_Change"] = prod_summary["Delta"] / prod_summary["Spend_A"].replace(0, np.nan)

            self._print("\n  Product Changes:")
            for _, row in prod_summary.iterrows():
                arrow = "↑" if row["Delta"] > 0 else "↓" if row["Delta"] < 0 else "—"
                self._print(f"    {row['dimension_value']:15s}  "
                            f"{format_currency(row['Spend_A']):>12s} → {format_currency(row['Spend_B']):>12s}  "
                            f"{arrow} {format_currency(abs(row['Delta']))}")

        return comparison

    def export_snapshot(self, snapshot_id: int, output_path: str = None):
        """Export a snapshot to Excel."""
        conn = sqlite3.connect(self.db_path)
        header = pd.read_sql_query(
            "SELECT * FROM snapshots WHERE snapshot_id = ?",
            conn, params=(snapshot_id,)
        )
        details = pd.read_sql_query(
            "SELECT * FROM snapshot_details WHERE snapshot_id = ?",
            conn, params=(snapshot_id,)
        )
        conn.close()

        if len(header) == 0:
            self._print(f"Snapshot #{snapshot_id} not found", "ERROR")
            return

        out = output_path or f"snapshot_{snapshot_id}.xlsx"
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            header.to_excel(writer, sheet_name="Summary", index=False)
            details.to_excel(writer, sheet_name="Details", index=False)

            # Pivot: product × month
            prod_details = details[details["dimension_type"] == "product"]
            if len(prod_details) > 0:
                pivot = prod_details.pivot_table(
                    index="dimension_value", columns="month",
                    values="abs_spend", fill_value=0
                )
                pivot.to_excel(writer, sheet_name="Product × Month")

        self._print(f"Snapshot #{snapshot_id} exported: {out}", "OK")


def main():
    parser = argparse.ArgumentParser(description="P&L Snapshot Manager")
    parser.add_argument("--file", "-f", default=SOURCE_FILE)
    parser.add_argument("--db", default=SNAPSHOT_DB)

    sub = parser.add_subparsers(dest="command")

    p_save = sub.add_parser("save", help="Save a snapshot")
    p_save.add_argument("--label", "-l", default=None)

    sub.add_parser("list", help="List snapshots")

    p_cmp = sub.add_parser("compare", help="Compare two snapshots")
    p_cmp.add_argument("--a", type=int, required=True, help="First snapshot ID")
    p_cmp.add_argument("--b", type=int, required=True, help="Second snapshot ID")

    p_exp = sub.add_parser("export", help="Export a snapshot")
    p_exp.add_argument("--id", type=int, required=True)
    p_exp.add_argument("--output", "-o", default=None)

    args = parser.parse_args()
    mgr = SnapshotManager(file_path=args.file, db_path=args.db)

    if args.command == "save":
        mgr.save_snapshot(label=args.label)
    elif args.command == "list":
        mgr.list_snapshots()
    elif args.command == "compare":
        mgr.compare_snapshots(args.a, args.b)
    elif args.command == "export":
        mgr.export_snapshot(args.id, output_path=args.output)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
