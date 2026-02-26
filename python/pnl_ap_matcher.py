#!/usr/bin/env python3
"""
pnl_ap_matcher.py — AP Invoice Matching Engine
================================================

PURPOSE: Match GL transactions to vendor invoices using fuzzy vendor name
         matching, amount matching, and date proximity. Flags unmatched
         items for review during audit prep.

USAGE:
    python pnl_ap_matcher.py
    python pnl_ap_matcher.py --threshold 85 --export matches.xlsx
    python pnl_ap_matcher.py --vendor "Amazon" --month 3

    from pnl_ap_matcher import APMatcher
    matcher = APMatcher("KeystoneBenefitTech_PL_Model.xlsx")
    result = matcher.match_all()
"""

import os
import sys
import argparse
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass

import pandas as pd
import numpy as np

try:
    from pnl_config import *
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pnl_config import *

try:
    from thefuzz import fuzz, process
    HAS_FUZZ = True
except ImportError:
    HAS_FUZZ = False


@dataclass
class MatchResult:
    gl_index: int
    vendor: str
    amount: float
    date: str
    match_type: str       # exact, fuzzy, amount_only, unmatched
    match_score: float    # 0-100
    matched_vendor: str
    matched_amount: float
    amount_diff: float
    department: str
    product: str


class APMatcher(PnLBase):
    """AP invoice matching and vendor reconciliation engine."""

    def __init__(self, file_path: str = None, fuzzy_threshold: int = 85, verbose: bool = True):
        super().__init__(verbose)
        self.file_path = file_path or SOURCE_FILE
        self.fuzzy_threshold = fuzzy_threshold
        self.gl = None
        self.results = []

    def load(self) -> "APMatcher":
        """Load GL data."""
        self.gl = self._load_gl(self.file_path)
        return self

    # ─────────────────────────────────────────────────────────
    # VENDOR NAME NORMALIZATION
    # ─────────────────────────────────────────────────────────

    @staticmethod
    def normalize_vendor(name: str) -> str:
        """Normalize vendor names for matching."""
        if not name:
            return ""
        name = str(name).strip().lower()
        # Remove common suffixes
        for suffix in [" inc.", " inc", " llc", " ltd", " corp.", " corp",
                       " co.", " co", " international", " services", " consulting"]:
            name = name.replace(suffix, "")
        return name.strip()

    # ─────────────────────────────────────────────────────────
    # DUPLICATE DETECTION
    # ─────────────────────────────────────────────────────────

    def find_potential_duplicates(self, month: int = None, tolerance: float = 0.01) -> pd.DataFrame:
        """
        Find GL transactions that are potential duplicates based on:
        - Same vendor (fuzzy match)
        - Same amount (within tolerance)
        - Close dates (within 7 days)
        """
        if self.gl is None:
            self.load()

        gl = self.gl.copy()
        if month:
            gl = gl[gl["Month"] == month]

        gl = gl[gl["Vendor"] != ""].reset_index(drop=True)
        gl["norm_vendor"] = gl["Vendor"].apply(self.normalize_vendor)

        duplicates = []

        for i in range(len(gl)):
            for j in range(i + 1, min(i + 50, len(gl))):  # Check next 50 records
                # Amount match
                amt_i = gl.loc[i, "Abs_Amount"]
                amt_j = gl.loc[j, "Abs_Amount"]
                if amt_i == 0:
                    continue
                amt_diff = abs(amt_i - amt_j) / amt_i
                if amt_diff > tolerance:
                    continue

                # Vendor match
                v_i = gl.loc[i, "norm_vendor"]
                v_j = gl.loc[j, "norm_vendor"]

                if v_i == v_j:
                    vendor_score = 100
                elif HAS_FUZZ:
                    vendor_score = fuzz.ratio(v_i, v_j)
                else:
                    vendor_score = 100 if v_i == v_j else 0

                if vendor_score < self.fuzzy_threshold:
                    continue

                # Date proximity
                date_i = gl.loc[i, "Date"]
                date_j = gl.loc[j, "Date"]
                day_diff = abs((date_i - date_j).days) if pd.notna(date_i) and pd.notna(date_j) else 999

                if day_diff <= 7:
                    duplicates.append({
                        "Txn_A": gl.loc[i, "ID"] if "ID" in gl.columns else i,
                        "Txn_B": gl.loc[j, "ID"] if "ID" in gl.columns else j,
                        "Vendor_A": gl.loc[i, "Vendor"],
                        "Vendor_B": gl.loc[j, "Vendor"],
                        "Amount_A": gl.loc[i, "Amount"],
                        "Amount_B": gl.loc[j, "Amount"],
                        "Date_A": str(date_i.date()) if pd.notna(date_i) else "",
                        "Date_B": str(date_j.date()) if pd.notna(date_j) else "",
                        "Vendor_Score": vendor_score,
                        "Amount_Diff": abs(amt_i - amt_j),
                        "Day_Diff": day_diff,
                        "Department": gl.loc[i, "Department"],
                        "Risk": "HIGH" if vendor_score >= 95 and amt_diff < 0.001 else "MEDIUM",
                    })

        return pd.DataFrame(duplicates)

    # ─────────────────────────────────────────────────────────
    # VENDOR CONSOLIDATION ANALYSIS
    # ─────────────────────────────────────────────────────────

    def vendor_consolidation(self) -> pd.DataFrame:
        """Find vendor names that are likely the same entity (fuzzy matches)."""
        if self.gl is None:
            self.load()

        vendors = self.gl[self.gl["Vendor"] != ""]["Vendor"].unique()
        normalized = {v: self.normalize_vendor(v) for v in vendors}

        # Group by normalized name
        groups = {}
        for v, norm in normalized.items():
            if norm not in groups:
                groups[norm] = []
            groups[norm].append(v)

        # Find groups with multiple original names
        consolidation = []
        for norm, originals in groups.items():
            if len(originals) > 1:
                for orig in originals:
                    spend = self.gl[self.gl["Vendor"] == orig]["Abs_Amount"].sum()
                    txns = len(self.gl[self.gl["Vendor"] == orig])
                    consolidation.append({
                        "Normalized_Name": norm,
                        "Original_Name": orig,
                        "Total_Spend": spend,
                        "Transactions": txns,
                        "Group_Size": len(originals),
                    })

        # Also find fuzzy matches between normalized names
        if HAS_FUZZ and len(vendors) < 200:  # Only for reasonable sizes
            norm_list = list(set(normalized.values()))
            for i, n1 in enumerate(norm_list):
                for n2 in norm_list[i+1:]:
                    score = fuzz.ratio(n1, n2)
                    if score >= self.fuzzy_threshold and n1 != n2:
                        consolidation.append({
                            "Normalized_Name": f"{n1} ↔ {n2}",
                            "Original_Name": f"Fuzzy match (score: {score})",
                            "Total_Spend": 0,
                            "Transactions": 0,
                            "Group_Size": 0,
                        })

        return pd.DataFrame(consolidation)

    # ─────────────────────────────────────────────────────────
    # VENDOR SPEND ANALYSIS
    # ─────────────────────────────────────────────────────────

    def vendor_spend_analysis(self, vendor: str = None, month: int = None) -> pd.DataFrame:
        """Analyze spending patterns for a specific vendor or all vendors."""
        if self.gl is None:
            self.load()

        gl = self.gl.copy()
        if month:
            gl = gl[gl["Month"] == month]
        if vendor:
            if HAS_FUZZ:
                matches = process.extract(vendor, gl["Vendor"].unique(), limit=5, scorer=fuzz.partial_ratio)
                match_names = [m[0] for m in matches if m[1] >= self.fuzzy_threshold]
                gl = gl[gl["Vendor"].isin(match_names)]
            else:
                gl = gl[gl["Vendor"].str.contains(vendor, case=False, na=False)]

        analysis = gl[gl["Vendor"] != ""].groupby("Vendor").agg(
            Total_Spend=("Amount", "sum"),
            Abs_Spend=("Abs_Amount", "sum"),
            Txn_Count=("Amount", "count"),
            Avg_Txn=("Abs_Amount", "mean"),
            Max_Txn=("Abs_Amount", "max"),
            Departments=("Department", "nunique"),
            Products=("Product", "nunique"),
            First_Date=("Date", "min"),
            Last_Date=("Date", "max"),
            Months_Active=("Month", "nunique"),
        ).sort_values("Abs_Spend", ascending=False).reset_index()

        # Add concentration %
        grand_total = analysis["Abs_Spend"].sum()
        analysis["Pct_of_Total"] = analysis["Abs_Spend"] / grand_total if grand_total > 0 else 0
        analysis["Cumulative_Pct"] = analysis["Pct_of_Total"].cumsum()

        return analysis

    # ─────────────────────────────────────────────────────────
    # MAIN RUNNER
    # ─────────────────────────────────────────────────────────

    def run(self, month: int = None, vendor: str = None, export: str = None):
        """Run the full AP matching analysis."""
        self.load()

        self._section("AP MATCHING ANALYSIS")

        # 1. Vendor spend analysis
        self._print("\n▶ Vendor Spend Analysis")
        spend = self.vendor_spend_analysis(vendor=vendor, month=month)
        top_10 = spend.head(10)
        for _, row in top_10.iterrows():
            self._print(f"  {row['Vendor']:30s}  {format_currency(row['Abs_Spend']):>12s}  "
                        f"({row['Txn_Count']:,} txns, {row['Pct_of_Total']:.1%})")

        # 2. Potential duplicates
        self._print("\n▶ Potential Duplicates")
        dups = self.find_potential_duplicates(month=month)
        if len(dups) > 0:
            self._print(f"  Found {len(dups)} potential duplicate pairs:", "WARN")
            for _, d in dups.head(10).iterrows():
                self._print(f"  ⚠ {d['Vendor_A']} | {format_currency(d['Amount_A'])} | "
                            f"{d['Date_A']} ↔ {d['Date_B']} | Score: {d['Vendor_Score']} | {d['Risk']}")
        else:
            self._print("  ✓ No potential duplicates found")

        # 3. Vendor consolidation
        self._print("\n▶ Vendor Name Consolidation")
        consol = self.vendor_consolidation()
        if len(consol) > 0:
            self._print(f"  Found {len(consol)} vendor name variations:", "WARN")
            for _, c in consol.head(10).iterrows():
                self._print(f"  ⚠ '{c['Original_Name']}' → '{c['Normalized_Name']}'")
        else:
            self._print("  ✓ No consolidation opportunities found")

        # 4. Summary
        self._section("SUMMARY")
        total_vendors = self.gl["Vendor"].nunique()
        total_spend = self.gl["Abs_Amount"].sum()
        top_5_spend = spend.head(5)["Abs_Spend"].sum()
        top_5_pct = top_5_spend / total_spend if total_spend > 0 else 0

        self._print(f"  Total vendors:       {total_vendors:,}")
        self._print(f"  Total spend:         {format_currency(total_spend)}")
        self._print(f"  Top 5 vendor spend:  {format_currency(top_5_spend)} ({top_5_pct:.1%})")
        self._print(f"  Duplicate flags:     {len(dups)}")
        self._print(f"  Name variations:     {len(consol)}")

        # Export
        if export:
            self._export(spend, dups, consol, export)

    def _export(self, spend: pd.DataFrame, dups: pd.DataFrame,
                consol: pd.DataFrame, output_path: str):
        """Export all results to Excel."""
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            spend.to_excel(writer, sheet_name="Vendor Spend", index=False)
            if len(dups) > 0:
                dups.to_excel(writer, sheet_name="Potential Duplicates", index=False)
            if len(consol) > 0:
                consol.to_excel(writer, sheet_name="Name Consolidation", index=False)

        self._print(f"\nExported: {output_path}", "OK")


def main():
    parser = argparse.ArgumentParser(description="AP Invoice Matching Engine")
    parser.add_argument("--file", "-f", default=SOURCE_FILE)
    parser.add_argument("--month", "-m", type=int, default=None)
    parser.add_argument("--vendor", "-v", default=None, help="Filter to specific vendor")
    parser.add_argument("--threshold", "-t", type=int, default=85, help="Fuzzy match threshold (0-100)")
    parser.add_argument("--export", "-e", default=None, help="Export to Excel")
    args = parser.parse_args()

    matcher = APMatcher(file_path=args.file, fuzzy_threshold=args.threshold)
    matcher.run(month=args.month, vendor=args.vendor, export=args.export)


if __name__ == "__main__":
    main()
