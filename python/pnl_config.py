#!/usr/bin/env python3
"""
pnl_config.py — Shared Configuration Module for Keystone BenefitTech P&L Toolkit
==================================================================================

PURPOSE: Single source of truth for all constants, paths, and shared utilities.
         Every other script imports from here instead of defining its own constants.

FIXES:   B1 (wrong filename), B2 (constant duplication), A1 (helper duplication)

USAGE:
    from pnl_config import *
    # or
    from pnl_config import PRODUCTS, REVENUE_SHARES, PnLBase
"""

import os
import sys
import warnings
from datetime import datetime
from typing import Dict, List, Optional, Any

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# =============================================================================
# FILE PATHS
# =============================================================================

SOURCE_FILE = "KeystoneBenefitTech_PL_Model.xlsm"
DB_PATH = "keystone_pnl.db"
OUTPUT_DIR = "./output"
CHART_DIR = "./charts"
LOG_DIR = "./logs"

# =============================================================================
# FISCAL YEAR — CHANGE THESE ANNUALLY
# =============================================================================

FISCAL_YEAR = "25"          # 2-digit for tab names
FISCAL_YEAR_4 = "2025"      # 4-digit for column headers
FY_LABEL = "FY2025"         # Display label


# =============================================================================
# PRODUCTS
# =============================================================================

PRODUCTS = ["iGO", "Affirm", "InsureSight", "DocFast"]

REVENUE_SHARES = {
    "iGO":         0.55,
    "Affirm":      0.28,
    "InsureSight":  0.12,
    "DocFast":      0.05,
}

AWS_COMPUTE_SHARES = {
    "iGO":         0.45,
    "Affirm":      0.25,
    "InsureSight":  0.20,
    "DocFast":      0.10,
}

HEADCOUNT_SHARES = {
    "iGO":         0.50,
    "Affirm":      0.25,
    "InsureSight":  0.15,
    "DocFast":      0.10,
}


# =============================================================================
# DEPARTMENTS
# =============================================================================

DEPARTMENTS = [
    "NetOps", "Security", "Support", "Partners",
    "Content", "R&D", "Product Management"
]

ALLOCATION_METHODS = {
    "NetOps":             "blended",       # 70% AWS compute + 30% revenue
    "Security":           "revenue_share", # 100% pro-rata by revenue
    "Support":            "blended",       # product-tied + vendor-tied + headcount proxy
    "Partners":           "blended",       # InsureSight 22/78 split + revenue share
    "Content":            "blended",       # product-tied + pro-rata revenue
    "R&D":                "blended",       # product-tied + pro-rata revenue
    "Product Management": "blended",       # product-tied + pro-rata revenue
}


# =============================================================================
# EXPENSE CATEGORIES
# =============================================================================

EXPENSE_CATEGORIES = [
    "AWS",
    "Employee Expenses",
    "Professional Fees and Outsourcing Expenses",
    "Software and Maintenance Expense",
    "Advertising and Promotion Expense",
    "Data Center Fees",
    "Connectivity Charges",
    "Office Expenses & Rent Allocation",
    "Overhead Allocation & Chargebacks",
    "Reselling Expenses",
    "Depreciation & Amortization Expense",
]


# =============================================================================
# SHEET NAMES — Must match Excel tab names exactly
# =============================================================================

SHEET_NAMES = {
    "gl":           "CrossfireHiddenWorksheet",
    "assumptions":  "Assumptions",
    "data_dict":    "Data Dictionary",
    "aws":          "AWS Allocation",
    "report":       "Report-->",
    "pnl_trend":    "P&L - Monthly Trend",
    "product":      "Product Line Summary",
    "func_trend":   "Functional P&L - Monthly Trend",
    "func_jan":     "Functional P&L Summary - Jan 25",
    "func_feb":     "Functional P&L Summary - Feb 25",
    "func_mar":     "Functional P&L Summary - Mar 25",
    "natural":      "US January 2025 Natural P&L",
    "checks":       "Checks",
}

SHEET_GL = SHEET_NAMES["gl"]


# =============================================================================
# MONTH REFERENCES
# =============================================================================

MONTH_ABBREVS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

MONTH_FULL = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]

MONTH_MAP = {abbr: i+1 for i, abbr in enumerate(MONTH_ABBREVS)}


# =============================================================================
# ROW OFFSETS — Where data starts in each sheet
# =============================================================================

HEADER_ROW_TREND = 4       # Row with month column headers on trend sheets
DATA_ROW_TREND = 5         # First data row on trend sheets
DATA_ROW_CHECKS = 5        # First data row on Checks sheet
DATA_ROW_ASSUMPTIONS = 5   # First data row on Assumptions (after title rows)


# =============================================================================
# THRESHOLDS
# =============================================================================

VARIANCE_PCT = 0.15          # 15% MoM variance flag
RECON_TOLERANCE = 1.0        # $1 reconciliation tolerance
OUTLIER_ZSCORE = 2.5         # Z-score threshold for outlier detection
OUTLIER_IQR_MULT = 1.5       # IQR multiplier for outlier detection
MAX_LOG_ROWS = 5000          # Audit log auto-prune limit


# =============================================================================
# COLORS — For charts and reports
# =============================================================================

COLORS = {
    "navy":     "#1F4E79",
    "blue":     "#4472C4",
    "green":    "#70AD47",
    "red":      "#C00000",
    "amber":    "#ED7D31",
    "grey":     "#808080",
    "white":    "#FFFFFF",
    "lt_grey":  "#F2F2F2",
    "lt_blue":  "#D5E8F0",
    "lt_green": "#E2EFDA",
    "lt_red":   "#FFE0E0",
    "lt_amber": "#FFF3E0",
}

PRODUCT_COLORS = {
    "iGO":         "#1F4E79",
    "Affirm":      "#4472C4",
    "InsureSight":  "#70AD47",
    "DocFast":      "#ED7D31",
}


# =============================================================================
# GL COLUMN NAMES — Standardized column names after loading
# =============================================================================

GL_COLUMNS = {
    "id":          "ID",
    "date":        "Date",
    "department":  "Department",
    "product":     "Product",
    "category":    "Expense Category",
    "vendor":      "Vendor",
    "amount":      "Amount",
}


# =============================================================================
# APP IDENTITY
# =============================================================================

APP_NAME = "KBT P&L Toolkit"
APP_VERSION = "2.1.0"


# =============================================================================
# SHARED BASE CLASS
# =============================================================================

class PnLBase:
    """
    Base class providing shared utility methods for all toolkit classes.
    Eliminates the need for each class to redefine _print, _section, _load_gl.
    """

    def __init__(self, verbose: bool = True):
        self.verbose = verbose

    def _print(self, msg: str, level: str = "INFO"):
        """Print a message if verbose mode is on."""
        if self.verbose:
            prefix = {"INFO": "  ", "WARN": "⚠ ", "ERROR": "✗ ", "OK": "✓ "}.get(level, "  ")
            print(f"{prefix}{msg}")

    def _section(self, title: str):
        """Print a formatted section header."""
        if self.verbose:
            print(f"\n{'='*60}")
            print(f"  {title}")
            print(f"{'='*60}")

    def _load_gl(self, file_path: str = None) -> pd.DataFrame:
        """
        Load the GL sheet from the Excel file with standard cleaning.
        Returns a clean DataFrame with normalized columns.
        """
        fp = file_path or SOURCE_FILE
        if not os.path.exists(fp):
            raise FileNotFoundError(f"Source file not found: {fp}")

        self._print(f"Loading GL from: {fp}")
        df = pd.read_excel(fp, sheet_name=SHEET_GL, engine="openpyxl")

        # Normalize column names
        df.columns = [str(c).strip() for c in df.columns]

        # Ensure Amount is numeric
        if "Amount" in df.columns:
            df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)

        # Normalize date column
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Month"] = df["Date"].dt.month
            df["Month_Abbrev"] = df["Date"].dt.strftime("%b")
            df["Year"] = df["Date"].dt.year
            df["Quarter"] = df["Date"].dt.quarter

        # Fill NaN strings
        for col in ["Department", "Product", "Expense Category", "Vendor"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        # Add computed columns
        if "Amount" in df.columns:
            df["Abs_Amount"] = df["Amount"].abs()
            df["Is_Positive"] = (df["Amount"] > 0).astype(int)

        self._print(f"Loaded {len(df):,} GL records", "OK")
        return df

    def _load_sheet(self, sheet_key: str, file_path: str = None) -> Optional[pd.DataFrame]:
        """Load any sheet by its key from SHEET_NAMES."""
        fp = file_path or SOURCE_FILE
        sheet_name = SHEET_NAMES.get(sheet_key)
        if not sheet_name:
            self._print(f"Unknown sheet key: {sheet_key}", "WARN")
            return None
        try:
            df = pd.read_excel(fp, sheet_name=sheet_name, engine="openpyxl")
            self._print(f"Loaded sheet '{sheet_name}': {len(df)} rows", "OK")
            return df
        except Exception as e:
            self._print(f"Failed to load '{sheet_name}': {e}", "WARN")
            return None

    @staticmethod
    def ensure_dir(path: str):
        """Create directory if it doesn't exist."""
        os.makedirs(path, exist_ok=True)

    @staticmethod
    def timestamp() -> str:
        """Return a formatted timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def file_timestamp() -> str:
        """Return a filename-safe timestamp."""
        return datetime.now().strftime("%Y%m%d_%H%M%S")


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def format_currency(value: float, decimals: int = 0) -> str:
    """Format a number as currency: $1,234 or ($1,234)."""
    if pd.isna(value):
        return "—"
    if value < 0:
        return f"(${abs(value):,.{decimals}f})"
    return f"${value:,.{decimals}f}"


def format_pct(value: float, decimals: int = 1) -> str:
    """Format a number as percentage: 12.3%."""
    if pd.isna(value):
        return "—"
    return f"{value * 100:.{decimals}f}%"


def format_number(value: float, decimals: int = 0) -> str:
    """Format a number with commas."""
    if pd.isna(value):
        return "—"
    return f"{value:,.{decimals}f}"


def resolve_file_path(file_arg: str = None) -> str:
    """
    Resolve the source Excel file path.
    Priority: CLI argument > environment variable > default constant.
    """
    if file_arg and os.path.exists(file_arg):
        return file_arg
    env_path = os.environ.get("KBT_SOURCE_FILE")
    if env_path and os.path.exists(env_path):
        return env_path
    if os.path.exists(SOURCE_FILE):
        return SOURCE_FILE
    raise FileNotFoundError(
        f"Source file not found. Tried: {file_arg}, ${env_path}, {SOURCE_FILE}"
    )


# =============================================================================
# SELF-TEST
# =============================================================================

if __name__ == "__main__":
    print(f"{APP_NAME} v{APP_VERSION} — Configuration Module")
    print(f"{'='*50}")
    print(f"Source File:      {SOURCE_FILE}")
    print(f"Database:         {DB_PATH}")
    print(f"Fiscal Year:      {FY_LABEL}")
    print(f"Products:         {', '.join(PRODUCTS)}")
    print(f"Departments:      {', '.join(DEPARTMENTS)}")
    print(f"Expense Cats:     {len(EXPENSE_CATEGORIES)}")
    print(f"Sheets:           {len(SHEET_NAMES)}")
    print(f"Variance Thresh:  {format_pct(VARIANCE_PCT)}")
    print(f"Recon Tolerance:  {format_currency(RECON_TOLERANCE)}")
    print()

    # Verify shares sum to 1.0
    rev_sum = sum(REVENUE_SHARES.values())
    aws_sum = sum(AWS_COMPUTE_SHARES.values())
    hc_sum = sum(HEADCOUNT_SHARES.values())
    print(f"Revenue shares sum:   {rev_sum:.2f} {'✓' if abs(rev_sum - 1.0) < 0.01 else '✗ MISMATCH'}")
    print(f"AWS compute shares:   {aws_sum:.2f} {'✓' if abs(aws_sum - 1.0) < 0.01 else '✗ MISMATCH'}")
    print(f"Headcount shares:     {hc_sum:.2f} {'✓' if abs(hc_sum - 1.0) < 0.01 else '✗ MISMATCH'}")

    # Check if source file exists
    if os.path.exists(SOURCE_FILE):
        print(f"\n✓ Source file found: {SOURCE_FILE}")
    else:
        print(f"\n⚠ Source file not found: {SOURCE_FILE}")
        print(f"  Set KBT_SOURCE_FILE env var or pass --file argument")
