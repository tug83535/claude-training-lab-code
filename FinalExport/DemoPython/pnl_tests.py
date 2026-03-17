#!/usr/bin/env python3
"""
pnl_tests.py — Automated Test Suite v2.1
==========================================

PURPOSE: Unit and integration tests for the P&L toolkit.
         Targets: 100% pnl_config, 80% pnl_month_end,
         80% pnl_allocation_simulator, smoke tests for all others.

USAGE:
    pytest pnl_tests.py -v                    # All tests
    pytest pnl_tests.py -v -k "config"        # Only config tests
    pytest pnl_tests.py -v -k "not file"      # Skip file-dependent tests
    pytest pnl_tests.py -v --tb=short         # Short tracebacks
    pytest pnl_tests.py --cov=pnl_config      # With coverage
"""

import os
import sys
import pytest
import tempfile
import importlib
from datetime import datetime
from unittest.mock import patch, MagicMock
from io import StringIO

import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pnl_config import (
    PRODUCTS, DEPARTMENTS, REVENUE_SHARES, AWS_COMPUTE_SHARES,
    HEADCOUNT_SHARES, ALLOCATION_METHODS, EXPENSE_CATEGORIES,
    SHEET_NAMES, SHEET_GL, MONTH_ABBREVS, MONTH_FULL, MONTH_MAP,
    FISCAL_YEAR, FISCAL_YEAR_4, FY_LABEL,
    SOURCE_FILE, DB_PATH, OUTPUT_DIR, CHART_DIR, LOG_DIR,
    HEADER_ROW_TREND, DATA_ROW_TREND, DATA_ROW_CHECKS,
    VARIANCE_PCT, RECON_TOLERANCE, OUTLIER_ZSCORE, OUTLIER_IQR_MULT,
    MAX_LOG_ROWS, COLORS, PRODUCT_COLORS, GL_COLUMNS,
    APP_NAME, APP_VERSION, PnLBase,
    format_currency, format_pct, format_number, resolve_file_path,
)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

HAS_SOURCE = os.path.exists(SOURCE_FILE)

def skip_without_source(reason="Source file not found"):
    return pytest.mark.skipif(not HAS_SOURCE, reason=reason)


def make_fake_gl(n=100, months=None, products=None, departments=None):
    """Build a synthetic GL DataFrame for testing without the Excel file."""
    rng = np.random.default_rng(42)
    months = months or [1, 2, 3]
    products = products or PRODUCTS
    departments = departments or DEPARTMENTS

    rows = []
    for i in range(n):
        m = rng.choice(months)
        rows.append({
            "ID": i + 1,
            "Date": pd.Timestamp(f"2025-{m:02d}-{rng.integers(1, 29):02d}"),
            "Department": rng.choice(departments),
            "Product": rng.choice(products),
            "Expense Category": rng.choice(EXPENSE_CATEGORIES[:5]),
            "Vendor": f"Vendor_{rng.integers(1, 20):03d}",
            "Amount": round(rng.normal(-5000, 3000), 2),
        })
    df = pd.DataFrame(rows)
    df["Month"] = df["Date"].dt.month
    df["Month_Abbrev"] = df["Date"].dt.strftime("%b")
    df["Year"] = df["Date"].dt.year
    df["Quarter"] = df["Date"].dt.quarter
    df["Abs_Amount"] = df["Amount"].abs()
    df["Is_Positive"] = (df["Amount"] > 0).astype(int)
    return df


# =============================================================================
# 1. CONFIG TESTS — Target 100% coverage
# =============================================================================

class TestConfigConstants:
    """Tests for pnl_config.py constants."""

    def test_products_defined(self):
        assert len(PRODUCTS) == 4
        assert "iGO" in PRODUCTS
        assert "Affirm" in PRODUCTS
        assert "InsureSight" in PRODUCTS
        assert "DocFast" in PRODUCTS

    def test_products_are_unique(self):
        assert len(PRODUCTS) == len(set(PRODUCTS))

    def test_departments_defined(self):
        assert len(DEPARTMENTS) == 7
        assert "NetOps" in DEPARTMENTS
        assert "R&D" in DEPARTMENTS
        assert "Product Management" in DEPARTMENTS

    def test_departments_are_unique(self):
        assert len(DEPARTMENTS) == len(set(DEPARTMENTS))

    def test_expense_categories_defined(self):
        assert len(EXPENSE_CATEGORIES) >= 10
        assert "AWS" in EXPENSE_CATEGORIES
        assert "Employee Expenses" in EXPENSE_CATEGORIES

    def test_revenue_shares_sum_to_one(self):
        total = sum(REVENUE_SHARES.values())
        assert abs(total - 1.0) < 0.001, f"Revenue shares sum to {total}"

    def test_aws_shares_sum_to_one(self):
        total = sum(AWS_COMPUTE_SHARES.values())
        assert abs(total - 1.0) < 0.001, f"AWS shares sum to {total}"

    def test_headcount_shares_sum_to_one(self):
        total = sum(HEADCOUNT_SHARES.values())
        assert abs(total - 1.0) < 0.001, f"HC shares sum to {total}"

    def test_all_products_have_revenue_share(self):
        for prod in PRODUCTS:
            assert prod in REVENUE_SHARES, f"Missing revenue share for {prod}"
            assert 0 < REVENUE_SHARES[prod] <= 1.0

    def test_all_products_have_aws_share(self):
        for prod in PRODUCTS:
            assert prod in AWS_COMPUTE_SHARES, f"Missing AWS share for {prod}"

    def test_all_products_have_headcount_share(self):
        for prod in PRODUCTS:
            assert prod in HEADCOUNT_SHARES, f"Missing headcount share for {prod}"

    def test_all_products_have_color(self):
        for prod in PRODUCTS:
            assert prod in PRODUCT_COLORS, f"Missing color for {prod}"
            assert PRODUCT_COLORS[prod].startswith("#")

    def test_allocation_methods_cover_all_depts(self):
        for dept in DEPARTMENTS:
            assert dept in ALLOCATION_METHODS, f"Missing method for {dept}"
            assert ALLOCATION_METHODS[dept] in ("revenue_share", "blended", "headcount")

    def test_sheet_names_required_keys(self):
        required = ["gl", "assumptions", "pnl_trend", "product", "func_trend", "checks"]
        for key in required:
            assert key in SHEET_NAMES, f"Missing sheet key: {key}"

    def test_sheet_gl_matches(self):
        assert SHEET_GL == SHEET_NAMES["gl"]

    def test_month_abbrevs(self):
        assert len(MONTH_ABBREVS) == 12
        assert MONTH_ABBREVS[0] == "Jan"
        assert MONTH_ABBREVS[11] == "Dec"

    def test_month_full(self):
        assert len(MONTH_FULL) == 12
        assert MONTH_FULL[0] == "January"
        assert MONTH_FULL[11] == "December"

    def test_month_map(self):
        assert MONTH_MAP["Jan"] == 1
        assert MONTH_MAP["Dec"] == 12
        assert len(MONTH_MAP) == 12

    def test_fiscal_year_consistency(self):
        assert FISCAL_YEAR in FISCAL_YEAR_4
        assert FISCAL_YEAR in FY_LABEL
        assert len(FISCAL_YEAR) == 2
        assert len(FISCAL_YEAR_4) == 4

    def test_threshold_ranges(self):
        assert 0 < VARIANCE_PCT < 1
        assert RECON_TOLERANCE >= 0
        assert OUTLIER_ZSCORE > 0
        assert OUTLIER_IQR_MULT > 0
        assert MAX_LOG_ROWS > 0

    def test_row_offsets_positive(self):
        assert HEADER_ROW_TREND > 0
        assert DATA_ROW_TREND > HEADER_ROW_TREND
        assert DATA_ROW_CHECKS > 0

    def test_colors_are_hex(self):
        for name, color in COLORS.items():
            assert color.startswith("#"), f"Color '{name}' missing # prefix"
            assert len(color) == 7, f"Color '{name}' wrong length: {color}"

    def test_gl_columns_dict(self):
        required = ["id", "date", "department", "product", "category", "vendor", "amount"]
        for key in required:
            assert key in GL_COLUMNS

    def test_app_identity(self):
        assert APP_NAME is not None and len(APP_NAME) > 0
        assert APP_VERSION is not None
        assert "." in APP_VERSION

    def test_paths_defined(self):
        assert SOURCE_FILE.endswith(".xlsm")
        assert DB_PATH.endswith(".db")
        assert OUTPUT_DIR != ""
        assert CHART_DIR != ""
        assert LOG_DIR != ""


# =============================================================================
# 2. UTILITY FUNCTION TESTS — Target 100% coverage
# =============================================================================

class TestFormatCurrency:
    def test_positive(self):
        assert format_currency(1234.56, 2) == "$1,234.56"

    def test_negative(self):
        assert format_currency(-1234.56, 2) == "($1,234.56)"

    def test_zero(self):
        assert format_currency(0) == "$0"

    def test_nan(self):
        assert format_currency(float("nan")) == "\u2014"

    def test_large_number(self):
        assert format_currency(1234567890) == "$1,234,567,890"

    def test_small_decimal(self):
        assert format_currency(0.50, 2) == "$0.50"

    def test_no_decimals_default(self):
        assert format_currency(1234.99) == "$1,235"


class TestFormatPct:
    def test_basic(self):
        assert format_pct(0.155, 1) == "15.5%"

    def test_zero(self):
        assert format_pct(0) == "0.0%"

    def test_nan(self):
        assert format_pct(float("nan")) == "\u2014"

    def test_negative(self):
        assert format_pct(-0.15, 1) == "-15.0%"

    def test_two_decimals(self):
        assert format_pct(0.12345, 2) == "12.35%"

    def test_one_hundred_pct(self):
        assert format_pct(1.0) == "100.0%"


class TestFormatNumber:
    def test_integer(self):
        assert format_number(1234567) == "1,234,567"

    def test_decimals(self):
        assert format_number(1234.5678, 2) == "1,234.57"

    def test_nan(self):
        assert format_number(float("nan")) == "\u2014"

    def test_zero(self):
        assert format_number(0) == "0"

    def test_negative(self):
        assert format_number(-999.5, 1) == "-999.5"


class TestResolveFilePath:
    def test_direct_path(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(b"test")
            tmp.flush()
            result = resolve_file_path(tmp.name)
            assert result == tmp.name
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass  # Windows file locking — temp dir cleans up later

    def test_nonexistent_raises(self):
        with pytest.raises(FileNotFoundError):
            resolve_file_path("/nonexistent/path.xlsx")

    def test_env_fallback(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(b"test")
            tmp.flush()
            with patch.dict(os.environ, {"KBT_SOURCE_FILE": tmp.name}):
                result = resolve_file_path("/also/nonexistent.xlsx")
                assert result == tmp.name
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass  # Windows file locking — temp dir cleans up later


# =============================================================================
# 3. PnLBase CLASS TESTS — Target 100% coverage
# =============================================================================

class TestPnLBase:
    def test_init_verbose(self):
        base = PnLBase(verbose=True)
        assert base.verbose is True

    def test_init_quiet(self):
        base = PnLBase(verbose=False)
        assert base.verbose is False

    def test_print_verbose(self, capsys):
        base = PnLBase(verbose=True)
        base._print("hello", "INFO")
        captured = capsys.readouterr()
        assert "hello" in captured.out

    def test_print_quiet(self, capsys):
        base = PnLBase(verbose=False)
        base._print("hello", "INFO")
        captured = capsys.readouterr()
        assert captured.out == ""

    def test_print_levels(self, capsys):
        base = PnLBase(verbose=True)
        for level in ["INFO", "WARN", "ERROR", "OK"]:
            base._print(f"test_{level}", level)
        captured = capsys.readouterr()
        assert "test_INFO" in captured.out
        assert "test_WARN" in captured.out
        assert "test_ERROR" in captured.out
        assert "test_OK" in captured.out

    def test_section(self, capsys):
        base = PnLBase(verbose=True)
        base._section("Test Section")
        captured = capsys.readouterr()
        assert "Test Section" in captured.out
        assert "=" in captured.out

    def test_section_quiet(self, capsys):
        base = PnLBase(verbose=False)
        base._section("Test Section")
        captured = capsys.readouterr()
        assert captured.out == ""

    def test_timestamp_format(self):
        ts = PnLBase.timestamp()
        datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")

    def test_file_timestamp_format(self):
        ts = PnLBase.file_timestamp()
        assert " " not in ts
        assert ":" not in ts
        datetime.strptime(ts, "%Y%m%d_%H%M%S")

    def test_ensure_dir_creates(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            new_dir = os.path.join(tmpdir, "subdir", "nested")
            PnLBase.ensure_dir(new_dir)
            assert os.path.isdir(new_dir)

    def test_ensure_dir_existing(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            PnLBase.ensure_dir(tmpdir)
            assert os.path.isdir(tmpdir)

    def test_load_gl_missing_file(self):
        base = PnLBase(verbose=False)
        with pytest.raises(FileNotFoundError):
            base._load_gl("/nonexistent/file.xlsx")

    def test_load_sheet_unknown_key(self):
        base = PnLBase(verbose=False)
        result = base._load_sheet("nonexistent_key_xyz")
        assert result is None


# =============================================================================
# 4. DATA LOADING TESTS (require Excel file)
# =============================================================================

@skip_without_source()
class TestDataLoading:
    def test_load_gl(self):
        base = PnLBase(verbose=False)
        gl = base._load_gl(SOURCE_FILE)
        assert isinstance(gl, pd.DataFrame)
        assert len(gl) > 0

    def test_gl_has_required_columns(self):
        base = PnLBase(verbose=False)
        gl = base._load_gl(SOURCE_FILE)
        for col in ["Date", "Department", "Product", "Amount"]:
            assert col in gl.columns, f"Missing column: {col}"

    def test_gl_computed_columns(self):
        base = PnLBase(verbose=False)
        gl = base._load_gl(SOURCE_FILE)
        for col in ["Abs_Amount", "Is_Positive", "Month", "Quarter", "Year", "Month_Abbrev"]:
            assert col in gl.columns, f"Missing computed column: {col}"

    def test_gl_amount_is_numeric(self):
        base = PnLBase(verbose=False)
        gl = base._load_gl(SOURCE_FILE)
        assert gl["Amount"].dtype in [np.float64, np.int64]

    def test_gl_products_are_known(self):
        base = PnLBase(verbose=False)
        gl = base._load_gl(SOURCE_FILE)
        unknown = set(gl["Product"].unique()) - set(PRODUCTS) - {""}
        assert len(unknown) == 0, f"Unknown products: {unknown}"

    def test_gl_departments_are_known(self):
        base = PnLBase(verbose=False)
        gl = base._load_gl(SOURCE_FILE)
        unknown = set(gl["Department"].unique()) - set(DEPARTMENTS) - {""}
        assert len(unknown) == 0, f"Unknown departments: {unknown}"

    def test_load_sheet_pnl_trend(self):
        base = PnLBase(verbose=False)
        df = base._load_sheet("pnl_trend", SOURCE_FILE)
        assert df is not None
        assert len(df) > 0

    def test_load_sheet_checks(self):
        base = PnLBase(verbose=False)
        df = base._load_sheet("checks", SOURCE_FILE)
        assert df is not None


# =============================================================================
# 5. MONTH-END CLOSE TESTS — Target 80% coverage
# =============================================================================

class TestMonthEndCloseUnit:
    """Unit tests using synthetic data (no Excel required)."""

    def test_check_status_enum(self):
        from pnl_month_end import CheckStatus
        assert CheckStatus.PASS.value == "PASS"
        assert CheckStatus.FAIL.value == "FAIL"
        assert CheckStatus.WARN.value == "WARN"
        assert CheckStatus.SKIP.value == "SKIP"

    def test_close_check_dataclass(self):
        from pnl_month_end import CloseCheck, CheckStatus
        check = CloseCheck("Cat", "Name", CheckStatus.PASS, "Detail", value=42, threshold=50)
        assert check.category == "Cat"
        assert check.status == CheckStatus.PASS
        assert check.value == 42

    def test_close_report_properties(self):
        from pnl_month_end import CloseReport, CloseCheck, CheckStatus
        report = CloseReport(month=1, month_name="January", fiscal_year="FY2025")
        report.checks = [
            CloseCheck("A", "a1", CheckStatus.PASS, "ok"),
            CloseCheck("A", "a2", CheckStatus.PASS, "ok"),
            CloseCheck("B", "b1", CheckStatus.FAIL, "bad"),
            CloseCheck("C", "c1", CheckStatus.WARN, "hmm"),
        ]
        assert report.pass_count == 2
        assert report.fail_count == 1
        assert report.warn_count == 1
        assert report.is_clean is False

    def test_close_report_clean(self):
        from pnl_month_end import CloseReport, CloseCheck, CheckStatus
        report = CloseReport(month=1, month_name="January", fiscal_year="FY2025")
        report.checks = [
            CloseCheck("A", "a1", CheckStatus.PASS, "ok"),
        ]
        assert report.is_clean is True

    def test_determine_month_explicit(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(month=5, verbose=False)
        assert closer._determine_month() == 5

    def test_determine_month_from_gl(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(months=[1, 2, 3])
        assert closer._determine_month() == 3

    def test_check_gl_completeness_with_data(self):
        from pnl_month_end import MonthEndClose, CheckStatus
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=200, months=[1, 2, 3])
        checks = closer.check_gl_completeness(1)
        assert len(checks) >= 1
        assert any(c.name == "Transaction count" and c.status == CheckStatus.PASS for c in checks)

    def test_check_gl_completeness_empty_month(self):
        from pnl_month_end import MonthEndClose, CheckStatus
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=50, months=[2, 3])
        checks = closer.check_gl_completeness(1)
        assert any(c.name == "Transaction count" and c.status == CheckStatus.FAIL for c in checks)

    def test_check_allocation_balance(self):
        from pnl_month_end import MonthEndClose, CheckStatus
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=100, months=[1])
        checks = closer.check_allocation_balance(1)
        rev_check = [c for c in checks if "Revenue shares" in c.name]
        assert len(rev_check) == 1
        assert rev_check[0].status == CheckStatus.PASS

    def test_check_reconciliation_empty(self):
        from pnl_month_end import MonthEndClose, CheckStatus
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=50, months=[2])
        checks = closer.check_reconciliation(1)
        assert any(c.status == CheckStatus.SKIP for c in checks)

    def test_check_variances_no_prior(self):
        from pnl_month_end import MonthEndClose, CheckStatus
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=50, months=[1])
        checks = closer.check_variances(1)
        assert any(c.status == CheckStatus.SKIP for c in checks)

    def test_check_variances_with_prior(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=200, months=[1, 2, 3])
        checks = closer.check_variances(2)
        assert len(checks) >= 1

    def test_check_data_quality(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=100, months=[1])
        checks = closer.check_data_quality(1)
        assert len(checks) >= 1
        names = [c.name for c in checks]
        assert "Potential duplicates" in names or "Statistical outliers" in names

    def test_generate_summary(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=100, months=[1])
        checks = closer.generate_summary(1)
        assert len(checks) >= 3
        names = [c.name for c in checks]
        assert "Total net spend" in names
        assert "Transaction count" in names

    def test_generate_summary_empty(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(verbose=False)
        closer.gl = make_fake_gl(n=50, months=[2])
        checks = closer.generate_summary(1)
        assert len(checks) == 0


@skip_without_source()
class TestMonthEndCloseIntegration:
    """Integration tests with real Excel file."""

    def test_close_runs(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(month=1, verbose=False)
        report = closer.run()
        assert report is not None
        assert len(report.checks) > 0

    def test_close_report_has_status(self):
        from pnl_month_end import MonthEndClose
        closer = MonthEndClose(month=1, verbose=False)
        report = closer.run()
        assert report.pass_count + report.fail_count + report.warn_count > 0

    def test_close_export(self):
        from pnl_month_end import MonthEndClose
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            closer = MonthEndClose(month=1, verbose=False)
            closer.run(export=True, output_path=tmp.name)
            assert os.path.exists(tmp.name)
            assert os.path.getsize(tmp.name) > 0
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass  # Windows file locking — temp dir cleans up later


# =============================================================================
# 6. ALLOCATION SIMULATOR TESTS — Target 80% coverage
# =============================================================================

class TestAllocationSimulatorUnit:
    """Unit tests using synthetic data."""

    def test_compute_metrics_returns_dataframe(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.gl = make_fake_gl(n=200)
        result = sim._compute_metrics(REVENUE_SHARES, "Test")
        assert isinstance(result, pd.DataFrame)
        assert len(result) == len(PRODUCTS)
        assert "Revenue_Share" in result.columns
        assert "CM_Dollar" in result.columns

    def test_compute_metrics_all_products(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.gl = make_fake_gl(n=200)
        result = sim._compute_metrics(REVENUE_SHARES, "Baseline")
        assert set(result["Product"]) == set(PRODUCTS)

    def test_simulate_normalizes_shares(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.gl = make_fake_gl(n=200)
        sim.baseline = sim._compute_metrics(REVENUE_SHARES, "Baseline")
        result = sim.simulate(
            {"iGO": 0.40, "Affirm": 0.30, "InsureSight": 0.20, "DocFast": 0.20},
            scenario_name="Over"
        )
        shares = result["New_Share"].sum()
        assert abs(shares - 1.0) < 0.01

    def test_simulate_preserves_unmodified(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.gl = make_fake_gl(n=200)
        sim.baseline = sim._compute_metrics(REVENUE_SHARES, "Baseline")
        result = sim.simulate({"InsureSight": 0.20}, scenario_name="Partial")
        assert isinstance(result, pd.DataFrame)
        assert len(result) == len(PRODUCTS)

    def test_compare_has_delta_columns(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.gl = make_fake_gl(n=200)
        sim.baseline = sim._compute_metrics(REVENUE_SHARES, "Baseline")
        scenario = sim._compute_metrics(
            {"iGO": 0.50, "Affirm": 0.25, "InsureSight": 0.15, "DocFast": 0.10},
            "Scenario"
        )
        comparison = sim._compare(sim.baseline, scenario)
        for col in ["Share_Change", "Revenue_Delta", "CM_Delta", "CM_Pct_Change"]:
            assert col in comparison.columns

    def test_run_scenarios_multiple(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.gl = make_fake_gl(n=200)
        sim.baseline = sim._compute_metrics(REVENUE_SHARES, "Baseline")
        scenarios = {
            "A": {"iGO": 0.60, "Affirm": 0.20, "InsureSight": 0.12, "DocFast": 0.08},
            "B": {"iGO": 0.40, "Affirm": 0.30, "InsureSight": 0.20, "DocFast": 0.10},
        }
        result = sim.run_scenarios(scenarios)
        assert "Scenario_Name" in result.columns
        assert set(result["Scenario_Name"].unique()) == {"A", "B"}

    def test_print_comparison_no_error(self, capsys):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=True)
        sim.gl = make_fake_gl(n=200)
        sim.baseline = sim._compute_metrics(REVENUE_SHARES, "Baseline")
        comparison = sim.simulate(
            {"InsureSight": 0.20, "DocFast": 0.10, "iGO": 0.45, "Affirm": 0.25}
        )
        sim.print_comparison(comparison)
        captured = capsys.readouterr()
        assert "iGO" in captured.out

    def test_export_creates_file(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.gl = make_fake_gl(n=200)
        sim.baseline = sim._compute_metrics(REVENUE_SHARES, "Baseline")
        comparison = sim.simulate({"InsureSight": 0.20})
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            sim.export(comparison, tmp.name)
            assert os.path.exists(tmp.name)
            assert os.path.getsize(tmp.name) > 0
            try:
                os.unlink(tmp.name)
            except PermissionError:
                pass  # Windows file locking — temp dir cleans up later


@skip_without_source()
class TestAllocationSimulatorIntegration:
    def test_load_and_baseline(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.load()
        assert sim.gl is not None
        assert len(sim.baseline) == len(PRODUCTS)

    def test_simulate_produces_comparison(self):
        from pnl_allocation_simulator import AllocationSimulator
        sim = AllocationSimulator(verbose=False)
        sim.load()
        result = sim.simulate({"InsureSight": 0.20, "DocFast": 0.10})
        assert len(result) == len(PRODUCTS)
        assert "CM_Delta" in result.columns


# =============================================================================
# 7. FORECAST TESTS
# =============================================================================

class TestForecasterUnit:
    """Unit tests for forecast methods (no Excel required)."""

    def test_sma_basic(self):
        from pnl_forecast import PnLForecaster
        fc = PnLForecaster(verbose=False)
        series = pd.Series([100, 110, 120, 130, 140])
        result = fc.forecast_sma(series, periods=3, window=3)
        assert len(result) == 3
        assert all(result["Method"] == "SMA")
        assert abs(result.iloc[0]["Forecast"] - 130.0) < 0.01

    def test_ets_basic(self):
        from pnl_forecast import PnLForecaster
        fc = PnLForecaster(verbose=False)
        series = pd.Series([100, 105, 110, 115, 120])
        result = fc.forecast_ets(series, periods=3)
        assert len(result) == 3
        assert result.iloc[0]["Forecast"] > 115

    def test_trend_basic(self):
        from pnl_forecast import PnLForecaster
        fc = PnLForecaster(verbose=False)
        series = pd.Series([100, 110, 120, 130, 140])
        result = fc.forecast_trend(series, periods=2)
        assert len(result) == 2
        assert abs(result.iloc[0]["Forecast"] - 150) < 5

    def test_forecast_confidence_interval(self):
        from pnl_forecast import PnLForecaster
        fc = PnLForecaster(verbose=False)
        series = pd.Series([100, 110, 105, 115, 120])
        result = fc.forecast_ets(series, periods=2)
        assert "Lower" in result.columns
        assert "Upper" in result.columns
        for _, row in result.iterrows():
            assert row["Lower"] <= row["Forecast"] <= row["Upper"]

    def test_sma_single_period(self):
        from pnl_forecast import PnLForecaster
        fc = PnLForecaster(verbose=False)
        series = pd.Series([10, 20, 30])
        result = fc.forecast_sma(series, periods=1, window=3)
        assert len(result) == 1
        assert abs(result.iloc[0]["Forecast"] - 20.0) < 0.01


# =============================================================================
# 8. SMOKE TESTS — Import + basic instantiation for all remaining modules
# =============================================================================

class TestSmokeImports:
    """Verify all modules import cleanly."""

    def test_import_pnl_config(self):
        mod = importlib.import_module("pnl_config")
        assert hasattr(mod, "PRODUCTS")
        assert hasattr(mod, "PnLBase")

    def test_import_pnl_month_end(self):
        mod = importlib.import_module("pnl_month_end")
        assert hasattr(mod, "MonthEndClose")
        assert hasattr(mod, "CloseReport")

    def test_import_pnl_allocation_simulator(self):
        mod = importlib.import_module("pnl_allocation_simulator")
        assert hasattr(mod, "AllocationSimulator")

    def test_import_pnl_forecast(self):
        mod = importlib.import_module("pnl_forecast")
        assert hasattr(mod, "PnLForecaster")

    def test_import_pnl_snapshot(self):
        mod = importlib.import_module("pnl_snapshot")
        assert hasattr(mod, "SnapshotManager")

    def test_import_pnl_ap_matcher(self):
        mod = importlib.import_module("pnl_ap_matcher")
        assert hasattr(mod, "APMatcher")

    def test_import_pnl_cli(self):
        mod = importlib.import_module("pnl_cli")
        assert hasattr(mod, "main") or hasattr(mod, "cli")

    def test_import_pnl_runner(self):
        mod = importlib.import_module("pnl_runner")
        assert hasattr(mod, "COMMANDS")
        assert hasattr(mod, "main")


@skip_without_source()
class TestSmokeSnapshot:
    def test_snapshot_manager_init(self):
        from pnl_snapshot import SnapshotManager
        mgr = SnapshotManager(verbose=False)
        assert mgr is not None

@skip_without_source()
class TestSmokeAPMatcher:
    def test_ap_matcher_init(self):
        from pnl_ap_matcher import APMatcher
        matcher = APMatcher(verbose=False)
        assert matcher is not None

# =============================================================================
# 9. RUNNER TESTS
# =============================================================================

class TestRunner:
    """Tests for pnl_runner.py command dispatch."""

    def test_commands_registry(self):
        from pnl_runner import COMMANDS
        assert "dashboard" in COMMANDS
        assert "month-end" in COMMANDS
        assert "forecast" in COMMANDS
        assert "allocate" in COMMANDS
        assert "test" in COMMANDS
        assert "config" in COMMANDS

    def test_show_help_no_error(self, capsys):
        from pnl_runner import show_help
        show_help()
        captured = capsys.readouterr()
        assert APP_NAME in captured.out
        assert "dashboard" in captured.out

    def test_unknown_command(self):
        from pnl_runner import main
        with patch("sys.argv", ["pnl_runner.py", "nonexistent_cmd"]):
            result = main()
            assert result == 1

    def test_help_flag(self):
        from pnl_runner import main
        with patch("sys.argv", ["pnl_runner.py", "--help"]):
            result = main()
            assert result == 0


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
