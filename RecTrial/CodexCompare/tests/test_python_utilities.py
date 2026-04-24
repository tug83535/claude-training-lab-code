#!/usr/bin/env python3
"""Basic unit tests for universal Python utilities."""

from __future__ import annotations

import csv
import importlib.util
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


def load_module(rel_path: str):
    path = ROOT / rel_path
    spec = importlib.util.spec_from_file_location(path.stem, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load module: {rel_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


profile_mod = load_module("python/universal/profile_workbook.py")
sanitize_mod = load_module("python/universal/sanitize_dataset.py")
compare_mod = load_module("python/universal/compare_workbooks.py")
summary_mod = load_module("python/universal/build_exec_summary.py")
bootstrap_mod = load_module("scripts/bootstrap_demo_workspace.py")
pnl_extract_mod = load_module("python/demo/pnl_data_extract.py")
variance_mod = load_module("python/demo/variance_classifier.py")
scenario_mod = load_module("python/demo/scenario_runner.py")
brief_mod = load_module("python/demo/export_brief_package.py")


class TestProfileWorkbook(unittest.TestCase):
    def test_build_profile_returns_sheet_count(self):
        workbook = ROOT / "samples/ExcelDemoFile_adv.xlsm"
        profile = profile_mod.build_profile(workbook)
        self.assertGreater(profile["sheet_count"], 0)
        self.assertTrue(profile["has_vba_project"])


class TestSanitizeDataset(unittest.TestCase):
    def test_sanitize_cell_date_and_number(self):
        self.assertEqual(sanitize_mod.sanitize_cell("01/15/2026"), "2026-01-15")
        self.assertEqual(sanitize_mod.sanitize_cell(" $1,200 "), "1200")

    def test_sanitize_csv_writes_output(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            in_path = Path(tmp_dir) / "in.csv"
            out_path = Path(tmp_dir) / "out.csv"
            in_path.write_text("Department,Amount\n Sales , $1,200 \n", encoding="utf-8")

            total_cells, changed_cells = sanitize_mod.sanitize_csv(in_path, out_path)
            self.assertGreater(total_cells, 0)
            self.assertGreater(changed_cells, 0)
            self.assertTrue(out_path.exists())


class TestCompareWorkbooks(unittest.TestCase):
    def test_compare_identical_workbooks_returns_no_diffs(self):
        wb = ROOT / "samples/Sample_Quarterly_ReportV2.xlsm"
        diffs = compare_mod.compare(wb, wb)
        self.assertEqual(len(diffs), 0)


class TestExecSummary(unittest.TestCase):
    def test_build_summary_contains_expected_sections(self):
        rows = [
            {"Department": "Sales", "Amount": "1000"},
            {"Department": "Finance", "Amount": "500"},
            {"Department": "Sales", "Amount": "250"},
        ]
        summary = summary_mod.build_summary(rows, "Amount", "Department")
        self.assertIn("# Executive Summary", summary)
        self.assertIn("## Top contributing groups", summary)

    def test_read_rows(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_path = Path(tmp_dir) / "sample.csv"
            with csv_path.open("w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Department", "Amount"])
                writer.writerow(["Finance", "100"])
            rows = summary_mod.read_rows(csv_path)
            self.assertEqual(len(rows), 1)


class TestBootstrapWorkspace(unittest.TestCase):
    def test_bootstrap_creates_workspace_copy(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_root = Path(tmp_dir) / "out"
            out_root.mkdir(parents=True, exist_ok=True)
            target_dir = bootstrap_mod.create_workspace(out_root, timestamp="test")

            self.assertTrue((target_dir / "ExcelDemoFile_adv.xlsm").exists())
            self.assertTrue((target_dir / "Sample_Quarterly_ReportV2.xlsm").exists())
            self.assertEqual(target_dir.name, "demo_workspace_test")

    def test_bootstrap_raises_when_workspace_exists(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_root = Path(tmp_dir) / "out"
            out_root.mkdir(parents=True, exist_ok=True)
            bootstrap_mod.create_workspace(out_root, timestamp="test")
            with self.assertRaises(FileExistsError):
                bootstrap_mod.create_workspace(out_root, timestamp="test")


class TestDemoUtilities(unittest.TestCase):
    def test_variance_classifier_outputs_expected_fields(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            in_csv = Path(tmp_dir) / "variance_in.csv"
            out_csv = Path(tmp_dir) / "variance_out.csv"
            in_csv.write_text("Item,Actual,Baseline\nRevenue,1200,1000\nCost,900,1000\n", encoding="utf-8")
            rows = variance_mod.classify_csv(
                in_csv,
                out_csv,
                actual_col="Actual",
                baseline_col="Baseline",
                materiality_abs=100,
                materiality_pct=0.05,
            )
            self.assertEqual(rows, 2)
            content = out_csv.read_text(encoding="utf-8")
            self.assertIn("Direction", content)
            self.assertIn("Materiality", content)

    def test_scenario_runner_writes_summaries(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            in_csv = Path(tmp_dir) / "scenario_in.csv"
            out_csv = Path(tmp_dir) / "scenario_out.csv"
            in_csv.write_text("Department,Amount\nSales,1000\nFinance,500\n", encoding="utf-8")

            count = scenario_mod.run_scenarios(
                in_csv,
                out_csv,
                metric_col="Amount",
                scenarios=[("base", 0.0), ("upside", 0.1)],
            )
            self.assertEqual(count, 2)
            content = out_csv.read_text(encoding="utf-8")
            self.assertIn("ScenarioTotal", content)
            self.assertIn("upside", content)

    def test_brief_builder_includes_artifact_links(self):
        payload = brief_mod.build_brief(
            "My Brief",
            "# Executive Summary\n\nHello",
            Path("artifacts/scenario.csv"),
            Path("artifacts/variance.csv"),
        )
        self.assertIn("# My Brief", payload)
        self.assertIn("Scenario output", payload)
        self.assertIn("Variance output", payload)

    def test_pnl_extract_writes_requested_sheet(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_dir = Path(tmp_dir) / "extract"
            workbook = ROOT / "samples/ExcelDemoFile_adv.xlsm"
            outputs = pnl_extract_mod.extract_sheets(workbook, out_dir, ["P&L - Monthly Trend"])
            self.assertEqual(len(outputs), 1)
            self.assertTrue(outputs[0].exists())


if __name__ == "__main__":
    unittest.main()
