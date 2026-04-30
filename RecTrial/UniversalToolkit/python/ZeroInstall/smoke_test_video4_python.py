#!/usr/bin/env python3
"""Smoke test for Video 4 ZeroInstall scripts (5 checks)."""

from __future__ import annotations

import csv
import subprocess
import sys
import tempfile
from pathlib import Path

VERSION = "1.0.0"


def run(cmd: list[str]) -> bool:
    p = subprocess.run(cmd, capture_output=True, text=True)
    return p.returncode == 0


def main() -> None:
    base = Path(__file__).resolve().parent
    py = sys.executable
    results = []

    with tempfile.TemporaryDirectory() as td:
        d = Path(td)
        in_csv = d / "input.csv"
        in_csv.write_text("Department,Amount,Actual,Baseline\nA,100,110,100\nB,200,180,200\n", encoding="utf-8")

        out_clean = d / "clean.csv"
        out_var = d / "var.csv"
        out_scn = d / "scn.csv"
        out_md = d / "summary.md"

        # 1 sanitize
        results.append(run([py, str(base / "sanitize_dataset.py"), str(in_csv), str(out_clean)]) and out_clean.exists())
        # 2 variance
        results.append(run([py, str(base / "variance_classifier.py"), str(in_csv), str(out_var)]) and out_var.exists())
        # 3 scenario
        results.append(run([py, str(base / "scenario_runner.py"), str(in_csv), str(out_scn)]) and out_scn.exists())
        # 4 summary
        results.append(run([py, str(base / "build_exec_summary.py"), str(in_csv), "--out", str(out_md)]) and out_md.exists())
        # 5 help checks for workbook tools
        results.append(
            run([py, str(base / "compare_workbooks.py"), "--help"]) and
            run([py, str(base / "sheets_to_csv.py"), "--help"])
        )

    passed = sum(1 for x in results if x)
    total = len(results)
    print(f"Smoke results: {passed}/{total} PASS")
    if passed != total:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
