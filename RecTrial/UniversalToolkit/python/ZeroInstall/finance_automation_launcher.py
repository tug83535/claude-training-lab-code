#!/usr/bin/env python3
"""Menu launcher for Video 4 ZeroInstall finance scripts."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

VERSION = "1.0.0"

SCRIPTS = {
    "1": ("sanitize_dataset.py", "Sanitize messy CSV data for clean downstream use."),
    "2": ("variance_classifier.py", "Classify variance rows into material/favorable buckets."),
    "3": ("scenario_runner.py", "Run base/optimistic/conservative scenario calculations."),
    "4": ("build_exec_summary.py", "Generate an executive summary markdown from CSV."),
    "5": ("compare_workbooks.py", "Compare two workbooks and export cell-level diffs."),
    "6": ("sheets_to_csv.py", "Extract selected workbook sheets into CSV files."),
}


def print_menu() -> None:
    print(f"Finance Automation Launcher v{VERSION}")
    for key, (script, desc) in SCRIPTS.items():
        print(f"{key}. {script} - {desc}")
    print("Q. Quit")


def main() -> None:
    base = Path(__file__).resolve().parent
    while True:
        print_menu()
        choice = input("Select a menu number and press Enter: ").strip().upper()
        if choice == "Q":
            print("Exiting launcher.")
            return
        if choice not in SCRIPTS:
            print("Invalid selection. Choose 1-6 or Q.")
            continue

        script = base / SCRIPTS[choice][0]
        print(f"Launching: {script.name}")
        print("Tip: run with --help to see required inputs.")
        subprocess.run([sys.executable, str(script), "--help"], check=False)
        print("Run the script manually with your file paths after reviewing help.\n")


if __name__ == "__main__":
    main()
