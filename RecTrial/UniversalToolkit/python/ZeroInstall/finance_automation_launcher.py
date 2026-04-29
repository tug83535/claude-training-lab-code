# Finance Automation Toolkit v1.0 — iPipeline
# finance_automation_launcher.py — numbered menu for all Finance Automation tools
#
# You don't have to remember command-line arguments. Pick a number and press Enter.
# Each tool runs in sample mode by default so you can see results before using real files.
#
# Usage:
#   python finance_automation_launcher.py
#
# Then pick a number from the menu.

import subprocess
import sys
from pathlib import Path

_ROOT = Path(__file__).parent

_TOOLS = [
    {
        "number":      1,
        "label":       "Revenue Leakage Finder",
        "description": "Find billing gaps, stale contracts, and data quality anomalies",
        "script":      "revenue_leakage_finder.py",
        "sample_flag": "--sample",
    },
    {
        "number":      2,
        "label":       "Data Contract Checker",
        "description": "Validate a CSV file structure before analysis begins",
        "script":      "data_contract_checker.py",
        "sample_flag": "--sample",
    },
    {
        "number":      3,
        "label":       "Exception Triage Engine",
        "description": "Score and rank exceptions — know what to review first",
        "script":      "exception_triage_engine.py",
        "sample_flag": "--sample",
    },
    {
        "number":      4,
        "label":       "Control Evidence Pack",
        "description": "Create a tamper-evident evidence bundle from any analysis run",
        "script":      "control_evidence_pack.py",
        "sample_flag": "--sample",
    },
    {
        "number":      5,
        "label":       "Workbook Dependency Scanner",
        "description": "Map cross-sheet formula references inside any .xlsx file",
        "script":      "workbook_dependency_scanner.py",
        "sample_flag": "--sample",
    },
]

_SAFETY_RULES = """
Finance Automation Toolkit — Safety Rules
==========================================

1.  Your input files are NEVER modified. All scripts open files read-only.
2.  All outputs go to the /outputs/ folder. Nothing is written back to your file.
3.  Each run creates a new timestamped subfolder. Previous runs are never overwritten.
4.  Start with sample mode. Run a tool in sample mode first before using your real files.
5.  Do not run these tools on files containing Social Security Numbers, passwords,
    card numbers, or other regulated personal data.
6.  Do not run these tools on files marked CONFIDENTIAL or RESTRICTED unless you
    are the file owner or have explicit approval.
7.  Revenue, billing, and contract data for real customers is sensitive.
    Keep outputs in the /outputs/ folder and do not share them publicly.
8.  Never paste tool output into a public chat, shared document, or ticket
    visible outside the Finance & Accounting team.
9.  Python scripts bundled here use no network connections. No data leaves your machine.
10. Scripts use only Python standard library (no pip install required).
11. If a script fails with an error, stop and contact Connor before retrying.
12. Do not modify the scripts unless you know what you are doing.
13. SHA-256 hashes in the Control Evidence Pack are read-only fingerprints.
    Do not use them as a substitute for your company's official audit controls.
14. Questions, bugs, or something unexpected? Contact Connor Atlee — Finance & Accounting.
"""


def _print_header() -> None:
    print()
    print("=" * 60)
    print("  Finance Automation Toolkit v1.0  |  iPipeline")
    print("=" * 60)
    print()


def _print_menu() -> None:
    print("  Available tools:")
    print()
    for t in _TOOLS:
        print(f"  {t['number']}.  {t['label']}")
        print(f"       {t['description']}")
        print()
    print("  6.  Show safety rules")
    print("  7.  Open output folder in Explorer")
    print("  8.  Exit")
    print()


def _run_tool(tool: dict) -> None:
    script = _ROOT / tool["script"]
    if not script.exists():
        print(f"\n  ERROR: Script not found: {script}")
        return

    print(f"\n  Running: {tool['label']} (sample mode)...")
    print(f"  Command: python {tool['script']} {tool['sample_flag']}")
    print()
    print("-" * 60)

    result = subprocess.run(
        [sys.executable, str(script), tool["sample_flag"]],
        cwd=str(_ROOT),
    )

    print("-" * 60)
    if result.returncode == 0:
        print(f"\n  Done. Check the outputs/ folder for your report.")
    else:
        print(f"\n  Tool exited with errors (code {result.returncode}).")
        print("  Check the message above. Contact Connor if you need help.")


def _open_outputs() -> None:
    outputs_dir = _ROOT / "outputs"
    outputs_dir.mkdir(exist_ok=True)
    try:
        import os
        os.startfile(str(outputs_dir))
        print(f"\n  Opened: {outputs_dir}")
    except Exception:
        print(f"\n  Output folder is at: {outputs_dir}")


def main() -> None:
    _print_header()

    while True:
        _print_menu()
        try:
            raw = input("  Enter a number (1-8): ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n\n  Exiting. Goodbye.")
            break

        if not raw.isdigit():
            print("\n  Please enter a number.\n")
            continue

        choice = int(raw)

        if choice in range(1, len(_TOOLS) + 1):
            tool = next(t for t in _TOOLS if t["number"] == choice)
            _run_tool(tool)
            input("\n  Press Enter to return to the menu...")
            _print_header()

        elif choice == 6:
            print(_SAFETY_RULES)
            input("  Press Enter to return to the menu...")
            _print_header()

        elif choice == 7:
            _open_outputs()
            input("\n  Press Enter to return to the menu...")
            _print_header()

        elif choice == 8:
            print("\n  Goodbye.")
            break

        else:
            print(f"\n  '{choice}' is not a valid option. Enter a number from 1 to 8.\n")


if __name__ == "__main__":
    main()
