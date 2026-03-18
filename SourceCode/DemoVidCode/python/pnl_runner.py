#!/usr/bin/env python3
"""
pnl_runner.py — Unified CLI Entry Point
=========================================

PURPOSE: Single command to run any P&L toolkit operation.
         Replaces the need to remember individual script names.

USAGE:
    python pnl_runner.py dashboard              # Launch Streamlit dashboard
    python pnl_runner.py month-end              # Run month-end close
    python pnl_runner.py month-end --month 3    # Close specific month
    python pnl_runner.py forecast               # Run forecast
    python pnl_runner.py allocate               # Allocation simulator
    python pnl_runner.py allocate --presets     # Run preset scenarios
    python pnl_runner.py snapshot list          # List snapshots
    python pnl_runner.py match                  # AP matching
    python pnl_runner.py test                   # Run pytest suite
    python pnl_runner.py config                 # Show configuration
    python pnl_runner.py --help                 # Show all commands
"""

import os
import sys
import argparse
import subprocess
from typing import List, Optional

# Ensure local imports work
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from pnl_config import APP_NAME, APP_VERSION, FY_LABEL, SOURCE_FILE
except ImportError:
    APP_NAME = "KBT P&L Toolkit"
    APP_VERSION = "2.1.0"
    FY_LABEL = "FY2025"
    SOURCE_FILE = "../DemoFile/ExcelDemoFile_adv.xlsm"


BANNER = f"""
╔══════════════════════════════════════════════════╗
║  {APP_NAME} v{APP_VERSION}                    ║
║  {FY_LABEL} — Unified Command Runner            ║
╚══════════════════════════════════════════════════╝
"""


def cmd_dashboard(args: List[str]) -> int:
    """Launch the Streamlit interactive dashboard."""
    print("Launching Streamlit dashboard...")
    cmd = [sys.executable, "-m", "streamlit", "run", "pnl_dashboard.py"]
    if args:
        cmd.append("--")
        cmd.extend(args)
    return subprocess.call(cmd)


def cmd_month_end(args: List[str]) -> int:
    """Run the month-end close process."""
    from pnl_month_end import MonthEndClose

    parser = argparse.ArgumentParser(prog="pnl_runner month-end")
    parser.add_argument("--file", "-f", default=SOURCE_FILE, help="Source Excel file")
    parser.add_argument("--month", "-m", type=int, default=None, help="Month to close (1-12)")
    parser.add_argument("--export", "-e", action="store_true", help="Export close report")
    parser.add_argument("--output", "-o", default=None, help="Output file path")
    opts = parser.parse_args(args)

    closer = MonthEndClose(file_path=opts.file, month=opts.month)
    report = closer.run(export=opts.export, output_path=opts.output)
    return 0 if report.is_clean else 1


def cmd_forecast(args: List[str]) -> int:
    """Run the forecast engine."""
    from pnl_forecast import PnLForecaster

    parser = argparse.ArgumentParser(prog="pnl_runner forecast")
    parser.add_argument("--file", "-f", default=SOURCE_FILE, help="Source Excel file")
    parser.add_argument("--months", "-m", type=int, default=3, help="Months to forecast")
    parser.add_argument("--method", default="all", choices=["sma", "ets", "trend", "all"])
    parser.add_argument("--export", "-e", default=None, help="Export to Excel file")
    opts = parser.parse_args(args)

    fc = PnLForecaster(file_path=opts.file)
    fc.load()
    fc.run_all(periods=opts.months)
    if opts.export:
        fc.export(opts.export)
    return 0


def cmd_allocate(args: List[str]) -> int:
    """Run the allocation what-if simulator."""
    from pnl_allocation_simulator import AllocationSimulator

    parser = argparse.ArgumentParser(prog="pnl_runner allocate")
    parser.add_argument("--file", "-f", default=SOURCE_FILE, help="Source Excel file")
    parser.add_argument("--scenario", "-s", default=None, help='Overrides: "iGO=0.50,Affirm=0.30"')
    parser.add_argument("--presets", action="store_true", help="Run 3 preset scenarios")
    parser.add_argument("--export", "-e", default=None, help="Export to Excel file")
    opts = parser.parse_args(args)

    sim = AllocationSimulator(file_path=opts.file)
    sim.load()

    if opts.presets:
        scenarios = {
            "InsureSight Growth": {"InsureSight": 0.20, "DocFast": 0.08, "iGO": 0.47, "Affirm": 0.25},
            "iGO Consolidation": {"iGO": 0.65, "Affirm": 0.22, "InsureSight": 0.08, "DocFast": 0.05},
            "Balanced Portfolio": {"iGO": 0.35, "Affirm": 0.30, "InsureSight": 0.20, "DocFast": 0.15},
        }
        for name, shares in scenarios.items():
            result = sim.simulate(revenue_shares=shares, scenario_name=name)
            sim.print_comparison(result)
    elif opts.scenario:
        overrides = {}
        for pair in opts.scenario.split(","):
            k, v = pair.strip().split("=")
            overrides[k.strip()] = float(v.strip())
        result = sim.simulate(revenue_shares=overrides)
        sim.print_comparison(result)
        if opts.export:
            sim.export(result, opts.export)
    else:
        sim._section("BASELINE P&L BY PRODUCT")
        from pnl_config import format_currency
        for _, row in sim.baseline.iterrows():
            sim._print(f"  {row['Product']:15s}  Share: {row['Revenue_Share']:.0%}  "
                        f"Rev: {format_currency(row['Est_Revenue']):>12s}  "
                        f"CM: {format_currency(row['CM_Dollar']):>12s}")
        sim._print("\nUse --scenario or --presets for what-if simulations")
    return 0


def cmd_snapshot(args: List[str]) -> int:
    """Manage P&L snapshots."""
    from pnl_snapshot import SnapshotManager

    parser = argparse.ArgumentParser(prog="pnl_runner snapshot")
    parser.add_argument("action", nargs="?", default="list",
                        choices=["list", "save", "compare", "delete"],
                        help="Snapshot action")
    parser.add_argument("--file", "-f", default=SOURCE_FILE, help="Source Excel file")
    parser.add_argument("--name", "-n", default=None, help="Snapshot name")
    opts = parser.parse_args(args)

    mgr = SnapshotManager(file_path=opts.file)
    if opts.action == "list":
        mgr.list_snapshots()
    elif opts.action == "save":
        name = opts.name or f"snap_{PnLBase.file_timestamp()}"
        mgr.save_snapshot(name)
    elif opts.action == "compare":
        mgr.compare_latest()
    elif opts.action == "delete":
        if opts.name:
            mgr.delete_snapshot(opts.name)
        else:
            print("Error: --name required for delete")
            return 1
    return 0


def cmd_match(args: List[str]) -> int:
    """Run AP matching."""
    from pnl_ap_matcher import APMatcher

    parser = argparse.ArgumentParser(prog="pnl_runner match")
    parser.add_argument("--file", "-f", default=SOURCE_FILE, help="Source Excel file")
    parser.add_argument("--threshold", "-t", type=int, default=80, help="Match threshold (0-100)")
    parser.add_argument("--export", "-e", default=None, help="Export matches to Excel")
    opts = parser.parse_args(args)

    matcher = APMatcher(file_path=opts.file)
    matcher.load()
    results = matcher.run(threshold=opts.threshold)
    if opts.export:
        matcher.export(results, opts.export)
    return 0


def cmd_test(args: List[str]) -> int:
    """Run the pytest test suite."""
    cmd = [sys.executable, "-m", "pytest", "pnl_tests.py", "-v", "--tb=short"]
    cmd.extend(args)
    return subprocess.call(cmd)


def cmd_config(args: List[str]) -> int:
    """Show current configuration."""
    cmd = [sys.executable, "pnl_config.py"]
    return subprocess.call(cmd)


# Command registry
COMMANDS = {
    "dashboard":  ("Launch interactive Streamlit dashboard", cmd_dashboard),
    "month-end":  ("Run month-end close checklist", cmd_month_end),
    "forecast":   ("Run forecast engine", cmd_forecast),
    "allocate":   ("What-if allocation simulator", cmd_allocate),
    "snapshot":   ("Manage P&L snapshots", cmd_snapshot),
    "match":      ("AP fuzzy matching", cmd_match),
    "test":       ("Run automated test suite", cmd_test),
    "config":     ("Show current configuration", cmd_config),
}


def show_help():
    """Print usage information."""
    print(BANNER)
    print("Available commands:\n")
    for cmd_name, (desc, _) in COMMANDS.items():
        print(f"  {cmd_name:14s}  {desc}")
    print()
    print("Usage: python pnl_runner.py <command> [options]")
    print("       python pnl_runner.py <command> --help  (for command-specific help)")
    print()


def main() -> int:
    """Main entry point — dispatch to sub-commands."""
    if len(sys.argv) < 2 or sys.argv[1] in ("--help", "-h", "help"):
        show_help()
        return 0

    cmd_name = sys.argv[1]
    if cmd_name not in COMMANDS:
        print(f"Unknown command: '{cmd_name}'")
        print(f"Available commands: {', '.join(COMMANDS.keys())}")
        print("Run: python pnl_runner.py --help")
        return 1

    _, cmd_func = COMMANDS[cmd_name]
    remaining_args = sys.argv[2:]

    try:
        return cmd_func(remaining_args)
    except KeyboardInterrupt:
        print("\nInterrupted.")
        return 130
    except FileNotFoundError as e:
        print(f"\nFile not found: {e}")
        print("Ensure the Excel file is in the current directory or use --file")
        return 1
    except Exception as e:
        print(f"\nError in '{cmd_name}': {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
