#!/usr/bin/env python3
"""
pnl_cli.py — Master CLI Orchestrator
======================================

PURPOSE: Single command-line interface for the entire P&L toolkit.
         Replaces running 10 separate python commands.

USAGE:
    python pnl_cli.py status                       # Show toolkit status
    python pnl_cli.py validate                     # Run data validation
    python pnl_cli.py close --month 3              # Month-end close
    python pnl_cli.py report                       # Generate executive report
    python pnl_cli.py forecast --months 6          # Run forecast
    python pnl_cli.py simulate --presets           # What-if scenarios
    python pnl_cli.py dashboard                    # Launch Streamlit dashboard
    python pnl_cli.py run-all                      # Full pipeline
    python pnl_cli.py compare --old v1.xlsx --new v2.xlsx
"""

import os
import sys
import argparse
import time
from datetime import datetime

try:
    from pnl_config import *
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pnl_config import *


def cmd_status(args):
    """Show toolkit status and configuration."""
    print(f"\n{'='*55}")
    print(f"  {APP_NAME} v{APP_VERSION} — Status")
    print(f"{'='*55}")

    # File check
    fp = args.file
    exists = os.path.exists(fp)
    size = os.path.getsize(fp) if exists else 0
    print(f"\n  Source File:    {fp}")
    print(f"  File Exists:    {'✓ Yes' if exists else '✗ No'}")
    if exists:
        print(f"  File Size:      {size / 1024:.0f} KB")
        mod_time = datetime.fromtimestamp(os.path.getmtime(fp))
        print(f"  Last Modified:  {mod_time.strftime('%Y-%m-%d %H:%M')}")

    # DB check
    db_exists = os.path.exists(DB_PATH)
    print(f"\n  Database:       {DB_PATH}")
    print(f"  DB Exists:      {'✓ Yes' if db_exists else '— Not yet created'}")

    # Config
    print(f"\n  Fiscal Year:    {FY_LABEL}")
    print(f"  Products:       {', '.join(PRODUCTS)}")
    print(f"  Departments:    {len(DEPARTMENTS)}")
    print(f"  Variance Thr:   {VARIANCE_PCT:.0%}")

    # Module availability
    print(f"\n  Module Availability:")
    modules = {
        "pnl_config": "pnl_config",
        "pnl_data_loader": "pnl_data_loader",
        "pnl_toolkit": "pnl_toolkit",
        "pnl_validation": "pnl_validation",
        "pnl_analysis": "pnl_analysis",
        "pnl_report_generator": "pnl_report_generator",
        "pnl_database_export": "pnl_database_export",
        "pnl_month_end": "pnl_month_end",
        "pnl_dashboard": "pnl_dashboard",
        "pnl_allocation_simulator": "pnl_allocation_simulator",
        "pnl_forecast": "pnl_forecast",
        "pnl_email_report": "pnl_email_report",
        "pnl_ap_matcher": "pnl_ap_matcher",
        "pnl_snapshot": "pnl_snapshot",
        "advanced_analytics": "advanced_analytics",
        "chart_generator": "chart_generator",
        "excel_comparison_tool": "excel_comparison_tool",
    }
    for display, mod in modules.items():
        try:
            __import__(mod)
            print(f"    ✓ {display}")
        except ImportError:
            print(f"    — {display} (not found)")

    # Dependency check
    print(f"\n  Dependencies:")
    deps = ["pandas", "numpy", "openpyxl", "matplotlib", "streamlit", "plotly",
            "statsmodels", "sklearn", "thefuzz", "click"]
    for dep in deps:
        try:
            __import__(dep)
            print(f"    ✓ {dep}")
        except ImportError:
            print(f"    — {dep} (not installed)")

    print()


def cmd_validate(args):
    """Run data validation."""
    try:
        from pnl_validation import PnLValidator
    except ImportError:
        try:
            from pnl_validation__1_ import PnLValidator
        except ImportError:
            print("Error: pnl_validation module not found")
            return
    validator = PnLValidator(args.file, strict=args.strict)
    validator.run_all(export=args.export is not None,
                      output_path=args.export or "validation_report.xlsx")


def cmd_close(args):
    """Run month-end close."""
    from pnl_month_end import MonthEndClose
    closer = MonthEndClose(file_path=args.file, month=args.month)
    closer.run(export=args.export, output_path=args.output)


def cmd_report(args):
    """Generate executive report."""
    try:
        from pnl_report_generator import ReportGenerator
        gen = ReportGenerator(args.file, output_path=args.output or "pnl_executive_report.xlsx")
        gen.generate()
    except ImportError:
        print("Error: pnl_report_generator module not found")


def cmd_forecast(args):
    """Run forecast."""
    from pnl_forecast import PnLForecaster
    fc = PnLForecaster(file_path=args.file)
    fc.load()
    results = fc.forecast(periods=args.months, method=args.method,
                         product=args.product, department=args.department)
    if args.export:
        fc.export(results, args.export)


def cmd_simulate(args):
    """Run allocation simulator."""
    from pnl_allocation_simulator import AllocationSimulator
    sim = AllocationSimulator(file_path=args.file)
    sim.load()

    if args.presets:
        scenarios = {
            "InsureSight Growth": {"InsureSight": 0.20, "DocFast": 0.08, "iGO": 0.47, "Affirm": 0.25},
            "iGO Consolidation":  {"iGO": 0.65, "Affirm": 0.22, "InsureSight": 0.08, "DocFast": 0.05},
            "Balanced Portfolio":  {"iGO": 0.35, "Affirm": 0.30, "InsureSight": 0.20, "DocFast": 0.15},
        }
        for name, shares in scenarios.items():
            result = sim.simulate(revenue_shares=shares, scenario_name=name)
            sim.print_comparison(result)
    elif args.scenario:
        overrides = {}
        for pair in args.scenario.split(","):
            k, v = pair.strip().split("=")
            overrides[k.strip()] = float(v.strip())
        result = sim.simulate(revenue_shares=overrides)
        sim.print_comparison(result)
    else:
        sim._section("BASELINE — Use --scenario or --presets")
        for _, row in sim.baseline.iterrows():
            sim._print(f"  {row['Product']:15s}  Share: {row['Revenue_Share']:.0%}  CM%: {row['CM_Pct']:.1%}")


def cmd_dashboard(args):
    """Launch Streamlit dashboard."""
    import subprocess
    dashboard_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pnl_dashboard.py")
    if not os.path.exists(dashboard_path):
        print(f"Error: {dashboard_path} not found")
        return
    cmd = ["streamlit", "run", dashboard_path, "--", "--file", args.file]
    print(f"Launching dashboard: {' '.join(cmd)}")
    subprocess.run(cmd)


def cmd_email(args):
    """Send email report."""
    from pnl_email_report import EmailReporter
    recipients = args.to.split(",") if args.to else None
    reporter = EmailReporter(file_path=args.file)
    reporter.generate_and_send(
        recipients=recipients,
        attachments=args.attach,
        preview=args.preview
    )


def cmd_compare(args):
    """Compare two Excel files."""
    try:
        from excel_comparison_tool import ExcelCompare
        cmp = ExcelCompare(args.old, args.new)
        cmp.compare_all()
        if args.export:
            cmp.export_report(args.export)
    except ImportError:
        print("Error: excel_comparison_tool module not found")


def cmd_charts(args):
    """Generate charts."""
    try:
        from chart_generator import ChartGenerator
        gen = ChartGenerator(args.file, output_dir=args.output_dir or "./charts")
        gen.generate_all()
    except ImportError:
        print("Error: chart_generator module not found")


def cmd_run_all(args):
    """Run the full pipeline."""
    print(f"\n{'='*55}")
    print(f"  {APP_NAME} — FULL PIPELINE")
    print(f"{'='*55}")
    start = time.time()

    steps = [
        ("Validating data...", lambda: cmd_validate(args)),
        ("Running month-end close...", lambda: cmd_close(args)),
        ("Generating forecast...", lambda: cmd_forecast(args)),
        ("Generating report...", lambda: cmd_report(args)),
    ]

    for desc, fn in steps:
        print(f"\n▶ {desc}")
        try:
            fn()
            print(f"  ✓ Complete")
        except Exception as e:
            print(f"  ✗ Error: {e}")

    elapsed = time.time() - start
    print(f"\n{'='*55}")
    print(f"  Pipeline complete in {elapsed:.1f}s")
    print(f"{'='*55}\n")


def main():
    parser = argparse.ArgumentParser(
        prog="pnl",
        description=f"{APP_NAME} v{APP_VERSION} — Master CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python pnl_cli.py status
  python pnl_cli.py validate --strict
  python pnl_cli.py close --month 3 --export
  python pnl_cli.py forecast --months 6 --method ets
  python pnl_cli.py simulate --presets
  python pnl_cli.py dashboard
  python pnl_cli.py run-all
"""
    )
    parser.add_argument("--file", "-f", default=SOURCE_FILE, help="Source Excel file")

    sub = parser.add_subparsers(dest="command", help="Command to run")

    # status
    sub.add_parser("status", help="Show toolkit status")

    # validate
    p_val = sub.add_parser("validate", help="Run data validation")
    p_val.add_argument("--strict", action="store_true")
    p_val.add_argument("--export", default=None)

    # close
    p_close = sub.add_parser("close", help="Month-end close")
    p_close.add_argument("--month", "-m", type=int, default=None)
    p_close.add_argument("--export", action="store_true")
    p_close.add_argument("--output", "-o", default=None)

    # report
    p_rpt = sub.add_parser("report", help="Generate executive report")
    p_rpt.add_argument("--output", "-o", default=None)

    # forecast
    p_fc = sub.add_parser("forecast", help="Run forecast")
    p_fc.add_argument("--months", "-m", type=int, default=3)
    p_fc.add_argument("--method", default="ets", choices=PnLForecaster.METHODS if 'PnLForecaster' in dir() else ["sma","ets","trend","scenario"])
    p_fc.add_argument("--product", "-p", default=None)
    p_fc.add_argument("--department", "-d", default=None)
    p_fc.add_argument("--export", "-e", default=None)

    # simulate
    p_sim = sub.add_parser("simulate", help="Allocation what-if")
    p_sim.add_argument("--scenario", "-s", default=None)
    p_sim.add_argument("--presets", action="store_true")

    # dashboard
    sub.add_parser("dashboard", help="Launch Streamlit dashboard")

    # email
    p_email = sub.add_parser("email", help="Send email report")
    p_email.add_argument("--to", "-t", default=None)
    p_email.add_argument("--attach", "-a", nargs="*")
    p_email.add_argument("--preview", action="store_true")

    # compare
    p_cmp = sub.add_parser("compare", help="Compare two Excel files")
    p_cmp.add_argument("--old", required=True)
    p_cmp.add_argument("--new", required=True)
    p_cmp.add_argument("--export", default=None)

    # charts
    p_cht = sub.add_parser("charts", help="Generate charts")
    p_cht.add_argument("--output-dir", default=None)

    # run-all
    sub.add_parser("run-all", help="Run full pipeline")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    commands = {
        "status": cmd_status,
        "validate": cmd_validate,
        "close": cmd_close,
        "report": cmd_report,
        "forecast": cmd_forecast,
        "simulate": cmd_simulate,
        "dashboard": cmd_dashboard,
        "email": cmd_email,
        "compare": cmd_compare,
        "charts": cmd_charts,
        "run-all": cmd_run_all,
    }

    fn = commands.get(args.command)
    if fn:
        fn(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
