from __future__ import annotations

import argparse
from pathlib import Path
import time

try:
    from .tools import cfo_pulse_report, control_evidence_pack, data_contract_checker, exception_triage_engine
    from .tools.telemetry_logger import log_event
except ImportError:
    from tools import cfo_pulse_report, control_evidence_pack, data_contract_checker, exception_triage_engine
    from tools.telemetry_logger import log_event


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="finance_copilot",
        description="Finance Copilot CLI - production-ready finance automation starter",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    dc = subparsers.add_parser("data-contract", help="Run data contract validation on CSV input")
    dc.add_argument("--input", required=True, help="Path to input CSV")
    dc.add_argument("--contract", required=True, help="Path to data contract JSON")
    dc.add_argument("--output-dir", required=True, help="Output folder")

    tri = subparsers.add_parser("triage", help="Rank exception records by business priority")
    tri.add_argument("--input", required=True, help="Path to input CSV")
    tri.add_argument("--weights", required=True, help="Path to triage weights JSON")
    tri.add_argument("--output-dir", required=True, help="Output folder")
    tri.add_argument("--top-n", type=int, default=20, help="Top N rows to export")

    ev = subparsers.add_parser("evidence-pack", help="Create control evidence zip with hashes")
    ev.add_argument("--input-dir", required=True, help="Folder containing evidence files")
    ev.add_argument("--output-dir", required=True, help="Output folder")
    ev.add_argument("--pack-name", default="control_evidence_pack", help="Prefix for output zip")

    pulse = subparsers.add_parser("cfo-pulse", help="Build CFO one-page pulse KPI report")
    pulse.add_argument("--input", required=True, help="Path to KPI CSV with columns: kpi,value")
    pulse.add_argument("--thresholds", required=True, help="Path to KPI threshold JSON")
    pulse.add_argument("--output-dir", required=True, help="Output folder")

    parser.add_argument("--telemetry", default="./output/tool_usage.csv", help="Path to telemetry CSV log")

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    telemetry_path = Path(args.telemetry)

    start = time.perf_counter()
    try:
        if args.command == "data-contract":
            out = data_contract_checker.run(
                input_csv=Path(args.input),
                contract_json=Path(args.contract),
                output_dir=Path(args.output_dir),
            )
            print(f"[OK] Data contract report created: {out}")
            log_event(telemetry_path, "data-contract", "SUCCESS", int((time.perf_counter()-start)*1000), str(out))
            return 0

        if args.command == "triage":
            out = exception_triage_engine.run(
                input_csv=Path(args.input),
                weights_json=Path(args.weights),
                output_dir=Path(args.output_dir),
                top_n=args.top_n,
            )
            print(f"[OK] Exception triage output created: {out}")
            log_event(telemetry_path, "triage", "SUCCESS", int((time.perf_counter()-start)*1000), str(out))
            return 0

        if args.command == "evidence-pack":
            out = control_evidence_pack.run(
                input_dir=Path(args.input_dir),
                output_dir=Path(args.output_dir),
                pack_name=args.pack_name,
            )
            print(f"[OK] Evidence pack created: {out}")
            log_event(telemetry_path, "evidence-pack", "SUCCESS", int((time.perf_counter()-start)*1000), str(out))
            return 0

        if args.command == "cfo-pulse":
            j, m = cfo_pulse_report.run(
                input_csv=Path(args.input),
                thresholds_json=Path(args.thresholds),
                output_dir=Path(args.output_dir),
            )
            print(f"[OK] CFO pulse report created: {j} and {m}")
            log_event(telemetry_path, "cfo-pulse", "SUCCESS", int((time.perf_counter()-start)*1000), f"{j};{m}")
            return 0

        parser.error("Unknown command")
        return 2
    except Exception as exc:
        log_event(telemetry_path, args.command or "unknown", "FAIL", int((time.perf_counter()-start)*1000), error_message=str(exc))
        raise


if __name__ == "__main__":
    raise SystemExit(main())
