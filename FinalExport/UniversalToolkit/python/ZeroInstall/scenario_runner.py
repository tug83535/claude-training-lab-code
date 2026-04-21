#!/usr/bin/env python3
"""Run simple what-if scenarios against a baseline metric column."""

from __future__ import annotations

import argparse
import csv
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Apply scenario percentage shocks to a metric column.")
    parser.add_argument("input_csv", type=Path, help="Input CSV path")
    parser.add_argument("output_csv", type=Path, help="Output CSV path")
    parser.add_argument("--metric-col", default="Amount", help="Numeric metric column")
    parser.add_argument(
        "--scenarios",
        default="base:0,optimistic:0.05,conservative:-0.05",
        help="Comma-separated list of label:delta_decimal entries",
    )
    return parser.parse_args()


def parse_float(value: str | None) -> float | None:
    if value is None:
        return None
    cleaned = str(value).replace(",", "").replace("$", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def parse_scenarios(text: str) -> list[tuple[str, float]]:
    scenarios: list[tuple[str, float]] = []
    for item in text.split(","):
        label, sep, raw_delta = item.partition(":")
        if not sep:
            raise SystemExit(f"Invalid scenario entry: {item}")
        try:
            delta = float(raw_delta)
        except ValueError as exc:
            raise SystemExit(f"Invalid scenario delta in entry: {item}") from exc
        scenarios.append((label.strip(), delta))
    if not scenarios:
        raise SystemExit("No scenarios provided")
    return scenarios


def run_scenarios(input_csv: Path, output_csv: Path, metric_col: str, scenarios: list[tuple[str, float]]) -> int:
    with input_csv.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    with output_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Scenario", "Rows", "BaseTotal", "ScenarioTotal", "Delta"])
        count = 0
        for label, pct in scenarios:
            base_total = 0.0
            scenario_total = 0.0
            metric_rows = 0

            for row in rows:
                value = parse_float(row.get(metric_col))
                if value is None:
                    continue
                metric_rows += 1
                base_total += value
                scenario_total += value * (1 + pct)

            writer.writerow(
                [
                    label,
                    metric_rows,
                    f"{base_total:.2f}",
                    f"{scenario_total:.2f}",
                    f"{(scenario_total - base_total):.2f}",
                ]
            )
            count += 1

    return count


def main() -> None:
    args = parse_args()
    scenarios = parse_scenarios(args.scenarios)
    count = run_scenarios(args.input_csv, args.output_csv, args.metric_col, scenarios)
    print(f"Scenarios evaluated: {count}")
    print(f"Output: {args.output_csv}")


if __name__ == "__main__":
    main()
