#!/usr/bin/env python3
"""Classify variance rows into finance-friendly categories."""

from __future__ import annotations

import argparse
import csv
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Classify variance rows as material/non-material and favorable/unfavorable.")
    parser.add_argument("input_csv", type=Path, help="Input CSV containing actual and baseline columns")
    parser.add_argument("output_csv", type=Path, help="Output CSV with classification fields")
    parser.add_argument("--actual-col", default="Actual", help="Column name for actual values")
    parser.add_argument("--baseline-col", default="Baseline", help="Column name for baseline/plan values")
    parser.add_argument("--materiality-abs", type=float, default=1000.0, help="Absolute materiality threshold")
    parser.add_argument("--materiality-pct", type=float, default=0.1, help="Relative materiality threshold (decimal)")
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


def classify_row(actual: float | None, baseline: float | None, materiality_abs: float, materiality_pct: float) -> tuple[str, str, float | None, float | None]:
    if actual is None or baseline is None:
        return "unknown", "insufficient_data", None, None

    delta = actual - baseline
    pct = None if baseline == 0 else delta / baseline

    material_abs = abs(delta) >= materiality_abs
    material_rel = pct is not None and abs(pct) >= materiality_pct
    materiality = "material" if (material_abs or material_rel) else "non_material"

    direction = "favorable" if delta >= 0 else "unfavorable"
    return direction, materiality, delta, pct


def classify_csv(
    input_csv: Path,
    output_csv: Path,
    actual_col: str,
    baseline_col: str,
    materiality_abs: float,
    materiality_pct: float,
) -> int:
    with input_csv.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
        if not rows:
            raise SystemExit("Input CSV has no rows")

        fieldnames = list(rows[0].keys())
        extras = ["Variance", "VariancePct", "Direction", "Materiality"]

    with output_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames + extras)
        writer.writeheader()
        count = 0
        for row in rows:
            actual = parse_float(row.get(actual_col))
            baseline = parse_float(row.get(baseline_col))
            direction, materiality, delta, pct = classify_row(actual, baseline, materiality_abs, materiality_pct)
            out = dict(row)
            out["Variance"] = "" if delta is None else f"{delta:.2f}"
            out["VariancePct"] = "" if pct is None else f"{pct:.4f}"
            out["Direction"] = direction
            out["Materiality"] = materiality
            writer.writerow(out)
            count += 1

    return count


def main() -> None:
    args = parse_args()
    count = classify_csv(
        args.input_csv,
        args.output_csv,
        args.actual_col,
        args.baseline_col,
        args.materiality_abs,
        args.materiality_pct,
    )
    print(f"Classified rows: {count}")
    print(f"Output: {args.output_csv}")


if __name__ == "__main__":
    main()
