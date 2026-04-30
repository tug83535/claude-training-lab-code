#!/usr/bin/env python3
"""Build a plain-English executive summary from tabular CSV input."""

from __future__ import annotations

VERSION = "1.0.0"

import argparse
import csv
from pathlib import Path
from statistics import mean


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a markdown executive summary from CSV data.")
    parser.add_argument("input_csv", type=Path, help="Input CSV file")
    parser.add_argument("--metric-column", default="Amount", help="Preferred numeric metric column name")
    parser.add_argument("--group-column", default="Department", help="Preferred label/group column name")
    parser.add_argument("--out", type=Path, help="Optional markdown output path")
    return parser.parse_args()


def read_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def parse_float(value: str) -> float | None:
    if value is None:
        return None
    cleaned = value.replace(",", "").replace("$", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def pick_numeric_column(rows: list[dict[str, str]], preferred: str) -> str | None:
    if not rows:
        return None

    headers = list(rows[0].keys())
    if preferred in headers:
        return preferred

    for header in headers:
        if sum(1 for row in rows if parse_float(row.get(header, "")) is not None) >= max(3, len(rows) // 3):
            return header
    return None


def build_summary(rows: list[dict[str, str]], metric_col: str, group_col: str) -> str:
    metrics = [parse_float(row.get(metric_col, "")) for row in rows]
    values = [v for v in metrics if v is not None]

    if not values:
        return "No numeric data was detected for summary generation."

    total = sum(values)
    avg = mean(values)
    maximum = max(values)
    minimum = min(values)

    group_totals: dict[str, float] = {}
    for row in rows:
        label = (row.get(group_col) or "Unspecified").strip() or "Unspecified"
        val = parse_float(row.get(metric_col, ""))
        if val is None:
            continue
        group_totals[label] = group_totals.get(label, 0.0) + val

    top_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)[:3]

    lines = [
        "# Executive Summary",
        "",
        f"- Rows analyzed: **{len(rows)}**",
        f"- Metric used: **{metric_col}**",
        f"- Total: **${total:,.0f}**",
        f"- Average per row: **${avg:,.0f}**",
        f"- Range: **${minimum:,.0f} to ${maximum:,.0f}**",
        "",
        "## Top contributing groups",
    ]

    if top_groups:
        for idx, (name, value) in enumerate(top_groups, start=1):
            lines.append(f"{idx}. {name}: ${value:,.0f}")
    else:
        lines.append("No group-level totals were available.")

    lines.extend(
        [
            "",
            "## Suggested talking points",
            "- Review the top contributing groups for concentration risk or one-time effects.",
            "- Validate whether unusual low values reflect timing or data quality issues.",
            "- Use this summary as a starting point for variance commentary and next-step actions.",
        ]
    )

    return "\n".join(lines)



def require_existing_file(path: Path, label: str) -> None:
    if path is None:
        raise SystemExit(f"Error: missing {label} path.")
    if not path.exists():
        raise SystemExit(f"Error: {label} file was not found: {path}")


def main() -> None:
    args = parse_args()
    require_existing_file(args.input_csv, "input CSV")
    rows = read_rows(args.input_csv)
    metric_col = pick_numeric_column(rows, args.metric_column)

    if metric_col is None:
        raise SystemExit("No numeric column detected. Provide a CSV with at least one numeric measure.")

    summary = build_summary(rows, metric_col, args.group_column)

    if args.out:
        args.out.write_text(summary, encoding="utf-8")
        print(f"Summary written to: {args.out}")
    else:
        print(summary)


if __name__ == "__main__":
    main()
