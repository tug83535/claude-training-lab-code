from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd


def _load_thresholds(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _status_for_value(value: float, rule: dict[str, Any]) -> str:
    direction = rule.get("direction", "higher_is_better")
    green = float(rule.get("green", 0.0))
    yellow = float(rule.get("yellow", 0.0))

    if direction == "lower_is_better":
        if value <= green:
            return "GREEN"
        if value <= yellow:
            return "YELLOW"
        return "RED"

    if value >= green:
        return "GREEN"
    if value >= yellow:
        return "YELLOW"
    return "RED"


def build_pulse(df: pd.DataFrame, thresholds: dict[str, Any]) -> pd.DataFrame:
    required = ["kpi", "value"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required KPI columns: {missing}")

    rules = thresholds.get("kpis", {})
    records = []

    for _, row in df.iterrows():
        kpi = str(row["kpi"])
        value = float(row["value"])
        rule = rules.get(kpi, {"direction": "higher_is_better", "green": value, "yellow": value})
        status = _status_for_value(value, rule)
        records.append(
            {
                "kpi": kpi,
                "value": value,
                "status": status,
                "direction": rule.get("direction", "higher_is_better"),
                "green_threshold": rule.get("green", value),
                "yellow_threshold": rule.get("yellow", value),
            }
        )

    out = pd.DataFrame(records)
    order = {"RED": 0, "YELLOW": 1, "GREEN": 2}
    out["_sort"] = out["status"].map(order)
    out = out.sort_values(["_sort", "kpi"]).drop(columns=["_sort"]).reset_index(drop=True)
    return out


def run(input_csv: Path, thresholds_json: Path, output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    df = pd.read_csv(input_csv)
    thresholds = _load_thresholds(thresholds_json)
    pulse = build_pulse(df, thresholds)

    json_out = output_dir / "cfo_pulse_report.json"
    md_out = output_dir / "cfo_pulse_report.md"

    json_out.write_text(pulse.to_json(orient="records", indent=2), encoding="utf-8")

    lines = ["# CFO One-Page Pulse Report", "", "| KPI | Value | Status |", "|---|---:|---|"]
    for _, r in pulse.iterrows():
        lines.append(f"| {r['kpi']} | {r['value']:.4f} | {r['status']} |")

    md_out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return json_out, md_out
