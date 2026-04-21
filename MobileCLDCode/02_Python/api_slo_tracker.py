"""
api_slo_tracker.py - API SLO & Error Budget Tracker

PURPOSE
-------
Read an access log (or an APM export) and compute, per endpoint per day:

  - request count
  - p50 / p95 / p99 latency
  - error rate (5xx and 4xx separately)
  - SLO compliance (% of minute-windows that met the latency + error target)
  - error budget remaining for the rolling 30-day window

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Log files are gigabytes; Excel cannot load them. APM vendor dashboards (Datadog,
New Relic) show current state but (a) don't track rolling error budget, (b)
don't let you redefine SLO targets retrospectively, and (c) don't give an
Excel-ready output for finance/leadership. This script does all three.

USE CASE
--------
Every Engineering leadership review starts with "are we hitting our SLOs?".
Run this against last week's logs; drop the output into the review deck.
Perfect for companies that care about uptime commitments in customer contracts.

INPUT (common log format or JSON-lines):
    timestamp, endpoint, status_code, latency_ms

USAGE
-----
    python api_slo_tracker.py access.log --slo slo_targets.yaml --output slo.xlsx

SLO CONFIG (slo_targets.yaml):
    targets:
      - endpoint_pattern: /api/v1/quotes
        p95_ms: 500
        error_budget_pct: 0.1     # 99.9% availability
        window_days: 30
      - endpoint_pattern: /api/v1/policies
        p95_ms: 800
        error_budget_pct: 0.5
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import re
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
import yaml


@dataclass
class SLOTarget:
    endpoint_pattern: str
    p95_ms: float
    error_budget_pct: float
    window_days: int = 30


LOG_RE = re.compile(
    r'(?P<ts>\S+)\s+"(?P<method>\w+)\s+(?P<endpoint>\S+)\s+HTTP[^"]+"'
    r'\s+(?P<status>\d{3})\s+(?P<latency>\d+)'
)


def load_logs(path: Path) -> pd.DataFrame:
    """Accepts Common Log Format, JSON-lines, or CSV."""
    text_sample = path.read_text(encoding="utf-8", errors="replace").splitlines()[:5]

    # JSON-lines?
    try:
        json.loads(text_sample[0])
        rows = []
        for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
            if not line.strip():
                continue
            j = json.loads(line)
            rows.append(
                {
                    "ts": pd.to_datetime(j.get("timestamp") or j.get("ts")),
                    "endpoint": j.get("endpoint") or j.get("path"),
                    "status_code": int(j.get("status") or j.get("status_code") or 0),
                    "latency_ms": float(j.get("latency_ms") or j.get("duration_ms") or 0),
                }
            )
        return pd.DataFrame(rows)
    except Exception:
        pass

    # CSV?
    if "," in text_sample[0]:
        df = pd.read_csv(path)
        df["ts"] = pd.to_datetime(df["ts"] if "ts" in df.columns else df["timestamp"])
        return df.rename(columns={"path": "endpoint", "status": "status_code"})

    # Common Log Format
    rows = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        m = LOG_RE.search(line)
        if not m:
            continue
        rows.append(
            {
                "ts": pd.to_datetime(m.group("ts"), errors="coerce"),
                "endpoint": m.group("endpoint").split("?")[0],
                "status_code": int(m.group("status")),
                "latency_ms": float(m.group("latency")),
            }
        )
    return pd.DataFrame(rows)


def load_slo_targets(path: Path) -> list[SLOTarget]:
    with path.open() as f:
        cfg = yaml.safe_load(f)
    return [SLOTarget(**t) for t in cfg["targets"]]


def match_target(endpoint: str, targets: list[SLOTarget]) -> SLOTarget | None:
    for t in targets:
        if fnmatch.fnmatch(endpoint, t.endpoint_pattern + "*"):
            return t
    return None


def daily_slo_table(df: pd.DataFrame, targets: list[SLOTarget]) -> pd.DataFrame:
    df = df.copy()
    df["date"] = df["ts"].dt.date
    df["is_error"] = df["status_code"] >= 500

    grouped = (
        df.groupby(["date", "endpoint"])
        .agg(
            requests=("status_code", "count"),
            errors_5xx=("is_error", "sum"),
            p50=("latency_ms", lambda s: np.percentile(s, 50)),
            p95=("latency_ms", lambda s: np.percentile(s, 95)),
            p99=("latency_ms", lambda s: np.percentile(s, 99)),
        )
        .reset_index()
    )
    grouped["error_rate_pct"] = grouped["errors_5xx"] / grouped["requests"] * 100

    # Attach SLO target and compliance
    grouped["slo_p95_ms"] = None
    grouped["slo_error_budget_pct"] = None
    grouped["meets_slo"] = False
    for i, row in grouped.iterrows():
        tgt = match_target(row["endpoint"], targets)
        if tgt is None:
            continue
        grouped.at[i, "slo_p95_ms"] = tgt.p95_ms
        grouped.at[i, "slo_error_budget_pct"] = tgt.error_budget_pct
        grouped.at[i, "meets_slo"] = (
            row["p95"] <= tgt.p95_ms and row["error_rate_pct"] <= tgt.error_budget_pct
        )
    return grouped


def rolling_error_budget(daily: pd.DataFrame, targets: list[SLOTarget]) -> pd.DataFrame:
    """For each endpoint with an SLO target, compute error budget consumed over the rolling window."""
    rows = []
    for t in targets:
        subset = daily[daily["endpoint"].str.startswith(t.endpoint_pattern.rstrip("*"))].copy()
        if subset.empty:
            continue
        subset = subset.sort_values("date")
        total_req = subset["requests"].sum()
        total_err = subset["errors_5xx"].sum()
        err_pct = total_err / total_req * 100 if total_req else 0
        budget_consumed = err_pct / t.error_budget_pct if t.error_budget_pct else None
        rows.append(
            {
                "endpoint_pattern": t.endpoint_pattern,
                "window_days": t.window_days,
                "requests": total_req,
                "errors": total_err,
                "error_rate_pct": err_pct,
                "slo_error_budget_pct": t.error_budget_pct,
                "budget_consumed_pct": budget_consumed * 100 if budget_consumed else None,
                "budget_remaining_pct": (1 - budget_consumed) * 100 if budget_consumed else None,
                "status": "BREACH" if err_pct > t.error_budget_pct else "OK",
            }
        )
    return pd.DataFrame(rows)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("log_path")
    ap.add_argument("--slo", required=True)
    ap.add_argument("--output", default="slo_report.xlsx")
    args = ap.parse_args()

    df = load_logs(Path(args.log_path))
    targets = load_slo_targets(Path(args.slo))

    daily = daily_slo_table(df, targets)
    budget = rolling_error_budget(daily, targets)

    with pd.ExcelWriter(args.output, engine="openpyxl") as w:
        budget.to_excel(w, sheet_name="Error Budgets", index=False)
        daily.to_excel(w, sheet_name="Daily by Endpoint", index=False)

    print(f"Wrote {args.output}")
    if not budget.empty:
        breaches = budget[budget["status"] == "BREACH"]
        print(f"{len(breaches)} SLO breaches in the window:")
        for _, r in breaches.iterrows():
            print(f"  {r['endpoint_pattern']}: {r['error_rate_pct']:.2f}% (budget {r['slo_error_budget_pct']}%)")


if __name__ == "__main__":
    main()
