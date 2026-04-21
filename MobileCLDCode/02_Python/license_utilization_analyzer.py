"""
license_utilization_analyzer.py - SaaS License Utilization & Reclaim Engine

PURPOSE
-------
Read license assignment and usage logs for any SaaS tool (Salesforce, Zoom,
Adobe, Office 365, GitHub, Atlassian, etc.), then produce:

  - Per-user utilization (did they actually use the seat they were paying for?)
  - Reclaim candidates (unused > 60 days)
  - Downgrade candidates (active but only using features of a cheaper tier)
  - Annualized waste $ estimate and a CSV of seats IT can safely revoke

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Every SaaS vendor has an admin console that shows ONE tool's usage. Nobody
consolidates usage across 40 tools at a company. Doing that consolidation in
Excel means VLOOKUPs across 40 sheets, which is where reclaim projects go
to die. This script unifies them with a single mapping file.

USE CASE
--------
VP of Finance asks, "How much are we wasting on unused SaaS?"
Run this once a quarter across every major SaaS admin export. Typical find
at a 2,000-person software company: $500K-$1.5M in reclaimable annual spend.

INPUT (any number of files, each representing one tool):
    <tool>_assignments.csv : user_email, plan, seat_cost_monthly
    <tool>_activity.csv    : user_email, last_active_date, feature_tier_used

CONFIG FILE: tools.yaml
    tools:
      - name: Salesforce
        assignments: sf_assign.csv
        activity: sf_activity.csv
        inactivity_days: 45
      - name: Zoom
        assignments: zoom_seats.csv
        activity: zoom_usage.csv
        inactivity_days: 30

USAGE
-----
    python license_utilization_analyzer.py tools.yaml --output seats.xlsx
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import yaml


@dataclass
class ToolSpec:
    name: str
    assignments: Path
    activity: Path
    inactivity_days: int


def _load_spec(config_path: Path) -> list[ToolSpec]:
    with config_path.open() as f:
        cfg = yaml.safe_load(f)
    base = config_path.parent
    return [
        ToolSpec(
            name=t["name"],
            assignments=base / t["assignments"],
            activity=base / t["activity"],
            inactivity_days=t.get("inactivity_days", 45),
        )
        for t in cfg["tools"]
    ]


def analyze_tool(spec: ToolSpec, today: datetime) -> pd.DataFrame:
    """Join assignments + activity, classify each seat."""
    assign = pd.read_csv(spec.assignments)
    activity = pd.read_csv(spec.activity, parse_dates=["last_active_date"])

    df = assign.merge(activity, on="user_email", how="left")
    df["tool"] = spec.name
    df["inactivity_threshold_days"] = spec.inactivity_days

    # Days since last active
    df["days_inactive"] = (today - df["last_active_date"]).dt.days
    df["days_inactive"] = df["days_inactive"].fillna(9999).astype(int)

    # Classify
    df["classification"] = "Active"
    df.loc[df["days_inactive"] > spec.inactivity_days, "classification"] = "Reclaim Candidate"
    df.loc[df["last_active_date"].isna(), "classification"] = "Never Used"

    # Downgrade candidates - active but paying for a tier above what they use
    tier_rank = {"Free": 0, "Basic": 1, "Pro": 2, "Business": 3, "Enterprise": 4}
    df["plan_rank"] = df["plan"].map(tier_rank).fillna(0)
    df["used_rank"] = df["feature_tier_used"].map(tier_rank).fillna(0)
    overpaying = (df["classification"] == "Active") & (df["plan_rank"] > df["used_rank"] + 1)
    df.loc[overpaying, "classification"] = "Downgrade Candidate"

    # Annualized waste
    df["annual_waste"] = 0.0
    df.loc[df["classification"].isin(["Reclaim Candidate", "Never Used"]), "annual_waste"] = (
        df["seat_cost_monthly"] * 12
    )
    df.loc[df["classification"] == "Downgrade Candidate", "annual_waste"] = (
        df["seat_cost_monthly"] * 12 * 0.4  # assume 40% savings from downgrade
    )

    return df


def build_summary(all_seats: pd.DataFrame) -> pd.DataFrame:
    summary = (
        all_seats.groupby(["tool", "classification"])
        .agg(seats=("user_email", "count"), annual_waste=("annual_waste", "sum"))
        .reset_index()
    )
    totals = all_seats.groupby("tool").agg(
        total_seats=("user_email", "count"),
        total_spend=("seat_cost_monthly", lambda s: (s * 12).sum()),
    )
    pivot = summary.pivot_table(
        index="tool", columns="classification", values="seats", aggfunc="sum", fill_value=0
    )
    waste = summary.pivot_table(
        index="tool", columns="classification", values="annual_waste", aggfunc="sum", fill_value=0
    )
    waste.columns = [f"waste_{c}" for c in waste.columns]
    return pd.concat([pivot, waste, totals], axis=1).reset_index()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("config_yaml")
    ap.add_argument("--output", default="seat_utilization.xlsx")
    ap.add_argument("--asof", default=None, help="YYYY-MM-DD; defaults to today")
    args = ap.parse_args()

    today = datetime.strptime(args.asof, "%Y-%m-%d") if args.asof else datetime.today()

    specs = _load_spec(Path(args.config_yaml))
    all_frames = [analyze_tool(s, today) for s in specs]
    all_seats = pd.concat(all_frames, ignore_index=True)

    summary = build_summary(all_seats)
    reclaim = all_seats[all_seats["classification"].isin(
        ["Reclaim Candidate", "Never Used"])].sort_values("annual_waste", ascending=False)
    downgrade = all_seats[all_seats["classification"] == "Downgrade Candidate"].sort_values(
        "annual_waste", ascending=False)

    with pd.ExcelWriter(args.output, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="Summary", index=False)
        all_seats.to_excel(w, sheet_name="All Seats", index=False)
        reclaim.to_excel(w, sheet_name="Reclaim Candidates", index=False)
        downgrade.to_excel(w, sheet_name="Downgrade Candidates", index=False)

    total_waste = all_seats["annual_waste"].sum()
    print(f"Wrote {args.output}")
    print(f"Total annualized reclaim opportunity: ${total_waste:,.0f}")
    print(f"Reclaim candidates: {len(reclaim)}")
    print(f"Downgrade candidates: {len(downgrade)}")


if __name__ == "__main__":
    main()
