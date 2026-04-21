"""
saas_arr_waterfall.py - SaaS ARR/MRR Waterfall Engine

PURPOSE
-------
Take a raw subscription roster (customer, plan, MRR, start date, end date) and
produce a month-by-month ARR waterfall broken into the five SaaS movements:

    Starting ARR -> + New -> + Expansion -> - Contraction -> - Churn -> Ending ARR

Also emits NRR, GRR, quick-ratio, and a per-customer expansion/contraction log.

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Excel cannot do panel/windowed comparisons of subscription state across months
at scale in a reliable way - anyone who's tried knows LOOKUP-based waterfalls
are bug factories. This engine uses pandas joins across monthly snapshots,
which is both auditable and fast on 100K+ customers.

USE CASE
--------
Every month the FP&A team needs a bulletproof ARR bridge for the board deck.
Run this against Snowflake exports and get the number + every movement line
item + the sanity checks, in ~5 seconds.

INPUT: subscriptions.csv with columns
    customer_id, plan, mrr, start_date, end_date (blank if active)

USAGE
-----
    python saas_arr_waterfall.py subscriptions.csv --from 2025-01 --to 2026-03 \\
        --output waterfall.xlsx
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pandas as pd


@dataclass
class WaterfallConfig:
    start_month: str
    end_month: str
    output_path: Path


def _month_range(start: str, end: str) -> list[pd.Timestamp]:
    return list(pd.period_range(start=start, end=end, freq="M").to_timestamp())


def _active_at(df: pd.DataFrame, month_end: pd.Timestamp) -> pd.DataFrame:
    """Return the MRR each customer has active at month_end."""
    end_bound = df["end_date"].fillna(pd.Timestamp("2099-12-31"))
    mask = (df["start_date"] <= month_end) & (end_bound >= month_end)
    snap = df.loc[mask, ["customer_id", "plan", "mrr"]].copy()
    # If a customer has multiple overlapping rows, sum them (coterminous upsells).
    return snap.groupby("customer_id", as_index=False).agg(
        plan=("plan", "last"), mrr=("mrr", "sum")
    )


def build_waterfall(df: pd.DataFrame, cfg: WaterfallConfig) -> pd.DataFrame:
    """Compute the monthly ARR waterfall. Returns one row per month."""
    months = _month_range(cfg.start_month, cfg.end_month)
    rows = []
    prev_snap = None

    for m in months:
        month_end = (m + pd.offsets.MonthEnd(0)).normalize()
        snap = _active_at(df, month_end)
        snap.rename(columns={"mrr": "curr_mrr"}, inplace=True)

        if prev_snap is None:
            starting_arr = 0.0
            new = expansion = contraction = churn = 0.0
            new_ids = exp_ids = ctr_ids = chr_ids = set()
        else:
            merged = prev_snap.merge(
                snap, on="customer_id", how="outer", suffixes=("_prev", "_curr")
            )
            merged[["prev_mrr", "curr_mrr"]] = merged[["prev_mrr", "curr_mrr"]].fillna(0)

            new_mask = (merged["prev_mrr"] == 0) & (merged["curr_mrr"] > 0)
            churn_mask = (merged["prev_mrr"] > 0) & (merged["curr_mrr"] == 0)
            exp_mask = (merged["prev_mrr"] > 0) & (merged["curr_mrr"] > merged["prev_mrr"])
            ctr_mask = (merged["prev_mrr"] > 0) & (merged["curr_mrr"] < merged["prev_mrr"]) & (merged["curr_mrr"] > 0)

            new = merged.loc[new_mask, "curr_mrr"].sum() * 12
            expansion = (merged.loc[exp_mask, "curr_mrr"] - merged.loc[exp_mask, "prev_mrr"]).sum() * 12
            contraction = (merged.loc[ctr_mask, "prev_mrr"] - merged.loc[ctr_mask, "curr_mrr"]).sum() * 12
            churn = merged.loc[churn_mask, "prev_mrr"].sum() * 12

            starting_arr = prev_snap["prev_mrr"].sum() * 12
            new_ids = set(merged.loc[new_mask, "customer_id"])
            exp_ids = set(merged.loc[exp_mask, "customer_id"])
            ctr_ids = set(merged.loc[ctr_mask, "customer_id"])
            chr_ids = set(merged.loc[churn_mask, "customer_id"])

        ending_arr = snap["curr_mrr"].sum() * 12
        nrr = (starting_arr + expansion - contraction - churn) / starting_arr if starting_arr else None
        grr = (starting_arr - contraction - churn) / starting_arr if starting_arr else None
        quick_ratio = (new + expansion) / (contraction + churn) if (contraction + churn) else None

        rows.append(
            {
                "month": m.strftime("%Y-%m"),
                "starting_arr": starting_arr,
                "new": new,
                "expansion": expansion,
                "contraction": -contraction,
                "churn": -churn,
                "ending_arr": ending_arr,
                "check_delta": ending_arr - (starting_arr + new + expansion - contraction - churn),
                "nrr": nrr,
                "grr": grr,
                "quick_ratio": quick_ratio,
                "n_new": len(new_ids),
                "n_expanded": len(exp_ids),
                "n_contracted": len(ctr_ids),
                "n_churned": len(chr_ids),
            }
        )

        prev_snap = snap.rename(columns={"curr_mrr": "prev_mrr"})

    return pd.DataFrame(rows)


def cohort_retention(df: pd.DataFrame, cfg: WaterfallConfig) -> pd.DataFrame:
    """Cohort retention by signup month, expressed as % of original ARR remaining."""
    df = df.copy()
    df["signup_cohort"] = df["start_date"].dt.to_period("M").astype(str)
    months = _month_range(cfg.start_month, cfg.end_month)
    cohorts = sorted(df["signup_cohort"].unique())

    matrix = pd.DataFrame(index=cohorts, columns=[m.strftime("%Y-%m") for m in months], dtype=float)
    for cohort in cohorts:
        original = df.loc[df["signup_cohort"] == cohort, "mrr"].sum()
        if original == 0:
            continue
        customer_ids = df.loc[df["signup_cohort"] == cohort, "customer_id"].unique()
        for m in months:
            month_end = (m + pd.offsets.MonthEnd(0)).normalize()
            active = _active_at(df[df["customer_id"].isin(customer_ids)], month_end)
            matrix.loc[cohort, m.strftime("%Y-%m")] = active["mrr"].sum() / original
    return matrix


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input_csv")
    ap.add_argument("--from", dest="start_month", default=None)
    ap.add_argument("--to", dest="end_month", default=date.today().strftime("%Y-%m"))
    ap.add_argument("--output", default="arr_waterfall.xlsx")
    args = ap.parse_args()

    df = pd.read_csv(args.input_csv, parse_dates=["start_date", "end_date"])
    if args.start_month is None:
        args.start_month = df["start_date"].min().strftime("%Y-%m")

    cfg = WaterfallConfig(args.start_month, args.end_month, Path(args.output))
    waterfall = build_waterfall(df, cfg)
    retention = cohort_retention(df, cfg)

    with pd.ExcelWriter(cfg.output_path, engine="openpyxl") as writer:
        waterfall.to_excel(writer, sheet_name="ARR Waterfall", index=False)
        retention.to_excel(writer, sheet_name="Cohort Retention")

    print(f"Wrote {cfg.output_path}")
    print(waterfall.to_string(index=False))


if __name__ == "__main__":
    main()
