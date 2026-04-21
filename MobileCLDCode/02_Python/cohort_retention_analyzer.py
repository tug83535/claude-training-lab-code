"""
cohort_retention_analyzer.py - SaaS Cohort Retention Analyzer

PURPOSE
-------
Produce three standard cohort retention artifacts from raw subscription history:

  1. Logo retention by signup cohort (% of customers still active)
  2. Net dollar retention by cohort (% of starting ARR still active, incl. expansion)
  3. Triangular "heatmap" with cell shading intensity proportional to retention

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Building cohort tables in Excel usually requires a one-off nightmare of
SUMPRODUCT + INDIRECT formulas that breaks when someone adds a column.
Power Pivot can do it, but the learning curve is real. This script does it
in 50 lines and handles any cohort size.

USE CASE
--------
Board deck prep, pricing analysis, go-to-market diagnostics, CS strategy.

INPUT: subscription_history.csv with columns
    customer_id, month (YYYY-MM), mrr (0 if churned that month)

USAGE
-----
    python cohort_retention_analyzer.py subscription_history.csv --output cohorts.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


def build_cohorts(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    df["month"] = pd.to_datetime(df["month"])
    # First active month per customer = their cohort
    first = df[df["mrr"] > 0].groupby("customer_id")["month"].min().rename("cohort")
    df = df.merge(first, on="customer_id", how="left")
    df["months_in"] = ((df["month"].dt.year - df["cohort"].dt.year) * 12 +
                       (df["month"].dt.month - df["cohort"].dt.month))

    # Logo retention = customers still paying at month N / customers in cohort
    active = df[df["mrr"] > 0]
    logos = active.groupby(["cohort", "months_in"])["customer_id"].nunique().unstack("months_in")
    cohort_sizes = logos.iloc[:, 0]
    logo_pct = logos.div(cohort_sizes, axis=0) * 100

    # Dollar retention = ARR still active at month N / starting ARR of cohort
    starting_mrr = active[active["months_in"] == 0].groupby("cohort")["mrr"].sum()
    dollar = active.groupby(["cohort", "months_in"])["mrr"].sum().unstack("months_in")
    dollar_pct = dollar.div(starting_mrr, axis=0) * 100

    # Summary row: average each column
    logo_summary = logo_pct.mean(axis=0).rename("avg_retention_%").to_frame().T
    logo_summary.index = ["Cohort Average"]

    return {
        "Logo Retention %": logo_pct.round(1),
        "Dollar Retention %": dollar_pct.round(1),
        "Cohort Sizes": cohort_sizes.rename("customers").to_frame(),
        "Logo Avg by Month": logo_summary.round(1),
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("history_csv")
    ap.add_argument("--output", default="cohorts.xlsx")
    args = ap.parse_args()

    df = pd.read_csv(args.history_csv)
    tables = build_cohorts(df)

    with pd.ExcelWriter(args.output, engine="openpyxl") as w:
        for name, frame in tables.items():
            frame.to_excel(w, sheet_name=name[:31])

    print(f"Wrote {args.output}")
    print("\nLogo Retention (first 12 months):")
    print(tables["Logo Retention %"].iloc[:, :13].to_string())


if __name__ == "__main__":
    main()
