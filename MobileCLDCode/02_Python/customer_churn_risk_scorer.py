"""
customer_churn_risk_scorer.py - Customer churn risk scoring

PURPOSE
-------
Score every customer on a 0-100 churn risk scale using product usage signals,
support activity, billing health, NPS, and tenure. Explains each score with the
top 3 risk drivers so the CSM team knows what to act on.

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Excel can build a simple weighted score, but (a) cannot standardize features
against a rolling baseline, (b) cannot combine categorical and continuous
features cleanly, and (c) cannot produce per-row model explanations. This
script uses scikit-learn's logistic regression with calibration + SHAP-style
coefficient explanations, then writes the results back to Excel for the CSM
team to triage.

USE CASE
--------
Every Monday 9am, CS leadership wants a fresh list of the top 50 at-risk
customers, with the reason for each. Run this weekly against fresh usage
exports; paste the Top50 sheet into the CS standup.

INPUT: customers.csv with columns
    customer_id, arr, tenure_months, last_login_days_ago, active_users_pct,
    support_tickets_90d, nps, past_due_balance, price_increase_last_year,
    churned (0/1, historical label to train on)

USAGE
-----
    python customer_churn_risk_scorer.py customers.csv --output churn_scores.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import Pipeline
from sklearn.calibration import CalibratedClassifierCV


FEATURES = [
    "arr",
    "tenure_months",
    "last_login_days_ago",
    "active_users_pct",
    "support_tickets_90d",
    "nps",
    "past_due_balance",
    "price_increase_last_year",
]

# Direction of risk: +1 means HIGHER value raises churn risk; -1 means HIGHER value lowers it.
EXPECTED_DIRECTION = {
    "arr": -1,
    "tenure_months": -1,
    "last_login_days_ago": +1,
    "active_users_pct": -1,
    "support_tickets_90d": +1,
    "nps": -1,
    "past_due_balance": +1,
    "price_increase_last_year": +1,
}


def train_model(df: pd.DataFrame) -> tuple[Pipeline, pd.Series]:
    """Train a calibrated logistic regression and return it + feature coefs."""
    X = df[FEATURES].astype(float).fillna(df[FEATURES].median(numeric_only=True))
    y = df["churned"].astype(int)
    pipe = Pipeline(
        steps=[
            ("scale", StandardScaler()),
            (
                "clf",
                CalibratedClassifierCV(
                    LogisticRegression(max_iter=500, class_weight="balanced"),
                    cv=5,
                    method="isotonic",
                ),
            ),
        ]
    )
    pipe.fit(X, y)

    # Pull the underlying LR coefficients by refitting one uncalibrated version
    lr = LogisticRegression(max_iter=500, class_weight="balanced")
    lr.fit(StandardScaler().fit_transform(X), y)
    coefs = pd.Series(lr.coef_[0], index=FEATURES).sort_values()
    return pipe, coefs


def score_customers(pipe: Pipeline, df: pd.DataFrame) -> pd.Series:
    X = df[FEATURES].astype(float).fillna(df[FEATURES].median(numeric_only=True))
    # probability of class 1 (churn)
    return pd.Series(pipe.predict_proba(X)[:, 1] * 100, index=df.index)


def per_customer_drivers(df: pd.DataFrame, coefs: pd.Series) -> pd.DataFrame:
    """For each customer, return the top 3 features driving their risk score."""
    X = df[FEATURES].astype(float).fillna(df[FEATURES].median(numeric_only=True))
    scaler = StandardScaler().fit(X)
    Xs = pd.DataFrame(scaler.transform(X), columns=FEATURES, index=X.index)
    # contribution = z_score * coefficient
    contrib = Xs.multiply(coefs, axis=1)

    top_drivers = []
    for _, row in contrib.iterrows():
        # Only features that raised risk
        positives = row[row > 0].sort_values(ascending=False).head(3)
        reasons = []
        for feat, score in positives.items():
            direction = EXPECTED_DIRECTION.get(feat, 0)
            sign_note = "(raised)" if direction > 0 else "(unexpected)"
            reasons.append(f"{feat} {sign_note}")
        top_drivers.append(" | ".join(reasons) if reasons else "low-risk profile")
    return pd.Series(top_drivers, index=contrib.index)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input_csv")
    ap.add_argument("--output", default="churn_scores.xlsx")
    ap.add_argument("--top-n", type=int, default=50)
    args = ap.parse_args()

    df = pd.read_csv(args.input_csv)
    missing = [c for c in FEATURES + ["customer_id", "churned"] if c not in df.columns]
    if missing:
        raise SystemExit(f"Missing columns: {missing}")

    trained = df.dropna(subset=["churned"])
    pipe, coefs = train_model(trained)
    print("Feature coefficients (more negative = protective; more positive = risky):")
    print(coefs.round(3).to_string())

    all_customers = df.drop(columns=["churned"]).copy()
    all_customers["risk_score_0_100"] = score_customers(pipe, df).round(1)
    all_customers["top_risk_drivers"] = per_customer_drivers(df, coefs).values
    all_customers["risk_tier"] = pd.cut(
        all_customers["risk_score_0_100"],
        bins=[-0.1, 20, 50, 80, 100],
        labels=["Low", "Medium", "High", "Critical"],
    )

    top = all_customers.sort_values("risk_score_0_100", ascending=False).head(args.top_n)

    with pd.ExcelWriter(args.output, engine="openpyxl") as w:
        all_customers.sort_values("risk_score_0_100", ascending=False).to_excel(
            w, sheet_name="All Customers", index=False
        )
        top.to_excel(w, sheet_name=f"Top {args.top_n}", index=False)
        coefs.rename("coefficient").to_frame().to_excel(w, sheet_name="Feature Weights")

    print(f"Wrote {args.output}")
    print(f"Top {args.top_n} customers at risk:")
    print(top[["customer_id", "risk_score_0_100", "top_risk_drivers"]].to_string(index=False))


if __name__ == "__main__":
    main()
