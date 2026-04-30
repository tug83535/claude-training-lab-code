from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


REQUIRED_DEFAULT = ["id", "impact", "confidence", "days_open"]


def _load_weights(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _validate_inputs(df: pd.DataFrame, required: list[str]) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns for triage scoring: {missing}")

    for col in ["impact", "confidence", "days_open"]:
        coerced = pd.to_numeric(df[col], errors="coerce")
        if coerced.isna().any():
            raise ValueError(f"Column '{col}' contains non-numeric values")

    if ((pd.to_numeric(df["impact"], errors="coerce") < 0) | (pd.to_numeric(df["impact"], errors="coerce") > 1)).any():
        raise ValueError("Column 'impact' must be within [0, 1]")

    if ((pd.to_numeric(df["confidence"], errors="coerce") < 0) | (pd.to_numeric(df["confidence"], errors="coerce") > 1)).any():
        raise ValueError("Column 'confidence' must be within [0, 1]")

    if (pd.to_numeric(df["days_open"], errors="coerce") < 0).any():
        raise ValueError("Column 'days_open' must be non-negative")


def compute_scores(df: pd.DataFrame, weights: dict) -> pd.DataFrame:
    required = weights.get("required_columns", REQUIRED_DEFAULT)
    _validate_inputs(df, required)

    impact_w = float(weights.get("impact_weight", 0.5))
    confidence_w = float(weights.get("confidence_weight", 0.3))
    recency_w = float(weights.get("recency_weight", 0.2))
    days_scale = max(float(weights.get("days_scale", 30)), 1.0)

    out = df.copy()
    out["impact"] = pd.to_numeric(out["impact"], errors="raise")
    out["confidence"] = pd.to_numeric(out["confidence"], errors="raise")
    out["days_open"] = pd.to_numeric(out["days_open"], errors="raise")

    out["recency_score"] = 1.0 - (out["days_open"] / days_scale).clip(upper=1.0)
    out["priority_score"] = (
        out["impact"] * impact_w
        + out["confidence"] * confidence_w
        + out["recency_score"] * recency_w
    )

    out = out.sort_values("priority_score", ascending=False).reset_index(drop=True)
    out["priority_rank"] = out.index + 1

    bins = [-1, 0.33, 0.66, 1.1]
    labels = ["LOW", "MEDIUM", "HIGH"]
    out["priority_band"] = pd.cut(out["priority_score"], bins=bins, labels=labels)

    return out


def run(input_csv: Path, weights_json: Path, output_dir: Path, top_n: int = 20) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(input_csv)
    weights = _load_weights(weights_json)
    scored = compute_scores(df, weights)

    out_path = output_dir / "exception_triage_ranked.csv"
    scored.head(top_n).to_csv(out_path, index=False)
    return out_path
