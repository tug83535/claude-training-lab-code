from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


def _load_weights(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def compute_scores(df: pd.DataFrame, weights: dict) -> pd.DataFrame:
    required = weights.get("required_columns", ["id", "impact", "confidence", "days_open"])
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns for triage scoring: {missing}")

    impact_w = float(weights.get("impact_weight", 0.5))
    confidence_w = float(weights.get("confidence_weight", 0.3))
    recency_w = float(weights.get("recency_weight", 0.2))
    days_scale = max(float(weights.get("days_scale", 30)), 1.0)

    out = df.copy()
    out["impact"] = pd.to_numeric(out["impact"], errors="coerce").fillna(0.0)
    out["confidence"] = pd.to_numeric(out["confidence"], errors="coerce").fillna(0.0)
    out["days_open"] = pd.to_numeric(out["days_open"], errors="coerce").fillna(days_scale)

    out["recency_score"] = 1.0 - (out["days_open"].clip(lower=0) / days_scale).clip(upper=1.0)
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
