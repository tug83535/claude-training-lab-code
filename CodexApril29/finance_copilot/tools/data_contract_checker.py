from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


@dataclass
class ContractResult:
    status: str
    issues: list[str]
    summary: dict[str, Any]


def _load_contract(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _normalize_type(type_name: str) -> str:
    t = type_name.strip().lower()
    if t in {"string", "str", "text"}:
        return "string"
    if t in {"number", "numeric", "float", "int", "integer"}:
        return "number"
    if t in {"date", "datetime"}:
        return "date"
    return "unknown"


def _check_type(series: pd.Series, expected: str) -> bool:
    non_null = series.dropna()
    if non_null.empty:
        return True

    if expected == "string":
        return True
    if expected == "number":
        coerced = pd.to_numeric(non_null, errors="coerce")
        return coerced.notna().all()
    if expected == "date":
        coerced = pd.to_datetime(non_null, errors="coerce")
        return coerced.notna().all()
    return True


def evaluate_dataset(df: pd.DataFrame, contract: dict[str, Any]) -> ContractResult:
    issues: list[str] = []

    required_columns = contract.get("required_columns", [])
    missing_cols = [c for c in required_columns if c not in df.columns]
    if missing_cols:
        issues.append(f"Missing required columns: {missing_cols}")

    type_rules = contract.get("column_types", {})
    for col, expected_type in type_rules.items():
        if col not in df.columns:
            continue
        normalized = _normalize_type(expected_type)
        if not _check_type(df[col], normalized):
            issues.append(f"Column '{col}' failed type check for expected type '{normalized}'")

    threshold = float(contract.get("null_threshold_pct", 100.0))
    for col in required_columns:
        if col not in df.columns:
            continue
        null_pct = float(df[col].isna().mean() * 100)
        if null_pct > threshold:
            issues.append(
                f"Column '{col}' null percentage {null_pct:.2f}% exceeds threshold {threshold:.2f}%"
            )

    constraints = contract.get("constraints", {})
    unique_keys = constraints.get("unique_keys", [])
    if unique_keys and all(k in df.columns for k in unique_keys):
        dup_count = int(df.duplicated(subset=unique_keys).sum())
        if dup_count > 0:
            issues.append(f"Duplicate key rows found for unique_keys={unique_keys}: {dup_count}")

    non_negative_cols = constraints.get("non_negative_columns", [])
    for col in non_negative_cols:
        if col not in df.columns:
            continue
        numeric = pd.to_numeric(df[col], errors="coerce")
        neg_count = int((numeric < 0).sum())
        if neg_count > 0:
            issues.append(f"Column '{col}' contains {neg_count} negative values")

    allowed_values = constraints.get("allowed_values", {})
    for col, allowed in allowed_values.items():
        if col not in df.columns:
            continue
        observed = set(df[col].dropna().astype(str).unique())
        allowed_set = set(str(v) for v in allowed)
        unknown = sorted(observed - allowed_set)
        if unknown:
            issues.append(f"Column '{col}' has disallowed values: {unknown[:10]}")

    status = "PASS" if not issues else "FAIL"
    summary = {
        "rows": int(len(df)),
        "columns": int(len(df.columns)),
        "required_columns": required_columns,
        "issues_count": len(issues),
    }

    return ContractResult(status=status, issues=issues, summary=summary)


def run(input_csv: Path, contract_json: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(input_csv)
    contract = _load_contract(contract_json)
    result = evaluate_dataset(df, contract)

    report = {
        "dataset": str(input_csv),
        "contract": str(contract_json),
        "status": result.status,
        "summary": result.summary,
        "issues": result.issues,
    }

    out_path = output_dir / "data_contract_report.json"
    out_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return out_path
