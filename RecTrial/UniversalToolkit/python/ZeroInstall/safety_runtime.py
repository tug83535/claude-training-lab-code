#!/usr/bin/env python3
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any


def toolkit_root() -> Path:
    return Path(__file__).resolve().parent


def make_run_output(tool_name: str) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = toolkit_root() / "outputs" / f"{ts}_{tool_name}"
    out.mkdir(parents=True, exist_ok=True)
    return out


def write_run_logs(output_dir: Path, summary: str, payload: dict[str, Any]) -> None:
    (output_dir / "run_summary.txt").write_text(summary + "\n", encoding="utf-8")
    (output_dir / "run_log.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")


def fail_message(msg: str) -> None:
    raise SystemExit(f"Error: {msg}")


def require_existing_file(path: Path | None, label: str) -> None:
    if path is None:
        fail_message(f"missing {label} path")
    if not path.exists():
        fail_message(f"{label} file was not found: {path}")
