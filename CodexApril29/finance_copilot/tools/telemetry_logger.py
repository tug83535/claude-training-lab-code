from __future__ import annotations

import csv
from datetime import datetime, timezone
from pathlib import Path


def log_event(
    telemetry_path: Path,
    command: str,
    status: str,
    duration_ms: int,
    output_ref: str = "",
    error_message: str = "",
) -> Path:
    telemetry_path.parent.mkdir(parents=True, exist_ok=True)
    needs_header = not telemetry_path.exists() or telemetry_path.stat().st_size == 0

    with telemetry_path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if needs_header:
            writer.writerow(
                [
                    "timestamp_utc",
                    "command",
                    "status",
                    "duration_ms",
                    "output_ref",
                    "error_message",
                ]
            )

        writer.writerow(
            [
                datetime.now(timezone.utc).isoformat(),
                command,
                status,
                duration_ms,
                output_ref,
                error_message,
            ]
        )

    return telemetry_path
