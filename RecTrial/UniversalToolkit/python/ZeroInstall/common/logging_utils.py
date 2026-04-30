# Finance Automation Toolkit v1.0 — iPipeline
# common/logging_utils.py — run logging to run_log.json + run_summary.txt
#
# Every script run produces exactly two log files in the output folder:
#   run_log.json     — machine-readable; row counts, warnings, errors, findings
#   run_summary.txt  — plain English; what the script did and what it found
#
# Sensitive data is NOT stored in logs: only file names, row counts, column names,
# exception types, and statistics. Never the actual data values.

from datetime import datetime
from pathlib import Path
from .safe_io import write_json, write_text


class RunLogger:
    """Collects run metadata and writes both log files when finish() is called."""

    def __init__(self, script_name: str, output_dir: Path):
        self.script_name = script_name
        self.output_dir = output_dir
        self.start_time = datetime.now()
        self.end_time: datetime | None = None
        self.rows_read = 0
        self.rows_processed = 0
        self.warnings: list[str] = []
        self.errors: list[str] = []
        self.findings: list[dict] = []
        self._meta: dict = {}

    # --- recording helpers ---

    def set_meta(self, **kwargs) -> None:
        """Store arbitrary key/value metadata (file names, counts, modes)."""
        self._meta.update(kwargs)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)

    def error(self, msg: str) -> None:
        self.errors.append(msg)

    def finding(self, category: str, description: str, impact: str = "") -> None:
        """Record one detected exception or finding."""
        self.findings.append({
            "category": category,
            "description": description,
            "impact": impact,
        })

    # --- output ---

    def finish(self) -> None:
        """Compute elapsed time and write run_log.json + run_summary.txt."""
        self.end_time = datetime.now()
        elapsed = (self.end_time - self.start_time).total_seconds()

        log_data = {
            "script": self.script_name,
            "start": self.start_time.isoformat(),
            "end": self.end_time.isoformat(),
            "elapsed_seconds": round(elapsed, 2),
            "rows_read": self.rows_read,
            "rows_processed": self.rows_processed,
            "findings_count": len(self.findings),
            "warnings_count": len(self.warnings),
            "errors_count": len(self.errors),
            "warnings": self.warnings,
            "errors": self.errors,
            "findings": self.findings,
            **self._meta,
        }
        write_json(self.output_dir / "run_log.json", log_data)

        summary_lines = [
            "Finance Automation Toolkit v1.0 — Run Summary",
            "=" * 50,
            f"Script:    {self.script_name}",
            f"Run time:  {self.start_time.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Elapsed:   {elapsed:.1f} seconds",
            "",
            f"Rows read:       {self.rows_read}",
            f"Rows processed:  {self.rows_processed}",
            f"Findings:        {len(self.findings)}",
            f"Warnings:        {len(self.warnings)}",
            f"Errors:          {len(self.errors)}",
        ]

        if self._meta:
            summary_lines.append("")
            for k, v in self._meta.items():
                summary_lines.append(f"{k}: {v}")

        if self.warnings:
            summary_lines += ["", "WARNINGS:"] + [f"  - {w}" for w in self.warnings]

        if self.errors:
            summary_lines += ["", "ERRORS:"] + [f"  - {e}" for e in self.errors]

        if self.findings:
            summary_lines += ["", "FINDINGS:"]
            for f in self.findings:
                line = f"  [{f['category']}] {f['description']}"
                if f.get("impact"):
                    line += f" | Impact: {f['impact']}"
                summary_lines.append(line)

        summary_lines += [
            "",
            f"Output folder: {self.output_dir}",
            "",
            "If something looks wrong, check run_log.json for details.",
            "Questions? Contact Connor Atlee — Finance & Accounting.",
        ]

        write_text(self.output_dir / "run_summary.txt", "\n".join(summary_lines))
