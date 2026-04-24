#!/usr/bin/env python3
"""Build a markdown executive brief package from existing artifacts."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create a markdown executive brief package.")
    parser.add_argument("--title", default="Executive Brief", help="Document title")
    parser.add_argument("--summary", type=Path, required=True, help="Path to markdown summary content")
    parser.add_argument("--scenario", type=Path, help="Optional scenario CSV artifact")
    parser.add_argument("--variance", type=Path, help="Optional variance CSV artifact")
    parser.add_argument("--out", type=Path, required=True, help="Output markdown file")
    return parser.parse_args()


def build_brief(title: str, summary: str, scenario: Path | None, variance: Path | None) -> str:
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    lines = [f"# {title}", "", f"Generated: {timestamp}", "", "## Summary", "", summary.strip(), ""]

    lines.extend(["## Artifact Links", ""])
    if scenario is not None:
        lines.append(f"- Scenario output: `{scenario}`")
    if variance is not None:
        lines.append(f"- Variance output: `{variance}`")
    if scenario is None and variance is None:
        lines.append("- No supplemental artifacts were provided.")

    lines.extend(["", "## Reviewer Notes", "", "- Confirm major variance drivers.", "- Validate scenario assumptions before final sign-off."])
    return "\n".join(lines) + "\n"


def main() -> None:
    args = parse_args()
    summary = args.summary.read_text(encoding="utf-8")
    payload = build_brief(args.title, summary, args.scenario, args.variance)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(payload, encoding="utf-8")
    print(f"Brief package written: {args.out}")


if __name__ == "__main__":
    main()
