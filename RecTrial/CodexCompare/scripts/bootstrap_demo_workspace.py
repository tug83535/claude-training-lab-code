#!/usr/bin/env python3
"""Create a safe working copy of sample workbooks for local demo runs."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = [
    ROOT / "samples/ExcelDemoFile_adv.xlsm",
    ROOT / "samples/Sample_Quarterly_ReportV2.xlsm",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create timestamped working copies of sample workbooks.")
    parser.add_argument("--out-dir", type=Path, default=ROOT / "artifacts" / "working-copies", help="Output directory")
    parser.add_argument("--timestamp", help="Optional fixed timestamp suffix for reproducible runs")
    return parser.parse_args()


def create_workspace(out_dir: Path, timestamp: str | None = None) -> Path:
    workspace_timestamp = timestamp or datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    target_dir = out_dir / f"demo_workspace_{workspace_timestamp}"
    target_dir.mkdir(parents=True, exist_ok=False)

    for sample in SAMPLES:
        if not sample.exists():
            raise SystemExit(f"Missing sample workbook: {sample}")
        destination = target_dir / sample.name
        shutil.copy2(sample, destination)

    return target_dir


def main() -> None:
    args = parse_args()
    target_dir = create_workspace(args.out_dir, timestamp=args.timestamp)

    print(f"Created demo workspace: {target_dir}")
    print("Next steps:")
    print("1) Open copied workbooks from this folder.")
    print("2) Run command center macros on copied files only.")
    print("3) Keep original samples unchanged.")


if __name__ == "__main__":
    main()
