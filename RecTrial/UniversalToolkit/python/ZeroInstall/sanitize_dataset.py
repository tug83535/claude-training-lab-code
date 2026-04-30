#!/usr/bin/env python3
from __future__ import annotations

VERSION = "1.0.0"

import argparse
import csv
from datetime import datetime
from pathlib import Path
import re

from safety_runtime import make_run_output, require_existing_file, write_run_logs

DATE_FORMATS = ["%m/%d/%Y", "%Y-%m-%d", "%b %d %Y", "%m/%d/%y", "%B %d %Y"]


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Clean a CSV file and write results into toolkit outputs/.")
    p.add_argument("input_csv", nargs="?", type=Path, help="Input CSV file path")
    p.add_argument("--sample", action="store_true", help="Run using built-in synthetic sample data")
    return p.parse_args()


def sanitize_cell(v: str) -> str:
    if v is None:
        return ""
    t = re.sub(r"\s+", " ", str(v).replace("\r", " ").replace("\n", " ")).strip()
    if not t:
        return ""
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(t, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    c = t.replace(",", "").replace("$", "")
    try:
        n = float(c)
        return str(int(n)) if n.is_integer() else f"{n:.6f}".rstrip("0").rstrip(".")
    except ValueError:
        return t


def main() -> None:
    args = parse_args()
    out_dir = make_run_output("sanitize_dataset")
    try:
        if args.sample:
            src = out_dir / "sample_input.csv"
            src.write_text("Department,Amount\nA,100\nB, 200 \n", encoding="utf-8")
        else:
            require_existing_file(args.input_csv, "input CSV")
            src = args.input_csv

        out_csv = out_dir / "sanitized.csv"
        total = changed = 0
        with src.open("r", encoding="utf-8-sig", newline="") as inf, out_csv.open("w", encoding="utf-8", newline="") as outf:
            r = csv.reader(inf)
            w = csv.writer(outf)
            for row in r:
                nr = []
                for c in row:
                    total += 1
                    cc = sanitize_cell(c)
                    if cc != (c or ""):
                        changed += 1
                    nr.append(cc)
                w.writerow(nr)

        summary = f"Sanitized {changed} of {total} cells. Output: {out_csv.name}"
        write_run_logs(out_dir, summary, {"tool": "sanitize_dataset", "rows_cells_total": total, "changed_cells": changed, "output": out_csv.name})
        print(summary)
        print(f"Output folder: {out_dir}")
    except SystemExit:
        raise
    except Exception:
        write_run_logs(out_dir, "Run failed. Check run_log.json.", {"tool": "sanitize_dataset", "status": "failed"})
        raise SystemExit("Error: processing failed. See run_summary.txt in output folder.")


if __name__ == "__main__":
    main()
