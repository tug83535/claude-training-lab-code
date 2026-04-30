#!/usr/bin/env python3
"""Sanitize CSV data for downstream Excel workflows."""

from __future__ import annotations

VERSION = "1.0.0"

import argparse
import csv
from datetime import datetime
from pathlib import Path
import re

DATE_FORMATS = ["%m/%d/%Y", "%Y-%m-%d", "%b %d %Y", "%m/%d/%y", "%B %d %Y"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Sanitize CSV values and write a cleaned CSV.")
    parser.add_argument("input_csv", type=Path, help="Input CSV path")
    parser.add_argument("output_csv", type=Path, help="Output CSV path")
    return parser.parse_args()


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\r", " ").replace("\n", " ")).strip()


def normalize_number(text: str) -> str:
    candidate = text.replace(",", "").strip()
    if candidate.startswith("$"):
        candidate = candidate[1:]
    if candidate.endswith("%"):
        return text
    try:
        value = float(candidate)
    except ValueError:
        return text

    if value.is_integer():
        return str(int(value))
    return f"{value:.6f}".rstrip("0").rstrip(".")


def normalize_date(text: str) -> str:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text


def sanitize_cell(value: str) -> str:
    if value is None:
        return ""

    text = normalize_text(str(value))
    if not text:
        return ""

    date_value = normalize_date(text)
    if date_value != text:
        return date_value

    number_value = normalize_number(text)
    return number_value


def sanitize_csv(input_csv: Path, output_csv: Path) -> tuple[int, int]:
    total_cells = 0
    changed_cells = 0

    with input_csv.open("r", encoding="utf-8-sig", newline="") as infile:
        reader = csv.reader(infile)
        rows = []
        for row in reader:
            new_row = []
            for cell in row:
                total_cells += 1
                cleaned = sanitize_cell(cell)
                if cleaned != (cell or ""):
                    changed_cells += 1
                new_row.append(cleaned)
            rows.append(new_row)

    with output_csv.open("w", encoding="utf-8", newline="") as outfile:
        writer = csv.writer(outfile)
        writer.writerows(rows)

    return total_cells, changed_cells



def require_existing_file(path: Path, label: str) -> None:
    if not path.exists():
        raise SystemExit(f"Error: {label} file was not found: {path}")
    if not path.is_file():
        raise SystemExit(f"Error: {label} path is not a file: {path}")


def main() -> None:
    args = parse_args()
    require_existing_file(args.input_csv, "input CSV")
    total_cells, changed_cells = sanitize_csv(args.input_csv, args.output_csv)
    print(f"Sanitized {changed_cells} of {total_cells} cells.")
    print(f"Output: {args.output_csv}")


if __name__ == "__main__":
    main()
