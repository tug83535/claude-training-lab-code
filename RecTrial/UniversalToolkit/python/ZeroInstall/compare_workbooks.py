#!/usr/bin/env python3
"""Compare two xlsx/xlsm workbooks and emit cell-level diffs."""

from __future__ import annotations

VERSION = "1.0.0"

import argparse
import csv
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

NS_MAIN = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare workbook cells and export a diff CSV.")
    parser.add_argument("left_workbook", type=Path)
    parser.add_argument("right_workbook", type=Path)
    parser.add_argument("out_csv", type=Path)
    return parser.parse_args()


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for si in root.findall("main:si", NS_MAIN):
        text = "".join(node.text or "" for node in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
        values.append(text)
    return values


def sheet_map(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall("rel:Relationship", NS_REL)}

    mapping: dict[str, str] = {}
    for sheet in workbook.find("main:sheets", NS_MAIN):
        name = sheet.attrib.get("name", "")
        rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
        target = rel_map.get(rid, "")
        mapping[name] = "xl/" + target.lstrip("/")
    return mapping


def sheet_cells(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> dict[str, str]:
    if sheet_path not in zf.namelist():
        return {}

    root = ET.fromstring(zf.read(sheet_path))
    values: dict[str, str] = {}
    for cell in root.findall(".//main:sheetData/main:row/main:c", NS_MAIN):
        ref = cell.attrib.get("r", "")
        if not re.match(r"^[A-Z]+\d+$", ref):
            continue

        cell_type = cell.attrib.get("t")
        value_node = cell.find("main:v", NS_MAIN)
        inline_node = cell.find("main:is", NS_MAIN)

        if inline_node is not None:
            text = "".join(node.text or "" for node in inline_node.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            values[ref] = text
            continue

        if value_node is None or value_node.text is None:
            values[ref] = ""
            continue

        raw = value_node.text
        if cell_type == "s":
            try:
                values[ref] = shared[int(raw)]
            except (ValueError, IndexError):
                values[ref] = raw
        else:
            values[ref] = raw
    return values


def compare(left: Path, right: Path) -> list[tuple[str, str, str, str]]:
    with zipfile.ZipFile(left) as zleft, zipfile.ZipFile(right) as zright:
        left_shared = shared_strings(zleft)
        right_shared = shared_strings(zright)

        left_sheets = sheet_map(zleft)
        right_sheets = sheet_map(zright)

        all_sheet_names = sorted(set(left_sheets) | set(right_sheets))
        diffs: list[tuple[str, str, str, str]] = []

        for name in all_sheet_names:
            left_cells = sheet_cells(zleft, left_sheets.get(name, ""), left_shared)
            right_cells = sheet_cells(zright, right_sheets.get(name, ""), right_shared)
            refs = sorted(set(left_cells) | set(right_cells))

            for ref in refs:
                left_value = left_cells.get(ref, "")
                right_value = right_cells.get(ref, "")
                if left_value != right_value:
                    diffs.append((name, ref, left_value, right_value))

        return diffs


def write_diffs(out_csv: Path, diffs: list[tuple[str, str, str, str]]) -> None:
    with out_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["sheet", "cell", "left_value", "right_value"])
        writer.writerows(diffs)



def require_existing_file(path: Path, label: str) -> None:
    if path is None:
        raise SystemExit(f"Error: missing {label} path.")
    if not path.exists():
        raise SystemExit(f"Error: {label} file was not found: {path}")


def main() -> None:
    args = parse_args()
    require_existing_file(args.left_workbook, "left workbook")
    require_existing_file(args.right_workbook, "right workbook")
    diffs = compare(args.left_workbook, args.right_workbook)
    write_diffs(args.out_csv, diffs)
    print(f"Diff rows: {len(diffs)}")
    print(f"Output: {args.out_csv}")


if __name__ == "__main__":
    main()
