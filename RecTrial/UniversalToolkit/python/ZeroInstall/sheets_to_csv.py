#!/usr/bin/env python3
"""Extract selected worksheet values from an xlsx/xlsm into CSV files."""

from __future__ import annotations

import argparse
import csv
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

NS_MAIN = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
NS_DOC_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract selected workbook sheets to CSV files.")
    parser.add_argument("workbook", type=Path, help="Path to xlsx/xlsm workbook")
    parser.add_argument("--out-dir", type=Path, required=True, help="Output folder for CSV files")
    parser.add_argument(
        "--sheets",
        default="P&L - Monthly Trend,Product Line Summary,Functional P&L - Monthly Trend",
        help="Comma-separated sheet names to extract",
    )
    return parser.parse_args()


def _shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for si in root.findall("main:si", NS_MAIN):
        text = "".join(node.text or "" for node in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
        values.append(text)
    return values


def _sheet_map(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {item.attrib["Id"]: item.attrib["Target"] for item in rels_xml.findall("rel:Relationship", NS_REL)}

    result: dict[str, str] = {}
    sheets = workbook_xml.find("main:sheets", NS_MAIN)
    if sheets is None:
        return result

    for sheet in sheets:
        name = sheet.attrib.get("name", "")
        rid = sheet.attrib.get(NS_DOC_REL, "")
        target = rel_map.get(rid)
        if target:
            result[name] = "xl/" + target.lstrip("/")
    return result


def _extract_sheet_rows(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> list[list[str]]:
    if sheet_path not in zf.namelist():
        return []

    root = ET.fromstring(zf.read(sheet_path))
    rows: list[list[str]] = []
    for row in root.findall(".//main:sheetData/main:row", NS_MAIN):
        cells = []
        for cell in row.findall("main:c", NS_MAIN):
            t = cell.attrib.get("t")
            value_node = cell.find("main:v", NS_MAIN)
            inline_node = cell.find("main:is", NS_MAIN)

            if inline_node is not None:
                text = "".join(node.text or "" for node in inline_node.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
                cells.append(text)
                continue

            if value_node is None or value_node.text is None:
                cells.append("")
                continue

            raw = value_node.text
            if t == "s":
                try:
                    cells.append(shared[int(raw)])
                except (ValueError, IndexError):
                    cells.append(raw)
            else:
                cells.append(raw)
        rows.append(cells)
    return rows


def sanitize_filename(name: str) -> str:
    clean = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in name.strip())
    return clean.strip("_") or "sheet"


def extract_sheets(workbook: Path, out_dir: Path, sheet_names: list[str]) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    outputs: list[Path] = []

    with zipfile.ZipFile(workbook) as zf:
        shared = _shared_strings(zf)
        mapping = _sheet_map(zf)

        for name in sheet_names:
            sheet = name.strip()
            if not sheet:
                continue
            path = mapping.get(sheet)
            if not path:
                continue
            rows = _extract_sheet_rows(zf, path, shared)
            out_path = out_dir / f"{sanitize_filename(sheet)}.csv"
            with out_path.open("w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerows(rows)
            outputs.append(out_path)

    return outputs


def main() -> None:
    args = parse_args()
    sheet_names = [item.strip() for item in args.sheets.split(",")]
    outputs = extract_sheets(args.workbook, args.out_dir, sheet_names)
    print(f"Extracted files: {len(outputs)}")
    for path in outputs:
        print(path)


if __name__ == "__main__":
    main()
