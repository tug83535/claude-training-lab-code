#!/usr/bin/env python3
"""Generate a workbook structure profile for xlsx/xlsm files."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

NS_MAIN = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Profile workbook sheets, named ranges, and VBA presence.")
    parser.add_argument("workbook", type=Path, help="Path to .xlsx or .xlsm file")
    parser.add_argument("--out", type=Path, help="Optional JSON output path")
    return parser.parse_args()


def col_to_num(col: str) -> int:
    value = 0
    for char in col:
        if "A" <= char <= "Z":
            value = value * 26 + (ord(char) - 64)
    return value


def parse_dimension(dim: str | None) -> tuple[int | None, int | None]:
    if not dim:
        return None, None
    if ":" not in dim:
        match = re.match(r"([A-Z]+)(\d+)", dim)
        if not match:
            return None, None
        return int(match.group(2)), col_to_num(match.group(1))

    _, end_ref = dim.split(":", 1)
    match = re.match(r"([A-Z]+)(\d+)", end_ref)
    if not match:
        return None, None
    return int(match.group(2)), col_to_num(match.group(1))


def build_profile(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        workbook_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

        rel_map = {
            item.attrib["Id"]: item.attrib["Target"]
            for item in rels_xml.findall("rel:Relationship", NS_REL)
        }

        sheets = []
        for sheet in workbook_xml.find("main:sheets", NS_MAIN):
            name = sheet.attrib.get("name", "")
            state = sheet.attrib.get("state", "visible")
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
            target = rel_map.get(rid, "")
            sheet_path = "xl/" + target.lstrip("/")

            dim = None
            if sheet_path in zf.namelist():
                root = ET.fromstring(zf.read(sheet_path))
                dim_node = root.find("main:dimension", NS_MAIN)
                dim = dim_node.attrib.get("ref") if dim_node is not None else None

            rows, cols = parse_dimension(dim)
            sheets.append(
                {
                    "name": name,
                    "visibility": state,
                    "dimension": dim,
                    "approx_rows": rows,
                    "approx_cols": cols,
                }
            )

        defined_names_node = workbook_xml.find("main:definedNames", NS_MAIN)
        defined_names = []
        if defined_names_node is not None:
            for dn in defined_names_node:
                defined_names.append(
                    {
                        "name": dn.attrib.get("name", ""),
                        "refers_to": (dn.text or "").strip(),
                    }
                )

        return {
            "workbook": str(path),
            "sheet_count": len(sheets),
            "sheets": sheets,
            "defined_names": defined_names,
            "has_vba_project": "xl/vbaProject.bin" in zf.namelist(),
        }


def main() -> None:
    args = parse_args()
    profile = build_profile(args.workbook)
    payload = json.dumps(profile, indent=2)

    if args.out:
        args.out.write_text(payload, encoding="utf-8")
        print(f"Profile written to: {args.out}")
    else:
        print(payload)


if __name__ == "__main__":
    main()
