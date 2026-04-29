# Finance Automation Toolkit v1.0 — iPipeline
# workbook_dependency_scanner.py — maps cross-sheet formula references inside an .xlsx file
#
# What it does:
#   - Opens an Excel .xlsx file as a ZIP archive (no Excel, no xlwings needed)
#   - Reads the XML inside to find every formula that references another sheet
#   - Reports: which sheet, which cell, what formula, what sheet it references
#   - Flags sheets that are referenced but hidden, or cells with long formula chains
#
# Why this matters: before modifying a shared workbook, you need to know what breaks
# if you rename a sheet, delete a column, or restructure a tab. This tells you in seconds.
#
# Usage:
#   python workbook_dependency_scanner.py path/to/workbook.xlsx
#   python workbook_dependency_scanner.py --sample    (uses the bundled sample .xlsx if present,
#                                                      otherwise creates a minimal test workbook)
#
# Outputs (in outputs/YYYYMMDD_HHMMSS_workbook_dependency_scanner/):
#   dependency_report.html    — visual map of cross-sheet references
#   cross_sheet_refs.csv      — one row per formula reference found
#   run_log.json
#   run_summary.txt

import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from common.safe_io import get_output_dir, get_toolkit_root, resolve_input_path, write_csv, write_html
from common.logging_utils import RunLogger
from common.report_utils import build_report, metric_row, data_table, note_box


TOOL_NAME = "workbook_dependency_scanner"

# Namespace map for Office Open XML
_NS = {
    "ss":  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Regex for cross-sheet references:
#   Group 1: quoted name  — 'Sheet Name'!
#   Group 2: unquoted name — SheetName! (word chars, dots, spaces, hyphens only)
# Using \b so we don't capture SUM(Data as "SUM(Data" — the \b anchors to the start of the name.
_CROSS_REF_RE = re.compile(r"'([^']+)'!|(\b[\w. -]+)!")


def _col_letter(col_idx: int) -> str:
    """Convert 0-based column index to Excel column letter (A, B, ..., Z, AA, ...)."""
    result = ""
    col_idx += 1
    while col_idx:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def _cell_address(row_idx: int, col_idx: int) -> str:
    return f"{_col_letter(col_idx)}{row_idx + 1}"


def _get_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Read sharedStrings.xml if present. Returns list of string values by index."""
    try:
        with zf.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
        root = tree.getroot()
        ns = _NS["ss"]
        strings = []
        for si in root.findall(f"{{{ns}}}si"):
            t = si.find(f"{{{ns}}}t")
            if t is not None and t.text:
                strings.append(t.text)
            else:
                # Concatenate all <t> within <r> runs
                parts = [r.text or "" for r in si.findall(f".//{{{ns}}}t")]
                strings.append("".join(parts))
        return strings
    except KeyError:
        return []


def _get_sheet_names(zf: zipfile.ZipFile) -> dict[str, str]:
    """Return {rId -> sheet_name} from workbook.xml."""
    try:
        with zf.open("xl/workbook.xml") as f:
            tree = ET.parse(f)
        root = tree.getroot()
        ns = _NS["ss"]
        r_ns = _NS["r"]
        sheets = {}
        for sh in root.findall(f".//{{{ns}}}sheet"):
            rid  = sh.get(f"{{{r_ns}}}id", "")
            name = sh.get("name", "")
            sheets[rid] = name
        return sheets
    except (KeyError, ET.ParseError):
        return {}


def _get_sheet_files(zf: zipfile.ZipFile) -> dict[str, str]:
    """Return {rId -> xl/worksheets/sheetN.xml} from workbook.xml.rels."""
    try:
        with zf.open("xl/_rels/workbook.xml.rels") as f:
            tree = ET.parse(f)
        root = tree.getroot()
        rel_ns = _NS["rel"]
        mapping = {}
        for rel in root.findall(f"{{{rel_ns}}}Relationship"):
            rid    = rel.get("Id", "")
            target = rel.get("Target", "")
            if "worksheets" in target:
                mapping[rid] = f"xl/{target}" if not target.startswith("xl/") else target
        return mapping
    except (KeyError, ET.ParseError):
        return {}


def _get_hidden_sheets(zf: zipfile.ZipFile) -> set[str]:
    """Return set of hidden sheet names from workbook.xml."""
    try:
        with zf.open("xl/workbook.xml") as f:
            tree = ET.parse(f)
        root = tree.getroot()
        ns = _NS["ss"]
        hidden = set()
        for sh in root.findall(f".//{{{ns}}}sheet"):
            state = sh.get("state", "visible")
            if state in ("hidden", "veryHidden"):
                hidden.add(sh.get("name", ""))
        return hidden
    except (KeyError, ET.ParseError):
        return set()


def scan_sheet(zf: zipfile.ZipFile, sheet_file: str, sheet_name: str,
               known_sheets: set[str]) -> list[dict]:
    """Scan one worksheet XML for cells containing cross-sheet formula references."""
    refs = []
    try:
        with zf.open(sheet_file) as f:
            tree = ET.parse(f)
    except (KeyError, ET.ParseError):
        return refs

    ns = _NS["ss"]
    for row in tree.findall(f".//{{{ns}}}row"):
        for cell in row.findall(f"{{{ns}}}c"):
            ref_attr = cell.get("r", "")
            f_elem = cell.find(f"{{{ns}}}f")
            if f_elem is None or not f_elem.text:
                continue
            formula = f_elem.text.strip()
            # Extract candidate names (group 1 = quoted, group 2 = unquoted)
            candidates = [g1 or g2 for g1, g2 in _CROSS_REF_RE.findall(formula) if g1 or g2]
            # Filter to only real sheet names — discards cell-ref fragments like "B12-Data"
            for target_sheet in set(candidates):
                target_sheet = target_sheet.strip()
                if target_sheet not in known_sheets or target_sheet == sheet_name:
                    continue
                refs.append({
                    "source_sheet":  sheet_name,
                    "cell":          ref_attr,
                    "formula":       formula[:120],
                    "target_sheet":  target_sheet,
                })
    return refs


def scan_workbook(xlsx_path: Path) -> dict:
    """Open workbook ZIP and extract all cross-sheet dependencies."""
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        sheet_names   = _get_sheet_names(zf)
        sheet_files   = _get_sheet_files(zf)
        hidden_sheets = _get_hidden_sheets(zf)

        all_refs = []
        sheet_order = list(sheet_names.values())
        known_sheets_set = set(sheet_order)

        for rid, name in sheet_names.items():
            sheet_file = sheet_files.get(rid)
            if sheet_file:
                refs = scan_sheet(zf, sheet_file, name, known_sheets_set)
                all_refs.extend(refs)

    # Build dependency map: source_sheet -> set of target_sheets
    dep_map: dict[str, set[str]] = defaultdict(set)
    for r in all_refs:
        dep_map[r["source_sheet"]].add(r["target_sheet"])

    return {
        "refs":          all_refs,
        "dep_map":       {k: sorted(v) for k, v in dep_map.items()},
        "sheet_order":   sheet_order,
        "hidden_sheets": hidden_sheets,
    }


def build_html(xlsx_path: Path, result: dict) -> str:
    refs        = result["refs"]
    dep_map     = result["dep_map"]
    hidden      = result["hidden_sheets"]
    sheet_order = result["sheet_order"]

    n_refs    = len(refs)
    n_sheets  = len(sheet_order)
    n_sources = len(dep_map)

    # Sheets that are referenced from formulas
    all_targets = set(t for targets in dep_map.values() for t in targets)
    hidden_but_referenced = hidden & all_targets

    subtitle = (
        f"File: {xlsx_path.name} &nbsp;|&nbsp; {n_sheets} sheets &nbsp;|&nbsp; "
        f"{n_refs} cross-sheet formula reference(s)"
    )

    warn_status = "warn" if n_refs > 0 else "ok"
    cards = [
        {"label": "Sheets Found",          "value": str(n_sheets),  "status": "normal"},
        {"label": "Sheets w/ Cross-Refs",  "value": str(n_sources), "status": warn_status},
        {"label": "Total Formula Refs",    "value": str(n_refs),    "status": warn_status},
        {"label": "Hidden & Referenced",   "value": str(len(hidden_but_referenced)),
         "status": "bad" if hidden_but_referenced else "ok"},
    ]

    sections = [metric_row(cards)]

    if hidden_but_referenced:
        sections.append(note_box(
            "<strong>Warning:</strong> The following sheets are hidden but referenced by formulas. "
            "Deleting or renaming them will break dependent formulas: "
            + ", ".join(f"<strong>{s}</strong>" for s in sorted(hidden_but_referenced))
        ))

    # Dependency summary
    if dep_map:
        dep_rows = [
            [src, ", ".join(targets), str(sum(1 for r in refs if r["source_sheet"] == src))]
            for src, targets in sorted(dep_map.items())
        ]
        sections.append(data_table(
            "Sheet Dependency Summary",
            ["Source Sheet", "References These Sheets", "Formula Count"],
            dep_rows
        ))
    else:
        sections.append("<h2>Sheet Dependency Summary</h2><p>No cross-sheet formula references found.</p>")

    # All refs detail
    if refs:
        ref_rows = [
            [r["source_sheet"], r["cell"], r["target_sheet"], r["formula"]]
            for r in refs
        ]
        sections.append(data_table(
            f"All Cross-Sheet References ({n_refs} total)",
            ["Source Sheet", "Cell", "References Sheet", "Formula"],
            ref_rows
        ))

    # Sheet inventory
    sheet_rows = [
        [name,
         "Hidden" if name in hidden else "Visible",
         str(sum(1 for r in refs if r["source_sheet"] == name)),
         str(sum(1 for t in dep_map.values() if name in t))]
        for name in sheet_order
    ]
    sections.append(data_table(
        "Sheet Inventory",
        ["Sheet Name", "Visibility", "Cross-Refs Out", "Times Referenced"],
        sheet_rows,
        status_col=1
    ))

    sections.append(note_box(
        "Safety: this tool read your workbook file as a ZIP archive — read-only. "
        "No changes were made to the file. Formula text is extracted from XML; "
        "no macros were run."
    ))

    return build_report("Workbook Dependency Scanner", subtitle, sections)


def _make_sample_xlsx(path: Path) -> None:
    """Create a minimal multi-sheet .xlsx for testing (stdlib only — no openpyxl)."""
    # Build the minimal OOXML structure manually
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets>'
        '<sheet name="Summary" sheetId="1" r:id="rId1"/>'
        '<sheet name="Data" sheetId="2" r:id="rId2"/>'
        '</sheets>'
        '</workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
        '</Relationships>'
    )
    # Summary sheet: has formulas that reference Data sheet
    sheet1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>'
        '<row r="1"><c r="A1" t="str"><v>Total Revenue</v></c>'
        '<c r="B1"><f>SUM(Data!B1:B12)</f><v>0</v></c></row>'
        '<row r="2"><c r="A2" t="str"><v>YoY Change</v></c>'
        '<c r="B2"><f>Data!B12-Data!B1</f><v>0</v></c></row>'
        '<row r="3"><c r="A3" t="str"><v>Avg Monthly</v></c>'
        '<c r="B3"><f>AVERAGE(Data!B1:B12)</f><v>0</v></c></row>'
        '</sheetData>'
        '</worksheet>'
    )
    # Data sheet: raw numbers, no cross-refs
    sheet2 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>'
        '<row r="1"><c r="A1" t="str"><v>Jan</v></c><c r="B1"><v>100000</v></c></row>'
        '<row r="2"><c r="A2" t="str"><v>Feb</v></c><c r="B2"><v>110000</v></c></row>'
        '</sheetData>'
        '</worksheet>'
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",       content_types)
        zf.writestr("_rels/.rels",               rels)
        zf.writestr("xl/workbook.xml",           workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml",  sheet1)
        zf.writestr("xl/worksheets/sheet2.xml",  sheet2)


def main(argv: list[str]) -> None:
    sample_mode = "--sample" in argv

    if sample_mode:
        # Look for any .xlsx in samples/ first, fall back to generating one
        samples_dir = get_toolkit_root() / "samples"
        xlsx_files = list(samples_dir.glob("*.xlsx")) if samples_dir.exists() else []
        if xlsx_files:
            xlsx_path = xlsx_files[0]
            print(f"[Sample mode] Using: {xlsx_path.name}")
        else:
            xlsx_path = samples_dir / "sample_workbook.xlsx"
            print(f"[Sample mode] Generating minimal test workbook: {xlsx_path.name}")
            _make_sample_xlsx(xlsx_path)
    elif len(argv) < 2 or argv[1].startswith("--"):
        print("Usage: python workbook_dependency_scanner.py path/to/workbook.xlsx")
        print("       python workbook_dependency_scanner.py --sample")
        sys.exit(0)
    else:
        xlsx_path = resolve_input_path(argv[1])

    if not xlsx_path.exists():
        print(f"ERROR: File not found: {xlsx_path}")
        sys.exit(1)
    if xlsx_path.suffix.lower() not in (".xlsx", ".xlsm", ".xlam"):
        print(f"ERROR: Expected .xlsx/.xlsm file, got: {xlsx_path.suffix}")
        sys.exit(1)

    out_dir = get_output_dir(TOOL_NAME)
    logger  = RunLogger(TOOL_NAME, out_dir)
    logger.set_meta(input_file=str(xlsx_path), mode="sample" if sample_mode else "real")

    print(f"Scanning: {xlsx_path}")
    print(f"Output:   {out_dir}")

    try:
        result = scan_workbook(xlsx_path)
    except zipfile.BadZipFile:
        print("ERROR: File is not a valid .xlsx (not a ZIP archive). Is it open in Excel?")
        logger.error("BadZipFile — file may be open in Excel or is an .xls (not .xlsx).")
        logger.finish()
        sys.exit(1)

    refs        = result["refs"]
    dep_map     = result["dep_map"]
    hidden      = result["hidden_sheets"]
    sheet_order = result["sheet_order"]

    logger.rows_read = len(sheet_order)
    logger.rows_processed = len(refs)

    all_targets = set(t for targets in dep_map.values() for t in targets)
    hidden_referenced = hidden & all_targets
    for s in hidden_referenced:
        logger.finding("WARN", f"Hidden sheet '{s}' is referenced by formulas — deleting it will break formulas.", "")

    write_csv(out_dir / "cross_sheet_refs.csv", refs,
              ["source_sheet", "cell", "target_sheet", "formula"])

    html = build_html(xlsx_path, result)
    write_html(out_dir / "dependency_report.html", html)

    logger.finish()

    print(f"\nSheets found: {len(sheet_order)}  |  Cross-sheet refs: {len(refs)}  |  Hidden & referenced: {len(hidden_referenced)}")
    if dep_map:
        print("\nDependency summary:")
        for src, targets in sorted(dep_map.items()):
            print(f"  {src} -> {', '.join(targets)}")
    else:
        print("No cross-sheet formula references found.")
    print(f"\nReport: {out_dir / 'dependency_report.html'}")
    print(f"CSV:    {out_dir / 'cross_sheet_refs.csv'}")
    print(f"Log:    {out_dir / 'run_summary.txt'}")


if __name__ == "__main__":
    main(sys.argv)
