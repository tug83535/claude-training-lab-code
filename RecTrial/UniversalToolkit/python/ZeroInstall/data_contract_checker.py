# Finance Automation Toolkit v1.0 — iPipeline
# data_contract_checker.py — validates CSV structure and quality before analysis begins
#
# What it checks:
#   - File is readable and not empty
#   - No completely blank rows
#   - Columns that look like IDs or names have no blanks
#   - Columns that look like amounts/fees are numeric (no text sneaking in)
#   - Columns that look like dates parse correctly
#   - No exact duplicate rows
#   - Business rules specific to contract/billing files (base_quantity=0, etc.)
#   - Reports PASS / WARN / FAIL per check in an HTML report
#
# Usage:
#   python data_contract_checker.py path/to/your_file.csv
#   python data_contract_checker.py --sample        (uses contracts_sample.csv)
#
# Outputs (in outputs/YYYYMMDD_HHMMSS_data_contract_checker/):
#   contract_check_report.html    — check results with PASS/WARN/FAIL badges
#   issues_detail.csv             — one row per failed check
#   run_log.json
#   run_summary.txt

import sys
import re
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
from common.safe_io import get_output_dir, get_samples_dir, read_csv_safe, write_csv, write_html
from common.logging_utils import RunLogger
from common.report_utils import build_report, metric_row, data_table, note_box


TOOL_NAME = "data_contract_checker"

# Patterns for auto-detecting column semantics
_AMOUNT_COLS  = re.compile(r"\b(amount|fee|revenue|price|rate|cost|billed|charge)\b", re.I)
_DATE_COLS    = re.compile(r"\b(date|_start|_end|_period|_term)\b", re.I)
_ID_COLS      = re.compile(r"\b(id|name|customer|client|vendor|account)\b", re.I)
_NUMERIC_PATS = re.compile(r"^-?\$?[\d,]+(\.\d+)?$")


def _is_numeric(val: str) -> bool:
    return bool(_NUMERIC_PATS.match(val.replace(",", "").strip()))


def _is_date(val: str) -> bool:
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m", "%m-%Y"):
        try:
            datetime.strptime(val.strip(), fmt)
            return True
        except ValueError:
            continue
    return False


def run_checks(rows: list[dict], filepath: Path) -> list[dict]:
    """Run all checks. Returns list of check-result dicts."""
    results = []

    def chk(name: str, status: str, detail: str, count: int = 0):
        results.append({
            "check": name,
            "status": status,
            "detail": detail,
            "affected_rows": count,
        })

    if not rows:
        chk("File has data rows", "FAIL", "File is empty or header-only.", 0)
        return results

    headers = list(rows[0].keys())
    n = len(rows)

    # ── Check 1: row count ──────────────────────────────────────────────────
    if n < 5:
        chk("Minimum row count", "WARN",
            f"Only {n} data row(s). This may be a test file or a filtered export.", n)
    else:
        chk("Minimum row count", "PASS", f"{n} rows found.", n)

    # ── Check 2: blank rows ─────────────────────────────────────────────────
    blank = [i + 2 for i, r in enumerate(rows)
             if all(not v.strip() for v in r.values())]
    if blank:
        chk("No completely blank rows", "FAIL",
            f"{len(blank)} blank row(s) found at row(s): {blank[:5]}{'...' if len(blank)>5 else ''}",
            len(blank))
    else:
        chk("No completely blank rows", "PASS", "No blank rows detected.")

    # ── Check 3: ID/name columns not blank ──────────────────────────────────
    id_cols = [h for h in headers if _ID_COLS.search(h)]
    for col in id_cols:
        blanks = [i + 2 for i, r in enumerate(rows) if not r.get(col, "").strip()]
        if blanks:
            chk(f"No blanks in '{col}'", "FAIL",
                f"{len(blanks)} blank value(s) in '{col}'. First at row {blanks[0]}.",
                len(blanks))
        else:
            chk(f"No blanks in '{col}'", "PASS",
                f"All {n} rows have a value in '{col}'.")

    # ── Check 4: amount columns are numeric ─────────────────────────────────
    amt_cols = [h for h in headers if _AMOUNT_COLS.search(h)]
    for col in amt_cols:
        non_num = [i + 2 for i, r in enumerate(rows)
                   if r.get(col, "").strip() and not _is_numeric(r.get(col, ""))]
        if non_num:
            chk(f"'{col}' is numeric", "FAIL",
                f"{len(non_num)} non-numeric value(s) in '{col}'. First at row {non_num[0]}.",
                len(non_num))
        else:
            chk(f"'{col}' is numeric", "PASS",
                f"All non-blank values in '{col}' are numeric.")

    # ── Check 5: date columns parse ─────────────────────────────────────────
    date_cols = [h for h in headers if _DATE_COLS.search(h) and h not in amt_cols]
    for col in date_cols:
        bad_dates = [i + 2 for i, r in enumerate(rows)
                     if r.get(col, "").strip() and not _is_date(r.get(col, ""))]
        if bad_dates:
            chk(f"'{col}' parses as date", "WARN",
                f"{len(bad_dates)} value(s) in '{col}' do not parse as a date. "
                f"First at row {bad_dates[0]}. Check the format.",
                len(bad_dates))
        else:
            chk(f"'{col}' parses as date", "PASS",
                f"All non-blank values in '{col}' are valid dates.")

    # ── Check 6: duplicate rows ──────────────────────────────────────────────
    seen: set[tuple] = set()
    dup_rows: list[int] = []
    for i, r in enumerate(rows):
        key = tuple(r.values())
        if key in seen:
            dup_rows.append(i + 2)
        seen.add(key)
    if dup_rows:
        chk("No exact duplicate rows", "WARN",
            f"{len(dup_rows)} row(s) are exact duplicates. First at row {dup_rows[0]}.",
            len(dup_rows))
    else:
        chk("No exact duplicate rows", "PASS", "No exact duplicate rows found.")

    # ── Check 7: base_quantity = 0 (contract-specific business rule) ─────────
    if "base_quantity" in headers:
        zero_qty = [i + 2 for i, r in enumerate(rows)
                    if r.get("base_quantity", "").strip() == "0"]
        if zero_qty:
            chk("base_quantity not zero", "FAIL",
                f"{len(zero_qty)} contract(s) have base_quantity = 0. "
                f"Every transaction will bill at overage rates — confirm this is intentional. "
                f"Rows: {zero_qty[:5]}{'...' if len(zero_qty)>5 else ''}",
                len(zero_qty))
        else:
            chk("base_quantity not zero", "PASS",
                "All contracts have a non-zero base_quantity.")

    # ── Check 8: term_end not blank for active contracts ────────────────────
    if "term_end" in headers and "status" in headers:
        active_no_end = [i + 2 for i, r in enumerate(rows)
                         if r.get("status", "").strip().lower() == "active"
                         and not r.get("term_end", "").strip()]
        if active_no_end:
            chk("Active contracts have term_end", "WARN",
                f"{len(active_no_end)} active contract(s) are missing term_end. "
                f"Rows: {active_no_end[:5]}",
                len(active_no_end))
        else:
            chk("Active contracts have term_end", "PASS",
                "All active contracts have a term_end date.")

    return results


def build_html(filepath: Path, rows: list[dict], checks: list[dict]) -> str:
    n_pass = sum(1 for c in checks if c["status"] == "PASS")
    n_warn = sum(1 for c in checks if c["status"] == "WARN")
    n_fail = sum(1 for c in checks if c["status"] == "FAIL")
    overall = "FAIL" if n_fail > 0 else ("WARN" if n_warn > 0 else "PASS")
    overall_status = "bad" if overall == "FAIL" else ("warn" if overall == "WARN" else "ok")

    subtitle = f"File: {filepath.name} &nbsp;|&nbsp; {len(rows)} rows &nbsp;|&nbsp; {len(rows[0]) if rows else 0} columns"

    cards = [
        {"label": "Overall", "value": overall, "status": overall_status},
        {"label": "Checks Passed", "value": str(n_pass), "status": "ok" if n_pass > 0 else "normal"},
        {"label": "Warnings",      "value": str(n_warn), "status": "warn" if n_warn > 0 else "normal"},
        {"label": "Failures",      "value": str(n_fail), "status": "bad"  if n_fail > 0 else "normal"},
        {"label": "Data Rows",     "value": str(len(rows)), "status": "normal"},
    ]

    check_rows = [[c["check"], c["status"], str(c["affected_rows"]), c["detail"]] for c in checks]
    sections = [
        metric_row(cards),
        data_table("Check Results", ["Check", "Status", "Affected Rows", "Detail"],
                   check_rows, status_col=1),
        note_box(
            "Safety reminder: this tool opens your file read-only. "
            "No changes have been made to the input file. "
            "All output is in the outputs/ folder."
        ),
    ]
    return build_report("Data Contract Checker", subtitle, sections)


def main(argv: list[str]) -> None:
    sample_mode = "--sample" in argv
    if sample_mode:
        filepath = get_samples_dir() / "contracts_sample.csv"
        print(f"[Sample mode] Using: {filepath.name}")
    elif len(argv) < 2 or argv[1].startswith("--"):
        print("Usage: python data_contract_checker.py path/to/file.csv")
        print("       python data_contract_checker.py --sample")
        sys.exit(0)
    else:
        raw = " ".join(a for a in argv[1:] if not a.startswith("--"))
        from common.safe_io import resolve_input_path
        filepath = resolve_input_path(raw)

    out_dir = get_output_dir(TOOL_NAME)
    logger  = RunLogger(TOOL_NAME, out_dir)
    logger.set_meta(input_file=str(filepath), mode="sample" if sample_mode else "real")

    print(f"Checking: {filepath}")
    print(f"Output:   {out_dir}")

    try:
        rows = read_csv_safe(filepath)
        logger.rows_read = len(rows)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        logger.error(str(e))
        logger.finish()
        sys.exit(1)

    checks = run_checks(rows, filepath)
    logger.rows_processed = len(rows)

    n_fail = sum(1 for c in checks if c["status"] == "FAIL")
    n_warn = sum(1 for c in checks if c["status"] == "WARN")
    for c in checks:
        if c["status"] == "FAIL":
            logger.finding("FAIL", c["check"] + ": " + c["detail"][:120], str(c["affected_rows"]) + " rows")
        elif c["status"] == "WARN":
            logger.finding("WARN", c["check"] + ": " + c["detail"][:120], str(c["affected_rows"]) + " rows")

    html = build_html(filepath, rows, checks)
    write_html(out_dir / "contract_check_report.html", html)

    issues = [c for c in checks if c["status"] in ("FAIL", "WARN")]
    write_csv(out_dir / "issues_detail.csv", issues,
              ["check", "status", "affected_rows", "detail"])

    logger.finish()

    overall = "FAIL" if n_fail > 0 else ("WARN" if n_warn > 0 else "PASS")
    print(f"\nResult: {overall}  ({n_fail} failure(s), {n_warn} warning(s))")
    print(f"Report: {out_dir / 'contract_check_report.html'}")
    print(f"Log:    {out_dir / 'run_summary.txt'}")


if __name__ == "__main__":
    main(sys.argv)
