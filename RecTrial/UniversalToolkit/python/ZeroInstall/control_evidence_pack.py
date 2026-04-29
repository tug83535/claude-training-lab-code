# Finance Automation Toolkit v1.0 — iPipeline
# control_evidence_pack.py — creates a tamper-evident evidence bundle from analysis outputs
#
# What it does:
#   - Scans a folder of analysis outputs (or the most recent run automatically)
#   - Records each file: name, size, last-modified timestamp, SHA-256 hash
#   - Writes a manifest CSV and a one-page HTML evidence summary
#   - Produces an evidence_readme.txt that can be attached to a ticket or email
#
# Why this matters: if someone asks "what files did you analyze and when?", this folder
# answers that question precisely. The SHA-256 hash proves the file was not changed after
# the analysis ran.
#
# Usage:
#   python control_evidence_pack.py --sample
#       (scans the most recent Revenue Leakage Finder run)
#   python control_evidence_pack.py --input-dir path/to/output/folder
#   python control_evidence_pack.py --input-dir path/to/output/folder --control-name "Q2 Revenue Review"
#
# Outputs (in outputs/YYYYMMDD_HHMMSS_control_evidence_pack/):
#   manifest.csv          — one row per file: name, size, modified, sha256
#   evidence_summary.html — one-page iPipeline-branded summary
#   evidence_readme.txt   — plain-text summary suitable for email attachment
#   run_log.json
#   run_summary.txt

import hashlib
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from common.safe_io import get_output_dir, get_toolkit_root, resolve_input_path, write_csv, write_html, write_text
from common.logging_utils import RunLogger
from common.report_utils import build_report, metric_row, data_table, note_box


TOOL_NAME = "control_evidence_pack"


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:,.0f} {unit}"
        n //= 1024
    return f"{n:,.0f} TB"


def scan_folder(folder: Path) -> list[dict]:
    """Scan all files in folder (non-recursive). Returns list of file-info dicts."""
    records = []
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        stat = path.stat()
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        size_bytes = stat.st_size
        sha = _sha256(path)
        records.append({
            "file_name":   path.name,
            "size_bytes":  str(size_bytes),
            "size_human":  _human_size(size_bytes),
            "modified":    modified,
            "sha256":      sha,
            "full_path":   str(path),
        })
    return records


def build_html(control_name: str, input_folder: Path, records: list[dict],
               run_ts: str) -> str:
    n = len(records)
    total_bytes = sum(int(r["size_bytes"]) for r in records)

    subtitle = (
        f"Control: {control_name} &nbsp;|&nbsp; "
        f"Source folder: {input_folder.name} &nbsp;|&nbsp; "
        f"{n} file(s) &nbsp;|&nbsp; {_human_size(total_bytes)} total"
    )

    cards = [
        {"label": "Files Inventoried", "value": str(n), "status": "ok"},
        {"label": "Total Size",        "value": _human_size(total_bytes), "status": "normal"},
        {"label": "Hash Algorithm",    "value": "SHA-256", "status": "ok"},
        {"label": "Analysis Date",     "value": run_ts[:10], "status": "normal"},
    ]

    table_rows = [
        [r["file_name"], r["size_human"], r["modified"], r["sha256"][:16] + "..."]
        for r in records
    ]
    file_table = data_table(
        "File Inventory",
        ["File Name", "Size", "Last Modified", "SHA-256 (first 16 chars)"],
        table_rows
    )

    full_hash_rows = [[r["file_name"], r["sha256"]] for r in records]
    hash_table = data_table(
        "Full SHA-256 Hashes",
        ["File Name", "SHA-256"],
        full_hash_rows
    )

    guidance = note_box(
        "<strong>What SHA-256 proves:</strong> If a file's hash changes after this report was generated, "
        "someone modified the file. To verify: re-run SHA-256 on the file and compare to the value above. "
        "This report was generated read-only — no input files were modified."
    )

    sections = [metric_row(cards), guidance, file_table, hash_table]
    return build_report(f"Control Evidence Pack — {control_name}", subtitle, sections)


def build_readme(control_name: str, input_folder: Path, records: list[dict],
                 run_ts: str, out_dir: Path) -> str:
    lines = [
        f"Control Evidence Pack",
        f"=" * 50,
        f"Control name:   {control_name}",
        f"Analysis date:  {run_ts}",
        f"Source folder:  {input_folder}",
        f"Output folder:  {out_dir}",
        f"Files included: {len(records)}",
        f"",
        f"FILES AND HASHES (SHA-256)",
        f"-" * 50,
    ]
    for r in records:
        lines.append(f"  {r['file_name']}")
        lines.append(f"    Size:     {r['size_human']}")
        lines.append(f"    Modified: {r['modified']}")
        lines.append(f"    SHA-256:  {r['sha256']}")
        lines.append("")
    lines += [
        f"HOW TO VERIFY",
        f"-" * 50,
        f"To confirm a file was not changed after this evidence pack was created:",
        f"  Python: import hashlib; print(hashlib.sha256(open(FILE,'rb').read()).hexdigest())",
        f"  Compare the result against the SHA-256 value above.",
        f"",
        f"If the hash matches: the file is unchanged.",
        f"If the hash differs: the file was modified after this analysis ran.",
        f"",
        f"Questions? Contact Connor Atlee — Finance & Accounting.",
    ]
    return "\n".join(lines)


def main(argv: list[str]) -> None:
    sample_mode = "--sample" in argv

    # Extract optional control name
    control_name = "Revenue Leakage Review"
    for i, arg in enumerate(argv):
        if arg == "--control-name" and i + 1 < len(argv):
            control_name = argv[i + 1]

    if sample_mode:
        # Use most recent Revenue Leakage Finder run
        outputs_root = get_toolkit_root() / "outputs"
        leakage_dirs = sorted(
            [d for d in outputs_root.iterdir()
             if d.is_dir() and "revenue_leakage_finder" in d.name],
            reverse=True
        ) if outputs_root.exists() else []

        if leakage_dirs:
            input_folder = leakage_dirs[0]
            print(f"[Sample mode] Scanning: {input_folder.name}")
        else:
            print("No Revenue Leakage Finder output found. Run revenue_leakage_finder.py --sample first.")
            sys.exit(1)
    else:
        # Look for --input-dir flag
        input_folder = None
        for i, arg in enumerate(argv):
            if arg == "--input-dir" and i + 1 < len(argv):
                raw = argv[i + 1]
                input_folder = resolve_input_path(raw)
                break

        if input_folder is None:
            print("Usage: python control_evidence_pack.py --sample")
            print("       python control_evidence_pack.py --input-dir path/to/folder")
            print("       python control_evidence_pack.py --input-dir path/to/folder --control-name \"Q2 Review\"")
            sys.exit(0)

    if not input_folder.is_dir():
        print(f"ERROR: Not a directory: {input_folder}")
        sys.exit(1)

    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    out_dir = get_output_dir(TOOL_NAME)
    logger  = RunLogger(TOOL_NAME, out_dir)
    logger.set_meta(
        control_name=control_name,
        input_folder=str(input_folder),
        mode="sample" if sample_mode else "real",
    )

    print(f"Control:  {control_name}")
    print(f"Scanning: {input_folder}")
    print(f"Output:   {out_dir}")

    records = scan_folder(input_folder)
    logger.rows_read = len(records)
    logger.rows_processed = len(records)

    if not records:
        print("No files found in the input folder.")
        logger.warn("No files found.")
        logger.finish()
        sys.exit(0)

    for r in records:
        logger.finding("FILE", r["file_name"], f"{r['size_human']} | SHA-256: {r['sha256'][:16]}...")

    write_csv(out_dir / "manifest.csv", records,
              ["file_name", "size_bytes", "size_human", "modified", "sha256", "full_path"])

    html = build_html(control_name, input_folder, records, run_ts)
    write_html(out_dir / "evidence_summary.html", html)

    readme = build_readme(control_name, input_folder, records, run_ts, out_dir)
    write_text(out_dir / "evidence_readme.txt", readme)

    logger.finish()

    total_bytes = sum(int(r["size_bytes"]) for r in records)
    print(f"\nInventoried {len(records)} file(s)  |  Total: {_human_size(total_bytes)}")
    print(f"\nFiles hashed:")
    for r in records:
        print(f"  {r['file_name']:45s}  {r['size_human']:>8}  {r['sha256'][:16]}...")
    print(f"\nManifest:  {out_dir / 'manifest.csv'}")
    print(f"Report:    {out_dir / 'evidence_summary.html'}")
    print(f"Readme:    {out_dir / 'evidence_readme.txt'}")
    print(f"Log:       {out_dir / 'run_summary.txt'}")


if __name__ == "__main__":
    main(sys.argv)
