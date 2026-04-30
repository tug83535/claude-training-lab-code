# Finance Automation Toolkit v1.0 — iPipeline
# common/safe_io.py — read-only input handling, timestamped output folders, CSV/JSON/text helpers
#
# Safety guarantees this module enforces:
#   - Input files are opened read-only; never modified
#   - All outputs go to a new timestamped subfolder of outputs/
#   - Nothing is ever written to the folder where the input file lives

import csv
import json
import os
from datetime import datetime
from pathlib import Path


# Toolkit root = two levels up from this file (common/ → ZeroInstall/ → ...)
_TOOLKIT_ROOT = Path(__file__).resolve().parent.parent


def get_toolkit_root() -> Path:
    return _TOOLKIT_ROOT


def get_output_dir(tool_name: str) -> Path:
    """Create and return a new timestamped output folder for one tool run.

    Pattern: outputs/YYYYMMDD_HHMMSS_toolname/
    A new folder is always created — nothing is ever overwritten.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_root = _TOOLKIT_ROOT / "outputs"
    out_root.mkdir(exist_ok=True)
    run_dir = out_root / f"{timestamp}_{tool_name}"
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def get_samples_dir() -> Path:
    """Return the samples directory path inside the toolkit root."""
    return _TOOLKIT_ROOT / "samples"


def resolve_input_path(raw_path: str) -> Path:
    """Resolve a user-supplied path, handling OneDrive-redirected desktops and env vars."""
    return Path(os.path.expandvars(os.path.expanduser(raw_path))).resolve()


def read_csv_safe(path: str | Path) -> list[dict]:
    """Read a CSV file as a list of row-dicts. Read-only — never touches the input file.

    Handles UTF-8 BOM (common in Excel exports). Strips whitespace from values.
    Raises FileNotFoundError with a plain-English message if the file is missing.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(
            f"Input file not found: {path}\n"
            "Check that the path is correct and the file has not been moved."
        )
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append({k: (v.strip() if v else v) for k, v in row.items()})
    return rows


def write_csv(path: Path, rows: list[dict], fieldnames: list[str] | None = None) -> None:
    """Write a list of row-dicts to a CSV file in the output folder."""
    if not rows:
        path.write_text("(no records)\n", encoding="utf-8")
        return
    if fieldnames is None:
        fieldnames = list(rows[0].keys())
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def write_json(path: Path, data: dict | list) -> None:
    """Write a dict or list to a JSON file in the output folder."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str)


def write_text(path: Path, text: str) -> None:
    """Write a plain-text string to a file in the output folder."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def write_html(path: Path, html: str) -> None:
    """Write an HTML string to a file in the output folder."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
