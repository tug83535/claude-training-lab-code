# README_VIDEO4_PYTHON

Version: 1.0.0

This folder contains the **6 Video 4 ZeroInstall scripts** used in the launcher.

## 1) sanitize_dataset.py
- Plain English: Cleans messy CSV values (spaces, dates, numbers).
- Input: `input_csv` (required existing CSV file).
- Output: `output_csv` cleaned CSV.
- Limitation: CSV only (not Excel workbook files).

## 2) variance_classifier.py
- Plain English: Compares Actual vs Baseline and labels each row.
- Input: `input_csv` with Actual/Baseline columns.
- Output: `output_csv` with Variance, VariancePct, Direction, Materiality.
- Limitation: Needs numeric values in the comparison columns.

## 3) scenario_runner.py
- Plain English: Runs multiple what-if percentage scenarios on one metric column.
- Input: `input_csv` with metric column (default `Amount`).
- Output: `output_csv` summary of each scenario.
- Limitation: Only simple percentage shock modeling.

## 4) build_exec_summary.py
- Plain English: Produces a short executive summary markdown from CSV data.
- Input: `input_csv` with at least one numeric column.
- Output: Printed summary or `--out` markdown file.
- Limitation: Heuristic numeric-column detection.

## 5) compare_workbooks.py
- Plain English: Compares two workbooks and lists changed cells.
- Input: `left_workbook`, `right_workbook`.
- Output: `out_csv` containing sheet/cell differences.
- Limitation: Works on xlsx/xlsm XML structure only.

## 6) sheets_to_csv.py
- Plain English: Pulls selected sheets from a workbook into CSV files.
- Input: `workbook` and `--out-dir`.
- Output: One CSV per requested sheet.
- Limitation: If a requested sheet name is missing, it is skipped.

## Safety rules for coworkers
- Always run `--help` first.
- Keep source files read-only backups before running scripts.
- Confirm the input path exists before running.
- Use a dedicated output folder per run.
