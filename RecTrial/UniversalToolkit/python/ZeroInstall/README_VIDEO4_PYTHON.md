# README_VIDEO4_PYTHON

Version: 1.0.0

All six scripts are local-only and write to timestamped folders under `outputs/`.
Each run creates:
- `run_log.json`
- `run_summary.txt`

## Scripts
1. sanitize_dataset.py — cleans CSV values; input: `input_csv` or `--sample`; output: `outputs/<timestamp>_sanitize_dataset/`.
2. variance_classifier.py — adds variance fields; input: `input_csv` or `--sample`; output in timestamped run folder.
3. scenario_runner.py — scenario totals; input: `input_csv` or `--sample`; output in timestamped run folder.
4. build_exec_summary.py — summary markdown; input: `input_csv` or `--sample`; output in timestamped run folder.
5. compare_workbooks.py — cell diff CSV; input: two workbook paths; output in timestamped run folder.
6. sheets_to_csv.py — extract requested sheets to CSV; input: workbook path; output in timestamped run folder.

## Safety notes
- No internet/API calls.
- Input files are read-only.
- Outputs never write back to input file locations.
- Use `--help` before first run.
