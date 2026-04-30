# SMOKE TEST CHECKLIST (Manual)

Version: 1.0.0

1. sanitize_dataset.py — PASS if output CSV is created and prints "Output:" path.
2. variance_classifier.py — PASS if output CSV has new columns Direction and Materiality.
3. scenario_runner.py — PASS if output CSV includes at least base/optimistic/conservative rows.
4. build_exec_summary.py — PASS if markdown summary prints or file is created with "Executive Summary" header.
5. compare_workbooks.py — PASS if diff CSV is created and has header row.
6. sheets_to_csv.py — PASS if at least one sheet CSV is created in output folder.
