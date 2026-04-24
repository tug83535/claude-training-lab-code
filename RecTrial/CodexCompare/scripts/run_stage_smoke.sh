#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

echo "[1/3] Running repository smoke checks"
python tests/stage2_smoke_check.py

echo "[2/3] Running basic Python syntax checks"
python -m py_compile \
  python/universal/profile_workbook.py \
  python/universal/sanitize_dataset.py \
  python/universal/compare_workbooks.py \
  python/universal/build_exec_summary.py \
  python/demo/pnl_data_extract.py \
  python/demo/variance_classifier.py \
  python/demo/scenario_runner.py \
  python/demo/export_brief_package.py \
  tests/stage2_smoke_check.py

echo "[3/3] Running Python unit tests"
python -m unittest tests/test_python_utilities.py

echo "All stage smoke checks passed."
