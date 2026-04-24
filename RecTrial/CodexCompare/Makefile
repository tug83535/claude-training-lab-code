.PHONY: smoke unit py-compile inventory check

smoke:
	bash scripts/run_stage_smoke.sh

unit:
	python -m unittest tests/test_python_utilities.py

py-compile:
	python -m py_compile \
		python/universal/profile_workbook.py \
		python/universal/sanitize_dataset.py \
		python/universal/compare_workbooks.py \
		python/universal/build_exec_summary.py \
		python/demo/pnl_data_extract.py \
		python/demo/variance_classifier.py \
		python/demo/scenario_runner.py \
		python/demo/export_brief_package.py \
		scripts/update_code_inventory.py \
		tests/stage2_smoke_check.py \
		tests/test_python_utilities.py

inventory:
	python scripts/update_code_inventory.py

check: py-compile unit inventory
