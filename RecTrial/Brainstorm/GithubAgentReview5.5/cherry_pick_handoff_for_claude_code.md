# Cherry-Pick Handoff: AP2 → April23CLD
> Generated: 2026-05-06  
> Purpose: Paste the **"Prompt for Claude Code"** section below directly into Claude Code while working inside the `April23CLD` / `claude-training-lab-code` repo.

---

## Background

This document was created by GitHub Copilot after a cross-repo audit comparing `tug83535/AP2` and `tug83535/claude-training-lab-code` (April23CLD branch).  
The goal: identify what AP2 has built that April23CLD should adopt.

---

## What AP2 Has That April23CLD Needs

### 1. `shared/python/` — the utility layer (highest priority)

| File | What it provides |
|---|---|
| `output_manager.py` | Creates `outputs/<slug>/<UTC>/` folder + writes `output_manifest.txt` automatically |
| `run_logger.py` | Writes `run_log.csv` per run (tool, input path, output path, validation pass/fail, notes) |
| `validation_utils.py` | `validate_file_exists()` and `validate_required_columns()` — guards against silent crashes |
| `backup_utils.py` | `backup_file()` with size-check assertion before overwriting anything |
| `excel_io_helpers.py` | `assert_excel_path()` — prevents silently opening `.csv` as `.xlsx` |

**All five files are stdlib-only. No new pip dependencies required.**

### 2. `tools/python/safety/one_click_backup_utility/`
Full folder backup tool. April23CLD has no equivalent. Critical before running any script against real files.

### 3. `tools/python/reporting/csv_batch_combiner/`
Combines multiple CSVs into one output. April23CLD has no equivalent.

### 4. `tools/python/safety/run_log_generator/`
Standalone tool to write a run log on demand. Pairs with `shared/python/run_logger.py`.

### 5. `tools/tool_runner.py`
Central dispatcher. Lets you run `python tool_runner.py <tool_name>` instead of hunting for individual `.py` files. Worth adding if April23CLD will keep growing.

### 6. Test patterns from `tests/`
Especially `test_output_contracts.py` — verifies every tool produces a manifest + run_log. A lightweight CI safety net.

---

## Priority Order

| Priority | What to copy | Why |
|---|---|---|
| **1** | `shared/python/` (all 5 files) | Foundation everything else depends on |
| **2** | `one_click_backup_utility` | Safety before touching real files |
| **3** | `csv_batch_combiner` | Fills a real gap |
| **4** | `run_log_generator` | Traceability — pairs with shared/python |
| **5** | `tool_runner.py` | Quality of life as toolkit grows |
| **6** | Test patterns | If you want CI in April23CLD |
| Hold | `verify_repo_sync`, `salesforce_export_cleaner` | Situational |

---

## After Copying `shared/python/`

Each of April23CLD's ~31 scripts will need a small wrapper (~15 lines) to wire into `output_manager` + `run_logger`. This is the main work — not the file copy itself.

Pattern to add at the top of each script:
```python
from shared.python.output_manager import init_output_dir
from shared.python.run_logger import log_run
from shared.python.validation_utils import validate_file_exists, validate_required_columns

out_dir = init_output_dir("tool_slug_here")
# ... existing script logic ...
log_run("tool_slug_here", input_path, out_dir, validation_passed=True)
```

---

## Prompt for Claude Code
> Copy everything between the lines below and paste it into Claude Code while inside the `April23CLD` repo.

---
```
Hey Claude Code — I need you to cherry-pick files from a sibling repo (tug83535/AP2) into this repo (tug83535/claude-training-lab-code, April23CLD branch).

Here is the priority-ordered list of what to copy over:

STEP 1 — Copy the entire `shared/python/` folder from AP2 into this repo at the same path: `shared/python/`
Files to copy:
  - shared/python/output_manager.py
  - shared/python/run_logger.py
  - shared/python/validation_utils.py
  - shared/python/backup_utils.py
  - shared/python/excel_io_helpers.py
  - shared/python/approved_packages.md  (if it exists)
  - shared/python/__init__.py  (if it exists)

STEP 2 — Copy these tool folders from AP2 into this repo under `tools/python/`:
  - tools/python/safety/one_click_backup_utility/
  - tools/python/safety/run_log_generator/
  - tools/python/reporting/csv_batch_combiner/

STEP 3 — Copy `tools/tool_runner.py` from AP2 into this repo at `tools/tool_runner.py`

After copying, do NOT modify any existing scripts in April23CLD yet. 
Just confirm what was copied and ask me if I want to proceed with wiring the shared utilities into the existing scripts.

The AP2 repo is at: https://github.com/tug83535/AP2
The files to copy are on the default branch (main).
```
---

That's the full handoff. Good luck!
