# Architecture Overview

## 1) Two-prong model

This repository is organized into two major solution paths:

1. **Universal Toolkit (Prong 1)**
   - File-agnostic VBA and Python tools.
   - Intended for any workbook shape.
   - Main folders: `vba/universal/`, `python/universal/`, `sql/universal/`.

2. **Demo-Specific Automation (Prong 2)**
   - Built for the P&L demo workbook structure.
   - Includes reconciliation, narratives, scenarios, and executive brief outputs.
   - Main folder: `vba/demo/`.

## 2) Runtime flow

### Universal path
1. User launches `UTL_CommandCenter`.
2. Toolkit runs profile/sanitize/compare/intelligence/output actions.
3. Actions log into `UTL_RunLog`.

### Demo path
1. User launches `Demo_CommandCenter`.
2. Demo workflows run in sequence (reconciliation → narrative → scenario → brief).
3. Actions log to demo audit sheet plus universal log helper.

## 3) Python companion role

Python utilities are used for:
- workbook metadata profiling,
- CSV sanitation,
- workbook diffing,
- markdown summary generation.

Python runs outside Excel and feeds outputs back into finance review workflows.

## 4) SQL role

SQL templates provide reproducible extract patterns for GL and revenue source pulls.
Demo SQL queries model reconciliation and variance fact outputs for narrative-ready data shapes.

## 5) Validation strategy

- `tests/stage2_smoke_check.py`: repository integrity and artifact-level checks.
- `tests/test_python_utilities.py`: core Python behavior tests.
- `scripts/run_stage_smoke.sh`: single command to run smoke + syntax + unit tests.

## 6) Safety principles

- Non-destructive outputs by default where practical.
- Sample files remain immutable in `samples/`.
- Explicit status logging and clear completion/failure messaging.
