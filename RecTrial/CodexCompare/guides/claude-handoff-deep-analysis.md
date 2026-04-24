# Claude Code Handoff — Deep Analysis Package

This document is designed to give another AI coding system (Claude Code) full working context for deep comparative analysis.

## 1) Repository Intent

This project is a Finance & Accounting automation demo with two layers:

1. **Universal toolkit** (reusable in arbitrary Excel files):
   - VBA modules in `vba/universal/`
   - Python utilities in `python/universal/`
   - SQL templates in `sql/universal/`

2. **Demo-specific workflow** (tailored to sample P&L workbook):
   - VBA modules in `vba/demo/`
   - Python demo utilities in `python/demo/`
   - SQL demo scripts in `sql/demo/`

## 2) Key Architecture and Runtime Flow

### A) Universal workflow
- User runs command center / utilities in Excel (`vba/universal/*`) for sanitization, compare/consolidate, intelligence flags, and output packs.
- For larger data tasks or outside-Excel processing, user runs Python utilities (`python/universal/*`) to profile, sanitize, compare, and summarize.

### B) Demo workflow
- User runs demo command center and scripts (`vba/demo/*`) on demo workbook copies.
- Python demo tools (`python/demo/*`) support extraction/classification/scenario/brief generation.

### C) Validation & governance
- Repository integrity checks are centralized in `tests/stage2_smoke_check.py`.
- Unit tests for Python utilities are in `tests/test_python_utilities.py`.
- Local runner: `scripts/run_stage_smoke.sh`.
- CI runner: `.github/workflows/smoke-check.yml`.
- Code inventory: `scripts/update_code_inventory.py` → `CODE_INVENTORY.md`.

## 3) Important File Map (by area)

### Core project control
- `README.md` — project framing, stage status, run instructions.
- `PLAN.md` — staged implementation plan and rationale.
- `PROJECT_TODO.md` — active next-step checklist.
- `CHANGELOG.md` — milestone notes.
- `CONTRIBUTING.md` — contributor workflow.

### VBA (universal)
- `vba/universal/modUTL_Core.bas`
- `vba/universal/modUTL_DataSanitizer.bas`
- `vba/universal/modUTL_CommandCenter.bas`
- `vba/universal/modUTL_CompareConsolidate.bas`
- `vba/universal/modUTL_Intelligence.bas`
- `vba/universal/modUTL_OutputPack.bas`

### VBA (demo)
- `vba/demo/modDemo_Config.bas`
- `vba/demo/modDemo_AuditTrail.bas`
- `vba/demo/modDemo_CommandCenter.bas`
- `vba/demo/modDemo_ReconciliationEngine.bas`
- `vba/demo/modDemo_VarianceNarrative.bas`
- `vba/demo/modDemo_ExecBriefPack.bas`
- `vba/demo/modDemo_WhatIfScenario.bas`

### Python (universal)
- `python/universal/profile_workbook.py`
- `python/universal/sanitize_dataset.py`
- `python/universal/compare_workbooks.py`
- `python/universal/build_exec_summary.py`

### Python (demo)
- `python/demo/pnl_data_extract.py`
- `python/demo/variance_classifier.py`
- `python/demo/scenario_runner.py`
- `python/demo/export_brief_package.py`

### SQL
- Universal: `sql/universal/template_gl_extract.sql`, `sql/universal/template_revenue_extract.sql`
- Demo: `sql/demo/demo_pnl_reconciliation_view.sql`, `sql/demo/demo_variance_fact.sql`

### Tooling / quality
- `scripts/run_stage_smoke.sh`
- `scripts/bootstrap_demo_workspace.py`
- `scripts/update_code_inventory.py`
- `tests/stage2_smoke_check.py`
- `tests/test_python_utilities.py`
- `.github/workflows/smoke-check.yml`
- `Makefile`

### Guides and training
- `guides/*.md` (architecture, walkthrough, troubleshooting, release checklist, push quickstart, etc.)
- `videos/*.md` (video scripts)

## 4) What to Compare Against Claude-built Version

Ask Claude to compare along these dimensions:

1. **Feature parity**
   - Universal tooling coverage.
   - Demo workflow completeness.

2. **Code quality**
   - Error handling, assumptions, deterministic behavior.
   - Function boundaries and testability.

3. **Validation posture**
   - Presence and rigor of smoke checks.
   - Unit coverage depth.

4. **Operational readiness**
   - Onboarding clarity for non-developers.
   - Branch/push/PR usability.
   - Repeatability in CI.

5. **Maintainability**
   - Docs completeness and drift controls.
   - Inventory/update discipline.

## 5) Current Quality Gates (must-pass)

Run locally:

```bash
python scripts/update_code_inventory.py
bash scripts/run_stage_smoke.sh
```

Expected:
- smoke checks pass,
- py_compile pass,
- unit tests pass,
- inventory is up-to-date.

## 6) Known Constraints / Risks

- Excel-host runtime behavior is only partially verifiable in CLI/CI environment.
- Sample `.xlsm` files are treated as read-only source assets.
- Some checks are marker/presence-based by design (broad guardrails, not full semantic runtime proofs).

## 7) Suggested Claude Review Deliverables

Ask Claude to produce:
1. Scorecard table (feature parity, quality, test depth, docs, ops readiness).
2. Critical gaps list (must-fix).
3. High-impact improvements (quick wins vs bigger refactors).
4. Migration plan if you want to merge best ideas from both codebases.
