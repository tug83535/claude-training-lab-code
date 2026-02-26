# Changelog

All notable changes to the KBT P&L Automation Toolkit are documented in this file.

Format follows [Semantic Versioning](https://semver.org/): MAJOR.MINOR.PATCH

- **MAJOR** — Breaking changes that require user action
- **MINOR** — New features, backward compatible
- **PATCH** — Bug fixes, backward compatible

---

## [2.1.0] — 2026-02-20

### ⚠️ Breaking Changes
- **modConfig:** Added 13 new public constants (SH_GL, SH_TECH_DOC, SH_CHANGE_LOG, SH_TEST_REPORT, SH_ALLOC_OUT, SH_SENSITIVITY, SH_VARIANCE, SH_DQ_REPORT, SH_SEARCH, SH_VAL_REPORT, APP_BUILD_DATE, and helpers SafeDeleteSheet, StyleHeader, AutoFitWithMax, WriteSummaryRow, CenterHeader). Modules that previously used `Private Const` for generated sheet names should now use the centralized constants.
- **modNavigation:** Keyboard shortcuts changed from Ctrl+Letter to Ctrl+Shift+Letter to avoid overriding Excel builtins (ISSUE-004). New bindings: Ctrl+Shift+H (Home), Ctrl+Shift+J (Jump), Ctrl+Shift+R (Checks), Ctrl+Shift+M (Command Center).
- **modMasterMenu:** Expanded from 36 to 50 action items. Previous action numbers 1–36 are preserved. New items 37–50 added.

### Added

**VBA Modules — Foundation Fixes (Phase 1):**
- modConfig_v2.1.bas — 8 missing items added (ISSUE-001), 13 new constants, 5 new helper subs
- modPerformance_v2.1.bas — Midnight timer rollover fix (ISSUE-005/BUG-004)
- modNavigation_v2.1.bas — `Application.OnKey` shortcuts replacing `MacroOptions` (ISSUE-004/BUG-011)
- modMonthlyTabGenerator_v2.1.bas — Safe `UpdateHeaderText` preventing "Mar"→"Aprigin" corruption (ISSUE-002/BUG-013)
- modDataQuality_v2.1.bas — Pre-flagged `FixTextNumbers` with `m_TextNumberCells` collection (ISSUE-003/BUG-018)
- modPDFExport_v2.1.bas — Dynamic month sheet names from modConfig constants (ISSUE-007)

**VBA Modules — Menu System (Phase 2):**
- modMasterMenu_v2.1.bas — 50-item dual-path architecture (UserForm primary, InputBox fallback)
- modFormBuilder_v2.1.bas — Programmatic UserForm builder (BuildCommandCenter), ExecuteAction router for all 50 items, GetFormInstallGuide for manual install instructions
- frmCommandCenter code-behind — Category filtering, search, double-click execution

**VBA Modules — Advanced Features (Phase 3):**
- modDashboard_v2.1.bas — Added CreateExecutiveDashboard, WaterfallChart, ProductComparison (ISSUE-009)
- modVarianceAnalysis_v2.1.bas — Added GenerateCommentary with auto-narrative generation (ISSUE-010), BuildNarrative helper, expanded cost-line detection (BUG-023), fixed highlighting (BUG-024)
- modReconciliation_v2.1.bas — Added ValidateCrossSheet with 4 computed validations (ISSUE-011), WriteValidationRow and FindFYCol helpers
- modMonthlyTabGenerator_v2.1.bas — Added GenerateNextMonthOnly (single-month generation, ISSUE-012 Part 1), tab color coding, data entry highlighting
- modSearch_v2.1.bas — MAX_RESULTS cap warning with total match count (ISSUE-012 Part 2), uses modConfig SH_SEARCH constant, modLogger integration on all Public Subs

**Python Scripts — Ecosystem Consolidation (Phase 4):**
- pnl_config.py — APP_VERSION updated to 2.1.0, clean UTF-8
- pnl_dashboard.py — Clean UTF-8, verified Streamlit imports
- pnl_month_end.py — Clean UTF-8, 6 check categories with CloseReport dataclass
- pnl_allocation_simulator.py — Clean UTF-8, Greek Delta headers, color-coded exports
- pnl_ap_matcher.py — Clean UTF-8, arrow/triangle symbols fixed
- pnl_cli.py — Clean UTF-8
- pnl_email_report.py — Clean UTF-8, triangle status indicators fixed
- pnl_forecast.py — Clean UTF-8, box drawing section dividers fixed
- pnl_snapshot.py — Clean UTF-8
- pnl_enhancements.sql — Renamed from .py (was incorrectly extended), clean UTF-8
- pnl_runner.py (NEW) — Unified CLI entry point dispatching to 9 commands
- pnl_tests.py — Expanded from ~30 to 116 test methods across 17 classes
- requirements.txt — Version pins verified, UTF-8 header cleaned

**Documentation (Phase 5):**
- QUICK_START.md — 10-minute onboarding path with ASCII Command Center diagram
- IMPLEMENTATION_GUIDE.md — Trust Center, module import, UserForm build (Mode A/B), Python setup, named ranges, FY rollover, architecture diagram
- USER_TRAINING_GUIDE.md — All 50 commands documented with number, description, usage context, and troubleshooting
- OPERATIONS_RUNBOOK.md — Monthly cadence (open/mid/close), step-by-step procedures, 8 failure scenarios with resolution
- SANITIZATION_PLAYBOOK.md — 6 masking procedures, verification checklist, reversal instructions
- CHANGELOG.md — This file

### Fixed
- **ISSUE-001** (CRITICAL) — modConfig missing 8 required items → Added all constants and helpers
- **ISSUE-002** (CRITICAL) — "Mar" text corruption in UpdateHeaderText → Safe pattern-based replacement
- **ISSUE-003** (CRITICAL) — FixTextNumbers data bomb → Pre-flagged cell collection
- **ISSUE-004** (CRITICAL) — Keyboard shortcuts overriding Excel builtins → Application.OnKey
- **ISSUE-005** (HIGH) — Timer midnight rollover → +86400 correction
- **ISSUE-006** (HIGH) — Menu 14 items behind (36→50) → Full 50-item routing
- **ISSUE-007** (HIGH) — PDF hardcoded month names → Dynamic from modConfig
- **ISSUE-008** (HIGH) — Python UTF-8 encoding artifacts → ~940 mojibake fixes across all files
- **ISSUE-009** (MEDIUM) — Dashboard missing 3 advanced charts → Implemented
- **ISSUE-010** (MEDIUM) — Missing GenerateCommentary → Implemented with BuildNarrative
- **ISSUE-011** (MEDIUM) — Missing ValidateCrossSheet → Implemented with 4 validations
- **ISSUE-012** (MEDIUM) — Search MAX_RESULTS silent cap → Total count + warning message
- **BUG-023** — Variance cost-line detection too narrow → Expanded keyword list
- **BUG-024** — Variance highlighting used alternating rows instead of flag-based → Fixed

### Statistics
- VBA: 14 modules delivered, 5,962 lines total
- Python: 12 files delivered, 4,807 lines total
- Documentation: 6 files
- Issues resolved: 12 of 15 (ISSUE-013 through ISSUE-015 pending Phase 6-7)
- Test methods: 116 across 17 test classes

---

## [2.0.0] — 2025-12-15

### ⚠️ Breaking Changes
- Complete architecture redesign from single-module to 29-module system
- New modConfig centralized constants replace hardcoded values
- modPerformance TurboOn/TurboOff wrappers required on all Public Subs
- modLogger audit trail integration required on all Public Subs

### Added
- **Foundation:** modConfig (constants, helpers), modPerformance (TurboOn/Off, timer), modLogger (audit trail), ThisWorkbook events
- **Core Features (15 modules):** modNavigation, modMonthlyTabGenerator, modReconciliation, modDataQuality, modVarianceAnalysis, modSensitivity, modDashboard (3 charts), modPDFExport, modAWSRecompute, modMasterMenu (36 items), modValidation, modSnapshot, modConditionalFormat, modEmailSummary, modSearch
- **Advanced (10 modules):** modAdmin, modAllocation, modForecast, modFormatting, modFormBuilder, modImport, modIntegrationTest, modRefresh, modScenario, modSetup
- **Python (10 scripts):** pnl_config, pnl_dashboard, pnl_month_end, pnl_allocation_simulator, pnl_ap_matcher, pnl_cli, pnl_email_report, pnl_enhancements, pnl_forecast, pnl_snapshot, pnl_tests
- **Documentation:** KBT_Training_Program_v2.docx, KBT_Implementation_Walkthrough_v2.docx, KBT_Customization_Guide_v2.docx, KBT_SQLPython_Audit_v1.docx, KBT_SQLPython_Training_Guide.docx

### Known Issues (at time of release)
- 40 bugs identified in pre-audit (see Part1A_Modular_Audit.md)
- 15 issues categorized for systematic resolution in v2.1
- Python files contain UTF-8 encoding artifacts from cross-platform development
- modMasterMenu limited to 36 items (50 planned)

---

## [1.0.0] — 2025-09-01

### Added
- Initial workbook structure with 13 core sheets
- Manual P&L reporting workflow
- Basic formulas for monthly trending
- Product Line Summary calculations
- Functional P&L breakdowns (Jan, Feb, Mar)
- AWS Allocation sheet with manual percentage entry
- Checks sheet with manual validation formulas
- No VBA automation (manual Excel-only workflow)

### Limitations
- All processes manual — no macros
- No data quality checking
- No automated reconciliation
- No version control or snapshots
- Limited to 3 months (Jan-Mar) of functional P&L tabs
- No dashboard or visualization capability
