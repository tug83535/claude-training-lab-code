# Changelog

All notable changes to the KBT P&L Automation Toolkit are documented in this file.

Format follows [Semantic Versioning](https://semver.org/): MAJOR.MINOR.PATCH

- **MAJOR** — Breaking changes that require user action
- **MINOR** — New features, backward compatible
- **PATCH** — Bug fixes, backward compatible

---

## [2.1.1] — 2026-03-11

### Added

**VBA Modules — New v2.1 Modules (2026-03-01, 7 modules):**
- modDemoTools_v2.1.bas — AddControlSheetButtons, SetParameterizedPrintArea, CreatePrintableExecSummary
- modDataGuards_v2.1.bas — ValidateAssumptionsPresence, CheckSumOfDrivers, FindNegativeAmounts, FindZeroAmounts, FindSuspiciousRoundNumbers
- modDrillDown_v2.1.bas — AddReconciliationDrillLinks, AutoPopulateReconciliationChecks, ApplyReconciliationHeatmap, RunGoldenFileCompare
- modAuditTools_v2.1.bas — AppendChangeLogEntry, FindExternalLinks, FixExternalLinks, AuditHiddenSheets, CreateMaskedCopy, ExportErrorSummaryClipboard
- modETLBridge_v2.1.bas — TriggerETLLocally, ImportETLOutput
- modTrendReports_v2.1.bas — CreateRolling12MonthView, CreateReconciliationTrendChart, ArchiveReconciliationResults
- modSheetIndex_v2.1.bas — CreateHomeSheet, ListAllSheetsWithLinks

**VBA Modules — Optional Add-Ins (2026-03-11, 5 modules):**
- modTimeSaved_v2.1.bas (305 lines) — Time Saved Calculator: manual vs automated time for all 62 actions
- modSplashScreen_v2.1.bas — Branded welcome screen with UserForm + MsgBox fallback
- modProgressBar_v2.1.bas (270 lines) — Animated progress bar with %, ETA, elapsed time
- modWhatIf_v2.1.bas (558 lines) — Live What-If scenarios: 7 presets + custom + baseline restore
- modExecBrief_v2.1.bas (447 lines) — Executive Brief auto-generator scanning 5 workbook sections

**VBA Modules — ProjectRefresh Enhancements (2026-03-05):**
- modDashboardAdvanced_v2.1.bas (NEW) — Split from modDashboard (was 1,398 lines)
- modDataQuality_v2.1.bas — Added Data Quality Letter Grade (A-F scoring)
- modVarianceAnalysis_v2.1.bas — Added YoY Variance Analysis
- modDataSanitizer_v2.1.bas — Numeric-only sanitizer with header keyword skip protection

**Universal Toolkit Modules (2026-03-04 through 2026-03-11, 14 modules total):**
- modUTL_Core.bas — Shared utilities (UTL_BackupSheet, UTL_TurboOn/Off)
- modUTL_Branding.bas — ApplyiPipelineBranding, SetiPipelineThemeColors
- modUTL_SheetTools.bas — ListAllSheetsWithLinks, TemplateCloner, GenerateUniqueCustomerIDs
- modUTL_DataSanitizer.bas — RunFullSanitize, PreviewSanitizeChanges, FixFloatingPointTails, ConvertTextStoredNumbers
- modUTL_ProgressBar.bas (NEW 2026-03-11) — Standalone status bar progress bar, ASCII visual
- modUTL_SplashScreen.bas (NEW 2026-03-11) — Standalone MsgBox splash screen
- modUTL_ExecBrief.bas (NEW 2026-03-11, 253 lines) — Universal workbook scanner, zero dependencies
- Plus 7 previously existing universal modules (modUTL_Audit, modUTL_DataQuality, modUTL_Search, etc.)

**Python Enhancements:**
- pnl_forecast.py — Added Forecast Accuracy Scoring (MAPE calculation)
- date_format_unifier.py — Fixed missing `day_first` parameter in `detect_date_columns()`

**Training Guides (6 drafts in FinalRoughGuides/):**
- Dynamic-Chart-Filter-Setup-Guide.md
- Plus 5 additional training guides for coworkers

**CoPilot Prompt Guide v2.0 (2026-03-07):**
- AP_Copilot_PromptGuideHelpV2.md — Fixed all broken quick reference links, added working anchors

**Video Package (2026-03-07):**
- COMPILED_VIDEO_PACKAGE.md — Demo build checklist, tool counts, file stats
- Sample_Quarterly_Report.xlsx — Built for Video 3 universal tools demo

### Fixed

**Bug Fixes — Testing Phase (2026-03-03 through 2026-03-04, 21 bugs):**
- BUG-T2.01: 9 missing sheet-name constants in modConfig (commit af44453)
- BUG-T2.03a/b: CLR_NAVY and CLR_ALT_ROW wrong hex-to-decimal conversion (commit 19320db)
- BUG-T2.04a/b: TestUpdateHeaderText wrapper + NumberFormat text fix (commits 6f40f91, ed3276f)
- BUG-T4.04a: Windows PermissionError on temp file cleanup (commit 3024c44)
- BUG-T5.01a/b: ExecDashboard row 1 vs HDR_ROW_REPORT + Error 5 crash (commits 6c17bd5, 847a982)
- BUG-T5.02: WaterfallChart hardcoded row label (commit 304743b)
- SR-01 through SR-12: 2 critical logic bugs + 9 LogAction signature bugs + 1 constant fix (commit 22ba831)

**Bug Fixes — Pre-Delivery Code Review (2026-03-07, 7 bugs):**
- CR-01: LogAction instance #13 in modReconciliation
- CR-02/03: HDR_ROW_REPORT in modReconciliation (ValidateCrossSheet + FindFYCol)
- CR-04: modPDFExport hardcoded 7 sheets → dynamic discovery
- CR-05/06: SpecialCells rng reset in modDataSanitizer + modAuditTools
- CR-07: xlSheetVeryHidden → xlSheetHidden in modDrillDown

**Bug Fixes — 5-Pass Review (2026-03-11, 3 bugs):**
- BR-01/02: Chr(9472) crashes VBA in modSplashScreen + modUTL_SplashScreen (3 locations)
- BR-03: Unused SPLASH_BG/SPLASH_ACCENT constants removed

**Bug Fixes — Code Review (2026-03-05, 4 bugs):**
- 3 more LogAction signature bugs (modDataQuality, modReconciliation, modPDFExport)
- Python `detect_date_columns()` missing `day_first` parameter

### Changed
- modFormBuilder_v2.1.bas — ExecuteAction router expanded from 50 to 62 actions
- modMasterMenu_v2.1.bas — Menu expanded from 50 to 62 items (4 pages)
- modDashboard_v2.1.bas — Split into base + advanced modules
- `_internal/` moved to `OldRoughVersions/_internal/` for repo cleanup (2026-03-11)
- `LastCallOptionalAddIns/` folder created for future add-in guides (2026-03-11)

### Statistics (Updated)
- Demo VBA modules: **39 total** (24 original + 8 new v2.1 + 2 enhancements + 5 optional add-ins)
- Universal Toolkit VBA modules: **14 total** (~100+ tools)
- Python scripts: **14 total**
- Total bugs found and fixed: **35** (9 testing + 12 self-review + 7 code review + 4 code review #2 + 3 five-pass review)
- Training guides: 6 drafts complete
- CoPilot Prompt Guide v2.0 complete
- Video package draft complete

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
- pnl_forecast.py — Clean UTF-8, box drawing section dividers fixed
- pnl_snapshot.py — Clean UTF-8
- pnl_enhancements.sql — Renamed from .py (was incorrectly extended), clean UTF-8
- pnl_runner.py (NEW) — Unified CLI entry point dispatching to 8 commands
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
- **Python (10 scripts):** pnl_config, pnl_dashboard, pnl_month_end, pnl_allocation_simulator, pnl_ap_matcher, pnl_cli, pnl_enhancements, pnl_forecast, pnl_snapshot, pnl_tests
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
