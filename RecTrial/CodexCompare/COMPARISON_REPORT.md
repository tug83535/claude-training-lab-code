# Project A vs Project B — Cherry-Pick Comparison Report

**Author:** Claude (Opus 4.7, 1M context)
**Date:** 2026-04-20
**Scope:** Static, read-only comparison. Project A stays as-is. Goal is to harvest good ideas from Project B (Codex's from-scratch build in `CodexCompare/`) and flag them for possible port into Project A's Universal Toolkit or docs.

**Reading order for the busy reader:** skim sections 1–4, then jump straight to **Section 8 (Cherry-Pick List)** for the actionable items.

---

## 0. Working definitions

- **Project A** = Connor's real iPipeline F&A demo, spread across:
  - `C:\Users\connor.atlee\RecTrial\` (active working folder — authoritative for VBA)
  - `C:\Users\connor.atlee\.claude\projects\claude-training-lab-code\` (git repo — authoritative for finalized Python/SQL/guides)
- **Project B** = Codex's from-scratch build in `C:\Users\connor.atlee\RecTrial\CodexCompare\`
- "Cherry-pick" = *idea worth porting into Project A*, nothing more. No migration plan, no restructuring, no "winner."

A quick note on Codex's handoff doc (`CodexCompare\guides\claude-handoff-deep-analysis.md`): it is an accurate *navigation* map but undersells one gap (no header comments on any VBA module) and oversells one claim ("comprehensive validation" — some of it is real, some is presence-only marker checks). Verified against the code; details below.

---

## 1. Inventory tables (side by side)

### 1.1 Folder structure overview

**Project A — RecTrial (working folder)**
```
RecTrial\
├── VBAToImport\               # newest modDirector + WhatIf
├── DemoVBA\                   # 38 v2.1 demo modules
├── DemoPython\                # 13 pnl_* scripts + sql\ (4 files)
├── UniversalToolkit\
│   ├── vba\                   # 23 modules + NewTools\ (4 modules)
│   └── python\                # 18 utilities + NewTools\ (4 utilities)
├── DemoFile\ExcelDemoFile_adv.xlsm        # demo workbook (authoritative sample)
├── SampleFile\SampleFileV2\               # universal toolkit test workbook
├── Guide\                      # active recording guides (Video 3/4)
├── VideoScripts\               # finalized video scripts
├── AudioClips\                 # MP3 narration (for modDirector playback)
├── Guides\                     # 13 user-facing PDFs
└── VBABackup_*\                # point-in-time backups (ignored in this report)
```

**Project A — `claude-training-lab-code` (repo)**
```
claude-training-lab-code\
├── CLAUDE.md                   # 44 KB governance doc
├── README.md
├── Archive\
│   ├── tasks\lessons.md        # bug patterns / anti-patterns
│   ├── tasks\todo.md           # release phases
│   ├── docs\ipipeline-brand-styling.md
│   └── qa\ TEST_PLAN, BUG_LOG, VALIDATION_REPORT, …
├── FinalExport\
│   ├── DemoPython\             # finalized PnL scripts + sql\
│   ├── DemoVBA\                # finalized demo modules
│   ├── UniversalToolkit\       # finalized universal VBA+Python
│   └── VideoRecording\         # finalized video scripts
├── ExeTest\xlam_kit\           # xlam add-in prep
└── SourceCode\                 # legacy working set
```

**Project B — CodexCompare**
```
CodexCompare\
├── README.md  PLAN.md  CONTEXT.md  CONSTRAINTS.md  BRAND.md
│   CHANGELOG.md  CODE_INVENTORY.md  CONTRIBUTING.md
│   STARTER_PROMPT.md  START_HERE_PROMPT.md  PROJECT_TODO.md
│   Makefile
├── .github\workflows\smoke-check.yml
├── vba\
│   ├── universal\              # 6 modUTL_ modules
│   └── demo\                   # 7 modDemo_ modules
├── python\
│   ├── universal\              # 4 stdlib-only scripts
│   └── demo\                   # 4 stdlib-only scripts
├── sql\
│   ├── universal\              # 2 templates (T-SQL-ish)
│   └── demo\                   # 2 views (CTE + window fns)
├── tests\                      # stage2_smoke_check + test_python_utilities
├── scripts\                    # run_stage_smoke.sh + bootstrap + inventory
├── guides\                     # 11 markdown guides
└── videos\                     # 5 short video scripts
```

### 1.2 File counts by type

| Artifact | Project A (RecTrial + repo, deduped by folder) | Project B (CodexCompare) |
|---|---:|---:|
| `.bas` VBA modules (active, not backups) | **67** (1 Director + 38 DemoVBA + 28 UTL) | **13** (6 UTL + 7 Demo) |
| VBA total LOC (approx) | **~32,500** | **~1,893** |
| `.py` active scripts | **~37** (14 DemoPython + 22 UTL Python + 1 build script) | **12** (4 universal + 4 demo + 3 scripts + 1 bootstrap) |
| `.sql` scripts | **4** (SQLite dialect) | **4** (T-SQL dialect) |
| Markdown guides (non-video) | ~13 (plus 13 PDFs in `Guides\`) | **11** |
| Video scripts | **~9** (5 primary + 4 support) | **5** |
| Tests | `pnl_tests.py` only (863 LOC) | `stage2_smoke_check.py` (456) + `test_python_utilities.py` (167) |
| CI / Makefile | **none** | GitHub Actions + Makefile |
| Auto-generated code inventory | **none** | `scripts/update_code_inventory.py` → `CODE_INVENTORY.md` |
| Top-level governance docs | CLAUDE.md, README.md, lessons.md | PLAN.md, CONTEXT.md, CONSTRAINTS.md, BRAND.md, CHANGELOG.md, CODE_INVENTORY.md, CONTRIBUTING.md |

**One-line takeaway:** Project A is ~17× larger in VBA and ~3× larger in Python. Project B is a *tight, disciplined, much smaller repo* with infrastructure Project A lacks (CI, Makefile, auto-inventory, stdlib-only Python, top-level governance files).

### 1.3 Names of every VBA module, grouped by category

| Category | Project A (RecTrial authoritative) | Project B (CodexCompare) |
|---|---|---|
| **Video/demo orchestration** | [modDirector.bas](VBAToImport/modDirector.bas) (2,891 LOC MCI-driven puppeteer) | — *(no equivalent)* |
| **Splash / onboarding** | modSplashScreen_v2.1, modUTL_SplashScreen | — |
| **Command center / menu** | modUTL_CommandCenter (1,155 LOC, auto-discovery), modMasterMenu_v2.1 (62-item InputBox), modFormBuilder_v2.1 | modUTL_CommandCenter, modDemo_CommandCenter |
| **Core/shared utilities** | modUTL_Core, modUtilities_v2.1, modPerformance_v2.1, modProgressBar_v2.1, modUTL_ProgressBar | modUTL_Core |
| **Data sanitizer / quality / cleaning** | modDataSanitizer_v2.1, modUTL_DataSanitizer (3-phase), modUTL_DataCleaning, modUTL_DataCleaningPlus, modDataQuality_v2.1, modDataGuards_v2.1, modUTL_NumberFormat | modUTL_DataSanitizer |
| **Compare / diff** | modUTL_Compare (643 LOC, highlighting + report) | merged into modUTL_CompareConsolidate |
| **Consolidate** | modUTL_Consolidate, modConsolidation_v2.1 | merged into modUTL_CompareConsolidate |
| **Finance toolkit** | modUTL_Finance (1,033 LOC, **14 tools**: AP dupe, GL validator, trial balance, AR/AP aging, flux, ratios, FX, journal mapper…) | — *(not a discrete module)* |
| **Audit / compliance** | modUTL_Audit (8 tools), modUTL_AuditPlus (4 tools: boundary, header validator, formula errors, consistency), modAuditTools_v2.1 | — |
| **Duplicate detection** | modUTL_DuplicateDetection | — |
| **Reconciliation** | modReconciliation_v2.1 (PASS/FAIL + cross-sheet) | modDemo_ReconciliationEngine |
| **Variance / narrative** | modVarianceAnalysis_v2.1 (657 LOC, MoM + YoY + auto-commentary) | modDemo_VarianceNarrative, modUTL_Intelligence |
| **Executive brief / one-pager** | modExecBrief_v2.1 (465 LOC), modUTL_ExecBrief (universal; scans any workbook) | modUTL_OutputPack, modDemo_ExecBriefPack |
| **What-if / scenarios** | modWhatIf_v2.1 (9-preset menu), modUTL_WhatIf, modScenario_v2.1 (save/load/compare/delete), modSensitivity_v2.1 (1-way+2-way) | modDemo_WhatIfScenario |
| **Forecasting** | modForecast_v2.1 | — |
| **Dashboards / charts** | modDashboard_v2.1, modDashboardAdvanced_v2.1 (1,030 LOC; waterfall, small multiples), modTrendReports_v2.1 | — |
| **Drill-down / navigation** | modDrillDown_v2.1, modNavigation_v2.1, modSearch_v2.1, modSheetIndex_v2.1 | — |
| **Monthly tab generator** | modMonthlyTabGenerator_v2.1 | — |
| **Allocations** | modAllocation_v2.1 | — |
| **Pivot / formatting / validation builder** | modUTL_PivotTools, modUTL_Formatting, modUTL_Highlights, modUTL_ValidationBuilder, modUTL_LookupBuilder, modUTL_Branding, modUTL_ColumnOps, modUTL_Comments, modUTL_SheetTools, modUTL_TabOrganizer, modUTL_WorkbookMgmt | — |
| **Import / ETL** | modImport_v2.1, modETLBridge_v2.1 | — |
| **Testing / health** | modIntegrationTest_v2.1 (30-test suite) | — |
| **PDF / reporting output** | modPDFExport_v2.1 | (PDF via ExportAsFixedFormat inside modUTL_OutputPack & modDemo_ExecBriefPack) |
| **Admin / logging / version** | modAdmin_v2.1, modLogger_v2.1, modVersionControl_v2.1, modAWSRecompute_v2.1, modTimeSaved_v2.1, modDemoTools_v2.1 | modDemo_AuditTrail, modDemo_Config |
| **Intelligence (materiality, narrative, data quality score)** | *(spread across modVarianceAnalysis + modDataQuality + modExecBrief)* | **modUTL_Intelligence** (single module, very clean) |

### 1.4 Names of every Python script

| Category | Project A | Project B |
|---|---|---|
| **Workbook profile / metadata** | — | `profile_workbook.py` (stdlib, zipfile+XML) |
| **Data cleaner / sanitizer** | `UniversalToolkit/python/clean_data.py`, `NewTools/date_format_unifier.py` | `sanitize_dataset.py` (stdlib) |
| **Compare / diff** | `compare_files.py`, `NewTools/two_file_reconciler.py`, `NewTools/multi_file_consolidator.py` | `compare_workbooks.py` (stdlib) |
| **Exec summary builder** | *(Streamlit dashboard + `word_report.py`)* | `build_exec_summary.py` (CSV → markdown) |
| **P&L CLI / orchestrator** | `DemoPython/pnl_cli.py`, `pnl_runner.py`, `pnl_config.py` | — |
| **Forecasting** | `pnl_forecast.py` (518 LOC, 4 methods + MAPE), `forecast_rollforward.py` | — |
| **Monte Carlo / scenario** | `pnl_monte_carlo.py` (1,110 LOC, Dirichlet+shocks) | `scenario_runner.py` (47 LOC, delta %) |
| **Allocation** | `pnl_allocation_simulator.py` | — |
| **AP match / fuzzy** | `pnl_ap_matcher.py`, `fuzzy_lookup.py` | — |
| **Snapshot / month-end** | `pnl_snapshot.py`, `pnl_month_end.py` | — |
| **Dashboard** | `pnl_dashboard.py` (Streamlit + plotly) | — |
| **Chart builder** | `build_charts.py` (matplotlib) | — |
| **Model redesign** | `redesign_pl_model.py` (openpyxl) | — |
| **Aging / reconciliation** | `aging_report.py`, `bank_reconciler.py`, `gl_reconciliation.py`, `reconciliation_exceptions.py` | — |
| **Variance** | `variance_analysis.py`, `variance_decomposition.py` | `variance_classifier.py` (stdlib) |
| **Unpivot / regex / PDF extract** | `unpivot_data.py`, `regex_extractor.py`, `pdf_extractor.py`, `master_data_mapper.py`, `consolidate_budget.py`, `consolidate_files.py`, `batch_process.py`, `word_report.py`, `NewTools/sql_query_tool.py` | — |
| **Sample extract to CSV** | *(no equivalent)* | `pnl_data_extract.py` (stdlib sheet→CSV) |
| **Brief package bundler** | *(implicit via Streamlit or Word report)* | `export_brief_package.py` (markdown assembler) |
| **Workspace bootstrap** | — | `scripts/bootstrap_demo_workspace.py` |
| **Code-inventory generator** | — | `scripts/update_code_inventory.py` |
| **Tests** | `pnl_tests.py` (pytest; 99 pass / 15 skip) | `tests/test_python_utilities.py` (unittest); `tests/stage2_smoke_check.py` |

### 1.5 Names of every SQL script

| Project A (`DemoPython/sql/` — SQLite) | Project B (`sql/` — T-SQL-ish) |
|---|---|
| `staging.sql` — dims + raw GL + fact_gl + dedup (221 LOC) | `sql/universal/template_gl_extract.sql` — parameterized GL extract (23) |
| `transformations.sql` — allocation pivots, dept_product, variance views (276 LOC) | `sql/universal/template_revenue_extract.sql` — parameterized revenue (22) |
| `validations.sql` — NULL, dup, outlier, allocation, FK checks (362 LOC) | `sql/demo/demo_pnl_reconciliation_view.sql` — `CREATE OR ALTER VIEW` with null flags (17) |
| `pnl_enhancements.sql` — audit log + trigger + retained earnings + rolling (408 LOC) | `sql/demo/demo_variance_fact.sql` — CTE + `ROW_NUMBER()/COUNT() OVER` period-over-period (29) |

**Observation:** Project A's SQL is an implementation — 1,267 LOC, SQLite-specific, does real work. Project B's SQL is a *template set* — 91 LOC, ANSI/T-SQL-leaning, meant to be copied and adapted. These are different products, not competing versions.

### 1.6 Video scripts / training guides

| Project A (authoritative: RecTrial `VideoScripts\`, `Guide\`, repo `FinalExport\VideoRecording\`) | Project B (`videos/`, `guides/`) |
|---|---|
| `Video_1_Script_Whats_Possible.md` (~18–22 min narrative) | `video-1-executive-hook.md` (~7 min) |
| `Video_2_Script_Full_Demo_Walkthrough.md` | `video-2-demo-workbook-deep-dive.md` (~9 min) |
| `Video_3_Script_Universal_Tools.md` + `VIDEO_3_STEP_BY_STEP.md`, `VIDEO_3_INTERACTIVE_GUIDE.md`, `VIDEO_3_CLIP_TRACKER.md`, `VIDEO_3_GEMINI_REVIEW.md` | `video-3-universal-toolkit-in-action.md` (~8 min) |
| `VIDEO_4_NARRATION_SCRIPT.md`, `VIDEO_4_RECORDING_GUIDE.md`, `VIDEO_4_INTERACTIVE_GUIDE.md`, `Video4_Script_Feedback_for_ClaudeCode.md` | `video-4-python-sql-integration.md` (~8 min) |
| `Video_Demo_Master_Plan.md`, `MASTER_RECORDING_GUIDE.md`, `DIRECTOR_MACRO_SETUP_GUIDE.md`, `RECORDING_INSTRUCTIONS.md`, `COMPILED_VIDEO_PACKAGE.md` | `video-5-copilot-adaptation-lab.md` (~6 min) |
| 13 user-facing PDFs in `Guides\` (Start Here, Command Center, First-Time Setup, Leadership Overview, Quick Ref, Training, Universal Toolkit, Runbook, What-If Guide, CC Guide, VBA Module Reference, AP Copilot Prompt Guide, BrandStyling-CopilotPrompt, Dynamic-Chart-Filter) | `architecture-overview.md`, `universal-tool-catalog.md`, `universal-toolkit-user-guide.md`, `demo-walkthrough-guide.md`, `troubleshooting-reference.md`, `release-readiness-checklist.md`, `git-branch-push-quickstart.md`, `brand-styling-reference.md`, `copilot-prompt-guide.md`, `claude-review-prompt.md`, `claude-handoff-deep-analysis.md` |

**Observation:** Project A is video-production-heavy (multi-pass recording support for a specific 3-video series) with 13 polished PDFs for end users. Project B is light on video production (shorter scripts, no recording guide, no director macro) but richer on *operational governance docs* (release checklist, git quickstart, troubleshooting, architecture overview, CONTRIBUTING, CHANGELOG).

---

## 2. Feature parity matrix

Legend: ✅ present · ⚠️ partial · ❌ not present · N/A not applicable

### 2.1 Universal VBA tooling

| Feature | Project A | Project B | Notes |
|---|:-:|:-:|---|
| Core helpers (TurboOn/Off, SafeDelete, Last row/col) | ✅ [modUTL_Core.bas](../UniversalToolkit/vba/modUTL_Core.bas) | ✅ modUTL_Core | Both implement; A has richer helper set (`SafeNum`, `SafeStr`, `BackupSheet`) |
| Header-row auto-detection | ✅ (ad-hoc per module) | ✅ `UTL_DetectHeaderRow` scoring 25-row scan | B's is centralized and reused cleanly — small advantage |
| Data sanitizer (text→number, float tails) | ✅ 3-phase (modUTL_DataSanitizer) + Preview | ✅ 1-phase (RunFullSanitize) + Preview | A's is deeper |
| Cell-by-cell compare | ✅ modUTL_Compare (highlighting) | ✅ `CompareActiveSheetToSheet` | Roughly equivalent |
| Multi-sheet consolidate | ✅ modUTL_Consolidate + modConsolidation_v2.1 | ✅ `ConsolidateVisibleSheetsByHeader` | B adds "SourceSheet" tag column automatically; A can do the same via parameter |
| Finance toolkit (14 tools: AP dupe, GL validator, ageing, flux, ratios, FX, trial balance) | ✅ modUTL_Finance 1,033 LOC | ❌ | Project A unique |
| Audit toolkit (external links, circ refs, error scan, formula consistency) | ✅ modUTL_Audit + modUTL_AuditPlus | ❌ | Project A unique |
| Duplicate detector | ✅ modUTL_DuplicateDetection | ❌ | Project A unique |
| Materiality classifier (any sheet) | ⚠️ *(embedded in modVarianceAnalysis; not a standalone tool)* | ✅ `MaterialityClassifierActiveSheet` | B has a cleaner universal tool |
| Data-quality scorecard (numeric 0–100 score) | ⚠️ modDataQuality_v2.1 (flags issues; no single score) | ✅ `DataQualityScorecardActiveSheet` (100 − blanks·60% − errors·40%) | B's score metric is a nice summarization for exec one-pager |
| Exception narrative generator (any sheet) | ⚠️ modVarianceAnalysis has MoM/YoY narratives | ✅ `GenerateExceptionNarrativesActiveSheet` | B's is a generic, universal tool |
| Executive one-pager (any workbook) | ✅ modUTL_ExecBrief | ✅ `BuildExecutiveOnePagerFromActiveSheet` | B additionally produces a "Run Receipt" sheet (useful audit artifact) |
| PDF export | ✅ modPDFExport_v2.1 | ✅ `ExportExecutivePackPDF` | Both use `ExportAsFixedFormat` |
| Run receipt sheet (per-run audit artifact) | ❌ | ✅ `CreateRunReceiptSheet` | **Nice cherry-pick** |
| Dedicated UTL run-log sheet | ⚠️ modLogger_v2.1 writes to a log sheet | ✅ `UTL_EnsureRunLogSheet` (Timestamp, User, Module, Procedure, Status, Message, Sheets, Cells Changed) | B's schema is tighter; A logs but the format is more ad-hoc |
| Command center auto-discovery of UTL_* modules | ✅ modUTL_CommandCenter with registry + search + custom tools | ⚠️ modUTL_CommandCenter (static button grid, 7 actions) | Project A unique |
| Splash screen | ✅ modSplashScreen (UserForm + MsgBox fallback) | ❌ | Project A unique |
| Progress bar | ✅ modUTL_ProgressBar | ❌ (uses `StatusBar` only) | Project A unique |
| Branding (iPipeline colors, auto-detect headers/totals) | ✅ modUTL_Branding (7-color palette, alternating rows) | ⚠️ hard-coded RGB in each header builder | Project A is more reusable |
| Validation builder | ✅ modUTL_ValidationBuilder | ❌ | Project A unique |
| Lookup builder | ✅ modUTL_LookupBuilder | ❌ | Project A unique |
| Pivot tools | ✅ modUTL_PivotTools | ❌ | Project A unique |
| Column ops / tab organizer / workbook mgmt | ✅ modUTL_ColumnOps, modUTL_TabOrganizer, modUTL_WorkbookMgmt | ❌ | Project A unique |
| Comments tool | ✅ modUTL_Comments | ❌ | Project A unique |
| Highlights / conditional formatting | ✅ modUTL_Highlights | ❌ | Project A unique |
| Whitespace / non-printable / case cleaner | ✅ modUTL_DataCleaningPlus | ❌ | Project A unique |

### 2.2 Demo-specific VBA

| Feature | Project A | Project B |
|---|:-:|:-:|
| Workbook-state validator (required sheets/columns present) | ⚠️ done per module | ✅ modDemo_Config `DemoValidateWorkbookOrStop` |
| Dual audit trail (local `VBA_AuditLog` sheet + call to universal logger) | ⚠️ single logger | ✅ modDemo_AuditTrail |
| Demo command center (4-button UI) | ✅ modMasterMenu_v2.1 (62 items!) + modFormBuilder_v2.1 | ✅ modDemo_CommandCenter |
| Reconciliation engine | ✅ modReconciliation_v2.1 (526 LOC, PASS/FAIL + cross-sheet) | ✅ modDemo_ReconciliationEngine (141) |
| Variance narrative | ✅ modVarianceAnalysis_v2.1 (657 LOC) | ✅ modDemo_VarianceNarrative (130) |
| Exec brief pack (KPI + checks + trend + PDF) | ✅ modExecBrief_v2.1 (465 LOC) | ✅ modDemo_ExecBriefPack (177) |
| What-if scenarios (named presets) | ✅ modWhatIf_v2.1 (9 presets) + modScenario_v2.1 (save/load/compare) + modSensitivity_v2.1 (1-way/2-way) | ✅ modDemo_WhatIfScenario (4 fixed scenarios) |
| Forecasting (trend extrapolation) | ✅ modForecast_v2.1 | ❌ |
| Dashboards (line, bar, waterfall, small multiples) | ✅ modDashboard_v2.1 + modDashboardAdvanced_v2.1 | ❌ |
| Drill-down / navigation / search / sheet index | ✅ modDrillDown / modNavigation / modSearch / modSheetIndex | ❌ |
| Monthly tab generator | ✅ modMonthlyTabGenerator_v2.1 | ❌ |
| Allocations (product/department) | ✅ modAllocation_v2.1 | ❌ |
| Integration test suite (30 tests + health check) | ✅ modIntegrationTest_v2.1 | ❌ |
| Time-saved tracking / FTE hours | ✅ modTimeSaved_v2.1 | ❌ |
| Version control / change tracking | ✅ modVersionControl_v2.1 | ❌ |
| Admin / permissions / audit tools | ✅ modAdmin_v2.1 + modAuditTools_v2.1 | ❌ |
| Video demo puppeteer (modDirector with MCI audio) | ✅ 2,891 LOC | ❌ |

### 2.3 Python utilities

| Capability | Project A | Project B |
|---|:-:|:-:|
| Workbook profile (sheets, named ranges, VBA presence) | ⚠️ inside `pnl_month_end.py`/`redesign_pl_model.py` | ✅ `profile_workbook.py` (stdlib only, zipfile+XML) |
| CSV sanitizer (text/number/date normalize) | ✅ `clean_data.py` (pandas) | ✅ `sanitize_dataset.py` (stdlib only) |
| Workbook cell diff → CSV | ✅ `compare_files.py`, `two_file_reconciler.py` (pandas) | ✅ `compare_workbooks.py` (stdlib only) |
| Markdown exec summary from CSV | ⚠️ via `word_report.py` | ✅ `build_exec_summary.py` (stdlib only) |
| P&L CLI orchestrator | ✅ `pnl_cli.py` | ❌ |
| P&L month-end close | ✅ `pnl_month_end.py` (531 LOC) | ❌ |
| Allocation simulator (product/dept) | ✅ `pnl_allocation_simulator.py` | ❌ |
| Monte Carlo (Dirichlet + shocks, 10k iterations) | ✅ `pnl_monte_carlo.py` (1,110 LOC) | ⚠️ `scenario_runner.py` (deterministic delta %, 97 LOC) |
| Rolling forecast (SMA/ETS/Trend/Scenario + MAPE) | ✅ `pnl_forecast.py` | ❌ |
| AP fuzzy matcher | ✅ `pnl_ap_matcher.py`, `fuzzy_lookup.py` (thefuzz) | ❌ |
| Bank reconciliation | ✅ `bank_reconciler.py` | ❌ |
| Aging report (AR/AP 30/60/90+) | ✅ `aging_report.py` | ❌ |
| GL reconciliation | ✅ `gl_reconciliation.py` | ❌ |
| Unpivot wide→long | ✅ `unpivot_data.py` | ❌ |
| Regex / PDF / multi-file / SQL-on-CSV | ✅ `regex_extractor.py`, `pdf_extractor.py`, `multi_file_consolidator.py`, `sql_query_tool.py`, `date_format_unifier.py` | ❌ |
| Word report generator | ✅ `word_report.py` (python-docx) | ❌ |
| Streamlit dashboard | ✅ `pnl_dashboard.py` (plotly) | ❌ |
| Variance classifier (material/favorable) | ⚠️ inside `variance_analysis.py` | ✅ `variance_classifier.py` (stdlib, standalone) |
| Extract named sheets → CSV | ⚠️ inside P&L scripts | ✅ `pnl_data_extract.py` (stdlib, standalone) |
| Brief package assembler (markdown + links) | ⚠️ Word report is the analogue | ✅ `export_brief_package.py` |

### 2.4 SQL templates

| Capability | Project A | Project B |
|---|:-:|:-:|
| SQL dialect | SQLite (used in-process) | T-SQL-leaning, ANSI-portable |
| Parameterized extract templates | ❌ | ✅ GL + revenue templates with `DECLARE @params` |
| CTE + window function variance view | ❌ | ✅ `demo_variance_fact.sql` |
| `CREATE OR ALTER VIEW` for reconciliation with null flags | ❌ | ✅ `demo_pnl_reconciliation_view.sql` |
| Audit log + trigger-based change capture | ✅ `pnl_enhancements.sql` | ❌ |
| FK / referential integrity checks | ✅ `validations.sql` | ❌ |
| Dedup + staging normalization | ✅ `staging.sql` | ❌ |
| Allocation-share table + variance views | ✅ `transformations.sql` | ❌ |

### 2.5 Validation / testing / tooling

| Feature | Project A | Project B |
|---|:-:|:-:|
| Python unit tests | ⚠️ `pnl_tests.py` covers DemoPython only (99 pass) | ✅ `tests/test_python_utilities.py` (167 LOC unittest) |
| Repo-wide smoke check (markers, presence, structure) | ❌ | ✅ `tests/stage2_smoke_check.py` (456 LOC) |
| Makefile (`make check`, `make smoke`, `make unit`) | ❌ | ✅ 4 targets |
| GitHub Actions CI | ❌ | ✅ `smoke-check.yml` |
| Auto-regenerated code inventory | ⚠️ manual `VBA-Module-Reference-List.md` (goes stale) | ✅ `scripts/update_code_inventory.py` → `CODE_INVENTORY.md` |
| Bootstrap demo workspace (timestamped copy of samples) | ❌ (manual copy) | ✅ `scripts/bootstrap_demo_workspace.py` |
| VBA integration-test suite (30 in-Excel tests) | ✅ modIntegrationTest_v2.1 | ❌ |

### 2.6 Video scripts

| Feature | Project A | Project B |
|---|:-:|:-:|
| Video count | 4 recorded/planned (1–4 + supporting) | 5 planned |
| Avg length | 18–22 min | ~7–8 min |
| Tone | "Coworker-plain-English" + enthusiastic | Executive-hook + business-benefit-framed |
| Director macro orchestration (frame-perfect audio + action sync) | ✅ modDirector | ❌ (scripts only) |
| Clip tracker / step-by-step / Gemini review | ✅ `VIDEO_3_*` guides | ❌ |
| "On-screen action" callouts with specific sheet/column refs | ⚠️ present but scattered | ✅ each video has explicit "On-Screen Action Callouts" section |
| Business-impact CTA at end of each video | ⚠️ implicit | ✅ explicit per-video CTA ("reduce repetitive month-end effort", etc.) |
| 5th video: CoPilot adaptation lab | ❌ (but has CoPilot prompt guide PDF) | ✅ video-5 |

### 2.7 Training / user guides

| Feature | Project A | Project B |
|---|:-:|:-:|
| Start-here welcome | ✅ PDF | ✅ (README + STARTER_PROMPT + START_HERE_PROMPT) |
| Quick reference card | ✅ PDF | ⚠️ *(partial — architecture-overview)* |
| Universal Toolkit user guide | ✅ PDF | ✅ `universal-toolkit-user-guide.md` |
| VBA module reference (all modules, grouped) | ✅ PDF (38 DemoVBA) | ✅ `universal-tool-catalog.md` (160-item aspirational) |
| CoPilot prompt guide (how to safely adapt code) | ✅ `AP_Copilot_PromptGuideHelpV2.md` | ✅ `copilot-prompt-guide.md` (274 LOC, workbook-mapping template + 6 worked examples) |
| Brand styling reference | ✅ `Company-BrandStyling-CopilotPrompt.pdf` + `ipipeline-brand-styling.md` | ✅ `BRAND.md` (top-level) + `brand-styling-reference.md` (ops card) |
| Operations runbook | ✅ PDF | ⚠️ in `troubleshooting-reference.md` |
| Troubleshooting reference (symptom → fix, user-facing) | ⚠️ lessons.md is internal | ✅ `troubleshooting-reference.md` |
| Release-readiness checklist | ⚠️ `todo.md` phases | ✅ `release-readiness-checklist.md` (7-section pre-demo) |
| Git branch / push quickstart (non-dev) | ❌ | ✅ `git-branch-push-quickstart.md` |
| Architecture overview | ❌ | ✅ `architecture-overview.md` |
| CONTRIBUTING.md | ❌ | ✅ |
| CHANGELOG.md top-level | ⚠️ `Archive/qa/CHANGELOG.md` | ✅ |

---

## 3. Same-intent, different-execution

Where both projects tackled the same problem but took different paths. Cherry-pick flags (🍒) are candidates for section 8.

### 3.1 Universal command center
- **Project A** [`UniversalToolkit\vba\modUTL_CommandCenter.bas`](../UniversalToolkit/vba/modUTL_CommandCenter.bas) is 1,155 LOC: auto-discovers any `modUTL_*` module (via Trust Access), maintains a static registry fallback, supports user-registered custom tools, search, categories, up to 200 tools.
- **Project B** [`CodexCompare\vba\universal\modUTL_CommandCenter.bas`](vba/universal/modUTL_CommandCenter.bas) is 180 LOC: static 7-button grid, applies iPipeline header, delegates to sibling UTL modules by name. No discovery, no search, no custom tools.
- **Verdict:** Project A wins decisively on capability. Project B's module reads cleanly if you're teaching a beginner how the architecture works, but it's not a replacement.
- **Cherry-pick:** None from this one; Project A's is richer.

### 3.2 Exec one-pager
- **Project A** [`modUTL_ExecBrief.bas`](../UniversalToolkit/vba/modUTL_ExecBrief.bas) (252 LOC) scans *any* workbook: sheet counts, file size, formula/value ratio, error cells, blank rows/cols, chart/pivot inventory, plain-English output.
- **Project B** [`modUTL_OutputPack.bas`](vba/universal/modUTL_OutputPack.bas) does three things: `BuildExecutiveOnePagerFromActiveSheet` (simpler KPIs from active sheet) + `ExportExecutivePackPDF` + `CreateRunReceiptSheet` (per-run audit artifact with timestamp, user, workbook, feature name).
- **Verdict:** Project A's ExecBrief is more ambitious; Project B's `CreateRunReceiptSheet` is a small, focused gem.
- 🍒 **Cherry-pick:** the run-receipt pattern. After any meaningful run, drop a 6-row sheet with feature, user, timestamp, workbook path, cells changed, sheets touched. Great for compliance and recorded demos.

### 3.3 Compare / consolidate (two modules vs one)
- **Project A** splits into [`modUTL_Compare.bas`](../UniversalToolkit/vba/modUTL_Compare.bas) (643 LOC, highlighting + report) and [`modUTL_Consolidate.bas`](../UniversalToolkit/vba/modUTL_Consolidate.bas) (547 LOC, pattern-based picker).
- **Project B** fuses both in `modUTL_CompareConsolidate.bas` (165 LOC) using `Scripting.Dictionary` for pipe-delimited row signatures and automatically adding a `SourceSheet` tag column.
- **Verdict:** Project A's split is better (separation of concerns, more features). Project B's signature-hash approach is elegant and compact.
- 🍒 **Cherry-pick (minor):** signature-hash dedup idea via pipe-delimited concat as `Scripting.Dictionary` key — a quick way to detect "identical-row-across-sheets" without a full cell-by-cell compare. Could become a small helper in `modUTL_Compare`.

### 3.4 Variance narrative
- **Project A** [`modVarianceAnalysis_v2.1.bas`](../DemoVBA/modVarianceAnalysis_v2.1.bas) (657 LOC): MoM, YoY, FY vs Budget, auto-narrative ("Revenue increased 8.2%, favorable"), color-coded variance columns, alternating rows on report.
- **Project B** [`modDemo_VarianceNarrative.bas`](vba/demo/modDemo_VarianceNarrative.bas) (130 LOC) compares first vs last month, labels via `Select Case` (Material increase/decrease, Watch, Normal), writes to `Exec_Variance_Narrative` sheet.
- **Verdict:** Project A's is richer; Project B's is a clean minimum-viable implementation of the same idea. Project B also has a separate generic version `GenerateExceptionNarrativesActiveSheet` in `modUTL_Intelligence` that works on *any* sheet.
- 🍒 **Cherry-pick:** Project B's **`MaterialityClassifierActiveSheet`** + **`GenerateExceptionNarrativesActiveSheet`** inside `modUTL_Intelligence` are genuinely universal — they don't depend on the demo workbook shape, they take `materiality_abs` and `materiality_pct` as optional parameters, and they write their output to a dedicated sheet. Project A's variance narrative is stuck inside the demo modules. Port the generic-for-any-sheet versions into `UniversalToolkit/vba/` as a new `modUTL_Intelligence.bas` (or fold into `modUTL_Finance`).

### 3.5 What-if / scenarios
- **Project A** [`modWhatIf_v2.1.bas`](../DemoVBA/modWhatIf_v2.1.bas) (645 LOC, 9 presets incl. combo Best/Worst Case) + [`modScenario_v2.1.bas`](../DemoVBA/modScenario_v2.1.bas) (save/load/compare/delete named scenarios) + `modSensitivity_v2.1.bas` (1-way + 2-way tables) + `modUTL_WhatIf.bas` (universal +/- 5/10/25% on selected cells with baseline save/restore).
- **Project B** [`modDemo_WhatIfScenario.bas`](vba/demo/modDemo_WhatIfScenario.bas) (154 LOC, 4 fixed scenarios: Base, Growth Push, Margin Protection, Stress).
- **Verdict:** Project A dominates decisively. Project B's 4-scenario labeled narrative ("aggressive" at ≥60% margin, "monitor" at ≥50%, "escalate" at <50%) is tidy but only a sliver of Project A's capability.
- 🍒 **Cherry-pick (minor):** the "margin threshold narrative labels" pattern — if margin ≥ X% label aggressive; else monitor; else escalate. Add as one more preset/scenario option to `modWhatIf_v2.1.bas`.

### 3.6 CSV sanitize / clean data
- **Project A** `UniversalToolkit/python/clean_data.py` (153 LOC) uses **pandas** + **openpyxl** — needs pip install.
- **Project B** `CodexCompare/python/universal/sanitize_dataset.py` (100 LOC) uses **stdlib only** (`csv`, `datetime`, `re`) — runs on any Python without `pip`.
- **Verdict:** Different target users. Project A's covers more formats; Project B's can be emailed to a coworker with no setup friction.
- 🍒 **Cherry-pick:** drop Project B's `sanitize_dataset.py` into Project A's `UniversalToolkit/python/NewTools/` as a **"zero-install" companion** — not a replacement. Same for `profile_workbook.py`, `compare_workbooks.py`, `build_exec_summary.py`. Coworkers with corporate laptops that block `pip` can still use these four.

### 3.7 Exec summary / brief package
- **Project A** `word_report.py` (python-docx) writes a polished Word doc; `pnl_dashboard.py` (Streamlit) renders interactive dashboard.
- **Project B** `build_exec_summary.py` + `export_brief_package.py` produce a **markdown** brief (CSV → markdown with totals, top groups, suggested talking points; then wraps with reviewer notes + optional artifact links).
- **Verdict:** Markdown briefs are viewable everywhere (GitHub, Teams, email). The pattern of "one-liner talking points" is genuinely useful.
- 🍒 **Cherry-pick:** the **"suggested talking points"** output section concept in `build_exec_summary.py`. Add to Project A's `word_report.py` (or a new sibling) — auto-generate 3–5 one-sentence talking points for the CFO from the numbers. Pairs well with Project A's existing `modVarianceAnalysis.GenerateCommentary`.

### 3.8 Python tests
- **Project A** `pnl_tests.py` (863 LOC) — pytest, rich mocking, covers DemoPython deeply. No coverage for `UniversalToolkit/python/` 22 scripts.
- **Project B** `tests/test_python_utilities.py` (167 LOC) — unittest, tempfile isolation, assertions on outputs. Covers all 8 Python scripts.
- **Verdict:** A has depth, B has breadth. Project A's UTL python has ZERO tests today.
- 🍒 **Cherry-pick:** add a **small unittest suite** (`UniversalToolkit/python/tests/test_utl_python.py`) modelled on Project B's pattern — for each utility, one or two tests with tempfile-based input/output, asserting something real (row count, field names, non-empty output). Doesn't need to be deep; having *any* tests prevents silent breakage between recording sessions.

### 3.9 Structure / governance docs
- **Project A** puts everything in CLAUDE.md (44 KB) + Archive/tasks/lessons.md + scattered PDFs.
- **Project B** splits into PLAN.md (strategy), CONTEXT.md (framing), CONSTRAINTS.md (must/must-not), BRAND.md (style), CHANGELOG.md (history), CONTRIBUTING.md (workflow), PROJECT_TODO.md (current), CODE_INVENTORY.md (auto).
- **Verdict:** Project A's monolith is fine for one human but punishing for AI tools to navigate. Project B's split makes it easier to find anything.
- 🍒 **Cherry-pick (optional):** no need to break CLAUDE.md apart. But **add two small top-level files**: a `CONSTRAINTS.md` (one screen of must/must-not) and an auto-generated `CODE_INVENTORY.md`. Both reduce drift and help future AI sessions orient fast.

### 3.10 Video scripts
- **Project A** Video 1–4 are 18–22 min, with matching MASTER_RECORDING_GUIDE, DIRECTOR_MACRO_SETUP_GUIDE, RECORDING_INSTRUCTIONS, step-by-step clip trackers, Gemini review files.
- **Project B** 5 scripts of 6–9 min each, each with **timestamped outline**, **on-screen action callouts** (naming specific sheets/columns), **business impact** section, and a concrete next-action CTA.
- **Verdict:** Project A is *far* deeper on production infrastructure — don't change the plan. But Project B's per-script structure template is worth a look.
- 🍒 **Cherry-pick:** **Not for the main 4 videos**; keep those as planned. But for any *future* 5-to-10-minute explainer (e.g., Video 5 CoPilot lab if you ever do one, or shorter "chapter" clips), adopt Project B's four-part template: (1) hook, (2) timestamped outline, (3) on-screen action callouts, (4) business-impact CTA.

---

## 4. Unique to Project B (Codex) — the main cherry-pick hunting ground

These are things Project B has that Project A doesn't.

### 4.1 VBA (universal)
1. **`modUTL_Intelligence.bas` — three self-contained universal tools:**
   - `MaterialityClassifierActiveSheet(materiality_abs, materiality_pct)` — labels rows Material / Watch / Normal on *any* sheet
   - `GenerateExceptionNarrativesActiveSheet` — prose narrative per row
   - `DataQualityScorecardActiveSheet` — numeric 0–100 score (100 − blanks·60% − errors·40%), written to `UTL_QualityScorecard`
   Project A has these *concepts* but spread across demo modules. Consolidating and universalizing is the value.
2. **`CreateRunReceiptSheet`** inside `modUTL_OutputPack.bas` — each run produces a timestamped receipt sheet with user, workbook, feature name. Lightweight compliance artifact.
3. **`UTL_EnsureRunLogSheet` schema** inside `modUTL_Core.bas` — 8 columns (Timestamp, User, Module, Procedure, Status, Message, Sheets, Cells Changed). Tighter than Project A's logger.
4. **`modDemo_AuditTrail.DemoLog` dual-logging pattern** — writes to local `VBA_AuditLog` + also calls the universal run logger. Simple idea, useful discipline.

### 4.2 Python (stdlib-only utilities)
5. **`profile_workbook.py`** — stdlib-only xlsx/xlsm profiler (sheet list, named ranges, VBA presence via zipfile XML). Zero pip install.
6. **`sanitize_dataset.py`** — stdlib-only CSV normalizer (text, number, date).
7. **`compare_workbooks.py`** — stdlib-only workbook diff → CSV.
8. **`build_exec_summary.py`** — CSV → markdown exec summary with auto "talking points."
9. **`pnl_data_extract.py`** — extract named sheets from xlsx to CSV files (stdlib only).
10. **`variance_classifier.py`** — classify a variance CSV (Actual vs Baseline) into Direction + Materiality. Stdlib.
11. **`scenario_runner.py`** — apply percentage shocks to a metric column, summarize. Stdlib.
12. **`export_brief_package.py`** — assemble a markdown brief from a summary + optional artifact links + reviewer notes.

Why all these are interesting: **zero dependencies**. Project A's Python requires pandas/openpyxl/thefuzz/pdfplumber/streamlit/plotly. Many iPipeline corporate laptops won't let employees `pip install` anything. These eight scripts run on any Python 3.8+ out of the box.

### 4.3 SQL templates
13. **`template_gl_extract.sql`** and **`template_revenue_extract.sql`** — parameterized, vendor-neutral starter templates. Useful teaching artifacts for video 4 / copilot guide.
14. **`demo_variance_fact.sql`** — CTE + `ROW_NUMBER() OVER / COUNT() OVER` for period-over-period variance. Compact (29 LOC) and pedagogically clean.
15. **`demo_pnl_reconciliation_view.sql`** — `CREATE OR ALTER VIEW` with inline null-flag `CASE` expressions. Minimal pattern for a reconciliation data mart.

### 4.4 Tooling / validation / governance
16. **`tests/stage2_smoke_check.py`** (456 LOC) — repo-wide structural validator: VBA markers, SQL keyword presence, Python `--help`, sample workbook integrity, guide sections, video script structure, changelog entries, inventory freshness. *Caveat:* the bulk of this script is presence/marker checks ("does this file exist and contain this keyword") rather than behavioral tests that actually run the code. Real behavioral coverage lives in `test_python_utilities.py` (167 LOC). Useful, but don't over-trust the "comprehensive validation" framing from Codex's handoff doc.
17. **`tests/test_python_utilities.py`** (167 LOC) — genuine unittest coverage across all Python utilities with tempfile isolation.
18. **`scripts/bootstrap_demo_workspace.py`** — creates a timestamped copy of the two sample workbooks, refuses to overwrite existing workspace (safe). Perfect for demo re-runs.
19. **`scripts/update_code_inventory.py`** — auto-generates `CODE_INVENTORY.md` from every source file, with LOC + first-comment summary. Kills doc drift.
20. **`Makefile`** — 4 targets (`smoke`, `unit`, `py-compile`, `check`). Local parity with CI.
21. **`.github/workflows/smoke-check.yml`** — GitHub Actions CI (22 LOC). Runs the smoke script on every push/PR.

### 4.5 Docs / guides / governance
22. **`guides/release-readiness-checklist.md`** — 7-section pre-demo checklist (workbook, features, Python/SQL, branding, risk controls, sign-off).
23. **`guides/architecture-overview.md`** — 54-line two-prong system map with runtime flow and validation strategy. Compact enough to paste into a stakeholder email.
24. **`guides/troubleshooting-reference.md`** — 6 symptom-based error recovery workflows + escalation template. User-facing (Project A's `lessons.md` is internal).
25. **`guides/git-branch-push-quickstart.md`** — step-by-step for non-dev coworkers (Section 5 anticipates auth/permission friction).
26. **`guides/brand-styling-reference.md`** — one-screen operational card condensing BRAND.md rules.
27. **`guides/universal-tool-catalog.md`** — aspirational 160-item tool catalog grouped by workflow. Useful stakeholder buy-in artifact (not an implementation doc).
28. **`guides/copilot-prompt-guide.md`** (274 lines) — includes a **workbook-mapping template** + **6 per-feature worked examples** + validation checklist + troubleshooting loop. Project A's `AP_Copilot_PromptGuideHelpV2.md` is more Q&A flavored; the mapping-template pattern is a nice addition.
29. **`STARTER_PROMPT.md` / `START_HERE_PROMPT.md`** — short onboarding prompts to paste into a new AI session. Project A's CLAUDE.md is authoritative but too long for cold-starting an AI.
30. **Top-level `BRAND.md`, `CONSTRAINTS.md`, `CHANGELOG.md`, `CONTRIBUTING.md`, `PROJECT_TODO.md`** — standard repo hygiene that Project A mostly lacks at top level.
31. **`video-5-copilot-adaptation-lab.md`** — unique video angle (teaching safe code adaptation via CoPilot prompts). Project A has the *written* CoPilot prompt guide but no video equivalent.

### 4.6 Design patterns worth copying
32. **`On Error GoTo <NamedFail>` with named failure labels** (e.g., `CompareFail`, `ReconFail`) that log + MsgBox. Project A uses `On Error GoTo ErrHandler`; named labels read a bit better but this is stylistic.
33. **`Scripting.Dictionary` for pipe-delimited row signature hashing** in `BuildRowHashMap` (`modUTL_CompareConsolidate`). Compact duplicate-detection helper.
34. **Application-state batching** in `RunFullSanitize` (`modUTL_DataSanitizer`): disable ScreenUpdating + Calculation + Events in one block, restore together. Project A does this via `TurboOn/TurboOff` (better — fewer lines per call site).

---

## 5. Unique to Project A (Codex did NOT attempt)

Short by design — these just document the moat Project A already has. Lower priority for this report.

- **modDirector video puppeteer** — MCI-driven MP3 playback, runtime duration detection, 39 clips across 3 videos, preflight checks, abort-key detection. No Codex analogue.
- **Branded splash screen with auto-dismiss** (modSplashScreen + modUTL_SplashScreen).
- **Command-center auto-discovery + custom tool registry + search** (modUTL_CommandCenter).
- **Finance toolkit (14 tools)**: duplicate invoice detector, auto-balancing GL validator, trial balance, AR/AP aging, flux, corkscrew builder, financial period rollforward, multi-currency consolidation, ratio dashboard, GL-journal mapper (modUTL_Finance).
- **Audit toolkit**: external link finder, circular reference detector, error scanner, data quality scorecard, named range auditor, data validation checker, inconsistent formula auditor, external link severance (modUTL_Audit + modUTL_AuditPlus).
- **Validation builder, lookup builder, pivot tools, column ops, tab organizer, workbook mgmt, comments tool, highlights tool, whitespace/non-printable/case cleaner, branding auto-styler**.
- **Dashboards (incl. advanced: waterfall, small multiples), drill-down, navigation, search, sheet index, monthly tab generator, forecasting, sensitivity (1-way + 2-way), allocations, consolidation, AWS recompute, time-saved tracking, version control, form builder, admin tools**.
- **Integration test suite (30 tests + health check)** in VBA.
- **Python P&L suite** (13 scripts): CLI orchestrator, Monte Carlo (1,110 LOC, Dirichlet), rolling forecast (4 methods + MAPE), AP matcher, allocation simulator, snapshot, month-end close, Streamlit dashboard, model redesign, chart builder.
- **Universal Python toolkit** (22 scripts): aging, bank recon, clean data, compare, consolidate, forecast rollforward, fuzzy, GL recon, master data mapper, PDF extractor, regex, recon exceptions, unpivot, variance analysis, variance decomposition, word report, batch process + NewTools (date unifier, multi-consolidator, SQL-on-CSV, two-file reconciler).
- **SQL DDL: dims + fact + allocation shares + variance views + validations + audit log trigger + retained earnings** (1,267 LOC across 4 SQL files).
- **13 polished user-facing PDFs** (Start Here, Command Center, First-Time Setup, Leadership Overview, Quick Ref, Training, Universal Toolkit, Runbook, What-If Guide, CC Guide, VBA Module Reference, CoPilot Prompt Guide v2, Dynamic Chart Filter, BrandStyling-CopilotPrompt).
- **Path A video automation pattern** (Director* silent wrappers instead of SendKeys) — documented in lessons.md as a hard-won lesson.
- **Monte Carlo with Dirichlet allocation** — statistically appropriate, parameterized.
- **Pytest suite with 99 passes / 15 skips / 0 failures** over the P&L suite.

---

## 6. Code-quality observations

### 6.1 VBA

| Dimension | Project A | Project B |
|---|---|---|
| `Option Explicit` everywhere | ✅ 100% (67/67) | ✅ (verified on read modules) |
| Header comment blocks on every module | ✅ 100% — standard format (PURPOSE / PUBLIC SUBS / DEPENDENCIES / VERSION / CHANGES / AUTHOR) | ❌ **none** on any of 13 modules |
| Consistent error-handling pattern | ✅ `On Error GoTo ErrHandler` + state cleanup (TurboOff / DisplayAlerts) | ✅ `On Error GoTo NamedFail` + log + MsgBox. Named labels read nicely. |
| Naming conventions | ✅ PascalCase public, camelCase private, UPPER_SNAKE constants | ✅ Similar — PascalCase public, `DemoLog`-style prefixes for demo |
| Brand colour use | ✅ centralised in `modUTL_Branding.bas` (7-colour palette) + reused via constants | ⚠️ RGB values inlined at each call site (`RGB(11,71,121)`, `RGB(17,46,81)`) — same hex but duplicated across files |
| Application-state batching (ScreenUpdating off etc.) | ✅ `TurboOn/TurboOff` helpers | ✅ inline, only in DataSanitizer |
| Use of Scripting.Dictionary | ✅ heavy | ⚠️ one place (late binding, `CreateObject`) |
| Module-count per concern | 67 modules — some overlap (e.g., DemoVBA has its own ProgressBar AND UTL has one) | 13 modules — leaner, one module per concern, but blends Compare + Consolidate and merges Intelligence into one file |

Concrete citations:
- Project A header example: `UniversalToolkit/vba/modUTL_Core.bas:1–17` (PURPOSE / PUBLIC SUBS block).
- Project B missing header: `CodexCompare/vba/universal/modUTL_Core.bas:1` jumps straight to `Option Explicit` with no preceding comment. Same on all 13 modules.
- Project A `TurboOn/TurboOff` reuse: `modUTL_Core.bas:60–84`.
- Project B inline state batching: `modUTL_DataSanitizer.bas:23–38`.

### 6.2 Python

| Dimension | Project A | Project B |
|---|---|---|
| Type hints (`from __future__ import annotations`, PEP 604 unions) | ⚠️ inconsistent across 37 scripts | ✅ consistent across all 8 scripts |
| Docstrings | ✅ module-level everywhere; function-level spotty | ✅ module-level everywhere; function-level spotty |
| Dependencies | pandas, numpy, openpyxl, matplotlib, streamlit, plotly, pdfplumber, statsmodels, thefuzz, pytest | **stdlib only** (`csv`, `zipfile`, `xml.etree`, `argparse`, `datetime`, `re`, `pathlib`, `statistics`) |
| CLI (argparse) | ✅ most scripts | ✅ every script |
| Deterministic I/O paths | ✅ (configurable via args) | ✅ (configurable via args) |
| Test coverage | ⚠️ deep on DemoPython, none on UniversalToolkit python | ✅ broad but shallow |
| Subprocess / help-gate tests | ❌ | ✅ `--help` probes in smoke check |

### 6.3 SQL

- Project A SQL is *production* code (DDL, triggers, views) for a real SQLite model. Well-commented with PURPOSE blocks and inline rationale. Heavier and more specific.
- Project B SQL is *template* code, small, ANSI-portable, with CTE + window function patterns that teach cleanly. Heavily-commented for their size.
- Neither has SQL unit tests beyond presence checks.

### 6.4 Overall

- **Project A** reads like a mature tool that's been iterated many times: 100% header coverage, working version-control discipline (`VBABackup_*` snapshots), recorded lessons, and a battle-tested video pipeline. The trade-off is some duplication (DemoVBA vs UTL versions of ProgressBar, SplashScreen, DataSanitizer) and docs that have gone stale in places (VBA-Module-Reference-List.md).
- **Project B** reads like a freshly-scaffolded repo that prioritized *repo hygiene over feature breadth*: clean governance files, deterministic CI, auto-inventory, stdlib-only Python, but *no header comments on VBA*, thin feature set, no dashboards, no forecasting, no Monte Carlo, no splash screen, no command-center discovery.

### 6.5 What NOT to copy from Project B

Two habits to avoid when porting any Codex code or pattern into Project A — because on both axes Project A is already better and regression would hurt.

- **No header comment blocks on any VBA module.** Project B's 13 `.bas` files jump from `Attribute VB_Name` / `Option Explicit` straight into code. No `PURPOSE` / `PUBLIC SUBS` / `DEPENDENCIES` / `VERSION` block. Project A has 100% header coverage (67/67 modules) in a consistent format — see [`modUTL_Core.bas:1–17`](../UniversalToolkit/vba/modUTL_Core.bas) as the template. Any ported helper must be wrapped in a header block before it's merged in.
- **Inlined RGB values at every call site.** Project B hard-codes `RGB(11, 71, 121)` (iPipeline Blue) and `RGB(17, 46, 81)` (Navy) inside every header-building routine (e.g., [`modUTL_CommandCenter.bas`](vba/universal/modUTL_CommandCenter.bas) line ~123, [`modDemo_CommandCenter.bas`](vba/demo/modDemo_CommandCenter.bas) line ~56). Project A centralizes these in [`modUTL_Branding.bas`](../UniversalToolkit/vba/modUTL_Branding.bas) as a named 7-colour palette applied consistently across 20+ modules. Keep the abstraction — when porting any Codex routine, replace inline `RGB(...)` calls with Project A's named constants.

---

## 7. Brand / tone adherence

**Project A rules (authoritative):** iPipeline Blue `#0B4779` primary; `#112E51` navy; `#4B9BCB` innovation blue; Arial family only; plain English; no emoji in official output; no hype language; professional, CFO-ready.

### 7.1 Project B VBA
- ✅ Uses iPipeline Blue `RGB(11, 71, 121) = #0B4779` in every header builder (e.g., `modUTL_CommandCenter.bas:123`; `modDemo_CommandCenter.bas:56`; `modDemo_ExecBriefPack.bas:67`).
- ✅ Navy `RGB(17, 46, 81) = #112E51` used for accent (`modUTL_CommandCenter.bas:133`).
- ⚠️ Innovation Blue is not named — Codex inlined `RGB(75, 155, 203) = #4B9BCB` directly. Same colour, but not abstracted; a coworker tweaking colours would have to grep.
- ✅ Arial everywhere. No Comic Sans / Calibri sighted.
- ✅ No emoji.
- ✅ Tone in MsgBox and narrative is plain English ("increased materially and requires owner confirmation" — `modUTL_Intelligence.bas:~212`).

### 7.2 Project B Python
- ✅ No emoji, no hype words in variable names or output ("Executive Summary", "Top contributing groups", "Suggested talking points").
- ✅ `sanitize_dataset.py` / `profile_workbook.py` etc. all produce CFO-appropriate terminology.
- ✅ Error messages are factual ("Marker '{marker}' not found in {rel}").

### 7.3 Project B docs
- ✅ `BRAND.md` states the rules cleanly; `brand-styling-reference.md` is the operational card.
- ⚠️ One tone nit — `video-1-executive-hook.md:~20`: "The challenge is **not just** calculations." "Just" is flagged as condescending when describing steps in your BRAND rules. In context here it's rhetorical contrast, not step-description, so arguably acceptable, but strict compliance would say: "The challenge is **not only** calculations." Flag, not a blocker.
- ✅ No emoji across any guide / video.
- ✅ Active voice, short sentences.

### 7.4 Would it look professional to a CFO?
**Yes — with one caveat.** The docs and outputs are CFO-ready. The only thing that would embarrass is the `universal-tool-catalog.md` listing 160 tools when maybe a fifth are actually implemented — don't use that doc in an exec deck as-is. The VBA modules, Python scripts, and video-5 script are all polished enough for a finance audience.

Project B matches Project A on brand tone in output. Where Project A has an edge: `modUTL_Branding.bas` abstracts the palette into named constants and applies them consistently across 20+ modules. Project B inlines the RGB values at each call site.

---

## 8. CHERRY-PICK LIST (main deliverable)

Sorted by bang-for-buck. Effort: **S** = under ~2 hours, **M** = ~half-day, **L** = >1 day. Each item has: what / where in A / why / effort / risk.

### Tier 1 — Strongly recommend (low effort, durable value)

1. **Auto-generated `CODE_INVENTORY.md`** — adapt [`CodexCompare/scripts/update_code_inventory.py`](scripts/update_code_inventory.py) to walk Project A's folders and regenerate a single inventory table (module, LOC, first-comment summary).
   - **Where in A:** add `scripts/update_code_inventory.py` at repo root `claude-training-lab-code/`. Output `CODE_INVENTORY.md` at same level, replacing/augmenting the stale `VBA-Module-Reference-List.md`.
   - **Why:** your `VBA-Module-Reference-List.pdf` is hand-maintained and drifts. Auto-regen keeps it honest forever.
   - **Effort:** S (Codex's script is 145 LOC, mostly reusable; adjust folder globs for A's 67 modules + 37 Python files + 4 SQL).
   - **Risk:** none — it's a read-only reporter.

2. **`CreateRunReceiptSheet` pattern** — port `CodexCompare/vba/universal/modUTL_OutputPack.bas:~140–190` (`CreateRunReceiptSheet`) as a helper in [`UniversalToolkit/vba/modUTL_Audit.bas`](../UniversalToolkit/vba/modUTL_Audit.bas) or `modUTL_ExecBrief.bas`.
   - **Where in A:** `RecTrial\UniversalToolkit\vba\modUTL_Audit.bas` (and later also in repo `FinalExport\UniversalToolkit\vba\`).
   - **Why:** after any major run (sanitizer, consolidate, exec brief), drop a 6-row receipt (timestamp, user, workbook path, feature name, cells changed, sheets touched). CFOs love audit trails; it also makes video recordings prove "yes, the macro actually ran."
   - **Effort:** S (~30 lines of VBA).
   - **Risk:** minor — ensure the receipt sheet is not picked up by demo comparisons; name it `UTL_RunReceipt` and prefix-exclude.

3. **`modUTL_Intelligence.bas`-style universal narratives** — port three procedures from [`CodexCompare/vba/universal/modUTL_Intelligence.bas`](vba/universal/modUTL_Intelligence.bas):
   - `MaterialityClassifierActiveSheet(materiality_abs, materiality_pct)`
   - `GenerateExceptionNarrativesActiveSheet`
   - `DataQualityScorecardActiveSheet`
   - **Where in A:** new file `RecTrial\UniversalToolkit\vba\modUTL_Intelligence.bas`, or fold into existing `modUTL_Finance.bas` / `modUTL_Audit.bas`. Follow Project A header-block convention (Codex's lack headers — add them on port).
   - **Why:** Project A has *flavours* of these inside demo modules. A universal version works on any workbook and broadens the command center's Finance/Audit tab.
   - **Effort:** M (~half-day — port 236 LOC, add headers, test on demo + sample workbook).
   - **Risk:** overlap with existing `modDataQuality_v2.1` — make sure you're offering a *generic* tool, not re-implementing demo-specific logic.

4. **Release-readiness checklist** — copy [`CodexCompare/guides/release-readiness-checklist.md`](guides/release-readiness-checklist.md) structure and adapt to Project A's 4-video release.
   - **Where in A:** `claude-training-lab-code\FinalExport\Guides\RELEASE_READINESS_CHECKLIST.md` (new), and reference it from `Archive/tasks/todo.md`.
   - **Why:** your `todo.md` tracks phases but doesn't have the 7-section pre-demo/pre-share checklist format (workbook integrity, features tested, Python/SQL green, branding spot-check, risk controls, sign-off). Useful before each video recording and before final SharePoint upload.
   - **Effort:** S (one file, ~50 lines).
   - **Risk:** none.

5. **User-facing `TROUBLESHOOTING.md`** — model on [`CodexCompare/guides/troubleshooting-reference.md`](guides/troubleshooting-reference.md): symptom → likely cause → fix → escalation.
   - **Where in A:** `claude-training-lab-code\FinalExport\Guides\TROUBLESHOOTING.md` and consider rendering to PDF alongside existing 13 user PDFs.
   - **Why:** `lessons.md` is internal / AI-facing. Coworkers who hit an error during the demo need a short, symptom-indexed doc that doesn't require them to know what LogAction is.
   - **Effort:** S–M (~1–2 hours; mine `lessons.md` for the user-visible symptoms — Excel macro errors, missing sheet, pivot table error, VBA import failure).
   - **Risk:** none.

6. **`bootstrap_demo_workspace.py`** — port [`CodexCompare/scripts/bootstrap_demo_workspace.py`](scripts/bootstrap_demo_workspace.py).
   - **Where in A:** `RecTrial\scripts\bootstrap_demo_workspace.py` (or in repo). Point it at your two samples: `RecTrial\DemoFile\ExcelDemoFile_adv.xlsm` and `RecTrial\SampleFile\SampleFileV2\Sample_Quarterly_ReportV2.xlsm`.
   - **Why:** protects your sample files — each recording attempt gets a fresh timestamped working copy; the originals never get accidentally saved-over.
   - **Effort:** S (~50 lines; adjust two paths).
   - **Risk:** none — refuses to overwrite existing workspace by design.

7. **Stdlib-only "zero-install" companion tools** — drop four Codex Python scripts into Project A as a "zero-install" sub-folder:
   - [`profile_workbook.py`](python/universal/profile_workbook.py)
   - [`sanitize_dataset.py`](python/universal/sanitize_dataset.py)
   - [`compare_workbooks.py`](python/universal/compare_workbooks.py)
   - [`build_exec_summary.py`](python/universal/build_exec_summary.py)
   - **Where in A:** `RecTrial\UniversalToolkit\python\ZeroInstall\` (new subfolder) + repo `FinalExport\UniversalToolkit\python\ZeroInstall\`. Add a `README.md` stating "these need only Python 3.8+, no pip install."
   - **Why:** coworkers with locked-down corporate laptops can't install pandas. These four give them basic profile / sanitize / diff / summary with zero setup.
   - **Effort:** S (copy files, write a 1-page README).
   - **Risk:** none; they're small and complement (don't replace) Project A's richer tools.

8. **Unit tests for `UniversalToolkit/python/`** — model on [`CodexCompare/tests/test_python_utilities.py`](tests/test_python_utilities.py). Add 1–2 unittest cases per utility (at minimum: "given a sample input CSV / xlsx, the script runs and produces expected output with right columns").
   - **Where in A:** new `RecTrial\UniversalToolkit\python\tests\test_utl_python.py` + `RecTrial\UniversalToolkit\python\tests\fixtures\small_sample.csv`.
   - **Why:** 22 scripts with **zero** tests today. One small test per script prevents silent breakage when you edit during recording.
   - **Effort:** M (~3–4 hours for shallow coverage of all 22 scripts; leaves deeper coverage as later work).
   - **Risk:** low — scripts are stateless CLI; just run them against fixtures.

9. **Top-level `CONSTRAINTS.md`** — a one-page must / must-not / must-never file modelled on [`CodexCompare/CONSTRAINTS.md`](CONSTRAINTS.md).
   - **Where in A:** repo root `claude-training-lab-code\CONSTRAINTS.md`. Anchor it from CLAUDE.md.
   - **Why:** CLAUDE.md (44 KB) is thorough but AI sessions routinely miss subtle rules (LogAction signature, Path A pattern, brand colour hex, no openpyxl pivot tables). A 1-screen CONSTRAINTS.md is what a new Claude / Codex / Gemini session reads first. Prevents the 13-time LogAction bug from becoming the 14th.
   - **Effort:** S (~30 lines, assembled from lessons.md's top anti-patterns).
   - **Risk:** none.

10. **Top-level `BRAND.md` (short, operational)** — move the core rules from `Archive/docs/ipipeline-brand-styling.md` to a top-level 1-page `BRAND.md`. Model on [`CodexCompare/BRAND.md`](BRAND.md) + [`CodexCompare/guides/brand-styling-reference.md`](guides/brand-styling-reference.md).
    - **Where in A:** repo root `claude-training-lab-code\BRAND.md`. Keep the full styling doc in Archive for reference.
    - **Why:** currently the brand doc is buried. Top-level, it's findable by AI and humans alike.
    - **Effort:** S.
    - **Risk:** none.

### Tier 1.5 — Small, focused adds (spotted on second pass)

Four narrow items that are cheap to port and don't fit neatly anywhere else. Lettered rather than numbered to keep Tier 2/3 numbering stable.

**A. `AI_REVIEW_PROMPT.md`** — adopt the meta-doc pattern from [`CodexCompare/guides/claude-review-prompt.md`](guides/claude-review-prompt.md) and [`claude-handoff-deep-analysis.md`](guides/claude-handoff-deep-analysis.md). One short doc that tells a future AI agent exactly how to review the project.
  - **Where in A:** repo root `claude-training-lab-code\AI_REVIEW_PROMPT.md` (~30 lines).
  - **Why:** ensures Gemini / Claude / Codex second-opinion reviews come back consistent in shape. No more re-explaining scope each time you want a fresh review.
  - **Effort:** S.
  - **Risk:** none.

**B. Centralized `UTL_DetectHeaderRow` helper** — port the scoring heuristic from [`CodexCompare/vba/universal/modUTL_Core.bas`](vba/universal/modUTL_Core.bas) (scans first 25 rows, returns the row with the highest non-empty-cell density) into Project A's [`modUTL_Core.bas`](../UniversalToolkit/vba/modUTL_Core.bas).
  - **Where in A:** `RecTrial\UniversalToolkit\vba\modUTL_Core.bas` (add the header block to match Project A's convention — Codex's has none).
  - **Why:** Project A does header detection ad-hoc across several modules (modVarianceAnalysis uses row 4, others assume row 1). One shared helper removes inconsistency and lets future UTL modules work on any-shape workbooks.
  - **Effort:** S (~20 lines + header comment).
  - **Risk:** low — additive; existing callers keep their hard-coded row numbers until deliberately refactored.

**C. Margin-threshold narrative preset for What-If** — steal the labelling logic from [`CodexCompare/vba/demo/modDemo_WhatIfScenario.bas`](vba/demo/modDemo_WhatIfScenario.bas) (≥60% = aggressive · ≥50% = monitor · <50% = escalate). Add as a new preset inside [`modWhatIf_v2.1.bas`](../DemoVBA/modWhatIf_v2.1.bas).
  - **Where in A:** `RecTrial\DemoVBA\modWhatIf_v2.1.bas` (and keep a matching `DirectorWhatIfPreset` silent wrapper per the Path A pattern).
  - **Why:** your 9 presets currently output numbers. Adding an automatic verdict in plain English ("aggressive — revisit owner confirmation") gives the CFO a built-in answer to "so is that OK or not?"
  - **Effort:** S.
  - **Risk:** low — purely additive preset.

**D. Row-signature dedup helper (`Scripting.Dictionary` fast-mode)** — port the pipe-delimited row-key pattern from `BuildRowHashMap` inside [`CodexCompare/vba/universal/modUTL_CompareConsolidate.bas`](vba/universal/modUTL_CompareConsolidate.bas) into [`modUTL_Compare.bas`](../UniversalToolkit/vba/modUTL_Compare.bas) as an optional fast entry point.
  - **Where in A:** `RecTrial\UniversalToolkit\vba\modUTL_Compare.bas`.
  - **Why:** quick "are these two sheets identical at the row level?" check without the full cell-by-cell pass + highlighting. Useful as a pre-check before a full compare on large sheets.
  - **Effort:** S.
  - **Risk:** low — additional entry point, not a replacement.

### Tier 2 — Recommend if time permits

11. **Workbook-mapping-template addition to CoPilot prompt guide** — take the template pattern from [`CodexCompare/guides/copilot-prompt-guide.md`](guides/copilot-prompt-guide.md) (Section 3 "workbook-mapping template" + 6 worked examples) and add it to Project A's `CoPilot-Quick-Start-Card.md` and/or `AP_Copilot_PromptGuideHelpV2.md`.
    - **Where in A:** `claude-training-lab-code\FinalExport\VideoRecording\Guides_v2\` area, or repo `FinalExport/Guides/`.
    - **Why:** forces the coworker to document (header row, amount column name, date column name, etc.) before prompting CoPilot. Fewer "CoPilot did the wrong thing" loops.
    - **Effort:** M.
    - **Risk:** none; purely additive.

12. **Git branch + push quickstart for non-devs** — port [`CodexCompare/guides/git-branch-push-quickstart.md`](guides/git-branch-push-quickstart.md) if you ever plan to let coworkers fork Project A and share back.
    - **Where in A:** `claude-training-lab-code\FinalExport\Guides\GIT_QUICKSTART.md`.
    - **Why:** if/when coworkers want to contribute their own adaptations, they'll need this. Not urgent today.
    - **Effort:** S.
    - **Risk:** none.

13. **Architecture overview one-pager** — adopt [`CodexCompare/guides/architecture-overview.md`](guides/architecture-overview.md) layout (system overview in under a page; two-prong diagram; runtime flow; validation strategy; key files by area).
    - **Where in A:** repo root `ARCHITECTURE.md` or `claude-training-lab-code\FinalExport\Guides\ARCHITECTURE.md`.
    - **Why:** currently the only architecture description is embedded in CLAUDE.md lines 16–21 (three bullets). A one-pager helps stakeholders and future AI sessions get oriented in under a minute.
    - **Effort:** S–M.
    - **Risk:** none.

14. **Makefile (`make check`, `make smoke`, `make unit`, `make inventory`)** — if you're willing to install make on Windows (Git Bash includes it). Model on [`CodexCompare/Makefile`](Makefile).
    - **Where in A:** repo root `claude-training-lab-code\Makefile`.
    - **Why:** single-command local verification before commits.
    - **Effort:** S.
    - **Risk:** requires make availability; can skip and use a PowerShell equivalent if preferred.

15. **GitHub Actions CI (`smoke-check.yml`)** — port [`CodexCompare/.github/workflows/smoke-check.yml`](.github/workflows/smoke-check.yml), scoped to running `py_compile` on Project A's Python + `pytest pnl_tests.py` on push.
    - **Where in A:** `claude-training-lab-code\.github\workflows\smoke-check.yml`.
    - **Why:** catches Python syntax breaks automatically. Today you'd only find out during a recording.
    - **Effort:** S.
    - **Risk:** keep it small — don't try to run VBA tests in CI (they can't run headless).

16. **Repo-wide structural smoke check** — lighter version of [`CodexCompare/tests/stage2_smoke_check.py`](tests/stage2_smoke_check.py). Checks Option Explicit on every .bas, LogAction signature isn't misused (grep anti-pattern), sample workbooks exist, guide files are non-empty.
    - **Where in A:** `claude-training-lab-code\tests\stage2_smoke_check.py` or similar.
    - **Why:** Codex's 456 LOC is overkill for Project A; but ~100 LOC of "hard-coded anti-patterns must not appear" would catch the most painful recurring bugs automatically (LogAction-with-Double, `ActiveWorkbook.Close SaveChanges:=False` in a director path, etc.).
    - **Bonus pattern worth including — guide/video section contract.** Codex's smoke check asserts every video script has a `Timestamped Outline`, `Narration`, and `CTA` section, and every guide has its required headings. Cheap add-on that stops half-finished recording guides sneaking into a release.
    - **Effort:** M.
    - **Risk:** false positives — start narrow (LogAction signature check + Option Explicit check).

17. **Run-log sheet schema standardization** — adopt Project B's 8-column `UTL_RunLog` schema inside `modLogger_v2.1.bas` and `modUTL_Audit.bas`: Timestamp | User | Module | Procedure | Status | Message | Sheets | Cells Changed.
    - **Where in A:** `RecTrial\DemoVBA\modLogger_v2.1.bas`.
    - **Why:** your existing logger output is harder to filter on. This schema is queryable and audit-friendly.
    - **Effort:** M.
    - **Risk:** moderate — existing code calls LogAction with positional args; you'd want to add (not replace) the sheet-writer, not break calls.

18. **SQL extract templates for teaching** — adopt [`CodexCompare/sql/universal/template_gl_extract.sql`](sql/universal/template_gl_extract.sql) and `template_revenue_extract.sql` as training artifacts for Video 4 / CoPilot lab.
    - **Where in A:** `RecTrial\UniversalToolkit\sql\` (new folder) or `claude-training-lab-code\FinalExport\DemoPython\sql\templates\`.
    - **Why:** your existing SQL is SQLite-specific; these ANSI-ish templates generalise better to coworkers on SQL Server / Oracle / Snowflake.
    - **Effort:** S.
    - **Risk:** none; additive.

19. **"Suggested talking points" output section** — steal the auto-generated talking-points block from [`CodexCompare/python/universal/build_exec_summary.py`](python/universal/build_exec_summary.py) and add it to `word_report.py`.
    - **Where in A:** `RecTrial\UniversalToolkit\python\word_report.py`.
    - **Why:** nice supplement to the numeric summary — "Revenue grew X%; top contributor Y; watch item Z." Plain English, CFO-grade.
    - **Effort:** S–M.
    - **Risk:** none.

20. **`STARTER_PROMPT.md` / short onboarding prompt** — a one-page prompt that summarises the project for a fresh Claude/Codex session, modelled on [`CodexCompare/STARTER_PROMPT.md`](STARTER_PROMPT.md) and [`START_HERE_PROMPT.md`](START_HERE_PROMPT.md).
    - **Where in A:** repo root `claude-training-lab-code\STARTER_PROMPT.md`.
    - **Why:** CLAUDE.md is too long to cold-start an agent. A 30-line starter dramatically reduces the "new AI session drifts on constraints" problem.
    - **Effort:** S.
    - **Risk:** none.

### Tier 3 — Lower priority, optional

21. **Variance-classifier stdlib script** — port [`variance_classifier.py`](python/demo/variance_classifier.py) into `UniversalToolkit\python\ZeroInstall\` (pairs with item #7).
    - **Effort:** S. **Risk:** none.

22. **Scenario-runner stdlib script** — port [`scenario_runner.py`](python/demo/scenario_runner.py) as a "quick-and-dirty no-dependencies what-if." Not a replacement for Monte Carlo.
    - **Effort:** S. **Risk:** none.

23. **`pnl_data_extract.py` (sheets → CSVs, stdlib)** — quick way to hand a CSV pack to a coworker who doesn't open xlsm.
    - **Where in A:** `UniversalToolkit\python\ZeroInstall\sheets_to_csv.py` (rename; this one is demo-specific in B but generalizes trivially).
    - **Effort:** S. **Risk:** none.

24. **Named-label error handling style** — gradually rename `ErrHandler:` to named labels (`CompareFail:`, `ReconFail:`, etc.) in new code. Not a bulk refactor; just adopt in new modules.
    - **Effort:** per-module S. **Risk:** none. (Honestly a taste call — A's `ErrHandler:` is fine.)

25. **Dual-log pattern for demo modules** — in demo modules, write a short line to a local `VBA_AuditLog` sheet *and* call the universal logger. Pattern from [`modDemo_AuditTrail.bas`](vba/demo/modDemo_AuditTrail.bas).
    - **Effort:** S per module. **Risk:** low (extra writes).

26. **T-SQL variance fact view for teaching** — adopt `demo_variance_fact.sql` as a teaching artifact in Video 4.
    - **Effort:** S. **Risk:** none; additive.

27. **Inline CHANGELOG discipline** — start writing a `CHANGELOG.md` top-level file modelled on [`CodexCompare/CHANGELOG.md`](CHANGELOG.md) (you have `Archive/qa/CHANGELOG.md` already — promote to top-level and keep current).
    - **Effort:** S. **Risk:** none.

28. **`CONTRIBUTING.md`** — only if you ever open the repo to coworkers. Project B's [`CONTRIBUTING.md`](CONTRIBUTING.md) is a useful template.
    - **Effort:** S. **Risk:** none.

29. **4-part video template for any future short explainer** — header pattern from Codex's 5 videos (timestamped outline + on-screen action callouts + business impact + CTA). **Not** for your existing 4 primary videos — those stay.
    - **Effort:** S per future video. **Risk:** none.

30. **Aspirational Universal Tool Catalog** — `universal-tool-catalog.md` pattern (a 160-item forward-looking menu) can serve as a *roadmap* artifact for Project A. Project A already has a better *existing* list (VBA-Module-Reference.pdf); an aspirational catalog is separate and complements.
    - **Effort:** M. **Risk:** don't put an aspirational list in exec materials — label clearly as "roadmap / future additions."

31. **Future-Video-5 option (only if you ever change your mind)** — Codex's [`video-5-copilot-adaptation-lab.md`](videos/video-5-copilot-adaptation-lab.md) is a 6-minute script teaching non-devs a 5-step Microsoft 365 CoPilot workflow for safely adapting demo code to their own workbooks. Names three specific failure patterns + recovery steps, ends with a concrete CTA. Pairs with your existing `AP_Copilot_PromptGuideHelpV2.pdf`. If future-you ever decides to add a fifth video, this is a ready skeleton — adapt, don't start from scratch.
    - **Where in A:** would live alongside your existing video scripts at `RecTrial\VideoScripts\` and `claude-training-lab-code\FinalExport\VideoRecording\`.
    - **Why it's here, not in Tier 1:** current plan explicitly excludes a 5th video. This item just captures the option so it's findable later if the plan changes.
    - **Effort:** M (recording + narration; research is done).
    - **Risk:** none unless/until recorded.

32. **Top-level `CONTEXT.md`** — adopt the one-page "what is this project" pattern from [`CodexCompare/CONTEXT.md`](CONTEXT.md). Distinct from `CONSTRAINTS.md` (rules), `STARTER_PROMPT.md` (AI onboarding), and `CLAUDE.md` (deep governance) — this is a five-minute orientation for a human walking in cold.
    - **Where in A:** repo root `claude-training-lab-code\CONTEXT.md` (~30 lines).
    - **Why:** the same framing info lives inside the first 60 lines of CLAUDE.md today, which is too long and too AI-centric for a human visitor. A short top-level `CONTEXT.md` covers the gap without duplicating CLAUDE.md.
    - **Effort:** S.
    - **Risk:** none; purely additive.

---

## 9. Summary in one paragraph

Project A is **the larger, more mature, feature-rich project** — 67 VBA modules (32,500 LOC), 37 Python scripts, 1,267 LOC of real SQL, a video-demo puppeteer (modDirector), 13 polished user PDFs, and a pytest suite that passes. Project B is **smaller but cleaner at the repo-infrastructure layer**, bringing things Project A doesn't have: a GitHub Actions smoke check, a Makefile, an auto-generated CODE_INVENTORY.md, stdlib-only Python utilities that run with zero pip installs, a release-readiness checklist, an architecture-overview page, a user-facing troubleshooting reference, a git-push quickstart for non-devs, and a `modUTL_Intelligence.bas` that exposes materiality classification / exception narratives / data-quality scoring as universal tools. The top ten cherry-picks (Tier 1) add real durability and coworker-friendliness to Project A with under two days of combined effort and near-zero risk.
