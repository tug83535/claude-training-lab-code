# What's in This Branch — Full Inventory

**Branch:** `claude/resume-ipipeline-demo-qKRHn`
**Last Updated:** 2026-03-09

This document explains every folder and file in the repo so you know exactly what you have.

---

## Folder-by-Folder Breakdown

---

### `vba/` — Demo File VBA Modules (34 files)

These are the VBA source code files that get imported into the Excel demo file. They power all 62 Command Center actions.

| File | What It Does |
|------|-------------|
| `modConfig_v2.1.bas` | All constants — sheet names, products, fiscal year, colors, thresholds. The "settings file" for everything else |
| `modFormBuilder_v2.1.bas` | Builds the Command Center pop-up form + routes all 62 actions to the right macro |
| `modMasterMenu_v2.1.bas` | Text-based fallback menu (InputBox) if the Command Center form isn't available |
| `modNavigation_v2.1.bas` | Table of Contents, GoHome, keyboard shortcuts (Ctrl+Shift combos), Toggle Executive Mode |
| `modDashboard_v2.1.bas` | Charts — Revenue Trend, Expense Breakdown, basic chart tools, LinkDynamicChartTitles |
| `modDashboardAdvanced_v2.1.bas` | Executive Dashboard, Waterfall Chart, Product Comparison, Small Multiples Grid |
| `modDataQuality_v2.1.bas` | 6 data quality scans + FixTextNumbers + Data Quality Letter Grade (A-F) |
| `modDataSanitizer_v2.1.bas` | Numeric-only sanitizer — fixes floating-point tails, text-stored numbers. Smart enough to skip dates/names/IDs |
| `modReconciliation_v2.1.bas` | Reconciliation Checks sheet — PASS/FAIL for 4 cross-sheet validations |
| `modVarianceAnalysis_v2.1.bas` | Month-over-Month variance analysis + auto-commentary + YoY Variance Analysis |
| `modPDFExport_v2.1.bas` | Batch PDF export — dynamically discovers all monthly tabs, professional headers/footers |
| `modPerformance_v2.1.bas` | TurboOn/TurboOff (speed optimization) + timer + status bar updates |
| `modMonthlyTabGenerator_v2.1.bas` | Clones monthly tabs (Mar→Apr, etc.) + AddNextMonthToModel (calendar-aware) |
| `modSearch_v2.1.bas` | Cross-sheet search — finds text across all sheets, 200-result cap, yellow highlight |
| `modUtilities_v2.1.bas` | 12 utility macros (actions 51-62) — freeze panes, zoom, hide/show sheets, etc. |
| `modLogger_v2.1.bas` | Runtime audit log — writes every action to hidden VBA_AuditLog sheet + ViewLog |
| `modSheetIndex_v2.1.bas` | Creates a Home sheet with Command Center button + sheet index with hyperlinks |
| `modDemoTools_v2.1.bas` | AddControlSheetButtons, SetParameterizedPrintArea, CreatePrintableExecSummary, CreateDisclaimerSheet |
| `modDataGuards_v2.1.bas` | ValidateAssumptionsPresence, CheckSumOfDrivers, FindNegativeAmounts, FindZeroAmounts, FindSuspiciousRoundNumbers |
| `modDrillDown_v2.1.bas` | Drill links from P&L to GL detail, reconciliation checks, heatmap, GoldenFileCompare |
| `modAuditTools_v2.1.bas` | ChangeLog, FindExternalLinks, FixExternalLinks, AuditHiddenSheets, CreateMaskedCopy, ExportErrorSummary |
| `modETLBridge_v2.1.bas` | TriggerETLLocally, ImportETLOutput — bridges VBA to Python data pipelines |
| `modTrendReports_v2.1.bas` | Rolling 12-Month View, Reconciliation Trend Chart, ArchiveReconciliationResults |
| `modSensitivity_v2.1.bas` | Sensitivity analysis on Assumptions sheet drivers |
| `modAWSRecompute_v2.1.bas` | AWS allocation validation and recalculation |
| `modImport_v2.1.bas` | CSV/Excel data import pipeline |
| `modForecast_v2.1.bas` | Rolling forecast + trend append |
| `modScenario_v2.1.bas` | Scenario save/load/compare/delete |
| `modAllocation_v2.1.bas` | Cost allocation engine + preview |
| `modConsolidation_v2.1.bas` | Multi-entity consolidation + intercompany eliminations |
| `modVersionControl_v2.1.bas` | Version save/compare/restore |
| `modAdmin_v2.1.bas` | Auto-documentation + change management |
| `modIntegrationTest_v2.1.bas` | 18-test integration suite + quick health check |
| `frmCommandCenter_code.txt` | The UserForm code for the Command Center pop-up (copy-paste into VBA Editor) |

---

### `python/` — Demo File Python Scripts (14 files)

Python scripts that support the demo P&L file. All 14 are complete and functional. 99 tests pass.

| File | What It Does |
|------|-------------|
| `pnl_config.py` | Central config — revenue shares, product names, fiscal year settings |
| `pnl_runner.py` | Main runner — executes the full P&L pipeline end to end |
| `pnl_cli.py` | Command-line interface — run scripts from terminal with flags |
| `pnl_tests.py` | Pytest test suite — 99 tests covering all scripts |
| `pnl_forecast.py` | Rolling forecast + Forecast Accuracy Scoring (MAPE/bias/hit rate) |
| `pnl_dashboard.py` | Dashboard data generator |
| `pnl_month_end.py` | Month-end close automation |
| `pnl_snapshot.py` | Point-in-time snapshot of P&L data |
| `pnl_allocation_simulator.py` | Cost allocation simulator |
| `pnl_ap_matcher.py` | AP invoice matching |
| `pnl_monte_carlo.py` | Monte Carlo P&L risk simulation |
| `build_charts.py` | Chart builder for Excel |
| `redesign_pl_model.py` | P&L model redesign script (used during initial build) |
| `requirements.txt` | Python dependencies |

---

### `sql/` — SQL Scripts (4 files)

| File | What It Does |
|------|-------------|
| `staging.sql` | Staging table creation |
| `transformations.sql` | Data transformation queries |
| `validations.sql` | Data validation checks |
| `pnl_enhancements.sql` | P&L-specific enhancements and joins |

---

### `excel/` — The Demo Excel File

| File | What It Does |
|------|-------------|
| `KeystoneBenefitTech_PL_Model.xlsx` | The main P&L demo file. All 34 VBA modules get imported into this. This is what you open and demo |

---

### `FinalRoughGuides/` — Training Guides (7 files, draft status)

These are the 6 main training guides + 1 extra guide. All written for non-technical Finance & Accounting coworkers. Currently in draft — move to `training/` after review and approval.

| File | What It Does |
|------|-------------|
| `01-How-to-Use-the-Command-Center.md` | Full walkthrough of all 62 actions, monthly close workflow, tips, troubleshooting, FAQ |
| `02-Getting-Started-First-Time-Setup.md` | Download, open, enable macros, trust center, verify it works, first 5 actions to try |
| `03-What-This-File-Does-Leadership-Overview.md` | CFO/CEO briefing — business impact, before/after, cost savings, rollout plan |
| `04-Quick-Reference-Card.md` | 1-page cheat sheet — all 62 actions, keyboard shortcuts, monthly close sequence |
| `05-Video-Demo-Script-and-Storyboard.md` | 18-22 min demo video script — 3 parts, shot lists, word-for-word narration |
| `06-Universal-Toolkit-Guide.md` | All ~100 universal tools (79 VBA + 22 Python), setup, playbooks, top 20 |
| `Dynamic-Chart-Filter-Setup-Guide.md` | How to add dropdown chart filters to any Excel file (bonus guide) |

---

### `CoPilotPromptGuide/` — CoPilot Prompt Guide (2 files)

| File | What It Does |
|------|-------------|
| `AP_Copilot_PromptGuideHelpV2.md` | Prompt library teaching coworkers how to use Microsoft CoPilot with VBA/Python code |
| `AP_Copilot_PromptGuideHelpV2.docx` | Same guide in Word format (original upload) |

---

### `videodraft/` — Video Demo Planning (3 files)

| File | What It Does |
|------|-------------|
| `COMPILED_VIDEO_PACKAGE.md` | The master video package — tool counts, demo file stats, build checklist, structure for all 3 videos |
| `VIDEO_DEMO_PLAN.md` | Original video demo plan — flow, recording tips, open questions |
| `AI_BRIEFING_VIDEO_REVIEW.md` | AI-generated briefing/review notes for the video content |

---

### `UniversalToolsForAllFiles/` — Universal Toolkit (~100 tools)

These are tools that work on ANY Excel file — not just the demo file. Intended for coworkers to use on their own spreadsheets (Scenario 2 — after the demo).

**VBA Modules (13 files):**

| File | Tools |
|------|-------|
| `vba/modUTL_Core.bas` | 9 shared helper functions used by other modules |
| `vba/modUTL_DataCleaning.bas` | 12 data cleaning tools |
| `vba/modUTL_DataSanitizer.bas` | 4 sanitizer tools (RunFullSanitize, PreviewChanges, FixFloatingPoint, ConvertTextNumbers) |
| `vba/modUTL_Formatting.bas` | 9 formatting tools |
| `vba/modUTL_Finance.bas` | 14 finance-specific tools |
| `vba/modUTL_Audit.bas` | 8 audit tools |
| `vba/modUTL_WorkbookMgmt.bas` | 15 workbook management tools |
| `vba/modUTL_Branding.bas` | 2 tools — apply iPipeline branding + set theme colors |
| `vba/modUTL_SheetTools.bas` | 4 tools — sheet index, template cloner, customer IDs, create folders |
| `vba/NewTools/modUTL_DataCleaningPlus.bas` | 3 extra cleaning tools |
| `vba/NewTools/modUTL_AuditPlus.bas` | 4 extra audit tools |
| `vba/NewTools/modUTL_DuplicateDetection.bas` | 1 tool — ExactDuplicateFinder |
| `vba/NewTools/modUTL_NumberFormat.bas` | 2 tools — EnhancedTextToNumberConverter, WorkbookMetadataReporter |

**Python Scripts (22 files):**

| Folder | Files |
|--------|-------|
| `python/` | 18 scripts — aging_report, bank_reconciler, batch_process, clean_data, compare_files, consolidate_budget, consolidate_files, forecast_rollforward, fuzzy_lookup, gl_reconciliation, master_data_mapper, pdf_extractor, reconciliation_exceptions, regex_extractor, unpivot_data, variance_analysis, variance_decomposition, word_report |
| `python/NewTools/` | 4 scripts — date_format_unifier, multi_file_consolidator, sql_query_tool, two_file_reconciler |
| `python/requirements.txt` | Python dependencies for universal tools |

**Docs:**
| File | What It Does |
|------|-------------|
| `README.md` | Overview of the universal toolkit |
| `UNIVERSAL_TOOLS_HOW_TO_GUIDE.md` | Full how-to guide for all tools — written for non-technical coworkers |
| `UniversalBuild/UNIVERSAL_BUILD_CANDIDATES.md` | The original 76 candidate list that was compiled and built |

---

### `docs/` — Project Documentation

| Folder / File | What It Does |
|---------------|-------------|
| `ipipeline-brand-styling.md` | Official iPipeline brand colors, fonts, and styling rules |
| `day-to-day/OPERATIONS_RUNBOOK.md` | Day-to-day operations reference |
| `day-to-day/SANITIZATION_PLAYBOOK.md` | Data sanitization procedures |
| `day-to-day/USER_TRAINING_GUIDE.md` | User training reference |
| `day-to-day/VBA-Re-Import-Guide.md` | Step-by-step guide for importing .bas files into Excel |
| `overview/ARCHITECTURE_DIAGRAM.md` | System architecture — how all the pieces connect |
| `overview/CODE_COMPARISON_REPORT.md` | Code comparison scorecard |
| `overview/EXECUTIVE_SUMMARY.md` | Executive summary of the project |
| `setup/IMPLEMENTATION_GUIDE.md` | Implementation guide |
| `setup/KBT_File_Map.pdf` | Visual file map (PDF) |
| `setup/QUICK_START.md` | Quick start guide |
| `setup/START_TO_FINISH_GUIDE.md` | Full start-to-finish setup walkthrough |
| `setup/WORKBOOK_SETUP_NOTES.md` | Workbook-specific setup notes |

---

### `qa/` — Quality Assurance (6 files)

| File | What It Does |
|------|-------------|
| `TEST_PLAN.md` | Master test plan — 69 tests across 8 categories (T1-T8) |
| `CHANGELOG.md` | Change log of all modifications |
| `INTEGRATION_TEST_GUIDE.md` | Integration testing procedures |
| `ISSUE_CLOSURE.md` | Issue resolution tracking |
| `VALIDATION_REPORT.md` | Validation results |
| `logging_template.csv` | CSV template for test logging |

---

### `tasks/` — Session Management (2 files)

| File | What It Does |
|------|-------------|
| `todo.md` | Running task list — updated every session with current status, completed items, and what's next |
| `lessons.md` | Log of mistakes and lessons learned — reviewed every session to avoid repeating errors |

---

### `review/` — External Review Materials

| File | What It Does |
|------|-------------|
| `PROJECT_OVERVIEW.md` | Full project overview written for an external Claude session to review |
| `CODE_INVENTORY.md` | Complete code inventory |
| `DemoWrapUp/` | Empty folder — was reserved for wrap-up review docs |

---

### `ProjectRefresh/` — Code Audit from Other Claude Session

| File | What It Does |
|------|-------------|
| `CODE_AUDIT_FINDINGS.md` | 120 tools cross-referenced between Claude sessions (34 exact, 30 overlap, 56 new ideas) |
| `Tool_Reference_All_120.md` | Full catalog of all 120 tools |
| `RefreshCompareMerge/Code_Audit_Final_Report.md` | Final audit report — Demo A-, Universal B+, 20 prioritized recommendations |
| `Universal_Toolkit_Chat1_Foundation.docx` | Original foundation doc from other Claude session |

---

### `NewTesting/` — Research & Ideas (3 files remaining)

| File | What It Does |
|------|-------------|
| `2026-02-28T223817Z.md` | Full audit doc — 15 issues, 10 VBA macros, Python ETL, Power Query M-Code |
| `Financial Model Correction Instructions.md` | 6-point fix checklist for the Excel model |
| `VBA Examples (200) — Name — Purpose.txt` | Catalog of 200 macro ideas (source for the 7 new modules built on 3/1) |

---

### `DemofileChartBuild/` — Chart Work in Progress

| File | What It Does |
|------|-------------|
| `chartexampleAP.xlsx` | Chart sheet example file (work in progress) |

---

### Other Root Files

| File | What It Does |
|------|-------------|
| `CLAUDE.md` | Master instructions file — project context, session summaries, rules, current status |
| `README.md` | Professional repo README |
| `.gitignore` | Git ignore rules |
| `TESTRUN/START_HERE.md` | Test run starting point doc |
| `Testing_Issues/TESTING_ISSUES_LOG.md` | Log of issues found during testing |

---

### `CompletePackageStorage/` — Final Production Files (empty)

| Folder | What It Does |
|--------|-------------|
| `production/` | Where the final ready-to-go files go after everything is approved |
| `backups/` | Where dated backup copies go |

---

### `training/` — Final Approved Guides (empty)

This is where guides move AFTER Connor reviews and approves them from `FinalRoughGuides/`.

---

## By the Numbers

| Category | Count |
|----------|-------|
| Demo VBA modules | 34 |
| Demo Python scripts | 14 |
| SQL scripts | 4 |
| Command Center actions | 62 |
| Universal VBA modules | 13 |
| Universal Python scripts | 22 |
| Universal tools total | ~100 |
| Training guides (draft) | 7 |
| Total tests in test plan | 69 |
| Tests passed | 15 |
| Bugs found and fixed (all sessions) | 30+ |
