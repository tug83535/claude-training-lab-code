# Project Todo — APCLDmerge (iPipeline P&L Demo)

## Current Status (2026-03-05)
- **Branch:** `claude/resume-ipipeline-demo-qKRHn` (active working branch)
- **VBA Modules:** 34 total (32 demo + modSheetIndex + modDashboardAdvanced) — need re-import
- **Python Scripts:** 14 complete and functional (main project) + pnl_forecast.py enhanced with MAPE accuracy
- **Universal Tools:** ~99 tools built (12 VBA modules + 22 Python scripts), code-reviewed, bugs fixed
- **Demo Enhancements (2026-03-05):** Data Quality Letter Grade, Forecast Accuracy (MAPE), YoY Variance Analysis, modDashboard split into 2 modules
- **Testing Phase:** T1 COMPLETE, T2 partially done (T2.01–T2.04 PASS, T2.05–T2.07 not yet run), T5.01+T5.02 PASS
- **Self-Review:** Full self-review of all remaining untested code completed — 12 bugs found and fixed preemptively
- **Overall:** Track B COMPLETE, Track C COMPLETE, Backlog Item 1 COMPLETE

### ⚠ IMPORTANT — RE-IMPORT NEEDED
Before continuing testing, re-import these 7 `.bas` files into the Excel workbook (VBA Editor → File → Import):
1. modConfig_v2.1.bas (color constant fixes)
2. modReconciliation_v2.1.bas (dateCol/amtCol constant fixes)
3. modVarianceAnalysis_v2.1.bas (GenerateCommentary row 1 → row 4 fix)
4. modDashboard_v2.1.bas (WaterfallChart fallbacks + ExecDashboard fixes + LogAction fixes)
5. modDemoTools_v2.1.bas (LogAction fix)
6. modTrendReports_v2.1.bas (LogAction fix)
7. modMonthlyTabGenerator_v2.1.bas (LogAction fixes + TestUpdateHeaderText wrapper)

### Project Refresh — Code Audit / Idea Review (NEW)
A separate Claude session (fresh, no context) independently built VBA and Python code for this same project. The goal is to review that code for ideas and improvements we may have missed — **NOT** to replace or change any of our existing code.

**Rules:**
- Do NOT modify any existing code based on this review
- Review only — look for new ideas, missed features, or better approaches
- Any ideas worth pursuing get added to the backlog as new items
- All uploaded code goes in `ProjectRefresh/` folder

**Steps:**
1. [x] Create `ProjectRefresh/` folder on the branch
2. [ ] Connor uploads the other Claude's code into `ProjectRefresh/`
3. [ ] Full code audit — review every file in `ProjectRefresh/` line by line
4. [ ] Cross-reference against our existing 215+ tools (33 demo VBA + 85 universal tools + 14 Python + extras)
5. [ ] Produce a findings report: new ideas, overlaps, and recommendations
6. [ ] Add any approved new ideas to the backlog

---

### ONE ACTIVE TRACK — Testing (T2.05 next)
1. Read this file, tasks/lessons.md, CLAUDE.md, and qa/TEST_PLAN.md first
2. Re-import the 7 fixed .bas files listed above
3. Resume testing at **T2.05** (FixTextNumbers requires scan), then T2.06, T2.07
4. Then proceed through T3, T4, T5.03–T5.06, T6, T7, T8 in order
5. Log all results in qa/TEST_PLAN.md Section 6 (Test Execution Results)
6. Any new issues found → add to Testing_Issues/TESTING_ISSUES_LOG.md

### Universal Tools — COMPLETE (Track B + Track C + Backlog Item 1)
~85 tools built (8 modules), code-reviewed, 9 bugs fixed, how-to guide written. 3 new modules added 2026-03-04 (Branding, SheetTools, DataSanitizer). No further action needed until after demo.

**Security note:** Profiles.md contains personal account emails/credentials and should be removed from the repo. Flag this to Connor.

---

**TRACK C — Universal Tools Code Review (Bug Verification) — COMPLETE (2026-03-03)**
All 76 tools reviewed against the full checklist. 9 bugs found and fixed:

VBA bugs fixed (8):
- [x] modUTL_Audit: CircularReferenceDetector crashed on sheets with no circular refs (For Each on Nothing)
- [x] modUTL_Audit: InconsistentFormulasAuditor used raw formulas instead of R1C1 — flagged every row as different
- [x] modUTL_Audit: NamedRangeAuditor used MacroType for scope detection — now checks for "!" in name
- [x] modUTL_Finance: JournalEntryValidator dr/cr not reset to 0 each loop — stale values corrupted balance
- [x] modUTL_Finance: FluxAnalysis wrote over existing columns — now inserts columns first
- [x] modUTL_Finance: FinancialPeriodRollForward iterated all 16K+ cells — now limited to used range
- [x] modUTL_WorkbookMgmt: BuildDistributionReadyCopy double-Replace created _DIST_DIST.xlsx for .xlsm files
- [x] modUTL_DataCleaning: Removed unused variable (dead code)

Python bugs fixed (1):
- [x] clean_data.py: Removed deprecated infer_datetime_format parameter (pandas 2.0+)

All fixes committed and pushed (commit a22dd76).

---

## Next Up — Demo Readiness (Priority Order)

### Phase 1: Make It Real (Import + Live Test)
- [x] Import all 31 .bas files into the Excel workbook via VBA Editor
- [x] Fix modAllocation VB_Name attribute mismatch (was importing as AuditTools — fixed)
- [x] Fix all compile errors — Debug > Compile now passes clean
- [x] **BUG FIXED:** modFormBuilder "Too many line continuations" — replaced single 62-item Array() with individual AddAction calls. No line continuations in generated code at all.
- [x] Re-import fixed modFormBuilder into Excel, re-ran LaunchCommandCenter → form built successfully
- [ ] Live test every Command Center action (1-62) in Excel — log pass/fail for each
- [ ] Fix any runtime issues discovered during testing
- [ ] Verify all hidden sheets are created properly (VBA_AuditLog, Scenarios, Version History, etc.)

### Phase 2: Script the Demo Video
- [ ] Write demo video storyboard — which features to show, in what order, talking points
- [ ] Identify the 10-15 most impressive actions to highlight (not all 62)
- [ ] Plan screen recording flow (open file → Command Center → run features → show results)
- [ ] Write speaker notes / narration script

### Phase 3: Training Materials
- [ ] **Guide Planning Session** — Build out guide plan in `FinalRoughGuides/GuidePlanOut/`
  - Create `GuidePlanOut/` folder and `GUIDE_IDEAS.md` to capture all planning
  - Planned guide categories (draft — refine with Connor before building):
    1. **"How to Use the Command Center"** — Main guide for all employees. Open the file, click the button, what every action does
    2. **"Getting Started / First Time Setup"** — Enable macros, trust settings, where to save, what to expect
    3. **"What This File Does" (Overview for Leadership)** — Non-technical explainer for CFO/CEO. What it is, why it matters, what it solves
    4. **"Universal Toolkit Guide"** — Install and use the 85 universal tools on your own files (Scenario 2 / post-demo)
    5. **"Video Demo Script / Storyboard"** — Script and flow for the walkthrough video
    6. **"Quick Reference Card"** — 1-page cheat sheet of top 10-15 most useful actions
  - Open questions to resolve before building:
    - Who gets which guide? All 2,000 employees or just Finance & Accounting?
    - Does CFO/CEO get a separate shorter version?
    - Video script — plan here or separate effort?
    - Any additional guides Connor has in mind?
  - All rough drafts go in `FinalRoughGuides/`, finals go in `training/`
- [ ] Build coworker training guide — step-by-step: how to open file, use Command Center, run reports
- [ ] Create quick-reference card of all 62 actions (1-page printable)
- [ ] Place completed guides in `training/` folder

### Phase 4: Lock Down the Demo File
- [ ] Save the final tested workbook as `.xlsm` (macros-enabled)
- [ ] Open it fresh on a different machine or clean Excel session — confirm it works out of the box
- [ ] Check that no personal file paths, test data, or debug code is left in the macros
- [ ] Copy final `.xlsm` to `CompletePackageStorage/production/`
- [ ] Copy a dated backup to `CompletePackageStorage/backups/` (e.g., `PL_Model_FINAL_2026-03-10.xlsm`)

### Phase 5: Convert Guides to PDF
- [ ] Convert coworker training guide to PDF (no .md files for coworkers)
- [ ] Convert quick-reference card to PDF
- [ ] Save PDFs in `training/` folder AND `CompletePackageStorage/production/`

### Phase 6: Record the Demo Video
- [ ] Do a dry run to practice timing and flow
- [ ] Record the screen + narration
- [ ] Save the video file in `CompletePackageStorage/production/`

### Phase 7: Upload to SharePoint
- [ ] Create `iPipeline P&L Demo/` folder on SharePoint
- [ ] Create 4 subfolders: `Demo File/`, `Training/`, `Universal Tools (Optional)/`, `Video/`
- [ ] Upload the `.xlsm` to `Demo File/`
- [ ] Upload the 2 training PDFs to `Training/`
- [ ] Upload the video to `Video/`
- [ ] Pin the folder or add to team Quick Links
- [ ] Set permissions — one group for the whole folder

### Phase 8: Universal Tools Upload (Later — Scenario 2)
- [ ] Convert the 18 Python scripts to `.exe` files (PyInstaller)
- [ ] Convert the Universal Tools how-to guide to PDF
- [ ] Upload `.bas` files, `.exe` files, and PDF to `Universal Tools (Optional)/` on SharePoint
- [ ] (Eventually) Package the 8 VBA tools into `KBT_UniversalTools.xlam` add-in

### Final Step
- [ ] Final PR to merge `claude/resume-apclmerge-project-V8WSj` → `main`

---

## Backlog (Nice to Have — After Demo)
- [ ] Dynamic Progress Bar KPI Shape (visual KPI indicator on dashboard)
- [ ] Financial Statement Generator from Trial Balance (requires account mapping design)
- [ ] VBA Outlook Email Integration (PDF → Email in one workflow)
- [ ] Build remaining monthly summary tabs (Apr-Dec) using modMonthlyTabGenerator
- [ ] **Scenario 2 — Universal Tools Add-In:** Package the 8 universal tools (Data Sanitizer, Find Negatives/Zeros/Round Numbers, Find External Links, Audit Hidden Sheets, Cross-Sheet Search) into `KBT_UniversalTools.xlam` so coworkers can use them on their own files. Source files staged in `UniversalToolsForAllFiles/`. Write coworker install guide when ready.
- [x] **Universal Tools — Coworker How-To Guide:** COMPLETE (2026-03-03) — Full guide at `UniversalToolsForAllFiles/UNIVERSAL_TOOLS_HOW_TO_GUIDE.md`. Covers all 76 tools with installation, step-by-step usage, examples, and quick reference table. Written for non-technical Finance & Accounting staff.
- [ ] **Universal Tools — Python .exe Conversion:** Convert all 18 Python scripts to standalone `.exe` files using PyInstaller (or similar) so coworkers can just double-click and run — no Python installation required. Package with a simple folder + README.

### Enhancement Ideas — Top 3 — ALL BUILT (2026-03-05)
1. **[x] Data Quality Letter Grade (A-F)** — BUILT. Added CalculateLetterGrade to modDataQuality. Grade badge (28pt, color-coded) at top of report. Grading: A (0 issues) through F (4+ critical).
2. **[x] Forecast Accuracy Scoring (MAPE/bias)** — BUILT. Added accuracy_report() to pnl_forecast.py. Leave-one-out backtest, MAPE/bias/hit rate metrics, letter grade. CLI flag: --accuracy.
3. **[x] YoY Variance Analysis** — BUILT. Added RunYoYVarianceAnalysis to modVarianceAnalysis. Smart column detection (Prior Year/PY/Budget fallback). Cost-line reversal logic. Creates styled report sheet.

---

## Dropped by User (Do Not Build)
- ~~Backup Workbook with Timestamp macro~~ — user declined (2026-02-28)
- ~~VBA Timestamp Audit Trail on Cell Changes~~ — user declined (2026-02-28)
- ~~Export All Charts to PowerPoint~~ — user dropped permanently (2026-02-28)

---

## Completed — This Session (2026-03-05)

### Demo Enhancements + Universal Tools Expansion
- [x] Fixed 3 demo bugs: duplicate constants in modConfig, GL sheet visibility leak, missing TurboOn/Off in scanning loops
- [x] Built Data Quality Letter Grade (A-F) in modDataQuality
- [x] Built Forecast Accuracy Scoring (MAPE) in pnl_forecast.py
- [x] Built YoY Variance Analysis in modVarianceAnalysis
- [x] Split modDashboard (1,398 lines) into modDashboard (533) + modDashboardAdvanced (650) — 34th VBA module
- [x] Built 14 new universal tools in NewTools/ folder:
  - VBA: modUTL_DataCleaningPlus (3 tools), modUTL_AuditPlus (4 tools), modUTL_DuplicateDetection (1 tool), modUTL_NumberFormat (2 tools)
  - Python: date_format_unifier.py, sql_query_tool.py, multi_file_consolidator.py, two_file_reconciler.py
- [x] SpecialCells performance fix for 8 slow universal tools across 3 modules
- [x] Created modUTL_Core.bas shared utilities module (9 public functions)
- [x] Added backup-before-destructive to 6 high-risk universal tools

---

## Completed — Previous Session (2026-03-04)

### Testing Bug Fixes + Self-Review — Branch V8WSj
- [x] Fixed T2.01: Added 9 missing sheet-name constants to modConfig
- [x] Fixed T2.03: CLR_NAVY and CLR_ALT_ROW color constants (VBA BGR byte order)
- [x] Fixed T2.04: Added TestUpdateHeaderText wrapper + NumberFormat text fix
- [x] Fixed T4.04: Windows PermissionError on temp file cleanup
- [x] Fixed T5.01: CreateExecutiveDashboard row 1 → row 4 + Error 5 crash + row/column detection
- [x] Fixed T5.02: WaterfallChart multi-variant row label fallbacks
- [x] Self-review of ALL remaining untested VBA against test plan pass criteria
- [x] Found and fixed 12 additional bugs preemptively (commit 22ba831):
  - modReconciliation: dateCol=5 → COL_GL_DATE, amtCol=7 → COL_GL_AMOUNT
  - modVarianceAnalysis: GenerateCommentary row 1 → HDR_ROW_REPORT
  - 9 LogAction calls: elapsed Double passed as status String → moved into message
- [x] Python pytest: 99 passed, 15 skipped, 0 failures
- [x] Added Pre-Delivery Self-Review Requirement to lessons.md

---

## Completed — Previous Session (2026-03-03)

### Track C — Universal Tools Code Review (Bug Verification) — COMPLETE
- [x] Read all 5 VBA modules line by line against the Track C checklist
- [x] Read all 18 Python scripts line by line against the Track C checklist
- [x] Found 9 bugs (8 VBA, 1 Python) — 4 critical, 4 moderate, 1 minor
- [x] Fixed all 9 bugs in-place
- [x] Committed and pushed (commit a22dd76)

### Backlog Item 1 — Coworker How-To Guide — COMPLETE
- [x] Wrote full how-to guide for all 76 Universal Tools
- [x] Saved at UniversalToolsForAllFiles/UNIVERSAL_TOOLS_HOW_TO_GUIDE.md
- [x] Committed and pushed (commit 199f983)

### Updated tasks/todo.md and tasks/lessons.md

---

## Completed — Previous Session (2026-03-02)

### Universal Tools — All Code Built (Track B Complete)
- [x] Reviewed GrokALL.md, PrelexALL.md, GemAll.md — curated 76 tool candidates
- [x] Created UniversalToolsForAllFiles/UniversalBuild/UNIVERSAL_BUILD_CANDIDATES.md
- [x] Built modUTL_DataCleaning.bas — 12 VBA tools
- [x] Built modUTL_Formatting.bas — 9 VBA tools
- [x] Built modUTL_WorkbookMgmt.bas — 15 VBA tools
- [x] Built modUTL_Finance.bas — 14 VBA tools
- [x] Built modUTL_Audit.bas — 8 VBA tools
- [x] Built 18 Python scripts covering all Tier 1 + Tier 2 Python candidates
- [x] Created requirements.txt for all Python dependencies
- [x] Created review/PROJECT_OVERVIEW.md — full project overview for external Claude review
- [x] Created review/DemoWrapUp/ folder — ready for Connor's review document

---

### T1 Testing — ALL PASS (T1.01 through T1.08)
- [x] T1.01 — PASS — VBA project compiled with zero errors
- [x] T1.02 — PASS — All 32 modules visible in Project Explorer
- [x] T1.03 — PASS — Option Explicit found in all modules
- [x] T1.04 — PASS — ?APP_VERSION returned "2.1.0"
- [x] T1.05 — PASS — pnl_config.py printed full config summary, all shares sum to 1.00
- [x] T1.06 — PASS — All 14 Python scripts imported successfully, printed "All imports OK"
- [x] T1.07 — PASS — All 14 files valid UTF-8 (non-ASCII is intentional Unicode, not mojibake)
- [x] T1.08 — PASS — pip install -r requirements.txt completed successfully (all packages installed)

### T2 Testing — T2.01 through T2.04 PASS (T2.05–T2.07 not yet run)
- [x] T2.01 — PASS (after fix: 9 missing constants added to modConfig)
- [x] T2.02 — PASS (SafeDeleteSheet works)
- [x] T2.03 — PASS (after fix: CLR_NAVY/CLR_ALT_ROW color constants corrected)
- [x] T2.04 — PASS (after fix: TestUpdateHeaderText wrapper + NumberFormat text)
- [ ] T2.05 — Not yet run
- [ ] T2.06 — Not yet run
- [ ] T2.07 — Not yet run

### T4 Testing — T4.04 PASS
- [x] T4.04 — PASS (after fix: PermissionError on temp file + email feature removed) — 99 passed, 15 skipped, 0 failures

### T5 Testing — T5.01 and T5.02 PASS
- [x] T5.01 — PASS (after fix: ExecDashboard row detection + Error 5 crash)
- [x] T5.02 — PASS (after fix: WaterfallChart multi-variant row label fallbacks)
- [ ] T5.03–T5.06 — Not yet run

### T3, T6, T7, T8 — Not yet started

---

## Completed — This Session (2026-03-01)

### NewTesting File Review
- [x] Reviewed 3 new files added to NewTesting/ (commit 075d457)
- [x] Created new Ideas branch: `claude/ideas-newtesting-wDuOY`

### 7 New VBA Modules Built (from VBA Examples 200 list)
- [x] modDemoTools_v2.1.bas — #17 AddControlSheetButtons, #63 SetParameterizedPrintArea, #64 CreatePrintableExecSummary
- [x] modDataGuards_v2.1.bas — #48 ValidateAssumptionsPresence, #49 CheckSumOfDrivers, #150 FindNegativeAmounts, #151 FindZeroAmounts, #155 FindSuspiciousRoundNumbers
- [x] modDrillDown_v2.1.bas — #18 AddReconciliationDrillLinks, #55 AutoPopulateReconciliationChecks, #56 ApplyReconciliationHeatmap, #90 RunGoldenFileCompare
- [x] modAuditTools_v2.1.bas — #93 AppendChangeLogEntry, #106 FindExternalLinks, #107 FixExternalLinks, #109 AuditHiddenSheets, #115 CreateMaskedCopy, #196 ExportErrorSummaryClipboard
- [x] modETLBridge_v2.1.bas — #119 TriggerETLLocally, #120 ImportETLOutput
- [x] modTrendReports_v2.1.bas — #77 CreateRolling12MonthView, #156 CreateReconciliationTrendChart, #163 ArchiveReconciliationResults
- [x] modDashboard_v2.1.bas updated — added #44 LinkDynamicChartTitles, #86 CreateSmallMultiplesGrid

### Data Sanitizer Module
- [x] modDataSanitizer_v2.1.bas — numeric-only sanitizer (never touches dates, names, customer IDs)
- [x] Updated SKIP_HEADER_KEYWORDS to include: customer, client, account, acct, company, vendor, contact, employee, entity, description, dept, product, type, status, label, region, country, city, address

### Calendar-Aware Month Expander
- [x] AddNextMonthToModel added to modMonthlyTabGenerator_v2.1.bas
  - Reads today's date to determine next calendar month automatically
  - Marks next month column yellow on P&L Monthly Trend
  - Marks next month column yellow on Functional P&L Monthly Trend
  - Clones current month's Functional P&L Summary tab to create next month's tab
  - Added MarkTrendColumn private helper

---

## Completed — This Session (2026-02-28)

### Branch Merge
- [x] Reviewed all 5 branches across 3 Claude accounts and mapped progress
- [x] Merged Track A: Excel redesign (Fortune 100 FP&A styling, 8 charts, redesigned workbook)
- [x] Merged Track B: Code improvements (Logger, Utilities, Monte Carlo, SQL fixes, repo cleanup)
- [x] Resolved merge conflict in `tasks/todo.md` (combined both tracks)
- [x] Pushed unified branch `claude/review-branch-progress-pP7Qf`

### Full Audit
- [x] Audited all 24 VBA modules — identified 11 working, 3 with bugs, 10 missing
- [x] Audited all 14 Python scripts — all complete and functional
- [x] Produced full inventory list with working/broken/unbuilt categorization

### 10 New VBA Modules Built
- [x] modSensitivity_v2.1.bas — Sensitivity analysis (Action 5)
- [x] modAWSRecompute_v2.1.bas — AWS allocation validation/recalc (Action 14)
- [x] modImport_v2.1.bas — Data import pipeline (Action 17)
- [x] modForecast_v2.1.bas — Rolling forecast + trend append (Actions 18-19)
- [x] modScenario_v2.1.bas — Scenario save/load/compare/delete (Actions 20-23)
- [x] modAllocation_v2.1.bas — Cost allocation engine + preview (Actions 24-25)
- [x] modConsolidation_v2.1.bas — Multi-entity consolidation (Actions 26-30)
- [x] modVersionControl_v2.1.bas — Version save/compare/restore (Actions 31-35)
- [x] modAdmin_v2.1.bas — Auto-documentation + change management (Actions 36-40)
- [x] modIntegrationTest_v2.1.bas — 18-test suite + quick health check (Actions 44-45)

### Bug Fixes (4)
- [x] modLogger: Added ViewLog procedure (Action 41 was missing its target)
- [x] modNavigation: Fixed Ctrl+Shift+R shortcut wiring + added ToggleExecutiveMode (Action 48)
- [x] modConfig: Added RECON_TOLERANCE constant (used by modReconciliation but not defined)
- [x] modReconciliation: Fixed StyleHeader call (was passing 4 args instead of 3)
- [x] modFormBuilder: Fixed install guide text from "50 actions" to "62 actions"

## Completed — Previous Sessions
- [x] Set up GitHub repo and folder structure
- [x] Created CLAUDE.md, tasks/todo.md, tasks/lessons.md
- [x] Created .gitignore at root (commit c31d0bb)
- [x] Created CompletePackageStorage/production/ and CompletePackageStorage/backups/
- [x] Repo structure audit (2026-02-26)
- [x] Full comprehensive audit of all code, docs, and NewTesting files (2026-02-27)
- [x] Redesigned P&L Model Excel to iPipeline Fortune 100 standard
- [x] Fixed reconciliation check failures (Checks 5-9, 12)
- [x] Built Executive Dashboard on Report sheet
- [x] Created Charts & Visuals with 8 interactive charts + dropdown selector
- [x] Redesigned Charts & Visuals to Fortune 100 dashboard layout
- [x] Fixed SQL bug: fact_gl_transactions → fact_gl in pnl_enhancements.sql
- [x] Built modLogger_v2.1.bas — VBA runtime audit log
- [x] Built modUtilities_v2.1.bas — 12 utility macros (actions 51-62)
- [x] Updated frmCommandCenter_code.txt — 62 actions, Sheet Tools category
- [x] Fixed revenue share mismatch: SQL synced to Python values
- [x] Built pnl_monte_carlo.py — Monte Carlo P&L risk simulation
- [x] Wired monte-carlo into pnl_cli.py
- [x] Rewrote README.md professionally
- [x] Updated CODE_COMPARISON_REPORT.md scorecard
