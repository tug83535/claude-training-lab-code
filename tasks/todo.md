# Project Todo — APCLDmerge (iPipeline P&L Demo)

## Current Status (2026-03-03)
- **Branch:** `claude/resume-apclmerge-project-CXWP5` (active working branch)
- **VBA Modules:** 32 imported into Excel workbook — Debug > Compile passes (greyed out = clean)
- **Python Scripts:** 14 complete and functional (main project)
- **Universal Tools:** 76 tools built, code-reviewed, bugs fixed, how-to guide written
- **Excel File:** Workbook open with all 32 modules imported and error-free
- **Testing Phase:** T1 in progress — T1.01 through T1.07 complete; T1.08 not yet run (Connor doing tomorrow)
- **Overall:** Track B complete, Track C complete, Backlog Item 1 complete

### ⚠ ACCOUNT SWITCH NEEDED — Usage at ~90%
The current Claude account is near its usage limit. The next Claude account must pick up from here.

**There are TWO active tracks. Pick up whichever the user directs:**

---

**TRACK A — Testing (T1.08 next)**
1. Read this file, tasks/lessons.md, CLAUDE.md, and Testing_Issues/TESTING_ISSUES_LOG.md first
2. Read qa/TEST_PLAN.md for the full test procedure
3. Resume at **T1.08**: run `pip install -r requirements.txt` → verify all packages install cleanly
4. Then proceed through T2, T3, T4 in order
5. Log all results in qa/TEST_RESULTS.md (create it if it doesn't exist yet)
6. Any new issues found → add to Testing_Issues/TESTING_ISSUES_LOG.md

---

**TRACK B — Universal Tools (UniversalBuild)**
Connor uploaded idea files to UniversalToolsForAllFiles/:
- GrokALL.md — Grok-generated list (~70 VBA + ~40 Python universal tools)
- PrelexALL.md — comprehensive catalog of 293 tools across 17 categories
- GemAll.md — Gemini-generated list (~35 VBA + ~30 Python, Finance/audit focus)
- Profiles.md — personal account reference (deleted from repo — was a security risk)

A curated candidate list has been created at:
`UniversalToolsForAllFiles/UniversalBuild/UNIVERSAL_BUILD_CANDIDATES.md`

**Status:** ALL CODE BUILT AND COMMITTED (2026-03-02)
- 24 Tier 1 VBA tools selected + 34 Tier 2 VBA = 58 VBA tools total
- 5 Tier 1 Python scripts + 13 Tier 2 Python = 18 Python scripts total
- 76 total candidates — ALL BUILT as actual code
- GemAll.md reviewed and folded in (2026-03-02) — 16 tools added
- review/PROJECT_OVERVIEW.md created — comprehensive overview for external Claude review
- review/DemoWrapUp/ folder created — ready for Connor's external review document

**What was built (2026-03-02):**
- UniversalToolsForAllFiles/vba/modUTL_DataCleaning.bas — 12 tools (unmerge, fill blanks, text-to-numbers, remove dupes, etc.)
- UniversalToolsForAllFiles/vba/modUTL_Formatting.bas — 9 tools (autofit, freeze rows, number/currency/date formats, etc.)
- UniversalToolsForAllFiles/vba/modUTL_WorkbookMgmt.bas — 15 tools (unhide all, PDF export, search, rename sheets, etc.)
- UniversalToolsForAllFiles/vba/modUTL_Finance.bas — 14 tools (duplicate invoice, GL validator, aging reports, variance, etc.)
- UniversalToolsForAllFiles/vba/modUTL_Audit.bas — 8 tools (external links, circular refs, error scanner, data quality, etc.)
- UniversalToolsForAllFiles/python/ — 18 scripts + requirements.txt (clean_data, compare_files, aging_report, bank_reconciler, variance_decomposition, fuzzy_lookup, and more)

**Next steps for Universal Tools:**
1. **BUG REVIEW — Code verify all 5 VBA modules and 18 Python scripts before Connor uses them** — see Track C below
2. Write coworker how-to/usage guide for all Universal Tools (VBA + Python) — see Backlog
3. Convert Python scripts to .exe files so coworkers can just click and run — see Backlog

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
- [ ] Final PR to merge `claude/resume-apclmerge-project-CXWP5` → `main`

---

## Backlog (Nice to Have — After Demo)
- [ ] Dynamic Progress Bar KPI Shape (visual KPI indicator on dashboard)
- [ ] Financial Statement Generator from Trial Balance (requires account mapping design)
- [ ] VBA Outlook Email Integration (PDF → Email in one workflow)
- [ ] Build remaining monthly summary tabs (Apr-Dec) using modMonthlyTabGenerator
- [ ] **Scenario 2 — Universal Tools Add-In:** Package the 8 universal tools (Data Sanitizer, Find Negatives/Zeros/Round Numbers, Find External Links, Audit Hidden Sheets, Cross-Sheet Search) into `KBT_UniversalTools.xlam` so coworkers can use them on their own files. Source files staged in `UniversalToolsForAllFiles/`. Write coworker install guide when ready.
- [x] **Universal Tools — Coworker How-To Guide:** COMPLETE (2026-03-03) — Full guide at `UniversalToolsForAllFiles/UNIVERSAL_TOOLS_HOW_TO_GUIDE.md`. Covers all 76 tools with installation, step-by-step usage, examples, and quick reference table. Written for non-technical Finance & Accounting staff.
- [ ] **Universal Tools — Python .exe Conversion:** Convert all 18 Python scripts to standalone `.exe` files using PyInstaller (or similar) so coworkers can just double-click and run — no Python installation required. Package with a simple folder + README.

---

## Dropped by User (Do Not Build)
- ~~Backup Workbook with Timestamp macro~~ — user declined (2026-02-28)
- ~~VBA Timestamp Audit Trail on Cell Changes~~ — user declined (2026-02-28)
- ~~Export All Charts to PowerPoint~~ — user dropped permanently (2026-02-28)

---

## Completed — This Session (2026-03-03)

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

### T1 Testing — T1.01 through T1.07
- [x] T1.01 — PASS — VBA project compiled with zero errors
- [x] T1.02 — PASS — All 32 modules visible in Project Explorer
- [x] T1.03 — PASS — Option Explicit found in all modules
- [x] T1.04 — PASS — ?APP_VERSION returned "2.1.0"
- [x] T1.05 — PASS — pnl_config.py printed full config summary, all shares sum to 1.00. Source file warning is cosmetic (file was renamed), not a bug.
- [x] T1.06 — PASS — All 14 Python scripts imported successfully, printed "All imports OK"
- [x] T1.07 — INVESTIGATED AND RESOLVED — Original scan flagged all 14 files for non-ASCII bytes. Investigation confirmed: all 14 files are valid UTF-8. Characters are intentional Unicode (em dashes, arrows, check marks, box-drawing, Greek letters, emoji) — NOT mojibake. TEST_PLAN.md updated with full explanation. This test is PASS.
- [x] TEST_PLAN.md updated — T1.07 and T4.01 pass criteria clarified
- [x] Testing_Issues/TESTING_ISSUES_LOG.md created — full log of all T1 issues found

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
