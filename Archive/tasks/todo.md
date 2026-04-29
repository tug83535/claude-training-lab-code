# Project Todo — APCLDmerge (iPipeline P&L Demo)

## Codex Cherry-Pick Campaign (2026-04-20 — active)

Porting select ideas from the parallel Codex build into Project A. Full tracker: `C:\Users\connor.atlee\RecTrial\CodexCompare\CHERRY_PICK_TRACKER.md`. Comparison report: `C:\Users\connor.atlee\RecTrial\CodexCompare\COMPARISON_REPORT.md`. Project A stays structurally as-is — no refactors, only additions.

**Status as of 2026-04-21:** Batches 1-3 complete and live in the .xlsm. Video 3 shipped. Batches 4-5 deferred until after Video 4 recording.

### Batch 1 — Small VBA wins ✅ LIVE
- [x] **K — MarginVerdict + AppendMarginVerdictRow** — `modWhatIf_v2.1.bas`
- [x] **B — CreateRunReceiptSheet** — `modUTL_Audit.bas`
- [x] **C — UTL_DetectHeaderRow** — `modUTL_Core.bas`

### Batch 2 — Capability ✅ LIVE
- [x] **A — modUTL_Intelligence.bas** — MaterialityClassifier + ExceptionNarratives + DataQualityScorecard. Registered in Command Center as static category "Intelligence" at position 6.
- [x] **D — UTL_QuickRowCompareCount + BuildRowHashMap** — `modUTL_Compare.bas`
- [x] Narrative-on-wrong-column bug fix (commit 8eff337) — `FindColumnByHeaderText` iterates candidates outer / columns inner so specific candidates beat generic ones

### Batch 3 — Zero-install Python ✅ PORTED
- [x] **E1-E4** — `profile_workbook.py`, `sanitize_dataset.py`, `compare_workbooks.py`, `build_exec_summary.py`
- [x] **F, G, H** — `variance_classifier.py`, `scenario_runner.py`, `sheets_to_csv.py`
- [x] **I/J — `--talking-points` opt-in flag in `word_report.py`**
- [ ] Optional: spot-test the 7 ZeroInstall scripts on a real file

### Batch 4 — Dual-logging pattern — DEFERRED until after Video 4
- [ ] **L — Dual-logging** (local `VBA_AuditLog` + universal logger in demo modules)

### Batch 5 — Docs — DEFERRED until after Video 4
- [ ] **9 — `CONSTRAINTS.md`** (repo root, next to `CLAUDE.md`)
- [ ] **10 — `BRAND.md`** (repo root)
- [ ] **4 — `RELEASE_READINESS_CHECKLIST.md`** (`RecTrial\Guide\`)
- [ ] **5 — `TROUBLESHOOTING.md`** (`RecTrial\Guide\`)

## Video 4 Replanning — LOCKED 2026-04-28

Original V4 (10 CMD-run Python scripts + ElevenLabs audio) was pulled 2026-04-22. Split 4a+4b plan also pulled after 5th-pass external review (2026-04-27). All 5 decisions locked by Connor 2026-04-28. 5-doc planning sprint complete 2026-04-28.

**Locked direction:** Single chaptered V4 ("Python Automation for Finance," 9–12 min), Revenue Leakage Finder as narrative hero + ARR waterfall closing artifact, `finance_automation_launcher.py` deliverable, 5–7 supported workflow doorway, adoption-grade framing (coworkers use on own files), 50–150 coworker audience near-term.

**Planning docs (source of truth — all complete 2026-04-28):**
- [x] `RecTrial\Brainstorm\VIDEO_4_REVIEW_DECISION_MEMO.md` — 5 locks, what got cut, stale-reference table
- [x] `RecTrial\Brainstorm\SUPPORTED_WORKFLOWS_V1.md` — 7 starter workflows mapped to modules + adoption guidance
- [x] `RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md` — 8-chapter outline + sample data design lock
- [x] `RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md` — 14 safety rules + adoption guidance
- [x] `RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md` — SharePoint zip, pilot plan, support intake, release gate

### Research foundation (complete — no further review needed)
- [x] 14 raw research files read + 6 AI-compiled synthesis docs produced (`RecTrial\Brainstorm\NewCodeResearch\`)
- [x] ~156 ideas inventoried, ~40-60 curated per doc
- [x] Claude Code subagent back-check of raw files — confirmed no new findings beyond synthesis
- [x] 5th-pass external review (`RecTrial\Brainstorm\NewCodeResearchExtra\`) — 5 pivot recommendations, all evaluated and locked

### V4 build list — Claude Code builds (all in `RecTrial\UniversalToolkit\python\ZeroInstall\`)

- [x] `common/safe_io.py` — DONE 2026-04-28
- [x] `common/logging_utils.py` — DONE 2026-04-28
- [x] `common/report_utils.py` — DONE 2026-04-28
- [x] `common/sample_data.py` + `samples/contracts_sample.csv` + `samples/billing_sample.csv` — DONE 2026-04-28 (123 contracts, 336 billing rows, 6 embedded exception classes)
- [x] `data_contract_checker.py` — DONE + tested (PASS sample mode, 1 failure found as expected)
- [x] `revenue_leakage_finder.py` — DONE + tested (12/10/2/9/5 exceptions across 5 classes)
- [x] `exception_triage_engine.py` — DONE + tested (38 exceptions scored, top-10 action list)
- [x] `control_evidence_pack.py` — DONE + tested (5 files hashed, manifest + evidence HTML)
- [x] `workbook_dependency_scanner.py` — DONE + tested (cross-sheet ref map, stdlib zipfile+xml)
- [x] `finance_automation_launcher.py` — DONE (numbered menu, 14 safety rules, Explorer shortcut)
- [x] `smoke_test_video4_python.py` — DONE (5/5 PASS)
- [x] `README_VIDEO4_PYTHON.md` — DONE
- [x] **Git commit** — all V4 Python files + narration script v1.1 + shot list committed 2026-04-28 (commits d9b7ea0, d9f2ad0, 48d133f)
- [x] **Update CLAUDE.md** — mark V4 Python build complete, update status section — DONE 2026-04-28
- [ ] `FinanceTools.xlsm` — Excel workbook with ONE VBA Shell() launcher button (Option A LOCKED 2026-04-28): single button → `finance_automation_launcher.py` → numbered CLI menu → coworker picks tool. No per-tool buttons in V1.
  - [ ] **Connor reviews `modFinanceToolsLauncher.bas`** — VBA Shell() launcher code drafted 2026-04-28 at `RecTrial\UniversalToolkit\python\ZeroInstall\modFinanceToolsLauncher.bas`. Review and provide feedback before importing into FinanceTools.xlsm.
- [ ] Test scripts with bundled Python 3.11 embeddable (confirm zero-install path works)
- [x] Write narration script for V4 (8 chapters, 9–12 min) — DONE 2026-04-28 (v1.1 at `RecTrial\Video4_V1\VIDEO_4_NARRATION_SCRIPT_v1.md` + `RecTrial\Brainstorm\VIDEO_4_NARRATION_SCRIPT.md`)
- [x] Write shot list / screen recording guide per chapter — DONE 2026-04-28 (`RecTrial\Video4_V1\VIDEO_4_SHOT_LIST_v1.md`)
- [ ] Connor reviews narration script + adjusts wording to match natural speaking style
- [ ] Connor generates ElevenLabs audio clips from narration script (9 clips: V4_C01 through V4_C08, with C03a + C03b)
- [ ] Confirm coworker pip access (Connor real-world task — affects stdlib-only requirement)
- [ ] Identify 10–20 pilot users (Connor real-world task per MINIMUM_DISTRIBUTION_PLAN.md)
- [ ] Assemble SharePoint zip distribution package
- [ ] Record + edit V4

### Future Automation Ideas doc (tracked separately)
Full parking lot at `RecTrial\Brainstorm\FUTURE_AUTOMATION_IDEAS.md`. Major parked categories: AI API ideas, Outlook automation, Task Scheduler, warehouse SQL, ML libs, infrastructure, 3rd-party platforms.

---

## Video 5 — "Getting Started" / Python Setup (Planned — post-V4)

Short supplemental video (3–5 min) covering the one-time setup experience for coworkers who want to use the Python tools after watching Video 4. Kept separate from Video 4 so V4 stays focused on the revenue leakage story without slowing down for logistics.

**Status:** Planned. Content, length, and recording approach to be discussed with Connor after Video 4 is locked in.

**Likely content (working idea — to be confirmed):**
- [ ] Download the zip from SharePoint
- [ ] Unzip it — show the folder structure
- [ ] Open `FinanceTools.xlsm` in Excel
- [ ] Click the Finance Tools button for the first time
- [ ] Run in sample mode to confirm everything works
- [ ] How to point a tool at your own file (brief example)
- [ ] Where your results go (outputs folder)
- [ ] Who to contact if something goes wrong (Connor)

**Note:** Since bundled Python 3.11 embeddable ships in the zip, there is NO Python installation step for coworkers. The "setup" video is really just "download, unzip, open Excel, click." Keep it under 5 min.

**Decisions to make when we get there:**
- [ ] Will it use ElevenLabs AI narration (like V1–V4) or a different format?
- [ ] Will it be a separate SharePoint video alongside V4, or linked from V4's description?
- [ ] Does it need a title card to match V1–V4 style?

## Post-Recording Fixes (2026-03-31) — Fix Before Sharing With Coworkers

### Bugs to Fix
- [ ] **RestoreBaseline (Action 65) errors out** — "Restore error:" with blank message. Coworkers using Action 65 manually will hit this error. Investigate SaveBaseline in modWhatIf.
- [ ] **BuildDashboard charts empty (Action 12)** — Bar/line charts on Report--> show axes but no data. Pie chart works. Root cause: chart data series reference formulas replaced with values.
- [ ] **Executive Dashboard waterfall chart empty** — Empty "Plot Area" box at bottom of Executive Dashboard. Low priority — KPI cards and table are the main visual.

### Fixed (2026-03-31)
- [x] **Executive Dashboard enhanced** — Added product breakdown, monthly trend, MoM growth, status indicators
- [x] **Executive Brief colored headers** — Added colored section header bars (green, orange, blue, purple, teal)
- [x] **Reconciliation Checks column A coloring** — PASS rows now green, FAIL rows now red in column A
- [x] **Sensitivity Analysis $0 values** — Fixed impact calculation to scale against baseline revenue
- [x] **Report--> scroll overshoots** — All scrolls capped at 2 steps
- [x] **Audio clipping** — WaitForAudioEnd replaces fragile timing math
- [x] **Checks sheet cleared by cleanup** — CleanupAllOutputSheets no longer wipes Checks data
- [x] **GL sheet name mismatch** — Director uses "CrossfireHiddenWorksheet" instead of "General Ledger"
- [x] **Executive Mode hides sheets permanently** — ForceUnhideAllSheets runs after toggle and during cleanup
- [x] **PDF Export page breaks** — Director resets view to Normal after PDF export
- [x] **SendKeys for all completion MsgBoxes** — All macros auto-dismiss their dialogs during recording
- [x] **Clip 9 sheet tour misaligned with script** — Timing and sheet order now matches narration
- [x] **Clip 10 Command Center too brief** — CC stays up 30+ seconds with slow category browsing

### ~~Video 4 — Ready to Record (Original CMD-Based Plan — PULLED 2026-04-22)~~
> **ARCHIVED 2026-04-28.** This was the original V4 recording plan (CMD-based, 10 clips, 8 Python scripts). Pulled 2026-04-22 and replaced by the locked V4 direction above. Preserved as historical reference only.
>
> Original demo files at `RecTrial\Video4DemoFiles\`, original audio at `RecTrial\AudioClips\Video4\`, original recording guide at `RecTrial\Guide\VIDEO_4_RECORDING_GUIDE.md`.

### Improvements for Coworker Experience (Post-Recording)
- [ ] **Add YoY Variance Analysis to Command Center** — Not assigned an action number.
- [ ] **Report--> TOC links point to old external file** — Fix to use internal sheet references.
- [ ] **Report--> Add Command Center launch button** — Add clickable button + auto-run AssignShortcuts.
- [ ] **TOC Refresh (Action 13) fails on merged cells** — Handle merged cells on Report--> page.
- [ ] **AssignShortcuts doesn't persist** — Add auto-run via Workbook_Open event.
- [ ] **Dashboard Charts (Action 12) chart data ranges** — Fix source data to match current cell values.

## Current Status (2026-03-12, updated)
- **Branch:** `claude/resume-ipipeline-demo-qKRHn` (active working branch)
- **Demo VBA Modules:** 39 total (34 core + 5 optional add-ins) — need re-import (18 files updated)
- **Python Scripts:** 14 complete and functional (main project) + pnl_forecast.py enhanced with MAPE accuracy
- **Universal Toolkit VBA Modules:** 23 total (14 previous + 9 new: modUTL_ColumnOps, modUTL_Compare, modUTL_Consolidate, modUTL_Highlights, modUTL_PivotTools, modUTL_TabOrganizer, modUTL_Comments, modUTL_ValidationBuilder, modUTL_LookupBuilder + modUTL_WhatIf + modUTL_CommandCenter), code-reviewed, 8 bugs fixed
- **Universal Toolkit Tools:** ~140+ total across 23 VBA modules + Python tools
- **Training Guides:** 6 guides in training/ (final), 8 guides in training/LastGuidesReview/ (pending review)
- **CoPilot Prompt Guide:** v2.0 complete + Quick Start Card + VBA Module Reference List added (2026-03-12)
- **Testing Phase:** T1 COMPLETE, T2 partially done (T2.01–T2.04 PASS, T2.05–T2.07 not yet run), T5.01+T5.02 PASS
- **Bug Reviews:** 5-pass review (2026-03-11), pre-delivery review (2026-03-07), universal toolkit review (2026-03-12) — 8 bugs found and fixed in 7 modules
- **Overall:** Track B COMPLETE, Track C COMPLETE, Backlog Item 1 COMPLETE, Training Guides COMPLETE (draft), ProjectRefresh COMPLETE

### ⚠ IMPORTANT — RE-IMPORT NEEDED
Before continuing testing, re-import these 13 `.bas` files into the Excel workbook (VBA Editor → File → Import):
1. modConfig_v2.1.bas (color constant fixes)
2. modReconciliation_v2.1.bas (dateCol/amtCol constant fixes + LogAction fix + trendLastCol row fix + FindFYCol row fix)
3. modVarianceAnalysis_v2.1.bas (GenerateCommentary row 1 → row 4 fix + YoY Variance)
4. modDashboard_v2.1.bas (split — base module only, charts moved to Advanced)
5. modDashboardAdvanced_v2.1.bas (NEW — ExecDashboard, Waterfall, ProductComp, SmallMultiples)
6. modDemoTools_v2.1.bas (LogAction fix)
7. modTrendReports_v2.1.bas (LogAction fix)
8. modMonthlyTabGenerator_v2.1.bas (LogAction fixes + TestUpdateHeaderText wrapper)
9. modDataQuality_v2.1.bas (Letter Grade + LogAction fix)
10. modPDFExport_v2.1.bas (LogAction fix + dynamic monthly tab discovery)
11. modDataSanitizer_v2.1.bas (rng reset fix in 2 worker functions)
12. modAuditTools_v2.1.bas (rng reset fix in FindExternalLinks)
13. modDrillDown_v2.1.bas (T8.19 fix — GL sheet xlSheetHidden instead of xlSheetVeryHidden)

### Project Refresh — Code Audit / Idea Review — COMPLETE
A separate Claude session independently built VBA and Python code for this same project. We reviewed that code for ideas and improvements.

**Status:** COMPLETE — All steps done. All recommendations implemented as code.

**Steps:**
1. [x] Create `ProjectRefresh/` folder on the branch
2. [x] CODE_AUDIT_FINDINGS.md — 120 tools cross-referenced (34 exact, 30 overlap, 56 new ideas)
3. [x] Tool_Reference_All_120.md — full catalog
4. [x] Code_Audit_Final_Report.md — Demo A-, Universal B+, 20 prioritized actions
5. [x] All Tier 1 recommendations implemented: 3 critical bug fixes, modDashboard split, modUTL_Core, SpecialCells perf, backup-before-destructive, 14 new tools, Data Quality Letter Grade, Forecast Accuracy MAPE, YoY Variance
6. [ ] Connor uploads the other Claude's actual code files into `ProjectRefresh/` (optional — audit is done from docs)

---

### ONE ACTIVE TRACK — Testing (T2.05 next)
1. Read this file, tasks/lessons.md, CLAUDE.md, and qa/TEST_PLAN.md first
2. Re-import the 7 fixed .bas files listed above
3. Resume testing at **T2.05** (FixTextNumbers requires scan), then T2.06, T2.07
4. Then proceed through T3, T4, T5.03–T5.06, T6, T7, T8 in order
5. Log all results in qa/TEST_PLAN.md Section 6 (Test Execution Results)
6. Any new issues found → add to Testing_Issues/TESTING_ISSUES_LOG.md

### Universal Tools — COMPLETE (Track B + Track C + Backlog Item 1 + 2026-03-12 Expansion)
~140+ tools built (23 VBA modules), code-reviewed, all bugs fixed, how-to guide written. 9 new modules added 2026-03-12 (ColumnOps, Compare, Consolidate, Highlights, PivotTools, TabOrganizer, Comments, ValidationBuilder, LookupBuilder). modUTL_WhatIf and modUTL_CommandCenter also added. Bug review agent found 8 bugs across 7 modules — all fixed (commit 63482a4). CoPilot Quick Start Card + VBA Module Reference List added to training/LastGuidesReview/.

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
- [x] Live test every Command Center action (1-62) in Excel — log pass/fail for each
- [x] Fix any runtime issues discovered during testing
- [x] Verify all hidden sheets are created properly (VBA_AuditLog, Scenarios, Version History, etc.)

### Phase 2: Script the Demo Video — COMPLETE (draft in FinalRoughGuides)
- [x] Write demo video storyboard — which features to show, in what order, talking points
- [x] Identify the 14 most impressive actions to highlight (not all 62)
- [x] Plan screen recording flow (open file → Command Center → run features → show results)
- [x] Write speaker notes / narration script (word-for-word in guide)

### Phase 3: Training Materials — COMPLETE (6 guides built, drafts in FinalRoughGuides)
- [x] **Guide 1:** `01-How-to-Use-the-Command-Center.md` — All 62 actions documented, monthly close workflow, troubleshooting, FAQ
- [x] **Guide 2:** `02-Getting-Started-First-Time-Setup.md` — Download, open, enable macros, trust center, verification, first 5 actions
- [x] **Guide 3:** `03-What-This-File-Does-Leadership-Overview.md` — CFO/CEO briefing: business impact, cost analysis, before/after, rollout plan
- [x] **Guide 4:** `04-Quick-Reference-Card.md` — 1-page cheat sheet: all 62 actions, shortcuts, monthly close workflow, troubleshooting
- [x] **Guide 5:** `05-Video-Demo-Script-and-Storyboard.md` — 18-22 min script, 3-part shot lists, word-for-word narration, checklists
- [x] **Guide 6:** `06-Universal-Toolkit-Guide.md` — All 100+ tools (79 VBA + 22 Python), setup, playbooks, top 20
- [x] Connor reviews all 6 guides → approves or requests changes
- [x] Move approved guides from `FinalRoughGuides/` to `training/`
- [x] Convert approved guides to PDF for distribution

### Phase 4: Lock Down the Demo File
- [ ] Save the final tested workbook as `.xlsm` (macros-enabled)
- [ ] Open it fresh on a different machine or clean Excel session — confirm it works out of the box
- [ ] Check that no personal file paths, test data, or debug code is left in the macros
- [ ] Copy final `.xlsm` to `CompletePackageStorage/production/`
- [ ] Copy a dated backup to `CompletePackageStorage/backups/` (e.g., `PL_Model_FINAL_2026-03-10.xlsm`)

### Phase 5: Convert Guides to PDF — COMPLETE
- [x] Convert coworker training guide to PDF (no .md files for coworkers)
- [x] Convert quick-reference card to PDF
- [x] Save PDFs in `training/` folder AND `CompletePackageStorage/production/`

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
- [ ] Final PR to merge `claude/resume-ipipeline-demo-qKRHn` → `main`

---

## Last Optional Adds
- [ ] **Time Saved Calculator** — Calculate how long each of the 62 Command Center actions would take manually vs. running the macro. Build a summary table showing per-action time savings and a grand total (e.g., "Manual: 47 hours/month → Automated: 2 hours/month"). Great talking point for managers and finance leadership.
- [ ] **Splash Screen on Open** — When the file opens, display a branded iPipeline welcome screen (logo colors, version number, "Launch Command Center" button). Auto-dismisses after 5 seconds or on click. Makes the file feel like a product, not a spreadsheet.
- [ ] **Animated Progress Bar for Long Macros** — Replace simple status bar messages with a UserForm progress bar showing percentage complete, estimated time remaining, and the iPipeline logo. Makes 10-second waits feel intentional instead of frozen.
- [ ] **Live "What If" Scenario Demo** — Build a one-click scenario demo: "What if revenue drops 15%?" or "What if AWS costs increase 10%?" — runs the scenario tool and instantly shows the P&L impact. Speaks the language a Finance manager will understand.
- [ ] **Anomaly Detector with Plain English Explanations** — Instead of just flagging data issues, generate plain English insights like: "March rent expense is $47,000 — 3x higher than the 6-month average of $15,200. Possible double-booking." Turns a data check into an AI-level insight.
- [ ] **One-Page Executive Brief Auto-Generator** — One button that scans the entire workbook and generates a plain English summary: "Revenue up 8% MoM. Three expense lines exceeded budget. Cash position strong. Two reconciliation items need attention." Ready to paste into an email or print.
- [ ] **Power BI Connection (Mention Only)** — During the demo, mention that the same data structure could feed a live Power BI dashboard as a Phase 2 initiative. Don't build it — just plant the seed to show forward thinking.
- [ ] **Teams Bot Integration (Mention Only)** — During the demo, mention the concept of typing `/pnl-status` in Teams to get an instant P&L summary. Don't build it — shows vision and where this could go next.
- [ ] **Copilot + Macros Integration (Mention Only)** — Demo how someone could ask Copilot a question and then use YOUR tools to verify or act on the answer. Position the macros as the "trust but verify" layer for AI outputs. Ties into the CoPilot Prompt Guide already built.

---

## Backlog (Nice to Have — After Demo)
- [ ] Dynamic Progress Bar KPI Shape (visual KPI indicator on dashboard)
- [ ] Financial Statement Generator from Trial Balance (requires account mapping design)
- [ ] VBA Outlook Email Integration (PDF → Email in one workflow)
- [ ] Build remaining monthly summary tabs (Apr-Dec) using modMonthlyTabGenerator
- [ ] **Scenario 2 — Universal Tools Add-In:** Package the 8 universal tools (Data Sanitizer, Find Negatives/Zeros/Round Numbers, Find External Links, Audit Hidden Sheets, Cross-Sheet Search) into `KBT_UniversalTools.xlam` so coworkers can use them on their own files. Source files staged in `UniversalToolsForAllFiles/`. Write coworker install guide when ready.
- [x] **Universal Tools — Coworker How-To Guide:** COMPLETE (2026-03-03) — Full guide at `UniversalToolsForAllFiles/UNIVERSAL_TOOLS_HOW_TO_GUIDE.md`. Covers all 76 tools with installation, step-by-step usage, examples, and quick reference table. Written for non-technical Finance & Accounting staff.
- [ ] **Universal Tools — Python .exe Conversion:** Convert all 18 Python scripts to standalone `.exe` files using PyInstaller (or similar) so coworkers can just double-click and run — no Python installation required. Package with a simple folder + README.
- [ ] **Print-Ready Formatting:** Add page setup, print areas, and page breaks to all major report sheets (P&L Trend, Functional P&L, Reconciliation Checks, Variance Analysis, YoY Variance). So any sheet can be printed or PDF'd in one click with professional headers/footers. Do this AFTER Connor reviews the training guides.

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

## Completed — This Session (2026-03-12)

### 9 New Universal Toolkit Modules Built (commit cc45970)
- [x] modUTL_ColumnOps.bas — Column insert/delete/move/split/merge/fill/swap (7 tools)
- [x] modUTL_Compare.bas — Sheet comparison with color-coded diff report (1 tool)
- [x] modUTL_Consolidate.bas — Multi-sheet data consolidation with source tracking (1 tool)
- [x] modUTL_Highlights.bas — Conditional highlighting: threshold, top/bottom N, duplicates (3 tools)
- [x] modUTL_PivotTools.bas — PivotTable creation, refresh all, style, drill-down (4 tools)
- [x] modUTL_TabOrganizer.bas — Sort/color/group/reorder/rename tabs in bulk (6 tools)
- [x] modUTL_Comments.bas — Extract/clear/convert comments/notes (3 tools)
- [x] modUTL_ValidationBuilder.bas — Data validation builder: lists, numbers, dates, custom (5 tools)
- [x] modUTL_LookupBuilder.bas — VLOOKUP/INDEX-MATCH formula builder with preview (2 tools)

### Universal Toolkit Bug Review — 8 Bugs Fixed (commit 63482a4)
- [x] CLR_HDR constant wrong (7930635 → 7948043) in 5 modules: Compare, Consolidate, PivotTools, Comments, ValidationBuilder
- [x] ReorderTabs index shifting bug — resolved indices to sheet names before moving
- [x] Consolidate source column inconsistency — pre-scan all sheets for max width
- [x] Highlights safety cap — added 500K cell limit on HighlightTopBottom and HighlightDuplicateValues

### CoPilot Guide Additions (commit 92b6cd3)
- [x] CoPilot-Quick-Start-Card.md — One-page cheat sheet pointing coworkers to the right prompt
- [x] VBA-Module-Reference-List.md — All 38 modules + 1 UserForm with descriptions, grouped by category

### Earlier Today (before context compaction)
- [x] modUTL_WhatIf.bas — Universal What-If scenario tool (commit 0cc3d0e)
- [x] modUTL_CommandCenter.bas — Universal Command Center for all toolkit tools (commit d063cca)
- [x] What-If Demo actions 63-65 added to demo Command Center (commit 2c72df8)
- [x] What-If guides for demo and universal tools (commits 5c367c8, ec203b7)
- [x] User Training Guide updated to 65 commands (commit 41400cc)
- [x] Operations Runbook updated (commit 41400cc)
- [x] TestingPreVid.md standalone testing guide (commit dfadfa3)
- [x] Demo Excel file renamed to ExcelDemoFile_adv.xlsm (commit 6309434)
- [x] Compile error fix: local lastRow renamed to lRow in AddNamedRanges (commit 351c9b2)
- [x] QA docs updated to reflect 39 modules, 5 add-ins, 35 bugs (commit 0359842)
- [x] Guide review and PDF conversion marked complete (commit 203361c)

---

## Completed — Previous Session (2026-03-11)

### Optional Add-Ins + Universal Expansions + Repo Cleanup
- See CLAUDE.md session summary for 2026-03-11 for full details

---

## Completed — Previous Session (2026-03-07)

### File Uploads (2026-03-06)
- [x] Connor uploaded CoPilot Prompt Guide files (AP_Copilot_PromptGuideHelpV2.docx + .md)
- [x] Removed .gitkeep from CoPilotPromptGuide/

### CoPilot Guide Improvements + Video Package + Sample File (2026-03-07, commit 74dc77e)
- [x] CoPilot Prompt Guide v2.0: Fixed all broken quick reference links (were empty), added working anchor links to every prompt, improved formatting with markdown tables, consistent headings, bold bracket fields
- [x] COMPILED_VIDEO_PACKAGE.md: Fixed tool counts (13 VBA modules/78 tools + 22 Python scripts = ~100 total), added demo file stats (34 VBA modules, 14 Python scripts), updated build checklist, fixed START HERE status, expanded Not Demoed tools list
- [x] Built Sample_Quarterly_Report.xlsx for Video 3 universal tools demo (via build_sample_file.py)

### Pre-Delivery Code Review — 6 Bugs Fixed (2026-03-07, commit 6818b01)
Full code review of all 46 live-Excel test pass criteria against the VBA source code. Found and fixed 6 bugs:
- [x] modReconciliation: LogAction 4th arg elapsed (Double) — instance #13 of this recurring bug
- [x] modReconciliation: ValidateCrossSheet trendLastCol used row 1 instead of HDR_ROW_REPORT — FY column search failed on report sheets
- [x] modReconciliation: FindFYCol scanned row 1 instead of HDR_ROW_REPORT for header matching
- [x] modPDFExport: GetReportSheetList was hardcoded to 7 sheets (Jan-Mar only) — now dynamically discovers all monthly tabs
- [x] modDataSanitizer: rng not reset before SpecialCells in 2 worker functions — could re-process previous sheet's cells
- [x] modAuditTools: rng not reset before SpecialCells in FindExternalLinks — could report duplicate links from wrong sheet

### T8.19 Drill Links Bug Fix (2026-03-07, commit b132885)
- [x] modDrillDown: HideGLSheet set xlSheetVeryHidden which blocked hyperlink navigation — changed to xlSheetHidden so drill links can navigate to GL sheet while keeping it hidden from tab bar

---

## Completed — Previous Session (2026-03-05)

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

### Training Guides — 6 Guides Built (2026-03-05, later session)
- [x] Built `01-How-to-Use-the-Command-Center.md` — ~750 lines, all 62 actions, monthly close workflow, tips, troubleshooting, FAQ
- [x] Built `02-Getting-Started-First-Time-Setup.md` — ~450 lines, download/open/enable macros/trust center/verification/first 5 actions/troubleshooting
- [x] Built `03-What-This-File-Does-Leadership-Overview.md` — ~400 lines, CFO/CEO briefing with business impact, before/after, cost analysis, rollout plan, future roadmap, FAQ
- [x] Built `04-Quick-Reference-Card.md` — ~300 lines, all 62 actions at-a-glance, keyboard shortcuts, monthly close sequence, top 10 actions, troubleshooting
- [x] Built `05-Video-Demo-Script-and-Storyboard.md` — ~500 lines, 3-part 18-22min demo script, shot lists, word-for-word narration, pre/post checklists, B-roll ideas
- [x] Built `06-Universal-Toolkit-Guide.md` — ~650 lines, all 100+ tools (79 VBA + 22 Python), setup instructions, complete reference, use case playbooks, top 20 list
- [x] Committed and pushed to branch (commit c0db49e)

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
