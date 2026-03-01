# Project Todo — APCLDmerge (iPipeline P&L Demo)

## Current Status (2026-02-28)
- **Branch:** `claude/review-branch-progress-pP7Qf` (unified — all 3 accounts merged)
- **VBA Modules:** 24 built (.bas files in repo), all 62 Command Center actions covered
- **Python Scripts:** 14 complete and functional
- **Excel File:** `KeystoneBenefitTech_PL_Model.xlsx` — iPipeline Fortune 100 redesign
- **Overall:** Code is complete. Next phase is import, test, and demo prep.

---

## Next Up — Demo Readiness (Priority Order)

### Phase 1: Make It Real (Import + Live Test)
- [ ] Import all 24 .bas files into the Excel workbook via VBA Editor (Alt+F11 → File → Import)
- [ ] Create the UserForm `frmCommandCenter` in the workbook using `frmCommandCenter_code.txt`
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

### Phase 4: Final Package
- [ ] Copy final tested Excel file to `CompletePackageStorage/production/`
- [ ] Copy training guides to `CompletePackageStorage/production/`
- [ ] Record demo video
- [ ] Final PR to merge `claude/review-branch-progress-pP7Qf` → `main`

---

## Backlog (Nice to Have — After Demo)
- [ ] Dynamic Progress Bar KPI Shape (visual KPI indicator on dashboard)
- [ ] Financial Statement Generator from Trial Balance (requires account mapping design)
- [ ] VBA Outlook Email Integration (PDF → Email in one workflow)
- [ ] Build remaining monthly summary tabs (Apr-Dec) using modMonthlyTabGenerator

---

## Dropped by User (Do Not Build)
- ~~Backup Workbook with Timestamp macro~~ — user declined (2026-02-28)
- ~~VBA Timestamp Audit Trail on Cell Changes~~ — user declined (2026-02-28)
- ~~Export All Charts to PowerPoint~~ — user dropped permanently (2026-02-28)

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
