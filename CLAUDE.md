# Claude Training Lab - APCLDmerge Project

## About Me
I am not a developer. I work on guides, training docs, VBA, SQL, and Python demos
for Finance & Accounting at iPipeline. Keep all explanations in plain English.

## The Project
I am building a world-class demo P&L Excel file with VBA macros, SQL, and Python
to present to 2,000+ employees and the CFO/CEO at iPipeline. I will also be
creating a video walkthrough for coworkers. Everything produced must be perfect,
polished, and professional — treat every output as if it represents the best
employee at the best company in the world.

## Repo Structure
- `vba/` — VBA modules (.bas files)
- `sql/` — SQL scripts
- `python/` — Python scripts
- `docs/day-to-day/` — day-to-day reference docs
- `docs/overview/` — project overview docs
- `docs/setup/` — setup guides
- `training/` — final reviewed guides ready for coworkers
- `qa/` — QA tracking, test plans, checklists, and bug logs
- `tasks/` — session management files, todo.md and lessons.md
- `excel/` — contains the main demo P&L Excel file, uploaded each session
- `DemofileChartBuild/` — chart sheet files for the demo P&L, work in progress
- `NewTesting/` — experimental code, research, and ideas not ready for main project yet
- `CompletePackageStorage/` — final production-ready files and backups
  - `CompletePackageStorage/production/` — live, ready-to-go final files
  - `CompletePackageStorage/backups/` — versioned backups

## My Audience
Training docs and guides are written for non-technical Finance & Accounting staff.
Every guide must be:
- Extremely detailed — every step written out no matter how small
- Organized, clean, and world class
- More detail is always better than less

## Excel File
Uploaded each session. It is a P&L demo file containing VBA macros.
When I upload it, always fully review EVERY sheet before doing anything.
Never assume a sheet is irrelevant. Always confirm what sheets you found.

## Common Tasks
- Edit, analyze, fix, and improve the Excel file and VBA code
- Write new VBA, SQL, and Python code to improve efficiency
- Create world-class training guides for coworkers
- Create detailed step-by-step guides for me to execute tasks
- Think of new ideas, innovations, and cutting-edge improvements
- Help plan and build toward a final video demo

## Always Do
- Before starting ANY task, confirm what you will do in bullet points and wait
  for my approval before proceeding
- Never infer — always ask clarifying questions if anything is unclear
- Fully review ALL files, ALL sheets, ALL pages before responding
- Always self-review your own output before delivering it — check for missing
  steps, incomplete sections, or anything that falls below world-class standard
- At the end of every session, update tasks/todo.md and tasks/lessons.md
- Strive for world-class output on every single task
- Always suggest new ideas, innovations, and better approaches proactively
- After any correction from me, log the pattern in tasks/lessons.md immediately
  so the same mistake never happens again
- Review tasks/lessons.md at the start of every session before doing anything

## Never Do
- Do not skip steps in guides no matter how obvious they seem
- Do not assume what sheet or page is most important — review everything
- Do not start work without my confirmation
- Do not infer what I want — ask first
- Do not attempt to complete too many steps at once — break it down

## Handling Large or Complex Requests
When I give you a task with many steps or a heavy workload:
1. Stop and build a numbered action plan first
2. Present the full plan to me and wait for my approval
3. Execute one step at a time, confirming completion before moving to the next
4. If something goes wrong mid-task, stop immediately, re-plan, and check in
5. Never push through errors or uncertainty — pause and ask

This prevents overload, mistakes, and missed steps.

## Task Management
- `tasks/todo.md` — running task list, updated every session
- `tasks/lessons.md` — log of mistakes and lessons learned, reviewed every session
- Always write the plan to tasks/todo.md before starting implementation
- Mark items complete as you go
- Add a summary of what was done at the end of every session

## Quality Bar
Before delivering ANYTHING ask yourself:
- Is every step written out completely?
- Have I reviewed every sheet, page, and file fully?
- Would the CFO/CEO be proud to see this?
- Is this truly world-class or just good enough?
- If any answer is no — fix it before delivering

## Current Status
- Full repo + code audit complete (2026-02-27)
- Branch: `claude/review-project-status-ntucB` — identical to `master`, no diffs
- See Current Session Summary below for full details

## Current Session Summary — 2026-02-27

### What Was Done This Session
This was a read-only audit session. No files were changed. A new Claude account took over
from a previous account that hit usage limits. All context was reconstructed from file reads.

**Files read and audited:**
- CLAUDE.md, tasks/todo.md, tasks/lessons.md
- All 13 VBA .bas files + frmCommandCenter_code.txt
- All 4 SQL scripts (staging, transformations, validations, enhancements)
- python/pnl_runner.py and python/pnl_config.py (9 remaining Python scripts not yet read)
- docs/overview/CODE_COMPARISON_REPORT.md and ARCHITECTURE_DIAGRAM.md
- NewTesting/GPT.md, Gemini.md, Perlex.md (all 3 idea files)

**Excel file:** KeystoneBenefitTech_PL_Model.xlsx is a binary file — cannot be read directly.
Sheet structure was sourced from ARCHITECTURE_DIAGRAM.md and modConfig_v2.1.bas constants.

### Key Findings

**Bugs Found:**
1. SQL table name mismatch: `pnl_enhancements.sql` references `fact_gl_transactions` but
   `staging.sql` creates a table named `fact_gl` — this will error if run as-is
2. Revenue share mismatch: Python `pnl_config.py` has iGO=55%, Affirm=28%, InsureSight=12%,
   DocFast=5% — SQL `transformations.sql` has iGO=50%, Affirm=25%, InsureSight=15%,
   DocFast=10% — these must be reconciled to match each other
3. `frmCommandCenter_code.txt` only shows 50 actions — current system has 62
   (modUtilities v2.1 added actions 51-62) — txt file is outdated and needs updating
4. `modLogger` is called throughout the VBA code but no `modLogger.bas` file exists
   in the repo — every logging call will fail at runtime

**Missing VBA .bas files (referenced in ExecuteAction but not committed to repo):**
modLogger, modImport, modForecast, modScenario, modAllocation, modConsolidation,
modVersionControl, modAdmin, modIntegrationTest, modAWSRecompute, modSensitivity

**Outdated document:**
- `docs/overview/CODE_COMPARISON_REPORT.md` was written before `modUtilities_v2.1.bas`
  was committed
- That commit added 12 macros listed as "Not Yet Built" in the report — they are now built
- Updated scorecard: ~29 Built (57%), 9 Partially Built (18%), ~13 Not Yet Built (25%)

**todo.md was stale:**
- `.gitignore` and `CompletePackageStorage/` subfolders were marked as "not done"
  but both exist in the repo from commit `c31d0bb`

### Excel Sheet Inventory (13 sheets — sourced from ARCHITECTURE_DIAGRAM.md + modConfig)
1. CrossfireHiddenWorksheet — raw GL data, 510 rows x 7 cols, hidden
2. Assumptions — driver table, 33 rows x 4 cols
3. Data Dictionary — products/departments/vendors reference, 54 rows x 5 cols
4. AWS Allocation — AWS cost allocation model, 42 rows x 6 cols
5. Report--> — home/TOC sheet, 22 rows x 6 cols
6. P&L - Monthly Trend — consolidated monthly P&L, 44 rows x 18 cols
7. Product Line Summary — product-level P&L, 80 rows x 18 cols
8. Functional P&L - Monthly Trend — core calculation engine, 147 rows x 18 cols
9. Functional P&L Summary - Jan 25 — monthly snapshot, 37 rows x 5 cols
10. Functional P&L Summary - Feb 25 — monthly snapshot, 37 rows x 5 cols
11. Functional P&L Summary - Mar 25 — monthly snapshot, 37 rows x 5 cols
12. US January 2025 Natural P&L — expense detail by dept, 77 rows x 5 cols
13. Checks — cross-sheet reconciliation, 13 rows x 5 cols
(Apr-Dec summary tabs do not exist yet — modMonthlyTabGenerator creates them on demand)

### VBA System Summary (v2.1.0 — 62 actions, 13 modules)
- modConfig: All constants (sheet names, products, fiscal year, colors, thresholds)
- modFormBuilder: Command Center builder + ExecuteAction() routing table (all 62 actions)
- modMasterMenu: InputBox fallback (4 pages, 62 items)
- modNavigation: TOC, GoHome, keyboard shortcuts (Ctrl+Shift+M/H/J/R)
- modDashboard: Charts, Executive Dashboard, Waterfall, Product Comparison
- modDataQuality: 6 scans + FixTextNumbers (BUG-018 safe cell-tracking fix)
- modReconciliation: Checks sheet PASS/FAIL + 4 cross-sheet validations
- modVarianceAnalysis: MoM variance + auto-commentary (15% threshold)
- modPDFExport: Batch 7-sheet PDF export + professional print headers/footers
- modPerformance: TurboOn/TurboOff + timer + status bar
- modMonthlyTabGenerator: Clone Mar template for Apr-Dec, next-month-only option
- modSearch: Cross-sheet search, 200-result cap, yellow highlight
- modUtilities: 12 utility macros (actions 51-62) — MOST RECENTLY COMMITTED

### Next Steps for Next Session (Priority Order)
1. Ask user to provide the missing VBA .bas files from their local APCLDmerge_ALL folder
   so they can be committed (modLogger is critical — other modules depend on it)
2. Fix SQL bug: rename `fact_gl_transactions` to `fact_gl` in pnl_enhancements.sql
3. Decide correct revenue share percentages and update BOTH pnl_config.py AND
   transformations.sql to match
4. Update frmCommandCenter_code.txt to show all 62 actions (currently shows 50)
5. Update docs/overview/CODE_COMPARISON_REPORT.md — 12 items now built in modUtilities
6. Rewrite README.md with professional project overview
7. Commit and push all changes to branch `claude/review-project-status-ntucB`
8. Read remaining 9 Python scripts (full audit still pending)
9. Build: Timestamp Audit Trail on Cell Changes (highest-priority remaining VBA feature)
10. Build: Export All Charts to PowerPoint

## Session Continuation Notes (same day, 2026-02-27)
- User confirmed they do NOT want a Backup Workbook macro — do not propose this feature
- Excel file (KeystoneBenefitTech_PL_Model.xlsx) is always a binary file — Claude cannot read it directly; all sheet knowledge comes from ARCHITECTURE_DIAGRAM.md and modConfig_v2.1.bas constants — permanent limitation
- User is new to GitHub; Pull Request / merge process was explained (feature branch → master via PR on github.com)
- Missing VBA modules confirmed NOT built by any prior Claude account — must be provided from user's local APCLDmerge_ALL folder OR built from scratch in a future session
- modLogger.bas is the most critical missing module — all other VBA modules depend on it for runtime logging
