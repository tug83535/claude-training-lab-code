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
- All VBA code complete — 24 modules, all 62 Command Center actions covered (2026-02-28)
- All Python scripts complete — 14 scripts, all functional (2026-02-28)
- Branch: `claude/review-branch-progress-pP7Qf` (unified — all 3 accounts merged)
- Next phase: Import .bas files into Excel workbook, live test, then demo prep

## Current Session Summary — 2026-02-28

### What Was Done This Session
Major session: merged all 3 Claude accounts' branches, audited everything, built all missing
VBA modules, and fixed all known bugs. The codebase is now complete.

**Branch Merge:**
- Discovered 5 branches across 3 Claude accounts via `git fetch --all`
- Merged Track A (Excel redesign) and Track B (code improvements) into unified branch
- Resolved merge conflict in tasks/todo.md — combined content from both tracks

**Full Audit:**
- Audited all 24 VBA modules — categorized as working/broken/unbuilt
- Audited all 14 Python scripts — all confirmed complete and functional
- Produced full inventory with status for every module

**10 New VBA Modules Built:**
- modSensitivity (Action 5) — sensitivity analysis on Assumptions drivers
- modAWSRecompute (Action 14) — AWS allocation validation and recalculation
- modImport (Action 17) — CSV/Excel data import pipeline
- modForecast (Actions 18-19) — rolling forecast + trend append
- modScenario (Actions 20-23) — scenario save/load/compare/delete
- modAllocation (Actions 24-25) — cost allocation engine + preview
- modConsolidation (Actions 26-30) — multi-entity consolidation + IC eliminations
- modVersionControl (Actions 31-35) — version save/compare/restore
- modAdmin (Actions 36-40) — auto-documentation + change management
- modIntegrationTest (Actions 44-45) — 18-test suite + quick health check

**Bug Fixes (5):**
1. modLogger: Added ViewLog procedure (Action 41 target was missing)
2. modNavigation: Fixed Ctrl+Shift+R shortcut + added ToggleExecutiveMode (Action 48)
3. modConfig: Added RECON_TOLERANCE constant (used but not defined)
4. modReconciliation: Fixed StyleHeader call (4 args → 3 args)
5. modFormBuilder: Fixed install guide text "50 actions" → "62 actions"

### VBA System Summary (v2.1.0 — 62 actions, 24 modules)
**Original 14 modules (from previous sessions):**
- modConfig: All constants (sheet names, products, fiscal year, colors, thresholds)
- modFormBuilder: Command Center builder + ExecuteAction() routing table (all 62 actions)
- modMasterMenu: InputBox fallback (4 pages, 62 items)
- modNavigation: TOC, GoHome, keyboard shortcuts, ToggleExecutiveMode
- modDashboard: Charts, Executive Dashboard, Waterfall, Product Comparison
- modDataQuality: 6 scans + FixTextNumbers (BUG-018 safe cell-tracking fix)
- modReconciliation: Checks sheet PASS/FAIL + 4 cross-sheet validations
- modVarianceAnalysis: MoM variance + auto-commentary (15% threshold)
- modPDFExport: Batch 7-sheet PDF export + professional print headers/footers
- modPerformance: TurboOn/TurboOff + timer + status bar
- modMonthlyTabGenerator: Clone Mar template for Apr-Dec, next-month-only option
- modSearch: Cross-sheet search, 200-result cap, yellow highlight
- modUtilities: 12 utility macros (actions 51-62)
- modLogger: Runtime audit log to hidden VBA_AuditLog sheet + ViewLog

**10 new modules (built this session):**
- modSensitivity, modAWSRecompute, modImport, modForecast, modScenario
- modAllocation, modConsolidation, modVersionControl, modAdmin, modIntegrationTest

### Excel File
- KeystoneBenefitTech_PL_Model.xlsx — iPipeline Fortune 100 redesigned version
- 13 sheets + Charts & Visuals (from Track A merge)
- Binary file — cannot be read directly by Claude
- The .bas files exist in the repo but have NOT been imported into the Excel workbook yet

### Next Steps for Next Session (Priority Order)
1. Import all 24 .bas files into the Excel workbook (Alt+F11 → File → Import)
2. Create frmCommandCenter UserForm in the workbook
3. Live test every Command Center action (1-62) — log pass/fail
4. Fix any runtime issues found during testing
5. Write demo video storyboard/script
6. Build coworker training guide (step-by-step for non-technical Finance staff)
7. Copy final files to CompletePackageStorage/production/
8. Record demo video

## Permanent Notes
- User confirmed they do NOT want: Backup Workbook macro, Timestamp Audit Trail, Export Charts to PowerPoint — do not propose these
- Excel file is always binary — Claude cannot read it; use ARCHITECTURE_DIAGRAM.md and modConfig constants as reference
- User is new to GitHub; explain PR/merge process in plain English
- .bas files in repo are source code only — must be manually imported into Excel to work
