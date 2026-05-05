# Claude Training Lab - APCLDmerge Project

## ⚡ CURRENT WORK (2026-05-01) — 4-VIDEO DEMO RECORDING PROJECT

The active work right now is recording Video 4. Videos 1-3 are fully done and shipped. Video 4 planning, Python build, narration script, shot list, FinanceTools.xlsm, and the full SharePoint package are all complete and tested. Next: Connor finalizes narration wording, records ElevenLabs audio, title cards for C01+C08, then screen recording.

### Working Folder
**`C:\Users\connor.atlee\RecTrial\`** — self-contained workspace with all audio clips, VBA files, sample Excel files, demo input files, and recording guides. This is NOT in the repo — it's a local working copy. The repo at this path holds the source of truth for commits.

### Video Status
- **Video 1** ("What's Possible") — DONE
- **Video 2** ("Full Demo Walkthrough") — DONE
- **Video 3** ("Universal Tools") — DONE (shipped 2026-04-21, Gemini 70/4 v2.4 result)
- **Video 4** ("Python Automation for Finance") — PACKAGE BUILT + TESTED (2026-05-01). All 5 planning docs done. All 6 V4 Python scripts built and smoke-tested. FinanceTools.xlsm built with VBA launcher button. FinanceTools_v1.0 package assembled and tested — all 4 demo tools pass (options 1–4 confirmed). Next: Connor finalizes narration wording → ElevenLabs audio → title cards → screen recording. Do NOT touch the 8 protected Python scripts until V4 records.

### Video 4 Locked Direction (as of 2026-04-28)

- **Single video** — not split into 4a/4b
- **Hero story:** "Python finds hidden revenue leakage your Excel can't see" — Revenue Leakage Finder as the story, ARR waterfall as the closing visual
- **Public face:** 5–7 workflows shown, not all 140 tools
- **Deliverable:** `finance_automation_launcher.py` CLI menu (renamed from finance_copilot.py)
- **Audience:** ~50–150 coworkers at iPipeline (Finance & Accounting staff)
- **Adoption goal:** Coworkers will actually USE the VBA, Python, and SQL tools on their own work files — not just watch the video
- **xlwings:** Parked for v2 after V4 ships
- **SOX Evidence Collector:** Out of V1 scope
- **8 protected scripts** (do NOT edit until V4 records): `aging_report.py`, `bank_reconciler.py`, `compare_files.py`, `forecast_rollforward.py`, `fuzzy_lookup.py`, `pdf_extractor.py`, `variance_analysis.py`, `variance_decomposition.py`
- **5 planning docs** (COMPLETE 2026-04-28, all in `RecTrial\Brainstorm\`): VIDEO_4_REVIEW_DECISION_MEMO.md, SUPPORTED_WORKFLOWS_V1.md, VIDEO_4_REVISED_PLAN.md, PYTHON_SAFETY.md, MINIMUM_DISTRIBUTION_PLAN.md
- **6 V4 Python scripts** (BUILT + SMOKE-TESTED 5/5 PASS, 2026-04-28, all in `RecTrial\UniversalToolkit\python\ZeroInstall\`): `revenue_leakage_finder.py`, `data_contract_checker.py`, `exception_triage_engine.py`, `control_evidence_pack.py`, `workbook_dependency_scanner.py`, `finance_automation_launcher.py` + `common/` utilities + `smoke_test_video4_python.py` + `README_VIDEO4_PYTHON.md`
- **Post-V4 backlog:** Universal Toolkit 2 (15 high-impact VBA tools) parked in `RecTrial\Brainstorm\FUTURE_AUTOMATION_IDEAS.md` Section 1.10
- **Video 5 (planned):** Short supplemental "Getting Started" video (3–5 min) showing coworkers how to download the zip, open FinanceTools.xlsm, and run their first tool. Content and format TBD after V4 records. Kept separate from V4 so V4 stays focused on the revenue leakage story.

### Key Architecture
**modDirector.bas** (v2.0) is the master VBA "puppeteer" module. It automates recording by:
- Playing AI narration (ElevenLabs MP3s) via Windows mciSendString API
- Navigating between sheets, running macros, scrolling, pausing
- User presses `RunVideo1`/`RunVideo2`/`RunVideo3` and watches hands-free
- Video 4 delivery model: VBA Shell() + bundled Python 3.11 embeddable. `FinanceTools.xlsm` Excel buttons launch `finance_automation_launcher.py`. One zip on SharePoint. Zero install for coworkers. (Excel button design LOCKED 2026-04-28: Option A — one launcher button that opens the CLI numbered menu. No per-tool buttons in V1.)

### The Path A Pattern (CRITICAL — applies to any new dialog-heavy VBA demo)
The original Video 3 approach used `Application.SendKeys` to auto-fill dialog boxes during recording. This was fragile and failed constantly. Gemini's review (8 PASS / 50 FAIL) confirmed. The fix: for each UTL macro that shows dialogs, add a `DirectorXxx` silent wrapper sub at the bottom of the .bas file that takes parameters directly and replicates the core logic with no InputBox/MsgBox. Then the Director clip calls `Application.Run "DirectorXxx", param1, param2, ...`. This is the same pattern used for Video 2 Clips 22 and 23 (SaveCopyAs direct call, RunWhatIfPreset, RestoreBaselineSilent). **Never use SendKeys against modal dialogs — always add a silent wrapper instead.**

### Backup of Pre-Path-A Code
`RecTrial\VBABackup_PrePathA\` — 10 files backed up before Path A refactor, in case we need to revert.

### UTL Files with Director Wrappers Added
- modUTL_DataSanitizer.bas (DirectorRunFullSanitize, DirectorPreviewSanitize)
- modUTL_Highlights.bas (DirectorHighlightThreshold, DirectorHighlightDuplicates, DirectorClearHighlights)
- modUTL_Comments.bas (DirectorExtractComments)
- modUTL_TabOrganizer.bas (DirectorColorTabsByKeyword, DirectorReorderTabs)
- modUTL_ColumnOps.bas (DirectorSplitColumn, DirectorCombineColumns)
- modUTL_SheetTools.bas (DirectorTemplateCloner, DirectorListAllSheetsWithLinks)
- modUTL_Compare.bas (DirectorCompareSheets)
- modUTL_Consolidate.bas (DirectorConsolidateSheets)
- modUTL_CommandCenter.bas (DirectorShowCommandCenter)

### Important Rules for This Work
1. Edit VBA files in BOTH places when needed: `RecTrial\VBAToImport\modDirector.bas` AND `RecTrial\DemoVBA\modDirector.bas` (sync with `cp` command).
2. Any new UTL macro that shows dialogs needs a DirectorXxx wrapper before Video 3 can use it.
3. Sample file for Video 3 is at `RecTrial\SampleFile\SampleFileV2\` — clean backup at `RecTrial\SampleFile\SampleFileBackup_nonMacroClean\`.
4. Feedback from Gemini AI reviews goes in `RecTrial\Feedback\Video3\` (and eventually Video4\).
5. All narration audio clips are in `RecTrial\AudioClips\Video1-4\`.
6. Pivot tables can NOT be created via openpyxl — must be created manually in Excel or via Copilot prompt.

### Files That Should Stay Synced
- `RecTrial\VBAToImport\modDirector.bas` ↔ `RecTrial\DemoVBA\modDirector.bas` (always same)
- `RecTrial\UniversalToolkit\vba\*.bas` — authoritative, edit here
- Memory folder at `C:\Users\connor.atlee\.claude\projects\c--Users-connor-atlee--claude-projects-claude-training-lab-code\memory\` — stays linked to this repo path

---

## About Me
I am not a developer. I work on guides, training docs, VBA, SQL, and Python demos
for Finance & Accounting at iPipeline. Keep all explanations in plain English.

## The Project
I am building a world-class demo + adoption-grade Finance automation package — Excel VBA, Python, and SQL — for Finance & Accounting at iPipeline. The near-term audience is **50–150 coworkers** in Finance, Accounting, and adjacent operations. Broader rollout and any CFO/CEO showcase are deferred, not cancelled. Coworkers are expected to actually adopt the tools on their own files and workflows, not just watch the demo videos. Everything produced must be perfect, polished, and professional — treat every output as if it represents the best employee at the best company in the world.

## iPipeline Brand Styling
- **Official brand guide:** `docs/ipipeline-brand-styling.md`
- All future training guides, documents, presentations, and any visual output MUST use the iPipeline brand colors, fonts, and styling rules defined in that file
- Primary color: iPipeline Blue `#0B4779` | Secondary: Navy `#112E51`, Innovation Blue `#4B9BCB`
- Accents: Lime Green `#BFF18C`, Aqua `#2BCCD3` | Neutrals: Arctic White `#F9F9F9`, Charcoal `#161616`
- Fonts: Arial family only (Arial Bold for headings, Arial Narrow for subheadings, Arial Regular for body)
- Before creating ANY guide, document, or styled output, review `docs/ipipeline-brand-styling.md` first
- Note: VBA modConfig color constants (CLR_NAVY etc.) predate this guide and use slightly different values — do NOT change working VBA code, but any NEW styling work should use the official brand colors

## Repo Structure
- `vba/` — VBA modules (.bas files) for the demo P&L file (39 modules)
- `sql/` — SQL scripts
- `python/` — Python scripts (14 scripts)
- `docs/day-to-day/` — day-to-day reference docs
- `docs/overview/` — project overview docs
- `docs/setup/` — setup guides
- `FinalRoughGuides/` — draft/rough versions of training guides (edit and revise here first)
- `training/` — final polished guides ready for coworkers (move here only after rough guide is fully reviewed and approved)
- `qa/` — QA tracking, test plans, checklists, and bug logs
- `tasks/` — session management files, todo.md and lessons.md
- `DemoVidCode/` — demo file source code (excel/, python/, sql/ grouped together)
- `UniversalToolsForAllFiles/` — future Excel Add-In package for tools that work on any file (23 VBA modules, ~140+ tools)
- `LastCallOptionalAddIns/` — folder for future guides related to the 5 optional add-in modules
- `OldRoughVersions/` — archived folders (includes `_internal/` dev-only folders moved here for cleanup)
- `CoPilotPromptGuide/` — CoPilot Prompt Guide v2.0 files
- `videodraft/` — video package draft and sample demo files

## Sharing Plan
- **Scenario 1 (Primary — Adoption + coworkers):** Distribute as a self-contained zip (`FinanceTools_v1.0.zip`) via SharePoint. Contains `FinanceTools.xlsm` (all VBA modules), bundled Python 3.11 (embeddable, zero install), Python scripts, sample data, and docs. Lead with the 5–7 supported starter workflows; full ~140-tool toolkit stays discoverable inside the workbook's Command Center. Pilot: 10–20 users in Finance/Accounting/Billing before broader rollout. See `RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md`.
- **Scenario 2 (Future — v2 post-V1):** Excel button UI via xlwings or signed `.xlam` add-in for tighter Excel integration. Parked until V1 ships and pilot completes. `UniversalToolsForAllFiles/` contains VBA source. Do this AFTER V1 pilot succeeds.

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
- Would a Finance & Accounting coworker — and eventually the CFO/CEO — be proud to see this?
- Is this truly world-class or just good enough?
- If any answer is no — fix it before delivering

## Current Status
- Original VBA system complete — 24 modules, all 62 Command Center actions covered (2026-02-28)
- All Python scripts complete — 14 scripts, all functional (2026-02-28)
- 7 new VBA modules added from NewTesting ideas (2026-03-01)
- 5 optional add-in modules built (2026-03-11): modTimeSaved, modSplashScreen, modProgressBar, modWhatIf, modExecBrief
- What-If Demo actions 63-65 added to Command Center (2026-03-12)
- **39 demo VBA modules total** (34 previous + 5 optional add-ins) — need re-import (18 files updated)
- **23 universal toolkit VBA modules** (14 previous + 9 new 2026-03-12: ColumnOps, Compare, Consolidate, Highlights, PivotTools, TabOrganizer, Comments, ValidationBuilder, LookupBuilder + WhatIf + CommandCenter)
- **~140+ universal toolkit tools total** across 23 VBA modules + Python tools
- T1 complete (T1.01–T1.08 all PASS), T2 partially tested (T2.01–T2.04 done, T2.05–T2.07 not yet run)
- T5.01 and T5.02 tested and fixed (ExecDashboard + WaterfallChart)
- Self-review of all remaining tests completed — 12 additional bugs found and fixed preemptively
- Python pytest: 99 passed, 15 skipped, 0 failures (T4.04 criteria met)
- ProjectRefresh COMPLETE — audit done, all recommendations implemented as working code
- Demo enhancements: Data Quality Letter Grade, Forecast Accuracy MAPE, YoY Variance, modDashboard split, modUTL_Core, backup-before-destructive, SpecialCells perf fixes
- 6 training guides finalized in training/ + 8 guides pending review in training/LastGuidesReview/
- CoPilot Prompt Guide v2.0 complete + Quick Start Card + VBA Module Reference List (2026-03-12)
- Video package draft (COMPILED_VIDEO_PACKAGE.md) + sample file in videodraft/ (2026-03-07)
- Track B COMPLETE, Track C COMPLETE, Backlog Item 1 COMPLETE, ProjectRefresh COMPLETE, Training Guides COMPLETE (draft)
- `_internal/` moved to `OldRoughVersions/_internal/` for repo cleanup (2026-03-11)
- `LastCallOptionalAddIns/` folder created for future add-in guides (2026-03-11)
- Branch: `claude/resume-ipipeline-demo-qKRHn` (active branch)
- Bug review (2026-03-12): Universal toolkit review — 8 bugs found and fixed across 7 modules (commit 63482a4)
- Bug review (2026-03-11): 5-pass review of all new code — 1 bug found (Chr(9472) crash) + unused constants cleanup
- Bug review (2026-03-07): Pre-delivery code review — 7 bugs found and fixed (6 VBA in commit 6818b01 + T8.19 drill links in b132885)
- Total LogAction signature bugs found to date: 13
- Next phase: Continue Track A testing (T2.05+, then T3–T8), then demo readiness — see tasks/todo.md

### Codex Cherry-Pick Campaign (2026-04-20 — Batches 1-3 complete)
Parallel Codex build at `tug83535/AP_CodexVersion`; local clone at `C:\Users\connor.atlee\RecTrial\CodexCompare\`. Comparison report at `CodexCompare\COMPARISON_REPORT.md`. Tracker at `CodexCompare\CHERRY_PICK_TRACKER.md`. Project A stays structurally as-is — only ADD code/docs, no refactors.

**Batches 1-3 — COMPLETE and live in `Sample_Quarterly_ReportV2.xlsm` (2026-04-21):**
- Batch 1: K (`MarginVerdict` + `AppendMarginVerdictRow` in modWhatIf_v2.1), B (`CreateRunReceiptSheet` in modUTL_Audit), C (`UTL_DetectHeaderRow` in modUTL_Core)
- Batch 2: A (new `modUTL_Intelligence.bas` — MaterialityClassifier + ExceptionNarratives + DataQualityScorecard), D (`UTL_QuickRowCompareCount` + `BuildRowHashMap` in modUTL_Compare)
- Batch 3: 7 stdlib-only Python scripts in `RecTrial\UniversalToolkit\python\ZeroInstall\` + `word_report.py --talking-points` opt-in flag

**Batches 4-5 — DEFERRED until after Video 4 recording:**
- Batch 4: Dual-logging pattern (demo file — careful integration with existing LogAction signature)
- Batch 5: Top-level docs — `CONSTRAINTS.md`, `BRAND.md`, `RELEASE_READINESS_CHECKLIST.md`, `TROUBLESHOOTING.md`

**Conventions used:** Project A header comment block on ported VBA (Codex's had none). iPipeline brand RGBs inline (`modUTL_Branding.bas` docs them at top, no named constants). Path A pattern for dialog-heavy subs. Output sheets prefixed `UTL_`.

**Intelligence module wired into Command Center:** registered as static category #6 "Intelligence (3 tools)" so coworkers see it in the visible list without needing to scroll. Auto-discovery also catches it as "Intelligence (Discovered)" at position 31 — harmless duplicate backup. Narrative-on-wrong-column bug found and fixed via `FindColumnByHeaderText` candidates-outer/columns-inner ordering (commit 8eff337).

**Video 4 protection rule still applies:** do not modify these 8 Python scripts until Video 4 is recorded — `aging_report.py`, `bank_reconciler.py`, `compare_files.py`, `forecast_rollforward.py`, `fuzzy_lookup.py`, `pdf_extractor.py`, `variance_analysis.py`, `variance_decomposition.py`. Anything else is fair game.

### Video 4 replanning — LOCKED 2026-04-28
Original V4 plan (10 ElevenLabs clips, 8 Python scripts run from CMD) was pulled 2026-04-22. Split 4a+4b plan was also pulled after a 5th-pass external review (2026-04-27). All 5 decisions locked by Connor 2026-04-28. 5-doc planning sprint complete.

**Locked V4 direction:**
- **Single chaptered Video 4, 9–12 min** — no 4a+4b split
- **Hero:** Revenue Leakage Finder (narrative: "Python found a possible billing problem") + ARR waterfall as closing visual artifact
- **Deliverable:** `finance_automation_launcher.py` (NOT `finance_copilot.py` — "Copilot" implies AI)
- **6 V4 scripts:** `revenue_leakage_finder.py`, `data_contract_checker.py`, `exception_triage_engine.py`, `control_evidence_pack.py`, `workbook_dependency_scanner.py`, `finance_automation_launcher.py`
- **Distribution:** Excel buttons via VBA Shell() + bundled Python 3.11 embeddable. One zip on SharePoint. Pilot: 10–20 users first.
- **Adoption-grade:** coworkers use on their own real files, not just watch
- xlwings parked as v2 post-V4

**Key planning docs (source of truth — all complete 2026-04-28):**
- `RecTrial\Brainstorm\VIDEO_4_REVIEW_DECISION_MEMO.md` — 5 locks + what got cut + stale-reference table
- `RecTrial\Brainstorm\SUPPORTED_WORKFLOWS_V1.md` — 7 supported starter workflows mapped to existing modules
- `RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md` — 8-chapter outline, sample data design lock, build effort
- `RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md` — 14 safety rules + adoption-grade guidance
- `RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md` — SharePoint zip, pilot plan, support intake, release gate
- `RecTrial\Brainstorm\VIDEO_4_CURRENT_PROPOSAL.md` — SUPERSEDED (historical: 4a+4b + ARR Waterfall hero)
- `RecTrial\Brainstorm\NewCodeResearchExtra\` — 5 files from 5th-pass review (research input)
- `RecTrial\Brainstorm\FUTURE_AUTOMATION_IDEAS.md` — parking lot (AI APIs, Outlook, Task Scheduler, warehouse SQL, ML libs)

Original V4 ElevenLabs audio (`RecTrial\AudioClips\Video4\`) + demo files (`RecTrial\Video4DemoFiles\`) remain untouched on disk.

## Session Summary — 2026-05-01 (Latest — V4 Package Build, Test, and Handoff)

### What Was Done This Session
Built and tested the complete Video 4 distribution package end-to-end. Wrote VBA launcher button code (3 iterations — debugged Dir() OneDrive false-positive and WScript.Shell quoting issues). Assembled FinanceTools_v1.0 package, confirmed all 4 demo tools PASS with bundled Python 3.11. Fixed assembly guide for samples path (must be scripts\samples\, not package root). Wrote two new guides (build FinanceTools.xlsm + assemble SharePoint zip), external review gameplan doc, and session handoff. Updated all tracking files and committed.

**Branch:** `April23CLD` — HEAD: `bc03ecc`

### Key Outcomes
- FinanceTools.xlsm — built with working VBA launcher button
- FinanceTools_v1.0 package — assembled + tested locally at `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\`
- Options 1–4 all confirmed PASS with bundled Python 3.11 embeddable
- Zero-install delivery model confirmed working on Connor's real machine

### Bugs Found and Fixed (2)
1. **Dir() false positives on OneDrive** — OneDrive Files On-Demand makes cloud-only files appear to exist. Fixed by switching to `FSO.FileExists()` in VBA
2. **WScript.Shell quoting failure** — complex paths with spaces failed in CMD. Fixed by setting `wsh.CurrentDirectory = ThisWorkbook.Path` and using relative paths — no quoting needed

### Key Architecture Decision (confirmed working)
- VBA launcher: `WScript.Shell` + `CurrentDirectory = ThisWorkbook.Path` + `cmd.exe /k python\python-embedded\python.exe scripts\finance_automation_launcher.py`
- samples\ must be inside scripts\ — safe_io.py resolves `_TOOLKIT_ROOT` as `parent.parent` of `common\safe_io.py` = the `scripts\` folder

### Files Added / Changed
- `RecTrial/Video4_V1/GUIDE_Build_FinanceTools_xlsm.md` — NEW
- `RecTrial/Video4_V1/GUIDE_Assemble_SharePoint_Zip.md` — fixed Steps 5 + 11 (samples path)
- `RecTrial/Video4_V1/VIDEO_4_GAMEPLAN_FOR_REVIEW.md` — NEW
- `RecTrial/HANDOFF_2026-05-02.md` — NEW
- `RecTrial/UniversalToolkit/python/ZeroInstall/modFinanceToolsLauncher.bas` — NEW
- `Archive/tasks/todo.md` — package tasks marked complete, recording tasks added
- `Archive/tasks/lessons.md` — 4 new V4 package build lessons
- `CLAUDE.md` — current status updated

### What's Next
Connor reads narration script → ElevenLabs audio (9 clips) → title cards → screen record → edit → SharePoint upload → pilot (10–20 users)

---

## Session Summary — 2026-03-12 (Universal Toolkit Expansion + Bug Review + CoPilot Guides)

### What Was Done This Session
Massive universal toolkit expansion: built 9 new VBA modules (38 tools), added modUTL_WhatIf and modUTL_CommandCenter, ran bug review agent that found 8 bugs across 7 modules — all fixed. Built CoPilot Quick Start Card and VBA Module Reference List. Updated demo Command Center with What-If actions 63-65. Multiple training guides and QA docs updated.

**Branch:** `claude/resume-ipipeline-demo-qKRHn`

### 9 New Universal Toolkit Modules (commit cc45970)
1. **modUTL_ColumnOps.bas** — Column insert/delete/move/split/merge/fill/swap (7 tools)
2. **modUTL_Compare.bas** — Sheet comparison with color-coded diff report
3. **modUTL_Consolidate.bas** — Multi-sheet data consolidation with source tracking
4. **modUTL_Highlights.bas** — Conditional highlighting: threshold, top/bottom N, duplicates (3 tools)
5. **modUTL_PivotTools.bas** — PivotTable creation, refresh all, style, drill-down (4 tools)
6. **modUTL_TabOrganizer.bas** — Sort/color/group/reorder/rename tabs in bulk (6 tools)
7. **modUTL_Comments.bas** — Extract/clear/convert comments/notes (3 tools)
8. **modUTL_ValidationBuilder.bas** — Data validation builder: lists, numbers, dates, custom (5 tools)
9. **modUTL_LookupBuilder.bas** — VLOOKUP/INDEX-MATCH formula builder with preview (2 tools)

### Additional Modules Built Earlier in Session
- **modUTL_WhatIf.bas** — Universal What-If scenario tool (commit 0cc3d0e)
- **modUTL_CommandCenter.bas** — Universal Command Center menu for all toolkit tools (commit d063cca)
- What-If Demo actions 63-65 added to demo Command Center (commit 2c72df8)

### Bug Review — 8 Bugs Found and Fixed (commit 63482a4)
Ran automated bug review agent across all 9 new modules. Found 1 CRITICAL, 5 MEDIUM, 2 LOW issues:
1. **CLR_HDR color constant wrong** (7930635 → 7948043) in 5 modules — green channel 3 instead of 71 (near-black vs iPipeline Blue)
2. **ReorderTabs index shifting** — sheet indices change during `.Move` operations; fixed by resolving to names first
3. **Consolidate source column inconsistency** — `srcLastCol + 1` varies per sheet width; fixed with max-width pre-scan
4. **Highlights overflow risk** — `ReDim vals(1 To rng.Cells.Count)` on large ranges; added 500K cell safety cap

### CoPilot Guide Additions (commit 92b6cd3)
- **CoPilot-Quick-Start-Card.md** — One-page cheat sheet: 5 scenarios → points to exact prompt section
- **VBA-Module-Reference-List.md** — All 38 demo modules + frmCommandCenter, grouped by category, with "easiest to adapt" section

### Training & QA Docs Updated
- User Training Guide updated to 65 commands (commit 41400cc)
- Operations Runbook updated (commit 41400cc)
- TestingPreVid.md standalone testing guide added (commit dfadfa3)
- QA docs updated to reflect 39 modules, 5 add-ins, 35 bugs (commit 0359842)
- Guide review and PDF conversion marked complete (commit 203361c)
- What-If guides moved to training/LastGuidesReview/ (commit ec203b7)

### Module Counts (Updated)
- Demo file VBA modules: **39 total** (34 core + 5 optional add-ins)
- Universal Toolkit VBA modules: **23 total** (14 previous + 9 new)
- Universal Toolkit tools: **~140+ total**

### Key Commits (14 total today)
- 92b6cd3 — CoPilot Quick Start Card + VBA Module Reference List
- 63482a4 — Fix 8 bugs across 7 universal toolkit modules
- cc45970 — Add 9 new universal toolkit modules (38 tools)
- d063cca — Add Universal Command Center
- 0cc3d0e — Add modUTL_WhatIf
- 2c72df8 — Add What-If Demo actions 63-65
- 5c367c8, ec203b7 — What-If guides
- 41400cc — Update User Training Guide + Operations Runbook
- dfadfa3 — TestingPreVid.md
- 6309434 — Rename demo Excel file
- 351c9b2 — Fix compile error in AddNamedRanges
- 0359842 — Update QA docs
- 203361c — Mark guide review complete

### Docs Updated
- CLAUDE.md — Full update: Repo Structure, Sharing Plan, Current Status, Session Summary
- tasks/todo.md — Updated current status, added completed work for 2026-03-12
- tasks/lessons.md — Added 4 new patterns: RGB constant verification, sheet index shifting, consolidation consistency, large range safety caps

---

## Session Summary — 2026-03-11 (Optional Add-Ins + Universal Expansions + Repo Cleanup)

### What Was Done This Session
Built 5 "Last Call Optional Add-Ins" from the todo.md backlog, created 3 universal toolkit versions, moved `_internal/` to `OldRoughVersions/` for cleanup, ran 5-pass bug review, and updated all docs.

**Branch:** `claude/resume-ipipeline-demo-qKRHn`

### Repo Cleanup
- Moved `_internal/` folder into `OldRoughVersions/_internal/` — dev-only folders no longer clutter root
- Created `LastCallOptionalAddIns/` folder for future guides related to the 5 new add-in modules

### 5 New Demo File Modules (Optional Add-Ins)
1. **`vba/modTimeSaved_v2.1.bas`** (305 lines) — Time Saved Calculator
   - Shows manual vs automated time for all 62 Command Center actions
   - Builds styled report sheet with per-action savings and Executive Summary box
   - Key output: "Manual: X hrs/month -> Automated: Y hrs/month -> Annual: Z hrs/year"
   - Demo-specific (depends on modConfig, modPerformance, modLogger)

2. **`vba/modSplashScreen_v2.1.bas`** — Branded Welcome Screen
   - Professional splash on workbook open with iPipeline branding
   - `ShowSplash` tries UserForm first, falls back to MsgBox
   - `BuildSplashForm` programmatically creates frmSplash UserForm

3. **`vba/modProgressBar_v2.1.bas`** (270 lines) — Animated Progress Bar
   - 3-call API: `StartProgress`, `UpdateProgress`, `EndProgress`
   - Shows %, ETA, elapsed time with animated bar fill
   - Falls back to status bar if frmProgress UserForm doesn't exist
   - `BuildProgressForm` creates frmProgress programmatically

4. **`vba/modWhatIf_v2.1.bas`** (558 lines) — Live "What If" Scenario Demo
   - 7 preset scenarios + custom + restore baseline
   - Presets: Revenue +/-15/10%, AWS +25%, Headcount +20%, Expenses -10%, Best/Worst Case
   - `SaveBaseline` saves Assumptions to hidden sheet (first run only)
   - `RestoreBaseline` restores originals and cleans up impact sheets
   - Demo-specific (depends on Assumptions sheet structure)

5. **`vba/modExecBrief_v2.1.bas`** (447 lines) — Executive Brief Auto-Generator
   - One-click executive brief scanning 5 sections: Revenue, Reconciliation, Assumptions, Products, Workbook Health
   - Builds styled report sheet with color-coded status indicators

### 3 New Universal Toolkit Modules
1. **`UniversalToolsForAllFiles/vba/modUTL_ProgressBar.bas`** — Standalone progress bar using status bar only (no UserForm dependency). ASCII visual bar: `[=========>          ] 45%`
2. **`UniversalToolsForAllFiles/vba/modUTL_SplashScreen.bas`** — Standalone splash screen using MsgBox only. Customizable title/subtitle.
3. **`UniversalToolsForAllFiles/vba/modUTL_ExecBrief.bas`** (253 lines) — Scans any workbook: Overview, Sheet Inventory, Data Quality (errors + formulas via SpecialCells), Hidden Sheets. Zero dependencies.

### 5-Pass Bug Review
**Pass 1 — Known patterns from lessons.md:** Checked LogAction signatures, SpecialCells guards, StyleHeader args, HDR_ROW_REPORT usage. All clean.
**Pass 2 — VBA-specific issues:** Checked On Error patterns, TurboOn/TurboOff pairing, variable resets. All clean.
**Pass 3 — Cross-module dependencies:** Verified all 19 modConfig constants referenced by new modules exist. Confirmed universal modules have zero dependencies.
**Pass 4 — Edge cases:** Checked midnight timer rollover, empty sheet handling, division by zero guards. All clean.
**Pass 5 — Character encoding linter:** Found `Chr(9472)` in 3 locations — crashes VBA (only handles 0-255). Fixed all to `String(50, "=")`. Also removed unused `SPLASH_BG` and `SPLASH_ACCENT` constants.

### Module Counts (Updated)
- Demo file VBA modules: **39 total** (34 previous + 5 optional add-ins)
- Universal Toolkit VBA modules: **14 total** (11 previous + 3 new)
- Universal Toolkit tools: **~100+ total**

### Re-Import Required (Updated — 18 files)
Previous 13 files + 5 new:
14. modTimeSaved_v2.1.bas (NEW)
15. modSplashScreen_v2.1.bas (NEW)
16. modProgressBar_v2.1.bas (NEW)
17. modWhatIf_v2.1.bas (NEW)
18. modExecBrief_v2.1.bas (NEW)

### Docs Updated
- CLAUDE.md — Full update: Repo Structure, Sharing Plan, Current Status, Session Summary
- tasks/lessons.md — Added Chr() range limitation lesson, CLAUDE.md update reminder lesson

---

## Session Summary — 2026-03-07 (Code Review + Bug Fixes + Video Package + CoPilot Guide)

### What Was Done This Session
Pre-delivery code review of all 46 live-Excel test pass criteria. Found and fixed 7 bugs across 5 VBA modules. Improved CoPilot Prompt Guide, built video package draft, created sample Excel file for demo.

**Branch:** `claude/resume-ipipeline-demo-qKRHn`

### File Uploads (2026-03-06)
- Connor uploaded CoPilot Prompt Guide files (AP_Copilot_PromptGuideHelpV2.docx + .md) to CoPilotPromptGuide/

### CoPilot Guide + Video Package (commit 74dc77e)
- CoPilot Prompt Guide v2.0: Fixed all broken quick reference links, added working anchor links, improved formatting
- COMPILED_VIDEO_PACKAGE.md: Fixed tool counts (13 VBA/78 tools + 22 Python = ~100), added demo file stats, updated build checklist
- Sample_Quarterly_Report.xlsx: Built for Video 3 universal tools demo (via build_sample_file.py)

### Pre-Delivery Code Review — 7 Bugs Found and Fixed
**Commit 6818b01 — 6 bugs:**
1. modReconciliation: LogAction 4th arg elapsed (Double) — instance #13
2. modReconciliation: ValidateCrossSheet trendLastCol used row 1 instead of HDR_ROW_REPORT
3. modReconciliation: FindFYCol scanned row 1 instead of HDR_ROW_REPORT
4. modPDFExport: GetReportSheetList hardcoded to 7 sheets — now dynamically discovers all monthly tabs
5. modDataSanitizer: rng not reset before SpecialCells in 2 worker functions
6. modAuditTools: rng not reset before SpecialCells in FindExternalLinks

**Commit b132885 — 1 bug:**
7. modDrillDown: HideGLSheet used xlSheetVeryHidden (blocks hyperlinks) — changed to xlSheetHidden

### Re-Import Required (Updated — 13 files)
See tasks/todo.md for the full list. Added 3 new files since last session: modDataSanitizer, modAuditTools, modDrillDown.

### Docs Updated
- tasks/todo.md — Updated current status, added 2026-03-06/07 completed work, expanded re-import list to 13 files
- tasks/lessons.md — Added 4 new patterns: SpecialCells rng reset, HDR_ROW_REPORT consistency, dynamic sheet discovery, xlSheetVeryHidden blocks hyperlinks
- CLAUDE.md — Updated current status and session summary

---

## Session Summary — 2026-03-05 (Earlier — Code Review + Bug Fixes + Doc Updates)

### What Was Done This Session
Full 3-pass code review of ALL new code built in the previous session, using known bug patterns from lessons.md. Found and fixed 4 bugs. Updated all docs to reflect current state.

**Branch:** `claude/resume-ipipeline-demo-qKRHn`

### Bug Review (3 Passes)
**Pass 1 — Known Bug Patterns from lessons.md:**
- Checked all LogAction signatures across all VBA modules
- Found 3 more instances of the recurring LogAction bug (Double passed as status String)

**Pass 2 — VBA Code Quality:**
- Checked SpecialCells Nothing guards, UsedRange iteration, loop variable resets, Chr vs ChrW, RGB colors
- All clean — no issues found

**Pass 3 — Python Code:**
- Found `detect_date_columns()` missing `day_first` parameter that was being passed by the caller
- Would crash with `TypeError` at runtime

### Bugs Found and Fixed (4 total)
1. **modDataQuality_v2.1.bas:150** — LogAction 4th arg = `ElapsedSeconds()` (Double → moved into message string)
2. **modReconciliation_v2.1.bas:128** — LogAction 4th arg = `ElapsedSeconds()` (Double → moved into message string)
3. **modPDFExport_v2.1.bas:102** — LogAction 4th arg = `ElapsedSeconds()` (Double → moved into message string)
4. **date_format_unifier.py:97+182** — `detect_date_columns()` missing `day_first` parameter (added to signature + passed through to `parse_date`)

### Docs Updated
- `tasks/todo.md` — Updated current status, ProjectRefresh marked COMPLETE, re-import list expanded to 10 files, PR branch corrected
- `tasks/lessons.md` — Added LogAction recurring bug pattern (now found 12 total times), added Python signature mismatch pattern
- `CLAUDE.md` — Updated current status and session summary

### Re-Import Required (Updated — 10 files)
1. modConfig_v2.1.bas
2. modReconciliation_v2.1.bas (LogAction fix)
3. modVarianceAnalysis_v2.1.bas (YoY + GenerateCommentary fix)
4. modDashboard_v2.1.bas (split — base only)
5. modDashboardAdvanced_v2.1.bas (NEW)
6. modDemoTools_v2.1.bas
7. modTrendReports_v2.1.bas
8. modMonthlyTabGenerator_v2.1.bas
9. modDataQuality_v2.1.bas (Letter Grade + LogAction fix)
10. modPDFExport_v2.1.bas (LogAction fix)

---

## Session Summary — 2026-03-05 (Earlier — ProjectRefresh Audit + Demo Enhancements + Training Guides)

### What Was Done
Full ProjectRefresh code audit completed. 120 tools cross-referenced. All Tier 1 recommendations implemented as working code:
- 3 critical demo bugs fixed (duplicate constants, GL visibility, TurboOn/Off)
- modDashboard split into base + advanced (was 1,398 lines)
- modUTL_Core shared utilities module created
- SpecialCells performance fix for 8 slow universal tools
- Backup-before-destructive added to 6 high-risk tools
- Data Quality Letter Grade (A-F) added to modDataQuality
- Forecast Accuracy Scoring (MAPE) added to pnl_forecast.py
- YoY Variance Analysis added to modVarianceAnalysis
- 14 new universal tools (7 VBA + 4 Python)
- 6 training guides built in FinalRoughGuides/

---

## Session Summary — 2026-03-04 (Earlier — New VBA Tools + Universal Toolkit Expansion)

### What Was Done This Session
Expanded both the demo file and Universal Toolkit with new VBA modules, created the first training guide draft, and set up project infrastructure.

**Branch:** `claude/resume-ipipeline-demo-qKRHn`

### New Demo File Module
- `vba/modSheetIndex_v2.1.bas` — 2 subs:
  - `CreateHomeSheet` — Creates a "Home" sheet at position 1 with a styled button that opens the Command Center (calls LaunchCommandCenter) + a "View Sheet Index" button
  - `ListAllSheetsWithLinks` — Lists all sheets in column A with clickable hyperlinks in column B. Safe to re-run (only adds sheets not already listed). Shows visibility status (Visible/Hidden/Very Hidden)

### New Universal Toolkit Modules (3 new .bas files)
- `UniversalToolsForAllFiles/vba/modUTL_Branding.bas` — 2 tools:
  - `ApplyiPipelineBranding` — Detects headers/totals on active sheet, applies iPipeline brand colors (iPipeline Blue headers, Navy totals, alternating rows, Arial font)
  - `SetiPipelineThemeColors` — Sets workbook theme colors to iPipeline brand palette so they appear in the Excel color picker. Falls back to legacy palette on older Excel versions
- `UniversalToolsForAllFiles/vba/modUTL_SheetTools.bas` — 3 tools:
  - `ListAllSheetsWithLinks` — Universal version: creates UTL_SheetIndex sheet with hyperlinks
  - `TemplateCloner` — Pick any sheet, type how many copies, get instant clones (1-50). Handles name conflicts and 31-char limit
  - `GenerateUniqueCustomerIDs` — Scans existing IDs, finds max, fills blank cells with sequential IDs (CUST-00001 format). Never duplicates. Custom prefix supported
- `UniversalToolsForAllFiles/vba/modUTL_DataSanitizer.bas` — 4 tools (ported from demo):
  - `RunFullSanitize` — All 3 fixes in one click (text-numbers, floating-point tails, integer formats)
  - `PreviewSanitizeChanges` — Dry-run report showing what WOULD change (no edits)
  - `FixFloatingPointTails` — Fix FP noise on all sheets
  - `ConvertTextStoredNumbers` — Convert text-numbers to real numbers
  - Smart detection: skips dates, names, IDs, labels, formulas via header keyword scanning

### Training Guide Draft
- `FinalRoughGuides/Dynamic-Chart-Filter-Setup-Guide.md` — Step-by-step guide for coworkers on how to add dropdown chart filters to their own Excel files. Covers Data Validation dropdowns, helper tables, PivotTables, Slicers, and troubleshooting

### Infrastructure
- Created `FinalRoughGuides/` folder for draft guides
- Added iPipeline brand styling reference to CLAUDE.md
- Updated tasks/lessons.md with brand styling lesson

### Bug Audit (2 passes)
All new code was audited twice for bugs. Bugs found and fixed:
1. **CRITICAL:** `rng` variable not reset to `Nothing` between loop iterations in modUTL_DataSanitizer worker functions — would cause double-processing of previous sheet's cells. Fixed by moving `Dim rng` outside loop + adding `Set rng = Nothing`
2. **MEDIUM:** Sheet names with apostrophes broke hyperlink SubAddress in modSheetIndex. Fixed with `Replace(name, "'", "''")`
3. **MEDIUM:** `Chr(8212)` only handles 0-255 in VBA — changed to plain dash
4. **MEDIUM:** Wrong RGB color `RGB(31,73,125)` in new universal modules — changed to brand-correct `RGB(11,71,121)`
5. **LOW:** `usedRng` not reset in PreviewSanitizeChanges loop — added `Set usedRng = Nothing`

### Module Counts
- Demo file VBA modules: 33 total (32 + modSheetIndex)
- Universal Toolkit VBA modules: 8 total (5 existing + 3 new)
- Universal Toolkit tools: ~85 total

### Already Existed (Not Rebuilt)
- Fix sheet links → `modAuditTools.FixExternalLinks` (demo) + `modUTL_Audit.ExternalLinkSeveranceProtocol` (universal)
- Show broken links → `modAuditTools.FindExternalLinks` (demo) + `modUTL_Audit.ExternalLinkFinder` (universal)

---

## Session Summary — 2026-03-04 (Earlier — Testing Bug Fixes + Self-Review)

### What Was Done This Session
This session focused on fixing bugs discovered during testing and then doing a full self-review of ALL remaining untested code against the test plan pass criteria — catching bugs before the user has to find them.

**Branch:** `claude/resume-apclmerge-project-V8WSj` (new branch, forked from CXWP5 — all CXWP5 work is included)

### Testing Bug Fixes (found by user during testing)
1. **T2.01 — PASS after fix:** Added 9 missing sheet-name constants to modConfig (commit af44453)
2. **T2.03 — PASS after fix:** CLR_NAVY and CLR_ALT_ROW color constants had wrong hex-to-decimal conversion (VBA BGR byte order) (commit 19320db)
3. **T2.04 — PASS after fix:** Added TestUpdateHeaderText wrapper (BUG-T2.04) + set NumberFormat to Text before writing "Mar 25" (BUG-T2.04b) (commits 6f40f91, ed3276f)
4. **T4.04 — PASS after fix:** Windows PermissionError on temp file cleanup + removed email report feature (commit 3024c44)
5. **T5.01 — PASS after fix:** CreateExecutiveDashboard read row 1 instead of row 4 for headers + Error 5 crash + row/column detection failures (commits 6c17bd5, 847a982)
6. **T5.02 — PASS after fix:** WaterfallChart row label fallbacks — searches for multiple label variants ("Total Revenue"/"Revenue"/"Net Revenue", etc.) (commit 304743b)

### Self-Review: 12 Bugs Found and Fixed Preemptively (commit 22ba831)
After the user asked "what can we do to limit bugs?", I ran a full self-review of all VBA modules against every remaining test's pass criteria. Found and fixed:

**2 Critical Logic Bugs:**
1. **modReconciliation line 292:** `dateCol = 5` was reading the Category column (E) instead of the Date column (B = COL_GL_DATE = 2). ValidateCrossSheet Check 2 would never find any January GL rows.
2. **modVarianceAnalysis line 221:** GenerateCommentary read row 1 (company title) instead of HDR_ROW_REPORT (row 4) for column headers. This made tLastCol = 1, so FY/Budget column search loops never ran — all variances would be zero.

**9 LogAction Signature Bugs** (across modDashboard, modDemoTools, modTrendReports, modMonthlyTabGenerator):
- 9 call sites passed `elapsed` (a Double) as 4th arg — the `status` field expects a String like "OK". This corrupted audit log Status column with numeric values like "0.547". Fixed by moving elapsed into the message string.

**1 Constant Consistency Fix:**
- modReconciliation `amtCol = 7` hardcoded → replaced with `COL_GL_AMOUNT` constant.

### What Passed Self-Review (No Bugs Found)
- T2.05 (FixTextNumbers guard) — correct
- T2.06 (Shortcuts use OnKey, Ctrl+H not overridden) — correct
- T2.07 (Timer midnight rollover) — correct
- T5.06 (Search 200-result cap) — correct
- T8.14/T8.30 (Button OnAction macro names) — correct
- T8.31 (ClearShortcuts) — correct
- All constant references across all T8 modules — verified against modConfig
- All StyleHeader calls — all pass exactly 3 arguments
- Python pytest: 99 passed, 15 skipped, 0 failures

### Pre-Delivery Self-Review Requirement Added
Added to tasks/lessons.md — all future code updates must be self-reviewed against the test plan before delivery.

### Files Modified This Session (6 VBA + 1 doc)
- `vba/modReconciliation_v2.1.bas` — dateCol/amtCol constants fix
- `vba/modVarianceAnalysis_v2.1.bas` — row 1 → HDR_ROW_REPORT fix
- `vba/modDashboard_v2.1.bas` — 4 LogAction fixes + WaterfallChart row label fallbacks + ExecDashboard fixes
- `vba/modDemoTools_v2.1.bas` — 1 LogAction fix
- `vba/modTrendReports_v2.1.bas` — 1 LogAction fix
- `vba/modMonthlyTabGenerator_v2.1.bas` — 3 LogAction fixes + TestUpdateHeaderText wrapper
- `tasks/lessons.md` — added Pre-Delivery Self-Review Requirement

### Re-Import Required
The user must re-import these 6 `.bas` files into the Excel workbook before continuing testing:
1. modConfig_v2.1.bas (from earlier commit — color constants)
2. modReconciliation_v2.1.bas
3. modVarianceAnalysis_v2.1.bas
4. modDashboard_v2.1.bas
5. modDemoTools_v2.1.bas
6. modTrendReports_v2.1.bas
7. modMonthlyTabGenerator_v2.1.bas

### Test Status Summary (as of end of session)
| Category | Tests | Passed | Failed→Fixed | Not Yet Run |
|----------|-------|--------|-------------|-------------|
| T1 Compilation | 8 | 8 | 0 | 0 |
| T2 Foundation | 7 | 4 | 0 | 3 (T2.05–T2.07) |
| T3 Menu/Command Center | 5 | 0 | 0 | 5 |
| T4 Python | 4 | 1 (T4.04) | 0 | 3 |
| T5 Advanced VBA | 6 | 2 (T5.01, T5.02) | 0 | 4 |
| T6 Data Integrity | 6 | 0 | 0 | 6 |
| T7 Integration | 4 | 0 | 0 | 4 |
| T8 New v2.1 Modules | 29 | 0 | 0 | 29 |
| **Total** | **69** | **15** | **0** | **54** |

### What's Left (Next Session)
1. **Re-import 7 fixed .bas files** into Excel workbook
2. **Continue testing:** T2.05–T2.07, then T3, T4, T5.03–T5.06, T6, T7, T8
3. **Demo Readiness:** After all tests pass → live test all 62 Command Center actions → script demo video → build training guide
4. **Backlog:** Python .exe conversion, Universal Tools Add-In packaging

### Branch
- Active branch: `claude/resume-apclmerge-project-V8WSj`
- Key commits this session: af44453, 19320db, 6f40f91, ed3276f, 3024c44, 6c17bd5, 847a982, 304743b, 22ba831

---

## Session Summary — 2026-03-03 (Earlier — Universal Tools Build)

### What Was Done
- Resumed from usage limit — picked up Track B (Universal Tools build)
- Reviewed GrokALL.md, PrelexALL.md, GemAll.md from UniversalToolsForAllFiles/
- Created UniversalToolsForAllFiles/UniversalBuild/UNIVERSAL_BUILD_CANDIDATES.md (76 total candidates)
- Built ALL 76 Universal Tool candidates as actual working code
- Created review/PROJECT_OVERVIEW.md — full overview doc for external Claude review
- All changes committed and pushed (commit accc11a)

---

## Session Summary — 2026-03-01

### What Was Done This Session
- Reviewed 3 new files added to NewTesting/ (commit 075d457 — "Add files via upload, New ideas")
- Created new Ideas branch: `claude/ideas-newtesting-wDuOY` (based on review-branch-progress-pP7Qf)
- Built 7 new VBA modules from the VBA Examples (200) idea list
- Updated modDashboard with 2 new subs

### New VBA Modules (7 new files in vba/)
- `modDemoTools_v2.1.bas` — #17 AddControlSheetButtons, #63 SetParameterizedPrintArea, #64 CreatePrintableExecSummary
- `modDataGuards_v2.1.bas` — #48 ValidateAssumptionsPresence, #49 CheckSumOfDrivers, #150 FindNegativeAmounts, #151 FindZeroAmounts, #155 FindSuspiciousRoundNumbers
- `modDrillDown_v2.1.bas` — #18 AddReconciliationDrillLinks, #55 AutoPopulateReconciliationChecks, #56 ApplyReconciliationHeatmap, #90 RunGoldenFileCompare
- `modAuditTools_v2.1.bas` — #93 AppendChangeLogEntry, #106 FindExternalLinks, #107 FixExternalLinks, #109 AuditHiddenSheets, #115 CreateMaskedCopy, #196 ExportErrorSummaryClipboard
- `modETLBridge_v2.1.bas` — #119 TriggerETLLocally, #120 ImportETLOutput
- `modTrendReports_v2.1.bas` — #77 CreateRolling12MonthView, #156 CreateReconciliationTrendChart, #163 ArchiveReconciliationResults
- `modDashboard_v2.1.bas` updated — added #44 LinkDynamicChartTitles, #86 CreateSmallMultiplesGrid

### Additional Modules (added same session)
- `modDataSanitizer_v2.1.bas` — numeric-only sanitizer; never touches dates, names, or customer IDs
  - SKIP_HEADER_KEYWORDS protects: id, date, name, code, customer, client, account, acct, company,
    vendor, contact, employee, entity, description, dept, product, type, status, region, address, etc.
- `modMonthlyTabGenerator_v2.1.bas` updated — new `AddNextMonthToModel` sub + `MarkTrendColumn` helper
  - Calendar-aware: reads today's date, determines next month automatically
  - Marks next month column yellow on both trend sheets (P&L Monthly Trend + Functional P&L Monthly Trend)
  - Clones current month's Functional P&L Summary tab to create next month's tab

### Total VBA Module Count: 32 modules (was 24 + 7 new + modDashboard updates + modDataSanitizer + modMonthlyTabGenerator update)

### NewTesting Files Reviewed
- `Financial Model Correction Instructions.md` — 6-point fix checklist for Excel model
- `2026-02-28T223817Z.md` — Full audit: 15 issues, 10 VBA macros, Python ETL, Power Query M-Code
- `VBA Examples (200) — Name — Purpose.txt` — Catalog of 200 macro ideas (source for new modules above)

### Key Overlap Notes (Do NOT double-import these)
- Audit doc's `FixTextNumbers` = already in `modDataQuality`
- Audit doc's `RunReconciliation` = already in `modReconciliation`
- Audit doc's `ExportChecksPDF` = already in `modPDFExport`

### Next Steps (Updated)
1. Import all 32 .bas files into Excel workbook (Alt+F11 → File → Import)
2. Create frmCommandCenter UserForm in the workbook
3. Live test every Command Center action (1-62) — log pass/fail
4. Run all new modDataGuards checks against real data (FindNegativeAmounts, etc.)
5. Run RunGoldenFileCompare to save baseline before any changes
6. Fix 6 Critical issues found in audit doc (floating-point, text-stored numbers, duplicates, etc.)
7. Write demo video storyboard/script
8. Build coworker training guide
9. Copy final files to CompletePackageStorage/production/
10. Record demo video

---

## Session Summary — 2026-02-28

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
- ExcelDemoFile_adv.xlsm — iPipeline Fortune 100 redesigned version
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
- iPipeline brand styling guide lives at `docs/ipipeline-brand-styling.md` — use it for ALL future visual/styled outputs
