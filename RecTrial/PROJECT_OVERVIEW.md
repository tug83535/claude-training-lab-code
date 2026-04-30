# iPipeline Finance Automation Demo Project — Master Overview

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Audience of this doc:** anyone (AI or human) giving the project a second-opinion review
**Snapshot date:** 2026-04-23
**Project status in one sentence:** Videos 1–3 recorded and shipped. Universal toolkit + zero-install Python pack live. V4 planning complete + Python build complete (2026-04-28); narration script and recording are next; post-video Batches 4–5 still parked.

---

## 1. Elevator pitch

A 4-video internal demo series for **50–150 iPipeline coworkers near-term** in Finance, Accounting, and adjacent operations, showing what's possible when a Finance & Accounting team combines **Excel + VBA + Python + SQL**. Built and recorded by a non-developer Finance analyst over ~6 weeks. Ships with an adoption-grade toolkit (~140 Excel tools + 28+ Python scripts) that coworkers can drop into their own files and use on their own data. Broader rollout and CFO/CEO showcase are deferred to v2.

The underlying thesis: Finance doesn't need to be a source of manual spreadsheet work anymore. With the right patterns, everyday analysts can automate the repetitive 80% and redirect time to judgment work.

---

## 2. Audience & voice

- **Primary viewers (near-term):** 50–150 iPipeline coworkers (Finance, Accounting, adjacent ops) — non-developers, Excel-literate, zero Python exposure. Expected to actually adopt and use the tools on their own files, not just watch.
- **Future audience (v2):** Broader company rollout + CFO/CEO showcase — deferred, not cancelled.
- **Voice:** plain English. Every tool must be explainable to a non-coder. Every output must be CFO-legible on one screen.
- **Brand:** iPipeline (Blue `#0B4779`, Navy `#112E51`, Arial fonts, clean/corporate)

---

## 3. The 4 videos — status and content

### Video 1 — "What's Possible" ✅ Shipped

- **Role:** fast highlight reel (~5 min)
- **Stack shown:** Excel + VBA
- **Recorded hands-free** via a custom VBA "Director" macro (plays ElevenLabs AI narration, navigates sheets, runs tools, pauses). User presses one button and watches.
- **Output:** one-take demo of the P&L workbook's most impressive 7 moments
- **Title card:** `VIDEO 1 OF 4 / What's Possible`

### Video 2 — "Full Demo Walkthrough" ✅ Shipped

- **Role:** deep tour (~18 min)
- **Stack shown:** Excel + VBA
- **Content:** all 62 automated actions on the demo P&L workbook — command center, data quality, reconciliation, variance analysis, dashboards, PDF export, exec brief, etc.
- **Also Director-automated** via `RunVideo2`
- **Title card:** `VIDEO 2 OF 4 / Full Demo Walkthrough`

### Video 3 — "Universal Tools" ✅ Shipped

- **Role:** plug-and-play toolkit demo (~8 min)
- **Stack shown:** Excel + VBA (universal toolkit on any file)
- **Content:** 10 universal tools demo'd on a fresh sample file — Data Sanitizer, Highlights, Comments, Tab Organizer, Column Ops, Sheet Tools, Compare, Consolidate, Pivot Tools, Command Center
- **Iterations:** went through 5 Gemini-assisted review cycles (v1 → v2.5) catching issues in silent Director wrappers, color rendering, sheet-name confusions. Final 70/4 Gemini result was accepted — remaining "bugs" were perception artifacts, not functional.
- **Delivered:** sample file `Sample_Quarterly_ReportV2.xlsm` + SHOW TOOLS button on Cover sheet launching the full Command Center
- **Title card:** `VIDEO 3 OF 4 / Universal Tools`

### Video 4 — "Python Automation for Finance" 🔄 Python build complete — recording next

- **Status:** Direction locked 2026-04-28. All 6 V4 Python scripts built and smoke-tested 5/5 PASS (2026-04-28). Narration script and ElevenLabs audio are the next blocking items.
- **Delivery model:** VBA Shell() + bundled Python 3.11 embeddable + `finance_automation_launcher.py` CLI menu. Excel buttons in `FinanceTools.xlsm`. One zip on SharePoint. Zero install for coworkers.
- **Original plan:** 10 ElevenLabs-narrated clips, manual recording, 8 Python scripts run from Command Prompt — pulled back 2026-04-22. Original audio + demo files remain on disk.
- **Title card:** `VIDEO 4 OF 4 / Python Automation for Finance` (already generated, matches V1–V3 style)

---

## 4. The codebase

### 4.1 VBA — 23 universal toolkit modules + ~140 tools

Plug-and-play on any `.xlsm` file. Each module is independent and can be imported on its own.

| Module | Purpose |
|---|---|
| `modUTL_Core` | Shared utilities (Turbo On/Off, LastRow, LastCol, detect header row, styled header, backup sheet) |
| `modUTL_Audit` | External link finder/severance, circular ref detector, error scanner, data quality, named range auditor, run receipt sheet |
| `modUTL_Branding` | Applies iPipeline colors, fonts, alt-row formatting to any sheet |
| `modUTL_DataCleaning` | 12 cleanup tools (unmerge/fill, text-to-numbers, remove spaces, delete blanks, error replacement, etc.) |
| `modUTL_DataSanitizer` | Full + preview sanitize, floating-point fix, integer format normalization |
| `modUTL_Highlights` | Threshold + duplicate + clear highlighters (saturated green / pure orange) |
| `modUTL_Comments` | Extract all comments into an inventory sheet |
| `modUTL_TabOrganizer` | Color tabs by keyword, reorder alphabetically, group |
| `modUTL_ColumnOps` | Split/combine columns (string-params version for Director automation) |
| `modUTL_SheetTools` | Template cloner, sheet index with hyperlinks |
| `modUTL_Compare` | Sheet-to-sheet diff + fast row-signature pre-check (`UTL_QuickRowCompareCount`) |
| `modUTL_Consolidate` | Multi-sheet consolidator with "Source Sheet" Column A tracking |
| `modUTL_PivotTools` | List pivots, refresh all, with silent DirectorListAllPivots wrapper |
| `modUTL_CommandCenter` | Menu-driven launcher for every toolkit tool (auto-discovers new modules) |
| `modUTL_LookupBuilder` | VLOOKUP/INDEX-MATCH helper |
| `modUTL_ValidationBuilder` | Data validation builder |
| `modUTL_Finance` | Finance-specific tools (invoice dup detector, GL validator, trial balance checker) |
| `modUTL_ExecBrief` | Workbook profiler that generates an exec brief |
| `modUTL_Intelligence` | **[NEW, Codex Batch 2]** Materiality Classifier, Exception Narratives, Data Quality Scorecard (3 universal subs) |
| `modUTL_Formatting` | Cell formatting helpers |
| `modUTL_WorkbookMgmt` | File-level management helpers |
| `modUTL_ProgressBar` | Status-bar progress helper |
| `modUTL_SplashScreen` | Workbook open splash |

Plus the core demo-file VBA (32 modules) and the Director automation macro `modDirector.bas` for Video recording.

### 4.2 Python — 28 scripts

| Category | Scripts |
|---|---|
| **Finance operations** | `aging_report`, `bank_reconciler`, `variance_analysis`, `variance_decomposition`, `pnl_forecast`, `pnl_dashboard` |
| **Data handling** | `compare_files`, `consolidate_files`, `multi_file_consolidator`, `date_format_unifier`, `two_file_reconciler`, `fuzzy_lookup`, `clean_data`, `unpivot_data`, `master_data_mapper`, `batch_process`, `regex_extractor` |
| **Extraction & reporting** | `pdf_extractor`, `forecast_rollforward`, `sql_query_tool`, `word_report` (with `--talking-points` flag from Codex Batch 3) |
| **Zero-install stdlib-only** (new subfolder `ZeroInstall/`) | `profile_workbook`, `sanitize_dataset`, `compare_workbooks`, `build_exec_summary`, `variance_classifier`, `scenario_runner`, `sheets_to_csv` |

### 4.3 SQL — 4 scripts

`staging.sql`, `transformations.sql`, `validations.sql`, `pnl_enhancements.sql`

### 4.4 Sample files

- `ExcelDemoFile_adv.xlsm` — demo P&L workbook (Videos 1 & 2)
- `Sample_Quarterly_ReportV2.xlsm` — universal toolkit demo (Video 3)
- Various Python demo input files in `RecTrial\Video4DemoFiles\`

---

## 5. The guides and training artifacts

### 5.1 Final training guides (shipped to coworkers)

At `FinalExport\Guides\` (PDFs) and `FinalExport\Guides_v2\` (Word):

1. `00-Start-Here-Welcome.pdf`
2. `01-How-to-Use-the-Command-Center.pdf`
3. `02-Getting-Started-First-Time-Setup.pdf`
4. `03-What-This-File-Does-Overview.pdf`
5. `04-Quick-Reference-Card.pdf`
6. `05-User-Training-Guide.pdf`
7. `06-Universal-Toolkit-Guide.pdf`
8. `07-Operations-Runbook.pdf`
9. `08-WhatIf-Scenario-Guide.pdf`
10. `09-Universal-CommandCenter-Guide.pdf`
11. `10-VBA-Module-Reference-List.pdf`
12. `AP-Copilot-Prompt-Guide.pdf`
13. `Company-BrandStyling-CopilotPrompt.pdf`
14. `Dynamic-Chart-Filter-Setup-Guide.pdf`
15. `Source-Code-vs-Universal-Toolkit.pdf`

### 5.2 Recording guides (for Connor only)

At `RecTrial\Guide\`: Master Recording Guide, V3 Gemini Review (v3.3), V3 Clip Tracker, V3 Interactive Guide, V3 Step-by-Step, V4 Interactive Guide, V4 Narration Script, V4 Recording Guide.

### 5.3 Brand + style

At `docs\ipipeline-brand-styling.md` — official RGB values, fonts, layout rules.

---

## 6. The Codex parallel build + cherry-pick campaign

**What it was:** Connor ran a parallel ChatGPT Codex session at `tug83535/AP_CodexVersion` that built a separate version of the Finance automation project from scratch. This was done as a comparison / second-opinion exercise.

**What came out of it:** A structured comparison report at `RecTrial\CodexCompare\COMPARISON_REPORT.md` identifying which Codex ideas were worth porting into Project A. Full tracker at `RecTrial\CodexCompare\CHERRY_PICK_TRACKER.md`.

### Cherry-pick campaign results — 3 batches shipped

**Batch 1 (2026-04-20, commit `fcd0211`):**
- `MarginVerdict` + `AppendMarginVerdictRow` in `modWhatIf_v2.1.bas` (aggressive/monitor/escalate classifier)
- `CreateRunReceiptSheet` in `modUTL_Audit.bas` (6-row audit receipt)
- `UTL_DetectHeaderRow` in `modUTL_Core.bas` (auto-detect header row)

**Batch 2 (2026-04-21, commit `39b4ce4`):**
- New `modUTL_Intelligence.bas` — MaterialityClassifier + ExceptionNarratives + DataQualityScorecard (3 universal tools)
- `UTL_QuickRowCompareCount` + `BuildRowHashMap` in `modUTL_Compare.bas`

**Batch 3 (2026-04-21, commit `282ccbb`):**
- 7 zero-install stdlib-only Python scripts in `RecTrial\UniversalToolkit\python\ZeroInstall\`
- `--talking-points` CLI flag in `word_report.py`

**Batches 4–5 — deferred post-V4:**
- Batch 4: dual-logging pattern (demo file)
- Batch 5: top-level docs (`CONSTRAINTS.md`, `BRAND.md`, `RELEASE_READINESS_CHECKLIST.md`, `TROUBLESHOOTING.md`)

---

## 7. Current Video 4 plan — LOCKED 2026-04-28

Full details in the 5-doc planning sprint (all complete 2026-04-28). **`VIDEO_4_CURRENT_PROPOSAL.md` is now SUPERSEDED** — see `VIDEO_4_REVIEW_DECISION_MEMO.md` for the locked direction.

**Direction:** Single chaptered Video 4, 9–12 min. No 4a+4b split.

### "Python Automation for Finance" — 8 chapters

- **Hero:** Revenue Leakage Finder — "Python found a possible billing problem" (narrative story)
- **Closing artifact:** ARR waterfall chart — executive-readable summary of the leakage analysis
- **Audience:** 50–150 Finance/Accounting coworkers who will use the tools, not just watch
- **Distribution:** Excel buttons via VBA Shell() + bundled Python 3.11 embeddable, one zip on SharePoint
- **Pilot:** 10–20 users before broader rollout (role mix: Finance + Accounting + Billing/RevOps + managers)

### Deliverable

`finance_automation_launcher.py` — simple menu launcher wrapping the V4 scripts. (NOT `finance_copilot.py` — "Copilot" implies AI this tool doesn't have.) xlwings Excel Button Edition parked as v2.

### Scripts built — ALL COMPLETE 2026-04-28 (smoke test: 5/5 PASS)

| # | Script | Status |
|---|---|---|
| 1 | `common/safe_io.py` + `logging_utils.py` + `report_utils.py` + `sample_data.py` | DONE |
| 2 | `data_contract_checker.py` | DONE |
| 3 | `revenue_leakage_finder.py` (hero) | DONE |
| 4 | `exception_triage_engine.py` | DONE |
| 5 | `control_evidence_pack.py` | DONE |
| 6 | `workbook_dependency_scanner.py` | DONE |
| 7 | `finance_automation_launcher.py` (menu) | DONE |
| 8 | `smoke_test_video4_python.py` + `README_VIDEO4_PYTHON.md` | DONE |

**Planning docs (source of truth):**
- `RecTrial\Brainstorm\VIDEO_4_REVIEW_DECISION_MEMO.md`
- `RecTrial\Brainstorm\SUPPORTED_WORKFLOWS_V1.md`
- `RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md`
- `RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md`
- `RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md`

---

## 8. Open items (remaining — not blockers for V4 build)

All major V4 decisions are LOCKED (2026-04-28). See `RecTrial\Brainstorm\VIDEO_4_REVIEW_DECISION_MEMO.md`.

**Remaining open items:**
1. **Pilot user list** — 10–20 specific people by role (Connor identifies; role mix in `MINIMUM_DISTRIBUTION_PLAN.md`)
2. **Coworker pip access** — determines whether stdlib-only is mandatory for all scripts (Connor confirms after a quick check with a coworker)
3. **IT endpoint scanner** — bundled `python.exe` may be flagged when distributed via SharePoint zip; Connor coordinates with IT before pilot launch

---

## 9. Research foundation for V4

Connor ran parallel sessions in other AI tools producing ~14 raw research files and 6 compiled synthesis docs (V1–V6). All at `RecTrial\Brainstorm\NewCodeResearch\`.

- **~156 unique ideas** inventoried across all raw files
- **~40–60 curated** per compiled doc into Sections A/B/C/D (Universal Toolkit / Video 4 candidates / Future / Skip)
- Full synthesis completed by Claude Code; confidence HIGH that all actionable ideas within constraints are captured
- Raw files back-checked via subagent 2026-04-23 — no new findings beyond what was already captured

---

## 10. Future ideas / parking lot

Full doc at `RecTrial\Brainstorm\FUTURE_AUTOMATION_IDEAS.md`. Summary of parked categories:

### External AI API ideas (parked until IT clarity)
AI variance narratives, LLM contract parsers, expense classifiers, AI anomaly explainers.

### Email / Outlook automation (parked)
Outlook mail merge, email-triggered reconciliations, scheduled email reports.

### Windows Task Scheduler / scheduled automation (parked)
Auto-refresh dashboards, auto-run month-end, watch folders.

### Warehouse-dependent SQL (future project)
Close Readiness Score View, Allocation Drift Tracker, Forecast Backtest Warehouse, Vendor Payment Velocity Baselines, JE Duplicate Ring Detection, Close Bottleneck Heatmap, SOD Audit Pack, Policy-as-Code Tables.

### ML-dependent Python (future approvals)
Forecast Ensemble, Close Calendar Risk Predictor, Isolation Forest anomaly detection, SARIMA/Prophet time-series, Splink entity resolution.

### Infrastructure (future project)
Airflow orchestration, Flask/FastAPI exception status API, dbt model layer, GitHub Actions CI, .NET signed add-in.

### Third-party platforms (discovery)
Power Automate workflows, Copilot Studio custom bots, Zapier / n8n, RPA (UiPath), Power BI / Tableau, Metabase, Streamlit / Dash apps, FloQast / BlackLine, Fireflies / Otter.ai, Adobe Acrobat batch, Azure Key Vault for credentials.

---

## 11. Current to-do list

Full list at `claude-training-lab-code\Archive\tasks\todo.md`. Summary:

### Immediate (blocking V4 recording)
- [x] Lock Video 4 plan — DONE (2026-04-28)
- [x] Build all 6 V4 Python scripts + launcher — DONE, smoke test 5/5 PASS (2026-04-28)
- [x] Sample data built into `common/sample_data.py` — DONE
- [ ] Git commit all V4 Python files ⚠️ NOT YET COMMITTED
- [ ] Design and build `FinanceTools.xlsm` Excel buttons (one launcher button OR per-tool — Connor decides)
- [ ] Test bundled Python 3.11 embeddable zero-install path on a real coworker machine
- [ ] Write V4 narration script — 8 chapters, 9–12 min (blocking for ElevenLabs and recording)
- [ ] Write shot list / screen recording guide per chapter
- [ ] Connor: generate ElevenLabs audio from narration script
- [ ] Connor: identify 10–20 pilot users (role mix in MINIMUM_DISTRIBUTION_PLAN.md)
- [ ] Connor: confirm coworker pip access
- [ ] Assemble SharePoint zip package
- [ ] Record Video 4

### Post-V4 wrap-up
- [ ] Batch 4: dual-logging pattern in demo file modules
- [ ] Batch 5: top-level `CONSTRAINTS.md`, `BRAND.md`, `RELEASE_READINESS_CHECKLIST.md`, `TROUBLESHOOTING.md`

### Optional polish (not blocking)
- [ ] Spot-test the 7 zero-install Python scripts
- [ ] Clean up "(Discovered)" duplicates in Command Center menu
- [ ] Port additional universal-toolkit items from the research backlog (Dependency Impact Preview, Workbook Policy Validator, Auto-Repair Suggestions, etc.)

---

## 12. File map — where to find everything

### Code sources (editing here first, then synced to repo)

- `C:\Users\connor.atlee\RecTrial\` — **active working folder**
  - `VBAToImport\modDirector.bas` — master Director macro
  - `DemoVBA\` — demo file VBA modules (32 files)
  - `UniversalToolkit\vba\` — 23 toolkit modules
  - `UniversalToolkit\python\` — 28 Python scripts (including `ZeroInstall\` subfolder)
  - `UniversalToolkit\sql\` — 4 SQL scripts
  - `SampleFile\SampleFileV2\Sample_Quarterly_ReportV2.xlsm` — V3 sample
  - `DemoFile\ExcelDemoFile_adv.xlsm` — V1/V2 demo
  - `Video4DemoFiles\` — 12 input files for Video 4 demos
  - `AudioClips\Video1\Video2\Video3\Video4\` — 40+ ElevenLabs narration MP3s
  - `Recordings\Video1-4\` — final recorded MP4s
  - `VideoTitleCards_v2\` — branded 5 title cards (V1–V4 + disclaimer)
  - `Guide\` — recording + Gemini review guides
  - `Feedback\Video3_*Feedback\` — 4 rounds of Gemini bug reports + v3.3 review prompt
  - `Brainstorm\` — **all planning docs**
    - `VIDEO_4_CURRENT_PROPOSAL.md` — locked-in-progress V4 plan
    - `VIDEO_4_DRAFT_IDEAS.md` — 17-idea initial draft
    - `FUTURE_AUTOMATION_IDEAS.md` — parking lot
    - `PROMPT_1_FULL_REVIEW.md` + `PROMPT_2_VIDEO_4_FOCUS.md` — research-prompts for claude.ai
    - `NewCodeResearch\ResearchComplied\` — 6 synthesis docs (V1–V6)
    - `NewCodeResearch\ResearchFiles\` — 14 raw research files
  - `CodexCompare\` — cherry-pick campaign tracker + comparison report
  - `VBABackup_PrePathA\` + `VBABackup_PreV2.2Fix\` — rollback safety nets

### Committed repository (source of truth for git)

- `C:\Users\connor.atlee\.claude\projects\claude-training-lab-code\` — git repo
  - `CLAUDE.md` — project instructions (long, for AI sessions)
  - `Archive\tasks\todo.md` — running task list
  - `Archive\tasks\lessons.md` — ~50 lessons learned documented
  - `FinalExport\` — what ships to coworkers (VBA bundle + Python bundle + PDFs)
  - `memory\` — auto-memory for Claude Code sessions

### GitHub

- `tug83535/claude-training-lab-code` — main project repo (branch `April19update`)
- `tug83535/AP_CodexVersion` — parallel Codex build (separate repo, read-only for reference)

---

## 13. History highlights

- **2026-02-28 → 2026-03-12:** Core VBA + Python built (~39 demo modules + 22 Python scripts). Universal Toolkit expanded to 23 modules, ~140 tools. 62 Command Center actions implemented. Multiple bug-review passes.
- **2026-03-12 → 2026-04-15:** Training guides finalized, CoPilot prompt guide v2 shipped, video package draft + sample file built, Director macro (v2.0) written.
- **2026-04-15 → 2026-04-21:** Videos 1–3 recorded. Video 3 went through 5 Gemini review cycles. Path A silent-wrapper refactor to eliminate dialog-based failures. All VBA fixes shipped to GitHub branch `April19update`.
- **2026-04-20 → 2026-04-22:** Codex cherry-pick campaign — comparison against parallel Codex build. Batches 1–3 shipped (5 VBA updates + new Intelligence module + 7 zero-install Python scripts + talking-points flag). SHOW TOOLS button on Cover sheet.
- **2026-04-22 → 2026-04-23:** Original Video 4 plan pulled. Brainstormed 17 alternatives, synthesized with 6 external AI research docs (156 ideas), converged on split 4a+4b plan with ARR Waterfall hero + Finance Copilot menu deliverable.

---

## 14. How to review this project

### If you're AI (giving a second opinion)

Useful angles to push on:
1. **V4 plan is locked.** Single video, 9–12 min, Revenue Leakage Finder hero, ARR waterfall closing artifact. Python build complete. This angle is closed — decisions already made.
2. **Is the universal toolkit actually adopt-able?** 140 tools is a lot. Which 20 matter most for coworker day-to-day? Is the Command Center UI intuitive enough?
3. **Cherry-pick completeness.** We ported 9 items from Codex. Are there obviously-better items we missed?
4. **Distribution strategy.** Excel buttons via VBA Shell() + bundled Python 3.11 embeddable. One zip on SharePoint. Coworkers click a button in `FinanceTools.xlsm`. Adoption plan in `MINIMUM_DISTRIBUTION_PLAN.md`.
5. **Post-demo roadmap.** The Future doc has lots of ideas. Which should happen first once V4 ships?
6. **Risk flags.** Is bundled `python.exe` blocked by iPipeline IT endpoint scanner? Coworker pip access still unconfirmed. Is the "no AI API" constraint right long-term?

### If you're Connor (personal review)

1. Does the V4 direction still feel right? If it doesn't excite you, don't build it.
2. Is the 5-day effort realistic given your other commitments?
3. Any decisions in Section 8 you want to nail down before lifting a finger?
4. Anything in the Future list (Section 10) you're actually excited about versus parking forever?

---

## 15. Key docs for deeper review

If a reviewer wants to dig into specific pieces:

| Topic | File |
|---|---|
| **V4 detailed plan (SUPERSEDED)** | `RecTrial\Brainstorm\VIDEO_4_CURRENT_PROPOSAL.md` — replaced by `VIDEO_4_REVIEW_DECISION_MEMO.md` |
| **V4 locked planning docs** | `RecTrial\Brainstorm\VIDEO_4_REVIEW_DECISION_MEMO.md`, `SUPPORTED_WORKFLOWS_V1.md`, `VIDEO_4_REVISED_PLAN.md`, `MINIMUM_DISTRIBUTION_PLAN.md`, `UniversalToolkit\python\PYTHON_SAFETY.md` |
| **V4 initial brainstorm (17 ideas)** | `RecTrial\Brainstorm\VIDEO_4_DRAFT_IDEAS.md` |
| **Future ideas parking lot** | `RecTrial\Brainstorm\FUTURE_AUTOMATION_IDEAS.md` |
| **Codex comparison** | `RecTrial\CodexCompare\COMPARISON_REPORT.md` |
| **Cherry-pick tracker** | `RecTrial\CodexCompare\CHERRY_PICK_TRACKER.md` |
| **Research synthesis V1–V6** | `RecTrial\Brainstorm\NewCodeResearch\ResearchComplied\` |
| **Raw research files** | `RecTrial\Brainstorm\NewCodeResearch\ResearchFiles\` |
| **Core project instructions** | `claude-training-lab-code\CLAUDE.md` |
| **To-do list** | `claude-training-lab-code\Archive\tasks\todo.md` |
| **Lessons learned** | `claude-training-lab-code\Archive\tasks\lessons.md` |
| **Brand guide** | `claude-training-lab-code\docs\ipipeline-brand-styling.md` |

---

*End of overview. If the project changes materially, regenerate this doc — it's a point-in-time snapshot, not a living spec.*
