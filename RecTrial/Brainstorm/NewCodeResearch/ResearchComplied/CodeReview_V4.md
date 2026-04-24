# Video 4 Planning — Consolidated Research Handoff

## 1. WHAT IS THIS FILE

This file contains three review passes — a full curated review (Sections A–D), a Video 4 finalists shortlist with scored table and camera narratives, and a consolidated Top 10 — drawn from 6 research files. Approximately 55 unique ideas were curated after de-duplication across source files, distilled to 8 Video 4 finalists and a ranked Top 10. Two combo recommendations and a split-video production recommendation are included.

---

## 2. PROJECT CONTEXT

**Who / what:** Connor, Finance & Accounting analyst at iPipeline (life insurance / financial services SaaS, ~2,000 employees). Building a 4-video internal demo series for 2,000+ coworkers + CFO/CEO showing Excel + VBA + Python + SQL + AI for Finance.

**Video status:**
- Video 1 — "What's Possible" (Excel + VBA highlight reel) — **RECORDED**
- Video 2 — "Full Demo Walkthrough" (62 automated actions on demo P&L) — **RECORDED**
- Video 3 — "Universal Tools" (VBA toolkit for any Excel file) — **RECORDED**
- Video 4 — "Python Automation for Finance" — **PLANNING NOW**, 5–8 min target, open to going longer or splitting

**Video 4 combos being weighed:**
1. **Finance Copilot menu** — numbered Python menu wrapping existing scripts
2. **Excel Button Edition (xlwings)** — Excel buttons trigger Python silently, results as new sheets
3. **Hero Demo + Cookbook** — one dramatic hero + 5 copy-pasteable recipe scripts

**Hard constraints:**
- No external AI API calls (OpenAI, Claude, Gemini)
- No Outlook or email automation
- No Windows Task Scheduler
- No internet scraping of company/paid data
- Non-developer audience — zero-coding-background explainable
- Plug-and-play — no hardcoded sheet names
- iPipeline branding: Blue `#0B4779`, Navy `#112E51`, Arial fonts

**Approved Python packages:** `pandas`, `openpyxl`, `pdfplumber`, `python-docx`, `thefuzz`, `numpy`, `matplotlib`, `xlwings`, stdlib only.

**Already built — Python (do not re-suggest):** `aging_report`, `bank_reconciler`, `compare_files`, `forecast_rollforward`, `fuzzy_lookup`, `pdf_extractor`, `variance_analysis`, `variance_decomposition`, `clean_data`, `consolidate_files`, `multi_file_consolidator`, `date_format_unifier`, `two_file_reconciler`, `sql_query_tool`, `word_report`, `batch_process`, `regex_extractor`, `unpivot_data`, `pnl_forecast`, `pnl_dashboard`, `master_data_mapper`, `profile_workbook`, `sanitize_dataset`, `compare_workbooks`, `build_exec_summary`, `variance_classifier`, `scenario_runner`, `sheets_to_csv`.

**Already built — VBA:** 23 modules, ~140 tools (sanitizer, compare, consolidate, highlights, pivot tools, tab organizer, column ops, sheet tools, comments, validation builder, lookup builder, command center, exec brief, finance tools, audit tools, branding).

**Already built — SQL:** 4 scripts (staging, transformations, validations, enhancements).

---

## 3. OPEN DECISIONS

### D1. Build status of VBA items A1–A8

**Decision:** Are the 8 VBA ideas in Section A1–A8 (Materiality Classifier, Exception Narrative Generator, Data Quality Scorecard, Header Row Auto-Detect, Quick Row Compare Count, Run Receipt Sheet, Cover Show Tools Button Installer, Pinned Intelligence Category in Command Center) actually already built?

**Options:** (a) Built — in `modUTL_*` modules already, drop from recommendations. (b) Not built — promote into active toolkit backlog.

**Findings this session:** The handoff file `CLAUDE_CODE_HANDOFF_BranchIdeas.md` marks all 8 as ✅ Built in `modUTL_*` modules. The project-level "already built" list references audit tools, command center, and finance tools but does not explicitly confirm these 8. **Resolution needed before actioning the Top 10.** If confirmed built, two swaps apply (see Section 5).

---

### D2. Video 4 — single video or split into 4A + 4B

**Decision:** Ship Video 4 as a single 5–8 min video, or split into 4A (hero, 6–7 min) + 4B (cookbook, 5–6 min)?

**Options:** (a) Single video, Combo 3 format. (b) Split into two videos, different audience focus per video.

**Findings this session:** Split recommended on audience-fit grounds — CFO/CEO want 2–3 wow tools, coworkers want take-home recipes; different videos, same material. Cost honestly doubles script-production burden (6 scripts vs 3), not "same tools, better organization." **Decision gates on:** recording runway, editor bandwidth, SharePoint rollout timing.

---

### D3. Video 4 hero selection

**Decision:** Which finalist is the hero tool?

**Options:** (a) SaaS ARR/MRR Waterfall Engine — M effort, visually dramatic, SaaS-native. (b) Revenue Recognition Engine ASC 606 — L effort, maximum CFO gravity. (c) Workbook Dependency Scanner — M effort, opens the "Python does what Excel can't" story.

**Findings this session:** ARR Waterfall recommended as primary hero (M effort, unbeatable visual, speaks iPipeline's language). Rev Rec recommended as its own future Video 5 rather than forced into V4. Dependency Scanner recommended as the **opening demo** regardless of hero choice — it earns the first 90 seconds of audience lean-forward.

---

### D4. xlwings IT feasibility

**Decision:** Can the 2,000-person audience actually install Python + xlwings on their machines?

**Options:** (a) Yes, IT will support — Combo 2 viable. (b) No, locked-down laptops — Combo 2 ships a tool only a small fraction can run.

**Findings this session:** Not verified this session. If Combo 2 is chosen, confirm with IT before building. If install is a barrier, Combo 2 is a demo-only story, not a shippable tool.

---

### D5. PDF generation capability for CFO Pack Pipeline

**Decision:** Does the CFO Pack Assembly Pipeline need to output `.pdf` in addition to `.docx`?

**Options:** (a) `.docx` only — all stdlib/approved packages, no hurdle. (b) `.pdf` required — needs `docx2pdf` (not on approved list) or MS Word COM automation (fragile on recording day).

**Findings this session:** **Assumption used in recommendations: `.docx` only.** If PDF is mandatory, it's a package-approval conversation with IT before building.

---

### D6. Data Contract Checker — YAML vs JSON schema format

**Decision:** What format defines the data contracts?

**Options:** (a) YAML — requires `PyYAML`, not on approved package list. (b) JSON — stdlib only, no new dependencies.

**Findings this session:** **Assumption used: JSON.** Source file (`video4_analysis.md`) specified YAML, but JSON avoids a package approval cycle and preserves the S effort rating.

---

### D7. Root Cause Assistant — effort depends on input data

**Decision:** Does Connor have a historical dataset of prior-resolved reconciliation breaks to match against?

**Options:** (a) Yes — M effort, build matcher. (b) No — L effort, build the matcher AND seed a starter dataset.

**Findings this session:** Not confirmed. Default assumption: M effort. If no historical data exists, rank drops and scope expands.

---

### D8. Production-hygiene items — separate checklist or ignore

**Decision:** Several items from source files are pre-release production tasks (pin package versions, add V3 sample file, update Video 1 hook, update README folder names, add recording-error note to guides, create QR code for closing card, produce 60-sec teaser clip). Claude Code should know these exist.

**Options:** (a) Generate a separate pre-release checklist doc. (b) Roll into `tasks/todo.md`. (c) Ignore — Connor tracks separately.

**Findings this session:** Not generated this session. Flagged in Section D of the full review. **Recommendation:** generate a dedicated `PRE_RELEASE_CHECKLIST.md` before FinalExport ships to SharePoint.

---

## 4. FULL RESULTS

### Full Review Results

#### Section A — Universal Toolkit Additions

*Plug-and-play ideas that work on any coworker's file.*

| # | Idea Name | What It Does | Language | Best Fit | Effort | Why It's Worth Including | Source |
|---|---|---|---|---|---|---|---|
| A1 | **Materiality Classifier** | Tags each row Material ↑ / Material ↓ / Watch / Normal using configurable $ and % thresholds; auto-detects Current vs Prior columns by header text. | VBA | Universal toolkit | S | Instant risk triage on any sheet with zero setup. ⚠️ May already be built — see D1. | BranchIdeasReview A1; CLAUDE_CODE_HANDOFF |
| A2 | **Exception Narrative Generator** | Writes plain-English row commentary based on materiality status; produces CFO-ready wording automatically. | VBA | Universal toolkit | S | Kills manual commentary drafting each close cycle. ⚠️ May already be built — see D1. | BranchIdeasReview A2; CLAUDE_CODE_HANDOFF |
| A3 | **Data Quality Scorecard** | Scores any sheet 0–100 based on blanks, errors, and type mismatches; writes a formatted quality report tab. | VBA | Universal toolkit | S | Creates a "data trust" signal non-technical leaders understand instantly. ⚠️ May already be built — see D1. | BranchIdeasReview A3; CLAUDE_CODE_HANDOFF |
| A4 | **Header Row Auto-Detect** | Scans the top 20 rows and picks the most-likely header row automatically, removing any need for hardcoded row numbers. | VBA | Universal toolkit (foundational helper) | S | Force multiplier — makes every other universal tool truly plug-and-play. ⚠️ May already be built — see D1. | BranchIdeasReview A4; CLAUDE_CODE_HANDOFF |
| A5 | **Quick Row Compare Count** | Fast pre-check that hashes rows across two sheets and returns a mismatch count before running a full compare. | VBA | Universal toolkit | S | Answers "are these files meaningfully different?" in under a second. ⚠️ May already be built — see D1. | BranchIdeasReview A5; CLAUDE_CODE_HANDOFF |
| A6 | **Run Receipt Sheet** | Writes a timestamped execution receipt to a `UTL_RunReceipt` tab on every macro run (what ran, when, by whom, row counts). | VBA | Universal toolkit | S | Instant control evidence and audit traceability out of the box. ⚠️ May already be built — see D1. | BranchIdeasReview A6; CLAUDE_CODE_HANDOFF |
| A7 | **Cover "Show Tools" Button Installer** | One-time macro that adds a branded blue launcher button to the Cover sheet pointing to Command Center. | VBA | Universal toolkit (UX) | S | Removes the "how do I start?" friction for non-technical users. ⚠️ May already be built — see D1. | BranchIdeasReview A7; CLAUDE_CODE_HANDOFF |
| A8 | **Pinned Intelligence Category in Command Center** | Pins Materiality / Narratives / Scorecard tools near the top of the Command Center instead of scrolling past item 29. | VBA | Universal toolkit (UX) | S | High-value tools become visible and discoverable for coworkers. ⚠️ May already be built — see D1. | BranchIdeasReview A8; CLAUDE_CODE_HANDOFF |
| A9 | **Word Report `--talking-points` Flag** | Adds an optional CLI flag to `word_report.py` that auto-builds 3–5 CFO-ready narrative bullets from the input data. | Python | Enhancement to existing `word_report` | S | AI-style output with zero AI calls — speeds exec prep while respecting the no-external-AI constraint. | BranchIdeasReview A10 |
| A10 | **Fiscal Year Startup Warning** | On workbook open, check `FISCAL_YEAR_4` against the current calendar year and show a one-time popup if they don't match. | VBA | Universal toolkit (maintenance) | S | Prevents the "every tab reference is wrong in January" trap for coworkers inheriting the file later. | GitClaudeReply2; GitAgentIdeas2 |
| A11 | **Error-Handling Standardization Pass** | Refactor all VBA modules to use consistent `On Error GoTo ErrorHandler` with a shared cleanup label; remove remaining `On Error Resume Next` silent-failure patterns. | VBA | Refactor (not a new tool) | M | Prevents mid-demo silent failures in front of the CFO. **Highest-leverage code-hygiene move before recording.** | GitClaudeReply2; GitAgentIdeas2 |
| A12 | **Quick Demo Mode Macro** | One button that auto-runs the 5 most impressive features back-to-back (Data Quality → Variance Commentary → Dashboard → PDF Export → Integration Test). | VBA | Universal toolkit (demo helper) | S | Answers "can you show me in 2 minutes?" perfectly — useful long after the videos are recorded. | GitAgentIdeas2 |

#### Section B — Video 4 Candidates

*Python-focused, demo-worthy, not already built.*

| # | Idea Name | What It Does | Language | Best Fit | Effort | Why It's Worth Including | Source |
|---|---|---|---|---|---|---|---|
| B1 | **Exception Triage Engine** | Ranks open exceptions by impact × confidence × recency using config-driven weights; exports a ranked Excel/CSV with top-priority items at the top. | Python | V4 candidate | M | Directly improves analyst workflow every single month-end. | video4_analysis #1; BranchIdeasReview B7; CLAUDE_CODE_HANDOFF |
| B2 | **Control Evidence Pack Generator** | Pulls approved outputs, macro logs, and validation results into a standardized zipped audit bundle with manifest and checksums. | Python | V4 candidate | M | Cuts audit-prep hours directly. | video4_analysis #2; BranchIdeasReview B8; CLAUDE_CODE_HANDOFF |
| B3 | **Finance Data Contract Checker** | Validates incoming files against required columns, data types, and quality rules (JSON-defined, per D6) before anyone reports from them. | Python | V4 candidate (doubles as universal toolkit post-V4) | S | Prevents bad data from polluting reports — a "guardrail" story CFOs love. | video4_analysis #3; BranchIdeasReview B9; CLAUDE_CODE_HANDOFF |
| B4 | **Root Cause Reconciliation Assistant** | Suggests likely cause categories for reconciliation breaks by fuzzy-matching against prior resolved patterns. | Python | V4 candidate | M (per D7) | Demonstrates "intelligence" without external AI APIs. | video4_analysis #4 |
| B5 | **Workbook Dependency Scanner** | Parses formulas and named ranges and outputs a change-impact map ("if you edit cell X, these 14 formulas and 3 charts break"). | Python | V4 candidate (doubles as universal toolkit post-V4) | M | Reduces breakage risk when coworkers edit shared workbooks — classic "Python does what Excel can't" story. | video4_analysis #5; BranchIdeasReview B10; CLAUDE_CODE_HANDOFF |
| B6 | **Narrative Variance Writer** | Converts variance outputs into deterministic, branded commentary drafts using approved templates — no AI calls, all rule-based. | Python | V4 candidate / cookbook recipe | S | Exec-pack-ready output while staying fully compliant with no-external-AI rule. | video4_analysis #6 |
| B7 | **CFO Pack Assembly Pipeline** | Combines approved tables, charts, and commentary into one locked release package (Word; PDF per D5) in a single run. | Python | V4 candidate | M | "One-click board pack" lands hard with CFO/CEO. | video4_analysis #7 |
| B8 | **SaaS ARR/MRR Waterfall Engine** | Converts a subscription roster into Starting ARR → New → Expansion → Contraction → Churn → Ending ARR, plus NRR/GRR and cohort retention. | Python | V4 candidate (**hero**) | M | Finance-native, high executive signal, visually dramatic — ideal hero demo for a SaaS company. | video4_analysis #8 |
| B9 | **Revenue Recognition Engine (ASC 606)** | From contract and billing inputs, produces period recognized revenue, deferred revenue rollforward, commission amortization, and exception tabs. | Python | V4 candidate (**hero**) / potential Video 5 | L | Extremely CFO-relevant and clearly beyond what Excel can do — production-grade output. | video4_analysis #9 |
| B10 | **Close Readiness Score View** | SQL view returning per-entity close-readiness score (0–100) aggregated from failed validations, missing feeds, and late postings. | SQL | V4 companion asset / better as future mini-video | M | One number per entity = CFO-level language. Fits V4 only as a supporting visual. | BranchIdeasReview B6; CLAUDE_CODE_HANDOFF |

#### Section C — Future Ideas (Parked)

| # | Idea Name | What It Does | Language | Effort | Parked Because | Source |
|---|---|---|---|---|---|---|
| C1 | **Allocation Drift Tracker** | Detects silent drift in cost-allocation percentages month-over-month with threshold flags. | SQL | M | Requires historical allocation history not in demo dataset; SQL idea in a Python video cycle. | BranchIdeasReview C1 |
| C2 | **Forecast Backtest Warehouse** | Stores every forecast run with assumptions and realized actuals so accuracy can be measured over time. | SQL | L | Needs multiple close cycles of data before demo-worthy; a year-2 asset. | BranchIdeasReview C2 |
| C3 | **Subledger Completeness Control Matrix** | Checks that all required upstream feeds are present and non-empty before close steps run. | SQL | M | Scope creep for V4 — close-process governance tool, not a Python-demo tool. | BranchIdeasReview C3 |
| C4 | **Workbook-to-Source Reconciliation Mart** | Reconciles workbook aggregates against warehouse source-of-truth tables. | SQL | M | Requires warehouse access coworkers may not have during demo window. | BranchIdeasReview C4 |
| C5 | **Vendor Payment Velocity Baselines** | Flags abnormal vendor payment timing or amount shifts using rolling medians. | SQL | L | AP-focused, not close-focused — distracts from Finance-close storyline. | BranchIdeasReview C5 |
| C6 | **JE Duplicate Ring Detection** | Finds near-duplicate journal entry patterns split across users, days, or entities. | SQL | L | Audit/forensics flavor — better suited to a future "Controls & Audit" video. | BranchIdeasReview C6 |
| C7 | **Close Bottleneck Heatmap Dataset** | Decomposes where close-cycle delays happen by step, entity, and user. | SQL | M | Needs timestamped process-log data the demo file doesn't include. | BranchIdeasReview C7 |
| C8 | **Segregation-of-Duties Audit Pack** | Flags conflicting role/action combinations in the transaction lifecycle. | SQL | M | Requires role/entitlement data not in scope for the public demo. | BranchIdeasReview C8 |
| C9 | **Formula Integrity Fingerprinting** | Hash-checks critical formula zones to catch silent changes over time. | VBA | M | Excellent control, visually boring on camera — park for a "controls" mini-video. | BranchIdeasReview C9 |
| C10 | **Exception Workbench Sheet** | Central Excel tab for assigning, tracking, and closing exceptions with owner and due-date workflow. | VBA | M | Strong post-demo adoption tool, but workflow feature takes time to explain — wrong fit for 5–8 min Python-focused video. | BranchIdeasReview C10 |
| C11 | **Macro Runtime Telemetry Dashboard** | Summarizes run times, error rates, and usage frequency by Command Center action. | VBA | M | Needs weeks of real-usage data before dashboard has anything meaningful to show. | BranchIdeasReview C11 |
| C12 | **Controlled Snapshot Sign-off** | Captures an approved monthly workbook state with checksum and approver metadata. | VBA | M | Close-process governance feature; better introduced once toolkit is in production. | BranchIdeasReview C12 |
| C13 | **"What's New" Sheet** | A tab that logs what changed in the workbook and when, so returning coworkers can see the delta. | VBA | S | Great idea, low camera impact, not V4-relevant; add to post-demo release. | GitAgentIdeas2 |
| C14 | **60-Second Teaser Clip** | Separate short video showing the single most jaw-dropping feature — builds buzz before official release. | Video asset | S | Not a code idea and not V4 itself — park as marketing add-on parallel to main rollout. | GitAgentIdeas2 |
| C15 | **QR Code on Closing Title Card** | Replaces the SharePoint path on the end card with a scannable QR code to boost adoption. | Video asset | S | Pure production polish, not a code idea; park on video-production checklist. | GitAgentIdeas2 |
| C16 | **Pin Exact Package Versions in `requirements.txt`** | Replaces unpinned package lines with exact versions so installs don't silently break later. | Python hygiene | S | Belongs on pre-release checklist, not the curated idea list — but do it before shipping. | GitClaudeReply2; GitAgentIdeas2 |

#### Section D — Skip

| Idea Name | Skip Reason | Source |
|---|---|---|
| Outlook Mail Merge with Attachments | Violates no-Outlook/no-email-automation constraint. | BranchIdeasReview D |
| Calendar Appointment Builder | Violates no-Outlook constraint. | BranchIdeasReview D |
| JIRA Bridge / Weekly Digest | External integration; out of scope for Finance-close storyline. | BranchIdeasReview D |
| Slack Webhook Notifier | External platform dependency. | BranchIdeasReview D |
| Teams Webhook Notifier (threshold alerts) | External platform dependency. | BranchIdeasReview D |
| AWS Cost Optimizer | Not Finance-close focused; wrong audience. | BranchIdeasReview D |
| Customer Churn Risk Scorer (ML) | Requires `scikit-learn` — not on approved package list. | BranchIdeasReview D |
| Support Ticket Triage (ML) | Requires `scikit-learn` — not on approved package list. | BranchIdeasReview D |
| PowerShell IT Admin Automations | Outside approved stack and target audience. | BranchIdeasReview D |
| Power Automate Flows | Different delivery surface; distracts from core story. | BranchIdeasReview D |
| Office Scripts (TypeScript) Flows | Different delivery surface; distracts from core story. | BranchIdeasReview D |
| Zero-Install Workbook Profiler | Already built (`profile_workbook.py`). | BranchIdeasReview A9 |
| Zero-Install Workbook Compare | Already built (`compare_workbooks.py`). | BranchIdeasReview B1 |
| Zero-Install Variance Classifier | Already built (`variance_classifier.py`). | BranchIdeasReview B2 |
| Zero-Install Scenario Runner | Already built (`scenario_runner.py`). | BranchIdeasReview B3 |
| Sheets-to-CSV Batch Export | Already built (`sheets_to_csv.py`). | BranchIdeasReview B4 |
| Executive Summary Builder | Already built (`build_exec_summary.py`). | BranchIdeasReview B5 |
| Sample File for Video 3 | Production task, not an idea — pre-release checklist item. | GitClaudeReply2 |
| Video 1 Opening Hook Edit | Production edit, not a code idea — pre-release checklist item. | GitClaudeReply2 |
| Time-Savings Overlay Made Mandatory | Production edit — pre-release checklist item. | GitClaudeReply2 |
| "If Something Goes Wrong" Recording Note | Production doc edit — recording-day checklist item. | GitClaudeReply2 |
| README / SharePoint Folder Name Reconciliation | Doc hygiene — pre-release checklist item. | GitClaudeReply2 |
| UserForm Import Note in README | Doc hygiene — pre-release checklist item. | GitClaudeReply2 |

---

### Video 4 Finalists

*8 finalists ranked by bang-for-buck. Two candidates from the source 9-list were cut: Narrative Variance Writer (templated mail-merge — kept as cookbook recipe only) and Close Readiness Score View (SQL in a Python video — wrong language fit).*

#### Scored Table

| # | Idea Name | What It Does | Why Perfect for V4 | Best Combo | CFO Wow | Coworker Use | Demo-ability | Effort | Packages | Source |
|:-:|---|---|---|:-:|:-:|:-:|:-:|:-:|---|---|
| 1 | **Workbook Dependency Scanner** | Parses all formulas and named ranges in a workbook and outputs a change-impact map: "if you edit cell B12, these 14 formulas and 3 charts break." | The only finalist that shows Python doing something Excel literally **cannot do**. Nobody in your audience will have seen this before. | 2 | 5 | 4 | 5 | M | openpyxl, pandas, stdlib | video4_analysis #5; BranchIdeasReview B10 |
| 2 | **Exception Triage Engine** | Ranks open exceptions by impact × confidence × recency with config-driven weights; exports a ranked Excel with top-priority items at the top. | Solves a real analyst problem every month-end. Output instantly legible to CFO — one look tells you what to fix first. | 1 | 5 | 5 | 5 | M | pandas, numpy, openpyxl, stdlib | video4_analysis #1; BranchIdeasReview B7 |
| 3 | **Finance Data Contract Checker** | Validates an incoming file against required columns, data types, and quality rules; exports a pass/fail report naming the exact rows and columns that failed. | Fastest-to-build finalist with the most dramatic before/after on camera (red fail → fix → green pass). | 1 | 4 | 5 | 5 | S | pandas, openpyxl, stdlib ⚠️ see D6 | video4_analysis #3; BranchIdeasReview B9 |
| 4 | **SaaS ARR/MRR Waterfall Engine** | Converts a subscription roster into Starting ARR → New → Expansion → Contraction → Churn → Ending ARR, plus NRR/GRR and cohort retention — exports a polished workbook. | SaaS metrics ARE the language of your business. Finance-native, visually dramatic, perfectly tailored to iPipeline. | 3 | 5 | 4 | 5 | M | pandas, openpyxl, matplotlib, stdlib | video4_analysis #8 |
| 5 | **Revenue Recognition Engine (ASC 606)** | From contracts and billings, produces period recognized revenue, deferred revenue rollforward, commission amortization, and exception tabs — in one run. | The most CFO-relevant tool in the pile. Clearly production-grade. Demonstrates Python handling real accounting complexity. | 3 | 5 | 5 | 4 | L | pandas, openpyxl, numpy, stdlib | video4_analysis #9 |
| 6 | **CFO Pack Assembly Pipeline** | Combines approved tables, charts, and commentary into one locked, branded release package (Word + PDF per D5) in a single command. | "One-click board pack" lands directly with CFO/CEO. The output IS the executive deliverable. | 2 | 5 | 3 | 5 | M | pandas, openpyxl, matplotlib, python-docx, stdlib | video4_analysis #7 |
| 7 | **Control Evidence Pack Generator** | Pulls outputs, macro logs, and validation results into a standardized zipped audit bundle with manifest and checksums. | Directly cuts audit-prep hours. Governance-flavored but tangible; coworkers reuse it every close cycle. | 3 | 4 | 4 | 4 | M | pandas, openpyxl, python-docx, stdlib (zipfile, hashlib) | video4_analysis #2; BranchIdeasReview B8 |
| 8 | **Root Cause Reconciliation Assistant** | Suggests likely cause categories for reconciliation breaks by fuzzy-matching against prior resolved patterns; outputs break file with cause tags + confidence. | "Intelligence without AI APIs" — the perfect answer to a nervous audience asking "is this AI?" No. Just good pattern matching. | 1 | 4 | 5 | 4 | M ⚠️ see D7 | pandas, thefuzz, numpy, openpyxl, stdlib | video4_analysis #4 |

#### 30–60 Second Camera Narratives

**1 — Workbook Dependency Scanner.** Open a shared Finance model on screen. Zoom in on one assumption cell — say, a revenue growth rate in B12. Say: *"This cell feeds five tabs, but I can't tell which ones. If I change it, what breaks?"* Switch to the terminal. Run one command. Cut to the output: a branded Excel workbook showing a tree of dependencies — "B12 → Sheet 'Forecast' cell D44 → Sheet 'Board Pack' chart 'Revenue Trend'." Close with: *"Now I can edit with confidence. I know exactly what will recalc — and what to double-check before I hit save."*

**2 — Exception Triage Engine.** Start with a spreadsheet dump of 100 unresolved exceptions — scroll to show it's overwhelming. Say: *"Monday morning. Hundred open items. Where do I start?"* Run the command. Cut to output file: the same 100 rows, sorted — top row is a $2.4M variance with 90% confidence flagged this morning. Bottom rows are $12 pennies from three months ago. Close with: *"This tells Monday-me exactly what to fix first — and my manager can see the same ranking I'm seeing."*

**3 — Finance Data Contract Checker.** On screen, open a broken input file — wrong date format in one column, missing a required column. Say: *"Bad data is how bad reports happen. Let me show you how to stop it at the door."* Run the checker. Cut to a red PASS/FAIL report showing the exact rows and columns that failed, with plain-English reasons. Then fix the file on camera — takes 10 seconds. Rerun. Cut to a green all-pass report. Close with: *"One guardrail. Every file. Before anyone reports from it."*

**4 — SaaS ARR/MRR Waterfall Engine.** Open a raw subscription export — messy, 40 columns, thousands of rows. Say: *"Every Monday morning, our CFO wants to know where ARR moved last week. Today that's a half-day of formulas. Watch this."* Run the command. Cut to a polished branded workbook with three tabs: the ARR waterfall chart, the NRR/GRR summary, the cohort retention grid. Linger on the waterfall chart. Close with: *"Half a day, gone. Plus we now have it for every week, not just the ones we had time for."*

**5 — Revenue Recognition Engine ASC 606.** Say: *"ASC 606 is the accounting rule that governs how SaaS companies recognize revenue. It's complicated. Most companies track it in a workbook nobody but one person understands."* Show inputs — contract file, billing file, commissions file. Run the engine. Cut to output: four tabs — Recognized Revenue by Period, Deferred Revenue Rollforward (with tie-out row), Commission Amortization Schedule, Exceptions. Zoom into the tie-out row to show it balances to the penny. Close with: *"This isn't a toy. This is what a production revenue engine looks like. Built by Finance, in Python, over a weekend."*

**6 — CFO Pack Assembly Pipeline.** Start with three separate files open — approved variance tables, approved charts, approved commentary. Say: *"Every month, someone spends four hours assembling the CFO pack. PDF-ing. Pasting. Re-formatting headers. Fixing fonts. Here's what one command does."* Run it. Cut to the final deliverable — a branded Word document, iPipeline blue headers, Arial, consistent spacing, page numbers, TOC. Scroll through slowly. Close with: *"Same inputs. Four hours, gone."*

**7 — Control Evidence Pack Generator.** Show a messy folder with 40 files — logs, outputs, validation results, cryptic names. Say: *"Audit season. You need to prove every control ran, every validation passed, every number tied. Normally this means three days of screenshotting and emailing."* Run the command. Cut to a clean zip file. Open it — index.html, MANIFEST with checksums, logs organized by control, outputs organized by report. Close with: *"Auditors get one zip. Checksums prove nothing was changed. We get our week back."*

**8 — Root Cause Reconciliation Assistant.** Load a reconciliation break file — 25 unexplained differences. Say: *"Staring at a blank 'Explanation' column every month is demoralizing. What if Python gave you a starting guess?"* Run the assistant. Cut to output: same file, now with a "Likely Cause" column populated for 19 of 25 rows — "Timing difference (87% match to prior patterns)," "FX adjustment (82%)," "Posted to wrong entity (78%)." Close with: *"No AI service. No data leaving the building. Just pattern matching on what your team has already solved before. Starting guesses, not blank pages."*

#### Combo Recommendations

**Combo 1 — Finance Copilot Menu**
1. Exception Triage Engine
2. Finance Data Contract Checker
3. Root Cause Reconciliation Assistant

*Unifying thread:* all three speak to the same analyst persona (Monday 8am, pile of unexplained items). Menu item 1 ranks them, item 2 stops bad new data joining them, item 3 suggests why the rest happened. Coherent workflow.

*Honest risk:* numbered menu looks DOS-era on camera. Dress with iPipeline branding in console header, colored output, clear breaks — or execution undercuts content.

**Combo 2 — Excel Button Edition (xlwings)**
1. Workbook Dependency Scanner
2. CFO Pack Assembly Pipeline
3. Finance Data Contract Checker

*Unifying thread:* each is about Excel files — scanning, assembling from, validating. xlwings's killer feature is "coworker never leaves Excel." These three respect that.

*Honest risk:* see D4 — xlwings needs Python + config on user machines. Great demo, possibly unshippable to 2,000 users.

**Combo 3 — Hero Demo + Cookbook**
- Hero: **SaaS ARR/MRR Waterfall Engine**
- Recipe 1: Finance Data Contract Checker (S)
- Recipe 2: Narrative Variance Writer (S)
- Recipe 3: Control Evidence Pack Generator (M)
- Recipe 4: Exception Triage Engine, simplified version (M)

*Unifying thread:* hero earns 3 min of exec attention; 4 recipes at 30–45 sec each give every viewer a take-home. Cookbook is forgiving — can drop recipes without breaking story.

*Honest risk:* ship-by-recording-day burden is 5 polished scripts, not 1. Most likely combo to slip on tight timeline.

#### Recommended Video Structure

**Split Combo 3 into two videos (see D2):**

- **Video 4A — "Python Shows You What Excel Can't" (6–7 min).** Hero: ARR Waterfall. Opener: Workbook Dependency Scanner as a 90-second "Excel literally can't do this" moment. Two tools, both dramatic, both M-effort.
- **Video 4B — "Your Python Cookbook" (5–6 min).** Four recipes at 60–75 sec each: Data Contract Checker, Narrative Variance Writer, Control Evidence Pack Generator, simplified Exception Triage.

**Rationale:** CFO/CEO want 2–3 tools that change their mind about Finance. Coworkers want take-home recipes. Different audiences, different videos, same material.

**Cost correction:** 6 production-polished scripts vs 3 for single-video Combo 3. Doubles script-production burden; worth it for audience impact but not free.

**Single-video fallback:** Combo 3 as written, but lead with Workbook Dependency Scanner as 60-sec opener before ARR Waterfall hero. Rev Rec parked for future Video 5.

---

## 5. TOP PICKS CONSOLIDATED

**Assumption for ranking:** VBA items A1–A8 are already built (per D1). If not, apply swap table below.

```
[1]  Workbook Dependency Scanner         | Python | M | V4 Finalist + post-demo toolkit | Only tool showing Python doing what Excel literally cannot
[2]  Finance Data Contract Checker       | Python | S | V4 Finalist + post-demo toolkit | Fastest build, most dramatic red→green demo, S effort
[3]  SaaS ARR/MRR Waterfall Engine       | Python | M | V4 Hero                         | Visually unbeatable, SaaS-native to iPipeline
[4]  Exception Triage Engine             | Python | M | V4 Finalist                     | Solves real Monday-morning analyst problem every cycle
[5]  Error-Handling Standardization Pass | VBA    | M | Toolkit refactor                | Pre-demo insurance against silent failure on CFO call
[6]  Word Report --talking-points Flag   | Python | S | Toolkit enhancement             | AI-style output, zero AI calls, S-effort upgrade
[7]  Quick Demo Mode Macro               | VBA    | S | Toolkit demo helper             | Answers "show me in 2 min" forever — lives past recording
[8]  Root Cause Reconciliation Assistant | Python | M | V4 Finalist (Combo 1)           | "Intelligence without AI" answer for nervous audiences
[9]  CFO Pack Assembly Pipeline          | Python | M | V4 Finalist                     | One-click board pack, CFO-native output
[10] Control Evidence Pack Generator     | Python | M | V4 Finalist                     | Audit-prep hours saved, right tone for CFO audience
```

**Swap table if A1–A8 are NOT already built (override):**

| Swap out | Swap in | Why |
|---|---|---|
| [7] Quick Demo Mode Macro | **Materiality Classifier (A1)** — VBA / S | Higher daily value than demo mode — every analyst, every sheet. |
| [8] Root Cause Reconciliation Assistant | **Data Quality Scorecard (A3)** — VBA / S | S vs M effort + visual appeal in-demo. |

**Notable exclusions:**
- Revenue Recognition Engine (ASC 606) — L effort makes it a bang-for-buck outlier. Deserves its own Video 5, not forced into V4.
- Close Readiness Score View — SQL in Python video cycle. Save for future SQL mini-video.
- Narrative Variance Writer — good cookbook recipe, not standalone Top 10 worthy.
