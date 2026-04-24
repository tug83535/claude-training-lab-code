# iPipeline Demo — AI Handoff

**File created:** 2026-04-22
**Purpose:** Self-contained handoff for Claude Code to review and act on curated code ideas for the iPipeline Finance & Accounting 4-video demo series.

---

## 1. WHAT IS THIS FILE

This file contains a full curated code-idea review compiled from 13 research files across 2 passes: (1) an inventory/intake pass flagging file quality and content overlap, and (2) a full curation pass producing Sections A (Universal Toolkit Additions), B (Video 4 Candidates), C (Future/Parked), and D (Skip). 56 curated picks are included across all sections, plus a Top 15 consolidated shortlist at the end.

---

## 2. PROJECT CONTEXT

**User:** Connor Atlee — Finance & Accounting analyst at iPipeline (SaaS for life insurance / financial services). Non-developer; reads VBA/SQL/Python at a working level but relies on AI to author new code.

**Project:** 4-video internal demo series for 2,000+ coworkers + CFO/CEO, showing Finance folks what's possible when you combine Excel + VBA + Python + SQL + AI. All outputs must be plain English, CFO-grade.

**Videos:**
- V1 "What's Possible" — Excel + VBA highlight reel — **RECORDED**
- V2 "Full Demo Walkthrough" — macro-enhanced P&L with 62 automated actions — **RECORDED**
- V3 "Universal Tools" — VBA toolkit on any Excel file — **RECORDED**
- V4 "Python Automation for Finance" — **PLANNING**; must ship a real downloadable tool coworkers can use the Monday after they watch

**Video 4 combos under consideration:**
- **Combo 1** — Tight 5–8 min: ZeroInstall arc only (B4 → B1 → B2 → B3 → B5)
- **Combo 2** — Extended 10–12 min: ZeroInstall arc + one advanced demo (B7 Exception Triage *or* B8 Evidence Pack)
- **Combo 3** — Split: V4a (ZeroInstall) + V4b (Advanced Automation with B6 / B7 / B8)

**Hard constraints:**
- No external AI API calls (OpenAI, Claude, Gemini, etc.)
- No Outlook / email automation
- No Windows Task Scheduler dependencies
- Non-developer audience — every feature must be explainable to someone with zero coding background
- Plug-and-play preferred (no hardcoded sheet names)
- iPipeline branding: Blue `#0B4779`, Navy `#112E51`, Arial fonts only

**Approved Python packages (exhaustive):** pandas, openpyxl, pdfplumber, python-docx, thefuzz, numpy, matplotlib, xlwings, and Python stdlib. Nothing else.

**Already built — do not duplicate:**
- **23 VBA modules / ~140 universal tools:** data sanitizer, compare, consolidate, highlights, pivot tools, tab organizer, column ops, sheet tools, comments, validation builder, lookup builder, command center, exec brief, finance-specific tools, audit tools, branding, and more
- **22+ Python scripts:** aging report, bank reconciler, compare files, forecast rollforward, fuzzy lookup, pdf extractor, variance analysis, variance decomposition, clean data, consolidate files, multi-file consolidator, date unifier, two-file reconciler, SQL query tool, word report, batch processor, regex extractor, unpivot, pnl forecast, pnl dashboard
- **7 new stdlib-only zero-install Python scripts** (covered under BranchIdeasReview April19update branch — these feed Video 4)
- **4 SQL scripts:** staging, transformations, validations, enhancements

---

## 3. OPEN DECISIONS

| # | Decision | Options | Findings from this session |
|---|---|---|---|
| 1 | Video 4 structure | Combo 1 (tight) / Combo 2 (extended) / Combo 3 (split V4a + V4b) | User stated openness to going longer or splitting if content warrants. ZeroInstall arc (B1–B5) is the strongest foundation; B7/B8 are plausible advanced demos. Combo 3 recommended if B7 or B8 are built by recording; otherwise Combo 1. |
| 2 | Does Connor's team own revenue recognition? | Yes → keep B13 Revenue Recognition Simulator in Section B / No → demote to Section C | iPipeline is a life-insurance SaaS so rev-rec is board-level, but rev-rec may sit in corporate accounting rather than Connor's F&A team. **Unresolved** — needs Connor's confirmation. |
| 3 | Section A trim to 12 or keep 14 | Trim (drop A11 "What's New" sheet + A14 Module Text Exporter) / Keep all 14 | Both are S-effort and both have value but neither made the Top 10. Low-stakes call. |
| 4 | "Already Covered" confirmations in Section D.4 | Confirm all / Override some | D.4 entries are inferred from script names only; actual contents not inspected. Specifically flagged: Monthly Billing Reconciler, Duplicate Customer Detector, AR/AP Aging, PDF Text Extraction, Basic VBA UserForm, Excel Dashboard Refresher. Needs a 10-minute spot check by Connor or Claude Code. |
| 5 | Section C ordering | By language (current) / By priority | CodexCodeIdeas.md has a phased roadmap (Phase 1 / 2 / 3) that can be adopted directly if priority ordering is preferred. |
| 6 | Next follow-up action | (a) apply flagged fixes / (b) scaffold code for a Top 10 item / (c) build V4 storyboard / (d) overlap audit with actual script contents / (e) package as branded Word/PDF | Not chosen in this session. |

---

## 4. FULL RESULTS

### Inventory Pass Results

13 files analyzed in `/mnt/project/`.

| # | Filename | File Type (actual) | One-sentence summary |
|---|---|---|---|
| 1 | `200_tools_catalog.md` | Markdown, 256 lines | Flat catalog of ~200 open-source Python and SQL libraries grouped by function; purely tool inventory, no iPipeline context. |
| 2 | `BranchIdeasReview_April2026.md` | Markdown, 111 lines | iPipeline-specific idea review dated 2026-04-22 organized into Section A/B/C/D tables with per-idea source citations back to specific `.bas`/`.py` files. |
| 3 | `CODE_CATALOG.md` | Markdown, 169 lines | Single-document reference for `MobileCLDCode/` folder (42 code files); orientation doc for cold-start Claude sessions. |
| 4 | `CodexCodeIdeas.md` | Markdown, 368 lines | Structured idea backlog from `codexreview2/` folder organized SQL/Python/VBA/Cross-stack with business outcome, use case, approach, and priority phase per entry. |
| 5 | `deep-research-report.md` | Markdown, 298 lines | Deep-research report framed around three-branch synthesis; delivers Python dedup harness plus external toolkit recommendation (Airflow, Great Expectations, RapidFuzz, Pydantic, dbt-utils) and 90-day roadmap. |
| 6 | `deep-research-report_2.md` | Markdown, 230 lines | Second deep-research report with similar harness but different external stack (sqlglot, great_expectations, pandera, dedupe, unstructured) and "shadow revenue journal" concept. |
| 7 | `Executive_Automation_Catalog_2.docx` | **Plain text, not a real Word doc** — 390 lines | Long-form enterprise catalog with 20-item Works Cited; enterprise-architect framing rather than iPipeline-specific. |
| 8 | `Executive_Automation_Catalog.md` | Markdown, 266 lines | Structural template: dedup methodology, naming conventions (SQL-RO-001 etc.), placeholder SQL/Python/VBA examples, open-source toolkit list, five-entry future-state roadmap. |
| 9 | `Executive_Automation_Catalog__Master_Reference_for___.docx` | **Plain text, not a real Word doc** — 171 lines | Shorter master reference: Compare-and-Classify SQL, Pydantic/Instructor gatekeeper, VBA-Web bridge, Splink/dbt-audit-helper/LangExtract toolkit. |
| 10 | `Exec_auto_master_fixed.pdf` | **ZIP archive, not real PDF** — 6 page-images + 6 text transcripts + manifest.json | Duplicate of file 9 in different container (rendered page images plus extracted text). |
| 11 | `GitAgentIdeas2.md` | Markdown, 128 lines | GitHub-Copilot-style review of `April19update/FinalExport` branch praising what's done and listing improvement recommendations for videos, Python, VBA, SQL. |
| 12 | `report.md` | Markdown, 356 lines | Reconstructed enterprise automation synthesis (branches inaccessible): SQL triggers, Python fuzzy matching and LLM extraction, VBA-Web and OLE Automation, next-gen strategy. |
| 13 | `report_extended.md` | Markdown, 336 lines | Expanded version of file 12 with concrete code templates per business function and fuller next-gen proposals section. |

**Inventory flags:**
- **3 extension mismatches:** `Exec_auto_master_fixed.pdf` is a ZIP; both `.docx` files are plain-text. All still readable.
- **Near-duplicates:** Files #7, #9, #10 cover overlapping enterprise-architect content. Together contributed ~4 unique ideas.
- **Two independent analyst opinions:** Files #5 and #6 tackle the same three-branch problem with different stack recommendations.
- **No truncation detected on any file.**
- **Content buckets:** (a) iPipeline-specific, grounded in code: #2, #3, #4, #11. (b) Generic enterprise essays, branches never accessed: #5, #6, #7, #8, #9, #10, #12, #13. (c) Pure tool inventory: #1.

---

### Video 4 Finalists

Distilled from Section B below. Recommended sequencing if Video 4 is recorded as a single 5–8 min video (Combo 1):

| Order | Item | Runtime slot | On-screen action |
|---|---|---|---|
| 1 | **B4 Sheets-to-CSV Batch Export** | 0:30 – 1:30 | Open a workbook, run script, show folder of CSVs. Establishes the "Excel → Python bridge." |
| 2 | **B1 Zero-Install Workbook Compare** | 1:30 – 3:00 | Compare two workbooks, show diff CSV. No pip install used — emphasize on camera. |
| 3 | **B2 Zero-Install Variance Classifier** | 3:00 – 4:30 | Label rows vs baseline, show Over/Under/On-target output. |
| 4 | **B3 Zero-Install Scenario Runner** | 4:30 – 6:00 | Apply -10% / +5% / +10% shocks to a metric column, show three scenario CSVs. This is the "wow" moment. |
| 5 | **B5 Executive Summary Builder** | 6:00 – 7:30 | Fold prior CSVs into Markdown exec summary. Ends the arc on a leadership-ready artifact. |

**If Combo 2 or 3 (extended / split):** Follow the above with **B7 Exception Triage Engine** or **B8 Control Evidence Pack Generator** as the "one advanced demo." B6 (Close Readiness Score View) is the strongest SQL moment if any SQL appears in V4.

**Flagship emphasis:** B3 or B5 are the strongest CFO demo beats — both end on visible business output.

---

### Best Ideas Curation

**Methodology:**
1. Used `BranchIdeasReview_April2026.md` as curated baseline (iPipeline-specific, pre-structured A/B/C/D).
2. Augmented with unique items from `CodexCodeIdeas.md`, `Executive_Automation_Catalog.md`, `GitAgentIdeas2.md`, `report.md`, `report_extended.md`.
3. Filtered against hard constraints — items that fail land in Section D.
4. Merged duplicates across sources into single entries with combined citations.
5. Discarded: files #7, #9, #10 mined only for unique ideas the `.md` files didn't cover; file #1 (library list) discarded entirely.

**Notation:** `[April19update]` = entries pulled directly from BranchIdeasReview. `[codexreview2]` = `CodexCodeIdeas.md`. `[ExecCatalog]` = `Executive_Automation_Catalog.md`. `[GitAgent]` = `GitAgentIdeas2.md`. `[report]` / `[report_ext]` = the two reports.

#### Section A — Universal Toolkit Additions

> Plug-and-play on any coworker file. Goes into `modUTL_*` or `UniversalToolkit/python/`.

| # | Idea | What it does | Language | Effort | Why include | Source |
|---|---|---|---|---|---|---|
| A1 | **Materiality Classifier** | Tags each row Material↑ / Material↓ / Watch / Normal using configurable $ and % thresholds; auto-detects Current vs Prior columns from headers. | VBA | S | Instant risk triage on any worksheet — zero setup. | `BranchIdeasReview_April2026.md` (A1) → `FinalExport/UniversalToolkit/vba/modUTL_Intelligence.bas`; also `CodexCodeIdeas.md` (PY-02) |
| A2 | **Exception Narrative Generator** | Reads Materiality Status column and writes a plain-English row-level narrative ("AR rose $412K, 18% above plan — investigate"). | VBA | S | CFO-ready commentary without manual drafting. | `BranchIdeasReview_April2026.md` (A2); concept echoed in `CodexCodeIdeas.md` (PY-02 Narrative Variance Writer) |
| A3 | **Data Quality Scorecard** | Scores any sheet 0–100 from blank %, error cells, inconsistent types; writes a formatted QualityReport tab. | VBA | S | Creates a "data trust" signal leaders grasp instantly — powerful live-demo moment. | `BranchIdeasReview_April2026.md` (A3); `report_extended.md` (Missing-Value Monitor concept) |
| A4 | **Header Row Auto-Detect** | Scans top rows and picks the real header row by detecting text vs numeric patterns — removes hardcoded `Row 1` assumptions. | VBA | S | Foundational — makes every other UT tool truly plug-and-play. | `BranchIdeasReview_April2026.md` (A4) → `modUTL_Core.bas` |
| A5 | **Quick Row Compare Count** | Hashes rows and returns mismatch count in seconds before running full compare. | VBA | S | Answers "are these files meaningfully different?" in one click. | `BranchIdeasReview_April2026.md` (A5) → `modUTL_Compare.bas` |
| A6 | **Run Receipt Sheet** | Writes a timestamped `UTL_RunReceipt` tab on every macro run (who / when / inputs / row counts). | VBA | S | Automatic audit evidence — SOX-friendly by default. | `BranchIdeasReview_April2026.md` (A6) → `modUTL_Audit.bas`; reinforced by `CodexCodeIdeas.md` (VBA-08 Macro Telemetry) |
| A7 | **"Show Tools" Launcher Button Installer** | One-time macro adds a branded blue button to the Cover sheet that jumps to Command Center. | VBA | S | Removes "how do I start?" friction for 2,000 non-technical coworkers. | `BranchIdeasReview_April2026.md` (A7) → `modUTL_CommandCenter.bas` |
| A8 | **Zero-Install Workbook Profiler** | Inventories sheets, ranges, VBA presence using stdlib only — no pip install required. | Python | S | Runs on locked-down laptops; a strong "first Python" moment. | `BranchIdeasReview_April2026.md` (A9) → `UniversalToolkit/python/ZeroInstall/profile_workbook.py` |
| A9 | **Word Report Talking Points** | Adds `--talking-points` flag to existing `word_report.py` that auto-generates 3–5 CFO narrative bullets from variance data. | Python | S | "AI-style" output with zero AI calls — perfect CFO demo story. | `BranchIdeasReview_April2026.md` (A10) → `word_report.py` |
| A10 | **Quick Demo Mode Macro** | Single button that runs the 5 most impressive features back-to-back (Data Quality → Narratives → Dashboard → PDF → Integration). | VBA | S | Answers "show me in 2 minutes" — a demo-disarming tool. | `GitAgentIdeas2.md` (New Ideas #1) |
| A11 | **"What's New" Change Log Sheet** | Dedicated tab that logs what changed in the workbook and when; updated by any macro that writes output. | VBA | S | Coworkers who open the file six months later know what's different. | `GitAgentIdeas2.md` (New Ideas #2) |
| A12 | **Template Enforcer** | Validates a submitted file matches required template (expected tabs, named ranges, column headers); writes compliance report. | VBA | S | Kills the "why is this file broken?" support loop. | `Executive_Automation_Catalog.md` (VBA-DI-01); echoed in `CodexCodeIdeas.md` (VBA-09 Workbook Policy Validator) |
| A13 | **Fiscal Year Sanity Check** | On workbook open, compares `FISCAL_YEAR` constant against current year and flashes warning if mismatched. | VBA | S | Prevents "it's January and every tab reference is wrong" disaster. | `GitAgentIdeas2.md` (VBA Recommendations #1) |
| A14 | **VBA Module Text Exporter** | Exports all `.bas`/`.cls` modules to a text folder so Git can track changes. | VBA | S | Gives toolkit real version control without IT involvement. | `Executive_Automation_Catalog.md` (VBA-RPT-03) |

**Assumption:** Section A items sourced from `April19update` branch treated as new additions to mainline, not already-merged code. `BranchIdeasReview` framing supports this.

#### Section B — Video 4 Candidates

> For the "Python Automation for Finance" video. Structure assumes potential split into 2 videos if content runs long.

| # | Idea | What it does | Language | Effort | Overlap with existing? | Why include | Source |
|---|---|---|---|---|---|---|---|
| B1 | **Zero-Install Workbook Compare** | Compares two workbooks row-by-row and exports diffs to CSV using only stdlib. | Python | S | Overlaps `compare_files` / two-file reconciler — **differentiated** because stdlib-only (no pandas/openpyxl), runs on locked laptops. | Strong "no-install automation" hook — demonstrates "you don't even need to install Python packages." | `BranchIdeasReview_April2026.md` (B1) → `ZeroInstall/compare_workbooks.py` |
| B2 | **Zero-Install Variance Classifier** | Labels rows Over / Under / On-target vs baseline using rules only — no ML. | Python | S | Overlaps `variance_analysis` — this version rules-based + stdlib only. | Easy-to-explain risk labeling with zero setup. | `BranchIdeasReview_April2026.md` (B2) → `ZeroInstall/variance_classifier.py` |
| B3 | **Zero-Install Scenario Runner** | Applies % shocks (-10%, +5%, +10%) to a metric column and exports each scenario as its own CSV. | Python | S | None | Real what-if automation; CFO/CEO will instantly see value. | `BranchIdeasReview_April2026.md` (B3) → `ZeroInstall/scenario_runner.py` |
| B4 | **Sheets-to-CSV Batch Export** | Exports every sheet in a workbook to its own CSV in one step. | Python | S | None | The "bridge" step from Excel-world to Python pipelines — relatable for Finance folks. | `BranchIdeasReview_April2026.md` (B4) → `ZeroInstall/sheets_to_csv.py` |
| B5 | **Executive Summary Builder** | Reads a folder of CSV outputs and builds a Markdown exec summary with stats and highlights. | Python | S | None | Turns raw data into leadership-ready narrative fast — closes Video 4 on a high note. | `BranchIdeasReview_April2026.md` (B5) → `ZeroInstall/build_exec_summary.py` |
| B6 | **Close Readiness Score View** | SQL view returning 0–100 readiness score per entity per day — weighted on failed validations, missing feeds, late postings. | SQL | M | None | CFO-level language in one number. Highest-ROI SQL idea in entire backlog. | `BranchIdeasReview_April2026.md` (B6); `CodexCodeIdeas.md` (SQL-02) — **appears in both sources, Phase 1 priority.** |
| B7 | **Exception Triage Engine** | Ranks exceptions by `impact × confidence × recency` using YAML-configurable weights; outputs ranked CSV. | Python | M | None | Every analyst's workflow improves immediately — they work highest-value break first. | `BranchIdeasReview_April2026.md` (B7); `CodexCodeIdeas.md` (PY-01) — **Phase 1 build order #2.** |
| B8 | **Control Evidence Pack Generator** | Reads macro run logs + validation outputs and zips them into timestamped audit evidence bundle with a manifest.json. | Python | M | None | Directly cuts audit-prep hours — exactly the language the CFO cares about. | `BranchIdeasReview_April2026.md` (B8); `CodexCodeIdeas.md` (PY-07) — **Phase 1 build order #5.** |
| B9 | **Finance Data Contract Checker** | Validates incoming CSVs against YAML schema (column names, types, ranges, non-null) before downstream use. | Python | M | Adjacent to `clean_data` but declarative, not procedural. | Prevents bad data quietly polluting reports and forecasts. | `BranchIdeasReview_April2026.md` (B9); `CodexCodeIdeas.md` (PY-03) |
| B10 | **Workbook Dependency Scanner** | Uses `openpyxl` to parse formulas and named ranges, then outputs JSON/HTML impact graph showing what breaks if you change cell X. | Python | M | None | Reduces change-breakage risk on shared workbooks — concrete SOX value. | `BranchIdeasReview_April2026.md` (B10); `CodexCodeIdeas.md` (PY-08) |
| B11 | **Root-Cause Reconciliation Assistant** | When recon breaks, looks up break signature against historical break/resolution log and proposes likely causes. | Python | M | Adjacent to `two_file_reconciler` — extends it. | Shortens time-to-resolution without AI — deterministic pattern matching only. | `CodexCodeIdeas.md` (PY-05) |
| B12 | **CFO Pack Assembly Pipeline** | Compiles approved charts, tables, and commentary into single release artifact (Word + Excel), locking content once approved. | Python | M | Adjacent to `word_report` and `pnl_dashboard` — this is the assembly layer above them. | Consistent monthly deliverables with release tagging — ends "we re-invent the CFO pack every month" problem. | `CodexCodeIdeas.md` (PY-09) |
| B13 | **Revenue Recognition Simulator** | Runs what-if simulations on revenue schedules under different recognition policies (useful in SaaS/life insurance context). | Python | M | None | Directly speaks to iPipeline's industry — insurance revenue recognition is a board-level topic. **See Open Decision #2.** | `Executive_Automation_Catalog.md` (PY-RO-02); `report_extended.md` (SaaS rev rec code block) |
| B14 | **Narrative Variance Writer (Python edition)** | Complements A2 — produces standalone Word doc of deterministic template-based variance commentary per segment, with governance checks before output. | Python | M | Complements A2 (A2 writes inline to sheet; B14 writes standalone deliverable). | Auditable language with no AI hallucination risk — same outcome, different delivery surface. | `CodexCodeIdeas.md` (PY-02) |

#### Section C — Future Ideas (Parked)

> Real business value but post-demo: either requires warehouse/infra access, pushes against time budget, or large enough to deserve its own workstream.

| # | Idea | What it does | Language | Effort | Why parked | Source |
|---|---|---|---|---|---|---|
| C1 | **Close Readiness Score View → Exception Severity Table** (SQL mart) | Extended version of B6 that also materializes companion `close_exceptions` table with severity tiers, drivers, owners. | SQL | M | Base view (B6) is in scope; extended mart is Phase 2 warehouse work. | `BranchIdeasReview_April2026.md` (B6 extension); `CodexCodeIdeas.md` (SQL-02 Phase 1 deliverable) |
| C2 | **Allocation Drift Tracker** | Monthly delta view of cost allocation percentages with tolerance bands; forces reason-code when drift exceeds threshold. | SQL | M | Requires allocation reference data centralized in warehouse first. | `BranchIdeasReview_April2026.md` (C1); `CodexCodeIdeas.md` (SQL-04) |
| C3 | **Forecast Backtest Warehouse** | Three tables (`forecast_run`, `forecast_assumption`, `forecast_actual`) that let you measure MAPE and bias over time. | SQL | L | Needs full close cycle of runs to be meaningful; ship after 3–6 months of captured data. | `BranchIdeasReview_April2026.md` (C2); `CodexCodeIdeas.md` (SQL-08) |
| C4 | **Subledger Completeness Control Matrix** | Control table with expected feed times and row-count bounds; blocks close steps until all upstream feeds present. | SQL | M | Needs SLA/timing agreements documented before control enforceable. | `BranchIdeasReview_April2026.md` (C3); `CodexCodeIdeas.md` (SQL-06) |
| C5 | **Workbook-to-Source Reconciliation Mart** | Standardized recon tables + variance reason taxonomy comparing workbook aggregates against warehouse truth. | SQL | M | Requires owner taxonomy agreement across teams — governance work. | `BranchIdeasReview_April2026.md` (C4); `CodexCodeIdeas.md` (SQL-05) |
| C6 | **Vendor Payment Velocity Baselines** | Rolling-median + MAD/z-score thresholds per vendor cohort to flag abnormal timing or amount shifts. | SQL | L | AP data scope and fraud-program alignment needed first. | `BranchIdeasReview_April2026.md` (C5); `CodexCodeIdeas.md` (SQL-03) |
| C7 | **Journal Entry Duplicate Ring Detection** | Graph-like grouping on amount/date/vendor/account to catch near-duplicate JE patterns split across users, days, or entities. | SQL | L | Sensitive — needs audit/internal-controls sponsor before rollout. | `BranchIdeasReview_April2026.md` (C6); `CodexCodeIdeas.md` (SQL-01) |
| C8 | **Close Bottleneck Heatmap Dataset** | Decomposes close-cycle lag by step × entity × user using event timestamps; output feeds a dashboard. | SQL | M | Requires consistent close-step event instrumentation across tools. | `BranchIdeasReview_April2026.md` (C7); `CodexCodeIdeas.md` (SQL-10) |
| C9 | **Segregation-of-Duties Audit Query Pack** | Role-action matrix joins + exception materialized views to flag conflicting role/action combos. | SQL | M | Needs role mapping signed off by IT/Controls before exception list credible. | `BranchIdeasReview_April2026.md` (C8); `CodexCodeIdeas.md` (SQL-09) |
| C10 | **Policy-as-Code Rule Engine Tables** | Metadata-driven rule catalog + dynamic execution procedure — finance policy checks maintained without editing SQL. | SQL | L | Powerful but abstract; needs a ruleset of real policies to justify framework. | `CodexCodeIdeas.md` (SQL-07) |
| C11 | **Deferred Revenue Waterfall Builder** | SQL generating monthly revenue recognition schedule from multi-year insurance/SaaS contracts. | SQL | M | iPipeline-relevant, but needs contract data model first. | `Executive_Automation_Catalog.md` (SQL-RO-02); `report_extended.md` (SaaS Rev Rec pattern) |
| C12 | **Cohort Retention Matrix Generator** | Produces cohort retention table by signup month × months-since-start for carrier/advisor segments. | SQL | M | Valuable but CS-owned, not Finance-owned. | `Executive_Automation_Catalog.md` (SQL-RPT-02) |
| C13 | **Idempotent Replay Ledger** | Tracks processed events from upstream systems to prevent double-processing if feed re-sent. | SQL | M | Needs eventing architecture — data engineering project. | `Executive_Automation_Catalog.md` (SQL-DI-04) |
| C14 | **Exception Workbench Sheet** | One unified Excel tab for all exceptions with owner / due-date / resolution columns; imports from B7 output. | VBA + sheet | M | The hub for B7 (Triage Engine) output — build once B7 lands and is proven. | `BranchIdeasReview_April2026.md` (C10); `CodexCodeIdeas.md` (VBA-04 Phase 1 deliverable) |
| C15 | **Formula Integrity Fingerprinting** | Hashes formulas in protected/critical ranges at baseline; compares on open or on demand to catch silent changes. | VBA | M | SOX-friendly but needs agreed "critical ranges" definition per workbook type. | `BranchIdeasReview_April2026.md` (C9); `CodexCodeIdeas.md` (VBA-02) |
| C16 | **Macro Runtime Telemetry Dashboard** | Summarizes runtime, error rates, usage frequency from existing `VBA_AuditLog` sheet into KPI dashboard tab. | VBA | M | Depends on audit log accumulating enough real runs — ship ~30 days post-launch. | `BranchIdeasReview_April2026.md` (C11); `CodexCodeIdeas.md` (VBA-08 Phase 1 deliverable) |
| C17 | **Controlled Snapshot Sign-off** | Locks workbook state at monthly sign-off with checksum of key ranges + approver name/timestamp in log sheet. | VBA | M | Post-launch; needs stable "final" workbook definition before locking meaningful. | `BranchIdeasReview_April2026.md` (C12); `CodexCodeIdeas.md` (VBA-07) |
| C18 | **Intelligent Rollforward Assistant** | Rolls month tabs forward with preflight checks + staged apply + undo if mapping validation fails. | VBA | M | Useful but risk-intensive — needs rollback story bulletproofed before coworker use. | `CodexCodeIdeas.md` (VBA-03) |
| C19 | **Guided Adjustment Wizard (UserForm)** | Step-by-step Excel UserForm walking finance users through reviewing and approving billing variances. | VBA | M | High UX polish needed; belongs to dedicated "last-mile UX" workstream post-demo. | `Executive_Automation_Catalog.md` (VBA-RO-01) |
| C20 | **Controlled Action Approvals** | Manager PIN / approval gate before high-impact macros execute, with signed action log. | VBA | L | Governance-heavy — requires approval policy and PIN management process. | `CodexCodeIdeas.md` (VBA-01) |
| C21 | **Data Entry Fraud Pattern Flags** | Event log of manual cell edits + rule windows scoring suspicious patterns (off-hours edits, threshold-adjacent overrides). | VBA | M | Sensitive — needs HR/Controls sponsor; starts political fast. | `CodexCodeIdeas.md` (VBA-10) |
| C22 | **Forecast Ensemble Manager** | Combines multiple forecast models via backtest-based weighting + champion/challenger registry. | Python | L | Meaningful only after C3 (Backtest Warehouse) has data; natural Phase 2 follow-up. | `CodexCodeIdeas.md` (PY-04) |

**Assumption:** C14 and C16 placed in Section C rather than Section B despite being Phase 1 in `CodexCodeIdeas.md` — both require upstream work to land first (C14 depends on B7's output format being stable; C16 depends on ~30 days of audit log data). Demo Video 4 ships sooner than both. Open for override.

#### Section D — Skip (and why)

**D.1 — Violates "no external AI API"**

| Idea | Why skip | Source |
|---|---|---|
| **LLM Contract Parser** (Python service OCRs PDFs then calls LLM for clauses) | Violates "no external AI API." Safer equivalent: existing `pdf_extractor` + `regex_extractor` + B14 Narrative Variance Writer for deterministic templated output. | `report_extended.md` Part 3; `Executive_Automation_Catalog.md` (3.2) |
| **LLM-Driven Unstructured-to-Structured ETL** (Pydantic + Instructor + OpenAI) | Violates "no external AI API." Constraint explicit. | `report.md` (Python); `Executive_Automation_Catalog__Master_Reference_for___.docx` (1.2); `Exec_auto_master_fixed.pdf` |

**D.2 — Uses disallowed Python packages**

| Idea | Why skip | Source |
|---|---|---|
| **Close Calendar Risk Predictor** (sklearn/lightgbm) | Requires `scikit-learn` — not approved. B6 Close Readiness Score View delivers same outcome deterministically. | `CodexCodeIdeas.md` (PY-06) |
| **Data Drift Monitor Service** (scipy/statsmodels PSI/KS tests) | Requires `scipy`/`statsmodels` — not approved. Pandas-only rolling-mean alternative possible but scoped for unavailable packages. | `CodexCodeIdeas.md` (PY-10) |
| **ML Churn Risk Scorer / Support Ticket Triage** | Requires sklearn — not approved, plus off-scope for Finance demo. | `BranchIdeasReview_April2026.md` (Section D); `CODE_CATALOG.md` (MobileCLDCode) |
| **SARIMA/Prophet Time-Series Forecasting Pipeline** | Requires `statsmodels` or `prophet` — not approved. Existing `pnl_forecast.py` covers forecasting within approved packages. | `report.md` (Time-Series Forecasting) |
| **Automated Slide Generator** (python-pptx) | `python-pptx` not approved. Word export via `python-docx` + `word_report.py` delivers executive artifact need. | `report_extended.md` (Reporting & Presentation) |
| **Great Expectations Data Validator** | `great_expectations` not approved. B9 Finance Data Contract Checker delivers same outcome using pandas + YAML. | `report_extended.md`; `deep-research-report.md`; `deep-research-report_2.md` |
| **Headless RPA** (pyautogui / selenium / pywinauto) | Disallowed packages. Non-developer coworkers would struggle to debug anyway. | `report.md` (Cross-Platform RPA) |
| **Anomaly Detection with Isolation Forest / Prophet** | Requires sklearn or prophet. Rules-based variance classifier (B2) covers demo-level need. | `report_extended.md` (Part 3 item 5) |

**D.3 — Violates "no Outlook / email" or "no Task Scheduler"**

| Idea | Why skip | Source |
|---|---|---|
| **Outlook Mail Merge with Attachments / Calendar Appointment Builder** | Explicit constraint violation. | `BranchIdeasReview_April2026.md` (Section D); `CODE_CATALOG.md` |
| **Slack / Teams / Email Distribution Bots** | Violates email constraint spirit; external messaging dependencies unsuited for demo audience. | `BranchIdeasReview_April2026.md` (Section D); `Executive_Automation_Catalog.md` (PY-RPT-02) |
| **Airflow-scheduled Monthly Billing Pipeline** | Violates "no Task Scheduler" and adds infrastructure coworkers cannot replicate. | `report_extended.md` (Part 2 Airflow example) |
| **Lightweight Internal API for Exception Status** (Flask/FastAPI) | Adds always-on server infrastructure — outside "runnable by a Finance person." | `CodexCodeIdeas.md` (OA-03) |
| **GitHub Actions Validation Bundle** | CI/CD pipeline work — different audience. Good for engineering team, not demo. | `CodexCodeIdeas.md` (OA-05) |

**D.4 — Already covered by existing code**

| Idea | What you already have | Source of the duplicate |
|---|---|---|
| **Monthly Billing Reconciler** (pandas + fuzzywuzzy) | Existing `two_file_reconciler.py` + `fuzzy_lookup.py` combo covers this. | `report_extended.md` (Revenue Ops Python example) |
| **AR/AP Aging Analysis script** | Existing `aging_report.py`. | `report_extended.md` |
| **Duplicate Customer Detector** (rapidfuzz) | Existing `fuzzy_lookup.py` + `clean_data.py`. | `report_extended.md`; `Executive_Automation_Catalog.md` |
| **Basic VBA UserForm with input validation** | Existing `modUTL_*` validation_builder and lookup_builder modules. | `report.md` (Advanced UserForm example) |
| **Excel Dashboard Refresher** (RefreshAll macro) | Too trivial for demo — likely already wired into Command Center refresh button. | `report_extended.md` |
| **PDF Text Extraction** | Existing `pdf_extractor.py` (via `pdfplumber`). | `report_extended.md` |

**D.5 — Off-scope for Finance & Accounting demo**

| Idea | Why skip | Source |
|---|---|---|
| **CRM/Sales Funnel VBA Categorizer** ("Lead"→"S1-Top" etc.) | Sales/CRM flavor, not Finance. Also too trivial. | `Executive_Automation_Catalog_2.docx` (VBA Sales Funnel) |
| **Office-to-Mainframe Integration** (SAP GUI / AS400 COM bridging) | Esoteric, audience-confusing, high-risk live demo. | `report.md` (VBA Legacy Bridge) |
| **PowerShell IT Admin Automations** (AD inactive users, SSL cert monitor) | Outside approved stack. IT ops, not Finance. | `CODE_CATALOG.md`; `BranchIdeasReview_April2026.md` (Section D) |
| **.NET / VSTO Add-In Rewrite** | Outside approved stack; requires IT deployment — defeats the "built by Finance" story. | `CodexCodeIdeas.md` (OA-02) |
| **AWS Cost Optimizer / JIRA Weekly Digest** | Neither Finance-close focused. | `CODE_CATALOG.md`; `BranchIdeasReview_April2026.md` (Section D) |
| **Office Scripts / Power Automate Close Trigger** | Different delivery surface; dilutes Excel+VBA+Python+SQL focus. | `CodexCodeIdeas.md` (OA-01); `CODE_CATALOG.md` |

#### Top 10 from Curation

**Scoring rubric:** value per hour of build effort, weighted toward (a) visible on-camera impact for CFO/CEO, (b) re-use frequency by 2,000 coworkers after launch, (c) foundation effect — does it make other tools better?

| Rank | Idea | Section | Effort | Why it wins |
|---|---|---|---|---|
| 1 | **Close Readiness Score View** | B6 | M | One number per entity = exact language CFO speaks. Highest-leverage SQL item in entire backlog. |
| 2 | **Data Quality Scorecard** | A3 | S | Any sheet → formatted 0–100 grade instantly. Strongest single "wow" demo moment. |
| 3 | **Exception Triage Engine** | B7 | M | Every analyst's workflow improves every close cycle. Config-driven weights = retuning without code changes. |
| 4 | **Materiality Classifier** | A1 | S | Row-level risk tiering automatically. Pairs with A2+A3 for 30-second triage narrative. |
| 5 | **Word Report Talking Points** | A9 | S | "AI-style" narrative output with zero AI calls. CFO demo-wrecker. |
| 6 | **Control Evidence Pack Generator** | B8 | M | Attacks audit-prep hours — leadership-metric CFOs and audit committees care about most. |
| 7 | **Exception Narrative Generator** | A2 | S | Eliminates manual variance-commentary drafting. Half-day saved every close cycle. |
| 8 | **Header Row Auto-Detect** | A4 | S | Foundational plumbing; low effort but makes every other UT tool work on real messy files. |
| 9 | **Zero-Install Workbook Compare** | B1 | S | Runs on any locked iPipeline laptop with no pip install. Proves "Python is not a developer thing." |
| 10 | **Quick Demo Mode Macro** | A10 | S | The "show me in 2 minutes" answer. Meta-tool that sells every other tool every time it runs. |

**One-paragraph wrap-up:** Top 10 is deliberately weighted toward Section A universal-toolkit adds because small-effort plug-and-play tools get used weekly by 2,000 people, which dominates any one-off demo moment. Medium-effort picks (B6, B7, B8) earn their spot because each unlocks a Phase-1 CFO-grade outcome — readiness scoring, triaged exception queues, compressed audit prep — that existing code doesn't deliver. Deliberately left off: all of Section C (great but post-demo), SQL-heavy warehouse items (valuable but infrastructure-bound), B10 Workbook Dependency Scanner (clever but specialist, not crowd-pleaser). If only five things can be built before Video 4 recording, build sequence: **A4 → A1 → A3 → A2 → A9** — they compound into a single ~45-second demo arc where each runs on a coworker's real file producing narrated CFO-ready output. Tightest possible Video 4 opening.

---

## 5. TOP PICKS CONSOLIDATED

15 items maximum. Format: `[#] Idea Name | Language | Effort | Best Fit | Why it wins`

| # | Idea | Lang | Effort | Best Fit | Why it wins |
|---|---|---|---|---|---|
| 1 | Close Readiness Score View | SQL | M | Video 4 (B6) | CFO-level single number per entity; highest-leverage SQL item in backlog |
| 2 | Data Quality Scorecard | VBA | S | Universal Toolkit (A3) | Any-sheet 0–100 grade; strongest single wow moment for CFO demo |
| 3 | Exception Triage Engine | Python | M | Video 4 (B7) | Config-driven impact×confidence×recency ranking improves every close |
| 4 | Materiality Classifier | VBA | S | Universal Toolkit (A1) | Row-level risk tiering in seconds; pairs with A2+A3 for triage arc |
| 5 | Word Report Talking Points | Python | S | Universal Toolkit (A9) | AI-style narrative with zero AI calls — CFO demo-wrecker |
| 6 | Control Evidence Pack Generator | Python | M | Video 4 (B8) | Compresses audit-prep hours — leadership-metric language |
| 7 | Exception Narrative Generator | VBA | S | Universal Toolkit (A2) | Eliminates manual variance commentary; half-day saved per close |
| 8 | Header Row Auto-Detect | VBA | S | Universal Toolkit (A4) | Foundational plumbing makes every other UT tool work on messy files |
| 9 | Zero-Install Workbook Compare | Python | S | Video 4 (B1) | Stdlib-only; runs on locked laptops; proves "Python is for everyone" |
| 10 | Quick Demo Mode Macro | VBA | S | Universal Toolkit (A10) | One-button 5-feature autoplay; meta-tool sells every other tool |
| 11 | Zero-Install Scenario Runner | Python | S | Video 4 (B3) | Real what-if automation with no deps; CFO/CEO sees instant value |
| 12 | Executive Summary Builder | Python | S | Video 4 (B5) | CSV folder → Markdown exec summary; closes Video 4 on a high note |
| 13 | Finance Data Contract Checker | Python | M | Video 4 (B9) | YAML-declared schema validation prevents bad data polluting reports |
| 14 | Template Enforcer | VBA | S | Universal Toolkit (A12) | Validates submitted files against template; kills support loop |
| 15 | "Show Tools" Launcher Button Installer | VBA | S | Universal Toolkit (A7) | Branded blue Cover-sheet button; removes "how do I start?" friction |

---

*End of handoff file. Claude Code should have everything needed to review, prioritize, and start building.*
