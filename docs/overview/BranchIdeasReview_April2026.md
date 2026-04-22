# Branch Ideas Review — April 2026
**Generated:** 2026-04-22 | **Reviewed by:** Copilot GitHub Agent  
**Branches reviewed (last 10 days):** `April19update`, `claude/document-codexreview2-DNgF3`, `claude/mobilecldcode-business-automation-zC3Jt`, `codex/create-codexreview2-folder-and-conduct-full-branch-review`, `codex/review-branch-and-suggest-new-ideas`, and others with no new commits in range.

---

## Project Context

Finance & Accounting enablement platform for iPipeline (life insurance / financial services SaaS).  
Goal: 4-video demo series for 2,000+ coworkers + CFO/CEO.

**Approved stack:** VBA · Python (pandas, openpyxl, pdfplumber, python-docx, thefuzz, numpy, matplotlib, xlwings, stdlib) · SQL  
**Constraints:** No external AI APIs · No Outlook/email automation · No Task Scheduler · No `scikit-learn`  
**Branding:** iPipeline Blue `#0B4779` · Navy `#112E51` · Arial fonts

---

## Section A — Universal Toolkit Additions

> Ideas that belong in `modUTL_*` or `UniversalToolkit/python/` and work on **any** coworker file.

| # | Idea Name | What It Does | Language | Effort | Why It's Worth Including | Source |
|---|-----------|--------------|----------|--------|--------------------------|--------|
| A1 | **Materiality Classifier** | Tags each row as Material increase/decrease, Watch, or Normal using configurable $ and % thresholds. Auto-detects Current/Prior columns by header text. | VBA | S | Gives analysts instant risk triage on any worksheet — no setup needed. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_Intelligence.bas` |
| A2 | **Exception Narrative Generator** | Writes plain-English row narratives based on Materiality Status column. Produces CFO-ready wording automatically. | VBA | S | Saves manual commentary drafting time every close cycle. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_Intelligence.bas` |
| A3 | **Data Quality Scorecard** | Scores a sheet 0–100 from blanks/errors and writes a formatted quality report tab. | VBA | S | Creates a simple "data trust" signal leaders can understand instantly. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_Intelligence.bas` |
| A4 | **Header Row Auto-Detect** | Scans top rows and picks the most-likely header row, removing the need for hardcoded row numbers. | VBA | S | Makes every tool truly plug-and-play on any coworker file. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_Core.bas` |
| A5 | **Quick Row Compare Count** | Fast pre-check that hashes rows and returns mismatch count before running a full compare. | VBA | S | Answers "are these files meaningfully different?" in seconds. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_Compare.bas` |
| A6 | **Run Receipt Sheet** | Writes a timestamped execution receipt to a `UTL_RunReceipt` tab on every macro run. | VBA | S | Improves control evidence and audit traceability out of the box. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_Audit.bas` |
| A7 | **Cover "Show Tools" Button Installer** | One-time macro that adds a branded blue launcher button to the Cover sheet pointing to Command Center. | VBA | S | Removes the "how do I start?" friction for non-technical users. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_CommandCenter.bas` |
| A8 | **Static Intelligence Category in Command Center** | Pins Materiality/Narratives/Scorecard tools near the top of the tool list instead of scrolling past item 29. | VBA | S | Ensures high-value tools are visible and discoverable for coworkers. | `April19update` — `FinalExport/UniversalToolkit/vba/modUTL_CommandCenter.bas` |
| A9 | **Zero-Install Workbook Profiler** | Inventories workbook sheets, ranges, and VBA flag with stdlib only — no pip install required. | Python | S | Works on locked-down laptops; great first Python demo for the crowd. | `April19update` — `FinalExport/UniversalToolkit/python/ZeroInstall/profile_workbook.py` |
| A10 | **Word Report Talking Points** | Adds optional `--talking-points` flag to `word_report.py` generating 3–5 auto-built CFO narrative bullets. | Python | S | Speeds exec narrative prep without any external AI calls. | `April19update` — `FinalExport/UniversalToolkit/python/word_report.py` |

---

## Section B — Video 4 Candidates

> Ideas best demonstrated in the "Python Automation for Finance" video.

| # | Idea Name | What It Does | Language | Effort | Overlap with existing? | Why Include | Source |
|---|-----------|--------------|----------|--------|------------------------|-------------|--------|
| B1 | **Zero-Install Workbook Compare** | Compares two workbooks row-by-row and exports diffs to CSV. | Python | S | Overlaps compare_files / two-file reconciler | Strong "no-install automation" demo for the whole audience. | `April19update` — `ZeroInstall/compare_workbooks.py` |
| B2 | **Zero-Install Variance Classifier** | Labels rows as Over/Under/On-target vs baseline using rules only. | Python | S | Overlaps variance_analysis | Easy-to-explain risk labeling — zero setup for coworkers. | `April19update` — `ZeroInstall/variance_classifier.py` |
| B3 | **Zero-Install Scenario Runner** | Applies percentage shocks to a metric column and exports all scenarios. | Python | S | None | Demonstrates real what-if automation with no dependencies. | `April19update` — `ZeroInstall/scenario_runner.py` |
| B4 | **Sheets-to-CSV Batch Export** | Exports every sheet in a workbook to its own CSV file. | Python | S | None | Bridge step from Excel to Python pipelines; relatable for Finance folks. | `April19update` — `ZeroInstall/sheets_to_csv.py` |
| B5 | **Executive Summary Builder** | Builds a Markdown executive summary from CSV outputs with stats and highlights. | Python | S | None | Turns raw data into leadership-ready output fast. | `April19update` — `ZeroInstall/build_exec_summary.py` |
| B6 | **Close Readiness Score View** | SQL view returning per-entity close readiness score (0–100) aggregated from failed checks and late postings. | SQL | M | None | One-metric close risk visibility — language the CFO already speaks. | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` |
| B7 | **Exception Triage Engine** | Ranks exceptions by impact × confidence × recency using config-driven weights. | Python | M | None | Helps teams work highest-value issues first every close cycle. | same as B6 |
| B8 | **Control Evidence Pack Generator** | Packages macro logs and validation results into a zipped audit evidence bundle with manifest. | Python | M | None | Directly cuts audit-prep hours; CFO-level governance message. | same as B6 |
| B9 | **Finance Data Contract Checker** | Validates incoming data against YAML schema/quality contracts before downstream use. | Python | M | None | Prevents bad data from quietly polluting reports and forecasts. | same as B6 |
| B10 | **Workbook Dependency Scanner** | Parses formulas and named ranges to map change-impact graph (JSON/HTML output). | Python | M | None | Reduces breakage risk when editing shared workbooks. | same as B6 |

---

## Section C — Future Ideas (Parked)

> Real value, but post-demo. Park until after Video 4 is recorded.

| # | Idea Name | What It Does | Language | Effort | Source |
|---|-----------|--------------|----------|--------|--------|
| C1 | **Allocation Drift Tracker** | Detects silent drift in cost allocation percentages month-over-month with threshold flags. | SQL | M | `codexreview2/02_new_automation_backlog.md` |
| C2 | **Forecast Backtest Warehouse** | Stores every forecast run, assumptions, and realized actuals for accuracy comparison. | SQL | L | same |
| C3 | **Subledger Completeness Control Matrix** | Checks that all required upstream feeds are present before close steps run. | SQL | M | same |
| C4 | **Workbook-to-Source Reconciliation Mart** | Reconciles workbook aggregates against warehouse source-of-truth tables. | SQL | M | same |
| C5 | **Vendor Payment Velocity Baselines** | Flags abnormal timing or amount shifts by vendor using rolling medians. | SQL | L | same |
| C6 | **JE Duplicate Ring Detection** | Finds near-duplicate journal entry patterns split across users, days, or entities. | SQL | L | same |
| C7 | **Close Bottleneck Heatmap Dataset** | Decomposes where close-cycle delays occur by step, entity, and user. | SQL | M | same |
| C8 | **Segregation-of-Duties Audit Pack** | Flags conflicting role/action combinations in the transaction lifecycle. | SQL | M | same |
| C9 | **Formula Integrity Fingerprinting** | Hash-checks critical formula zones to catch silent changes. | VBA | M | same |
| C10 | **Exception Workbench Sheet** | Central Excel tab for assigning, tracking, and closing exceptions with owner/due-date workflow. | VBA | M | same |
| C11 | **Macro Runtime Telemetry Dashboard** | Summarizes run times, error rates, and usage frequency by Command Center action. | VBA | M | same |
| C12 | **Controlled Snapshot Sign-off** | Captures approved monthly workbook state with checksum and approver metadata. | VBA | M | same |

---

## Section D — Skip

> Items found in branches that do **not** fit the project constraints or demo scope.

| Idea Name | Why Skip | Source |
|-----------|----------|--------|
| Outlook Mail Merge w/ Attachments | Violates "no Outlook/email automation" rule. | `claude/mobilecldcode-business-automation-zC3Jt` — `modMailMerge_WithAttachments.bas` |
| Calendar Appointment Builder | Violates "no Outlook/email automation" rule. | same — `modCalendarAppointmentBuilder.bas` |
| JIRA Bridge / Weekly Digest | Out-of-scope external integration; not Finance-close focused. | same — `modJiraBridge.bas`, `02_Python/jira_weekly_digest.py` |
| Slack / Teams Webhook Notifiers | External platform dependency; distracts from core demo. | same — `modSlackNotifier.bas`, `modTeamsNotifier.bas`, `05_OfficeScripts/TeamsWebhookOnThreshold.ts` |
| AWS Cost Optimizer | Not finance-close focused for current 4-video storyline. | same — `02_Python/aws_cost_optimizer.py` |
| ML Churn / Ticket Triage Scripts | Requires `scikit-learn` — violates approved-packages constraint. | same — `customer_churn_risk_scorer.py`, `support_ticket_triage.py` |
| PowerShell IT Admin Automations | Outside approved stack and audience for this finance demo. | same — `MobileCLDCode/04_PowerShell/*.ps1` |
| Power Automate / Office Scripts Flows | Different delivery surface; distracts from Excel+VBA+Python+SQL focus. | same — `MobileCLDCode/06_PowerAutomate/*`, `05_OfficeScripts/*` |

---

## Top 10 — Personal Picks (Ranked by Bang-for-Buck)

| Rank | Idea | Section | Effort | Why It Wins |
|------|------|---------|--------|-------------|
| 1 | Close Readiness Score View | B6 | M | One number per entity = CFO-level language. Highest ROI SQL item. |
| 2 | Exception Triage Engine | B7 | M | Directly improves analyst workflow every single month-end. |
| 3 | Data Quality Scorecard | A3 | S | Fast, visual, zero-setup. Great live demo moment. |
| 4 | Control Evidence Pack Generator | B8 | M | Cuts audit prep hours — leadership cares deeply about audit speed. |
| 5 | Materiality Classifier | A1 | S | Transforms any flat sheet into a risk-tiered view in seconds. |
| 6 | Word Report Talking Points | A10 | S | AI-style output with no AI calls — great story for the CFO demo. |
| 7 | Zero-Install Workbook Compare | B1 | S | Plug-and-play Python on a locked laptop = convincing for coworkers. |
| 8 | Header Row Auto-Detect | A4 | S | Foundational helper; makes all other tools more reliable. |
| 9 | Exception Workbench Sheet | C10 | M | Creates one action hub everyone can use after the demo. |
| 10 | Macro Runtime Telemetry Dashboard | C11 | M | Shows the toolkit is production-grade, not just a demo. |

---

*Last updated: 2026-04-22. Source branches reviewed with commits since 2026-04-12.*
