# Video 4 — Python Automation for Finance: Branch Analysis & Finalist Ranking

**Generated:** 2026-04-22  
**Branch reviewed:** All branches active in the last 10 days  
**Branches with activity (last 10 days):**
- `April19update` — 2026-04-21 (Memory + project doc updates for Video 3 shipped)
- `claude/document-codexreview2-DNgF3` — 2026-04-22 (CodexCodeIdeas.md added)
- `claude/mobilecldcode-business-automation-zC3Jt` — 2026-04-22 (MobileCLDCode CODE_CATALOG.md)
- `codex/create-codexreview2-folder-and-conduct-full-branch-review` — 2026-04-21
- `codex/review-branch-and-suggest-new-ideas` — 2026-04-20

**Assumption:** "Active in the last 10 days" = branch HEAD commit on or after 2026-04-12.

---

## Hard Constraints Applied (Skipped If Violated)

- No external AI API calls  
- No Outlook/email automation  
- No Windows Task Scheduler  
- No internet scraping of company/paid data  
- Packages: `pandas`, `openpyxl`, `pdfplumber`, `python-docx`, `thefuzz`, `numpy`, `matplotlib`, `xlwings`, `stdlib` only  
- Non-developer audience  
- iPipeline branding: Blue `#0B4779`, Navy `#112E51`, Arial fonts  

---

## Already Built — Excluded

`aging_report`, `bank_reconciler`, `compare_files`, `forecast_rollforward`, `fuzzy_lookup`,
`pdf_extractor`, `variance_analysis`, `variance_decomposition`, `clean_data`, `consolidate_files`,
`multi_file_consolidator`, `date_format_unifier`, `two_file_reconciler`, `sql_query_tool`,
`word_report`, `batch_process`, `regex_extractor`, `unpivot_data`, `pnl_forecast`, `pnl_dashboard`,
`master_data_mapper`, `profile_workbook`, `sanitize_dataset`, `compare_workbooks`,
`build_exec_summary`, `variance_classifier`, `scenario_runner`, `sheets_to_csv`

---

## Finalist Table (9 Candidates)

| # | Idea Name | What It Does | Why Perfect for Video 4 | Best Combo Fit | CFO Wow | Coworker Usefulness | Demo-ability | Effort | Packages Needed | Source (Branch + File) |
|---|---|---|---|---|:---:|:---:|:---:|---|---|---|
| 1 | Exception Triage Engine | Scores and ranks open exceptions so teams work highest-impact items first. Uses transparent rule weights (impact × confidence × recency) and exports a ranked list. | Shows Python doing priority logic Excel/VBA cannot do cleanly at scale. Ships as a practical "what should I fix first?" tool. | Combo 1 | 5 | 5 | 5 | M | pandas, numpy, openpyxl, stdlib | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` (PY-01) |
| 2 | Control Evidence Pack Generator | Pulls approved outputs/logs into one audit-ready folder with manifest and checksums. Produces a standardized evidence bundle in one run. | Very executive-friendly: faster audits, better control discipline, immediate business value. Great real-tool outcome after the video. | Combo 3 | 5 | 4 | 5 | M | pandas, openpyxl, python-docx, stdlib | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` (PY-07) |
| 3 | Finance Data Contract Checker | Validates incoming files against required columns, types, and quality rules before anyone reports from them. Exports a pass/fail report with exact issue rows. | Perfect non-dev guardrail story: "bad input blocked before it breaks finance reporting." Clear Python advantage over manual checking. | Combo 1 | 4 | 5 | 5 | S | pandas, openpyxl, stdlib | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` (PY-03) |
| 4 | Root Cause Reconciliation Assistant | Suggests likely cause categories for reconciliation breaks by matching against prior resolved patterns. Gives first-pass guidance so analysts troubleshoot faster. | Demonstrates intelligence without external AI APIs. High practical value and easy to reuse weekly/monthly. | Combo 1 | 4 | 5 | 4 | M | pandas, thefuzz, numpy, openpyxl, stdlib | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` (PY-05) |
| 5 | Workbook Dependency Scanner | Scans formulas/named ranges and outputs a dependency map + impact list for safe changes. Shows "if you edit this cell, these tabs break." | Great "Python on top of Excel" story for Video 4. Visual and reassuring for non-technical users. | Combo 2 | 5 | 4 | 5 | M | openpyxl, pandas, stdlib | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` (PY-08) |
| 6 | Narrative Variance Writer | Converts variance outputs into controlled, deterministic commentary drafts using approved templates. No AI calls; all rule-based text generation. | Strong for executive packs while staying compliant with the no-external-AI rule. Easy adoption via copy/paste-ready output. | Combo 3 | 4 | 4 | 4 | S | pandas, python-docx, stdlib | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` (PY-02) |
| 7 | CFO Pack Assembly Pipeline | Combines approved tables/charts/commentary into one locked release package. Standardizes month-end deliverables and reduces assembly mistakes. | "One-click board pack" lands well with CFO/CEO and is visibly useful to teams Monday morning. | Combo 2 | 5 | 4 | 5 | M | pandas, openpyxl, matplotlib, python-docx, stdlib | `claude/document-codexreview2-DNgF3` — `codexreview2/CodexCodeIdeas.md` (PY-09) |
| 8 | SaaS ARR/MRR Waterfall Engine | Converts a subscription roster into Starting ARR → New → Expansion → Contraction → Churn → Ending ARR, plus NRR/GRR and cohort retention. Exports polished workbook. | High executive signal and very finance-native. Excellent hero segment with immediate downloadable value. | Combo 3 | 5 | 4 | 5 | M | pandas, openpyxl, stdlib | `claude/mobilecldcode-business-automation-zC3Jt` — `MobileCLDCode/02_Python/saas_arr_waterfall.py` |
| 9 | Revenue Recognition Engine (ASC 606) | Builds period-level recognized revenue, deferred revenue rollforward, commission amortization, and exception tabs from contract/billing files. Handles proration and contract patterns. | Extremely CFO-relevant and clearly beyond Excel. Delivers a serious production-style finance tool. | Combo 3 | 5 | 5 | 4 | L | pandas, openpyxl, numpy, stdlib | `claude/mobilecldcode-business-automation-zC3Jt` — `MobileCLDCode/02_Python/revenue_recognition_engine.py` |

---

## 30–60 Second Camera Narratives Per Finalist

### 1 — Exception Triage Engine
Start with a messy "100 exceptions" file. Run one command and show a ranked output where the top 10 items are clearly highest dollar/risk. End by saying: *"This tells Monday-me exactly what to fix first."*

### 2 — Control Evidence Pack Generator
Show a folder with scattered logs and outputs. Run the script and reveal a clean evidence package with index + checksums. Close with: *"What took hours now takes minutes."*

### 3 — Finance Data Contract Checker
Intentionally break a sample input (missing column, wrong date format). Run the checker, show a red fail report with exact rows/columns, then rerun with corrected file and show the green pass. Great before/after demo moment.

### 4 — Root Cause Reconciliation Assistant
Feed a reconciliation break file and show suggested cause categories with confidence tags based on prior history. Highlight that analysts now start with probable causes, not a blank page.

### 5 — Workbook Dependency Scanner
Point at one key assumption cell, run the scanner, and show which sheets/formulas/charts depend on it. Emphasize: *"Safe edits, fewer accidental report breaks."*

### 6 — Narrative Variance Writer
Use a variance table as input and generate draft commentary paragraphs in a branded Word/Excel output. Show that the language is consistent, deterministic, and review-ready.

### 7 — CFO Pack Assembly Pipeline
Start with approved components in separate files, run one command, and open the final board-pack output. Show consistent structure and iPipeline branding.

### 8 — SaaS ARR/MRR Waterfall Engine
Run from raw subscriptions CSV to polished ARR waterfall workbook in seconds. Show the movement bridge + NRR/GRR + cohort retention tab — three executive-level visuals from one script.

### 9 — Revenue Recognition Engine (ASC 606)
Load contracts, billings, and commissions; run one period. Show the deferred rollforward tie-out, recognized revenue tab, commission amortization, and exceptions in one deliverable.

---

## Combo Recommendations

### Best 3 for Combo 1 — "Finance Copilot" Menu
1. Exception Triage Engine
2. Finance Data Contract Checker
3. Root Cause Reconciliation Assistant

### Best 3 for Combo 2 — "Excel Button Edition" (xlwings)
1. Workbook Dependency Scanner
2. CFO Pack Assembly Pipeline
3. Finance Data Contract Checker

### Best 3 for Combo 3 — "Hero Demo + Cookbook"
1. Revenue Recognition Engine (ASC 606)
2. SaaS ARR/MRR Waterfall Engine
3. Control Evidence Pack Generator

---

## Personal Pick

> **If this were my video, I'd pick Combo 3** because it gives you the strongest "wow" story in 5–8 minutes: one dramatic hero (Rev Rec or ARR Waterfall) plus a practical cookbook coworkers can immediately download and use Monday morning. The hero establishes credibility with the CFO/CEO, and the cookbook ensures every coworker in the audience leaves with something actionable.
