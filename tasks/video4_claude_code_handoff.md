# Video 4 Planning Handoff — Reply for Claude Code

> **How to use this file:** Copy everything from the "START COPY" line below and paste it
> into a new Claude Code session. Claude Code will have the full context it needs to begin
> building Video 4 without any repeated explanation.

---

<!-- START COPY -->

## Context — Who I Am and What This Project Is

I'm Connor, a Finance & Accounting analyst at iPipeline (life insurance / financial services SaaS,
~2,000 employees). I'm building a 4-video demo for 2,000+ coworkers + CFO/CEO.
Videos 1–3 are already recorded. **Video 4 is what I'm planning now.**

**Video 4 theme:** "Python Automation for Finance"  
**Goal:** Show what Python uniquely adds on top of Excel + VBA  
**Runtime target:** 5–8 minutes  
**Hard requirement:** Must ship a real, downloadable, coworker-usable tool by Monday after airing

---

## Repo Location and Structure

```
Repo: tug83535/claude-training-lab-code
Relevant Python scripts: /python/
Config/constants: /python/pnl_config.py
CLI runner: /python/pnl_runner.py
Month-end close: /python/pnl_month_end.py
Task tracking: /tasks/todo.md and /tasks/lessons.md
This analysis: /tasks/video4_analysis.md
```

**Brand colors:** Blue `#0B4779`, Navy `#112E51`, Arial fonts only

---

## Hard Constraints — Never Violate These

- No external AI API calls
- No Outlook/email automation
- No Windows Task Scheduler
- No internet scraping of company/paid data
- Python packages allowed: `pandas`, `openpyxl`, `pdfplumber`, `python-docx`,
  `thefuzz`, `numpy`, `matplotlib`, `xlwings`, `stdlib` only
- Audience is non-developers (Finance & Accounting staff)

---

## Already Built — Do NOT Re-Suggest or Re-Build These

`aging_report`, `bank_reconciler`, `compare_files`, `forecast_rollforward`, `fuzzy_lookup`,
`pdf_extractor`, `variance_analysis`, `variance_decomposition`, `clean_data`, `consolidate_files`,
`multi_file_consolidator`, `date_format_unifier`, `two_file_reconciler`, `sql_query_tool`,
`word_report`, `batch_process`, `regex_extractor`, `unpivot_data`, `pnl_forecast`, `pnl_dashboard`,
`master_data_mapper`, `profile_workbook`, `sanitize_dataset`, `compare_workbooks`,
`build_exec_summary`, `variance_classifier`, `scenario_runner`, `sheets_to_csv`

---

## The Three Combo Options I'm Weighing

1. **"Finance Copilot" menu** — One Python script with a numbered menu. User picks a number,
   Python walks them through the task. Wraps existing scripts in one entry point.
2. **"Excel Button Edition" (xlwings)** — Macro-enabled workbook where coworkers click Excel
   buttons; Python runs silently behind the scenes; results appear as new sheets.
3. **"Hero Demo + Cookbook"** — One dramatic hero demo + a 5-recipe cookbook of
   copy-pasteable Python scripts for coworkers to steal.

---

## Branch Review Results — Finalist Tools (Ranked and Scored)

A full analysis was completed across all branches active in the last 10 days. Here are the 9
finalists with scores and combo fits:

| # | Idea Name | CFO Wow | Coworker Use | Demo-ability | Effort | Best Combo |
|---|---|:---:|:---:|:---:|---|---|
| 1 | Exception Triage Engine | 5 | 5 | 5 | M | Combo 1 |
| 2 | Control Evidence Pack Generator | 5 | 4 | 5 | M | Combo 3 |
| 3 | Finance Data Contract Checker | 4 | 5 | 5 | S | Combo 1 |
| 4 | Root Cause Reconciliation Assistant | 4 | 5 | 4 | M | Combo 1 |
| 5 | Workbook Dependency Scanner | 5 | 4 | 5 | M | Combo 2 |
| 6 | Narrative Variance Writer | 4 | 4 | 4 | S | Combo 3 |
| 7 | CFO Pack Assembly Pipeline | 5 | 4 | 5 | M | Combo 2 |
| 8 | SaaS ARR/MRR Waterfall Engine | 5 | 4 | 5 | M | Combo 3 |
| 9 | Revenue Recognition Engine (ASC 606) | 5 | 5 | 4 | L | Combo 3 |

---

## Combo Recommendations (Best 3 Per Track)

### Combo 1 — "Finance Copilot" Menu
1. Exception Triage Engine
2. Finance Data Contract Checker
3. Root Cause Reconciliation Assistant

### Combo 2 — "Excel Button Edition" (xlwings)
1. Workbook Dependency Scanner
2. CFO Pack Assembly Pipeline
3. Finance Data Contract Checker

### Combo 3 — "Hero Demo + Cookbook"
1. Revenue Recognition Engine (ASC 606)
2. SaaS ARR/MRR Waterfall Engine
3. Control Evidence Pack Generator

---

## My Recommended Pick

> **Combo 3 is the strongest choice** because it gives the best "wow" story in 5–8 minutes:
> one dramatic hero tool (Revenue Recognition or ARR Waterfall) that impresses the CFO/CEO,
> plus a 5-recipe cookbook that gives every coworker in the audience something actionable to
> download and use on Monday. The hero establishes credibility. The cookbook drives adoption.

---

## What I Need You to Do Next

**Before starting any work, confirm this plan in bullet points and wait for my approval.**

Once I approve, here is the priority order:

1. Review `tasks/lessons.md` and `tasks/todo.md` before doing anything
2. Help me decide between the 3 combos (or pick the one I choose)
3. Build the selected script(s) following these standards:
   - Every script: full header with `PURPOSE`, `USAGE`, and `AUDIENCE` section
   - Output: branded `.xlsx` with iPipeline Blue `#0B4779` headers, Arial font
   - CLI: `argparse` with helpful `--help` text
   - Guard against bad input gracefully (show clear errors, not stack traces)
   - Include a `requirements.txt` if any new packages beyond the allowed list are needed
   - No hardcoded file paths — all passed as CLI arguments with sensible defaults
4. After building, update `tasks/todo.md` and `tasks/lessons.md`
5. All output files go in `/python/` folder to match existing project structure

---

## Quality Standard Reminder

Every piece of output must meet "world-class" standard:
- Would the CFO/CEO be proud to see this in a presentation?
- Is every step completely written out?
- Is the code clean, readable, and error-handled?
- Does the tool work on first run without developer help?

<!-- END COPY -->
