# Video 4 — Python + SQL Integration for Finance Workflows

## Target Length
8 minutes

## Timestamped Outline
- 00:00–00:20 Hook
- 00:20–01:30 Why hybrid workflows matter
- 01:30–03:20 Workbook profiling + sanitation pipeline
- 03:20–05:10 Workbook diff and exception output
- 05:10–06:50 Executive summary generation
- 06:50–08:00 CTA

## Full Narration Script

### 00:00–00:20 Hook
"When Excel gets heavy, this hybrid workflow keeps control in Finance while adding speed from Python and SQL-style extraction patterns."

### 00:20–01:30 Why hybrid workflows matter
"Excel is still the operating surface for Finance. But source quality checks, diffs, and narrative prep can run faster and more repeatably outside the workbook. The point is not replacing Excel. The point is strengthening it."

### 01:30–03:20 Workbook profiling + sanitation pipeline
"First, run `profile_workbook.py` to map sheets, dimensions, named ranges, and VBA presence. Then run `sanitize_dataset.py` on raw CSV exports to normalize dates, text-stored numbers, and spacing before data lands back in Excel workflows."

### 03:20–05:10 Workbook diff and exception output
"Use `compare_workbooks.py` to produce a structured diff file for two workbook versions. Instead of manual cell hunting, reviewers get a clear output they can filter and assign for review."

### 05:10–06:50 Executive summary generation
"Now run `build_exec_summary.py` against a cleaned dataset. The script generates a plain-English markdown summary with totals, ranges, top contributors, and talking points you can reuse in leadership prep."

### 06:50–08:00 CTA
"Adopt this pipeline in one close cycle: profile, sanitize, compare, summarize. Keep Excel as your delivery layer and let Python handle repeatable heavy lifting."

## On-Screen Action Callouts
- Run each Python utility from terminal with visible command and output.
- Open generated files (`profile.json`, `diff.csv`, `summary.md`).
- Highlight how each artifact feeds review or executive prep.

## Closing CTA
"In Video 5, we show how to adapt demo-specific macros to your own workbook with CoPilot prompts."
