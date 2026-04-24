# Video 2 — Deep Dive: P&L Demo Workbook Automation

## Target Length
9 minutes

## Timestamped Outline
- 00:00–00:15 Hook
- 00:15–01:20 Workbook map
- 01:20–03:20 Reconciliation engine deep dive
- 03:20–05:20 Variance narrative deep dive
- 05:20–06:50 Scenario runner deep dive
- 06:50–08:10 Executive brief pack deep dive
- 08:10–09:00 CTA

## Full Narration Script

### 00:00–00:15 Hook
"This is the full demo workflow your finance team can run every close cycle."

### 00:15–01:20 Workbook map
"We use the P&L trend, source transaction sheet, checks, and report tabs as the core model. The flow is intentionally non-destructive: calculations and summaries write to output sheets so source tabs remain stable."

### 01:20–03:20 Reconciliation engine deep dive
"We run `RunDemoReconciliation`. The macro refreshes row-count checks, amount completeness, date window extraction, and a revenue tie-out. The purpose is simple: trust the data before building commentary. If a check fails, status is explicit and timestamped."

### 03:20–05:20 Variance narrative deep dive
"Next we run `GenerateDemoVarianceNarrative`. The macro compares first and latest trend periods, calculates deltas and percentages, classifies movement by materiality thresholds, and writes plain-English narrative lines for leadership review."

### 05:20–06:50 Scenario runner deep dive
"Now we run `RunDemoWhatIfScenarios`. It generates base, growth push, margin protection, and stress cases. Each row includes adjusted revenue and cost, resulting margin, and an explanatory sentence. This gives a fast decision envelope for finance leadership."

### 06:50–08:10 Executive brief pack deep dive
"Finally, `BuildDemoExecutiveBriefPack` builds a branded one-page summary with KPI section, control status summary, and trend movement. If the workbook is saved, the macro exports a PDF for distribution."

### 08:10–09:00 CTA
"You now have a full close-cycle automation storyline. Next, we shift to universal tools that work on any coworker workbook."

## On-Screen Action Callouts
- Point to required sheets used by demo modules.
- Show `Checks` table refresh after macro run.
- Highlight materiality status and narrative columns.
- Highlight scenario narratives by case.
- Show `Exec_Brief` header and KPI/control sections.

## Closing CTA
"Open Video 3 to see how the same automation mindset works on plain, everyday files."
