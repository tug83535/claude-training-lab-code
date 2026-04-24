# iPipeline Finance Demo Series — Claude Code Handoff
**Date:** 2026-04-23
**Project:** Video 4 Planning & Automation Backlog

---

## 1. WHAT IS THIS FILE
This file is a consolidated, deduplicated master handoff document for Claude Code, compiling research and evaluating automation ideas for the iPipeline Finance & Accounting 4-video demo series. It synthesizes Section A (Universal Toolkit), Section B (Video 4 Candidates), parked/skipped concepts, and outputs a strict, scored top-picks list for immediate execution. 

## 2. PROJECT CONTEXT
* **User:** Connor Atlee, Finance contractor at iPipeline (producing deliverables for Michael Van Alstyne, Eric Morgan, and the C-suite), non-developer. 
* **Project:** 4-video internal demo series for 2,000+ coworkers + CFO/CEO.
* **Current Status:** Videos 1–3 recorded. Video 4 ("Python Automation for Finance") is currently in planning.
* **Hard Constraints:** * No external AI API calls (OpenAI, Claude, Gemini).
  * No Outlook / email automation.
  * No Windows Task Scheduler.
  * Approved Python packages ONLY: `pandas`, `openpyxl`, `pdfplumber`, `python-docx`, `thefuzz`, `numpy`, `matplotlib`, `xlwings`, stdlib. 
  * Must be plug-and-play (no hardcoded sheets) and explainable to a zero-coding audience.
* **Already Built:** 23 VBA modules (~140 tools), 28 Python scripts (including zero-install suite), and 4 SQL scripts.

## 3. OPEN DECISIONS
* **Video 4 Structure:** Deciding between a single 5–8 min video or splitting it into Video 4a (Hero Demo) and Video 4b (Cookbook). Splitting is strongly recommended to serve both leadership and analyst audiences properly.
* **VBA Duplication Check:** Need to confirm if Section A items like the Materiality Classifier and Data Quality Scorecard are already integrated into the existing `modUTL_*` modules before building them.
* **xlwings IT Feasibility:** Verify if the 2,000-person target audience can actually install Python and xlwings on locked-down machines, which impacts whether "Combo 2" is shippable or just a demo illusion.
* **Data Contract Format:** Data contracts must use JSON via stdlib instead of YAML, as `PyYAML` is not an approved package.

## 4. FULL RESULTS

**Section A: Universal Toolkit Additions**
* **Header Row Auto-Detect (VBA | S):** Scans top rows to dynamically find headers, removing hardcoded assumptions.
* **Data Quality Scorecard (VBA | S):** Scores sheets 0–100 based on blanks/errors for an instant data trust signal.
* **Exception Narrative Generator (VBA | S):** Writes CFO-ready row commentary based on materiality automatically.
* **Word Report `--talking-points` Flag (Python | S):** Adds AI-style narrative bullets to existing reports without AI APIs.
* **Dependency Impact Preview (VBA | S-M):** Shows downstream cell/chart breakage before a destructive macro runs.

**Section B: Video 4 Candidates**
* **SaaS ARR/MRR Waterfall Engine (Python | M):** Converts subscription rosters into a highly visual ARR cascade chart.
* **Exception Triage Engine (Python | M):** Ranks month-end exceptions by config-driven impact/confidence weights.
* **Workbook Dependency Scanner (Python | M):** Parses formulas to output a cross-sheet dependency map, showing Python doing what Excel cannot.
* **Finance Data Contract Checker (Python | S):** Validates incoming files against required columns and types to stop bad data.
* **Journal Entry Duplicate-Ring Detector (Python | M):** Uses fuzzy matching to group suspicious, split-day duplicate entries.

**Section C: Parked Ideas**
* **Revenue Recognition Engine (ASC 606):** Massive CFO value, but "L" effort is too heavy for a 5-minute video; park for Video 5.
* **Allocation Drift Tracker / Forecast Backtest:** Real value, but requires deep historical data missing from the demo set.
* **Exception Workbench Sheet:** Excel UI workflow that distracts from the core Python narrative.

**Section D: Skipped Ideas**
* **LLM Contract Parser:** Violates external AI API constraint.
* **Close Calendar Risk Predictor:** Violates package constraints (`scikit-learn`).
* **Outlook Mail Merge & Slack Bots:** Violates email and external platform constraints.
* **Airflow Pipelines:** Violates Task Scheduler / orchestration constraint.

## 5. TOP PICKS
[1] SaaS ARR/MRR Waterfall Engine | Python | M | Video 4a (Hero) | Highly visual, SaaS-native to iPipeline
[2] Workbook Dependency Scanner | Python | M | Video 4a (Opener) | Shows Python doing what Excel structurally cannot
[3] Exception Triage Engine | Python | M | Video 4b (Cookbook) | Cleanest chaos→clarity visual; config-driven weights
[4] Finance Data Contract Checker | Python | S | Video 4b (Cookbook) | Fastest build, ultimate guardrail story
[5] Control Evidence Pack Generator | Python | M | Video 4b (Cookbook) | Tangibly cuts audit-prep hours for leadership
[6] Journal Entry Duplicate-Ring Detector | Python | M | Video 4 (Alternate) | Forensic fraud/SOX story that lands hard
[7] Word Report --talking-points Flag | Python | S | Universal Toolkit | AI-style output, zero AI calls
[8] Header Row Auto-Detect | VBA | S | Universal Toolkit | Foundational plumbing making all other tools work
[9] Data Quality Scorecard | VBA | S | Universal Toolkit | Instant data trust signal for non-technical leaders
[10] Exception Narrative Generator | VBA | S | Universal Toolkit | Produces CFO-ready wording automatically
