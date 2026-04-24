# Video 4 — Current Proposal (snapshot)

**Snapshot date:** 2026-04-23
**Purpose:** Durable record of the current Video 4 plan + 6-bucket synthesis from the research-compiled docs, in case chat memory gets compressed or a new session starts.
**Status:** Proposal — not yet locked in. User reviewing.

Inputs that shaped this proposal:
- The 3 combos we weighed (Finance Copilot menu / xlwings Excel Button Edition / Hero Demo + Cookbook)
- 6 research-compiled docs at `RecTrial\Brainstorm\NewCodeResearch\ResearchComplied\` (CodeReview_V1.md through V6.md) — 156 unique ideas inventoried, 40–60 curated per doc
- Hard constraints: no AI APIs, no Outlook, no Task Scheduler, approved Python packages only, non-dev audience, iPipeline branding

---

## Final plan — Split Video 4 into 4a + 4b

### Video 4a — "Python Shows You What Excel Can't" (6–7 min, CFO-led)

| Beat | Duration | What happens |
|---|---|---|
| **Opener** | ~90 sec | **Workbook Dependency Scanner** — Python parses a messy inherited workbook, renders its cross-sheet dependency graph as HTML. Sets the "Python ≠ Excel++" frame. |
| **Hero** | ~3–4 min | **SaaS ARR/MRR Waterfall Engine** — subscription CSV → Starting ARR → New/Expansion/Contraction/Churn → Ending ARR as a branded matplotlib waterfall + roll-forward table. iPipeline-native SaaS story. |
| **Closer** | ~60 sec | One-liner teaser for 4b cookbook + download link. |

### Video 4b — "Your Python Cookbook" (5–6 min, coworker-led)

| Beat | Duration | What happens |
|---|---|---|
| **Recipe 1** | ~90 sec | **Finance Data Contract Checker** (JSON, not YAML) — red FAIL report → fix → green PASS. Fastest build, most dramatic before/after. |
| **Recipe 2** | ~90 sec | **Exception Triage Engine** — config-driven impact × confidence × recency scoring. Monday-morning priority list. |
| **Recipe 3** | ~90 sec | **Control Evidence Pack Generator** — audit bundle with hashes + manifest. Audit-prep hours → seconds. |
| **Closer** | ~60 sec | Download the Finance Copilot menu + cookbook folder. |

### Deliverable

A single downloadable package:
- `finance_copilot.py` — menu-driven launcher wrapping all 28 existing scripts + the 4–5 new ones
- `cookbook/` — folder with every Python script coworkers can copy-paste
- Optional v2: Excel Button Edition via xlwings (shipped later after IT check)

---

## 5 new scripts to build (all approved packages, all M-effort or smaller)

| # | Script | Language | Effort | Role |
|---|---|---|---|---|
| 1 | `saas_arr_waterfall.py` | Python (pandas + matplotlib + openpyxl) | M | V4a hero |
| 2 | `workbook_dependency_scanner.py` | Python (openpyxl + stdlib html) | M | V4a opener |
| 3 | `data_contract_checker.py` | Python (pandas + stdlib json) | S | V4b recipe |
| 4 | `exception_triage_engine.py` | Python (pandas + numpy + openpyxl) | M | V4b recipe |
| 5 | `control_evidence_pack.py` | Python (stdlib hashlib + zipfile + openpyxl + python-docx) | M | V4b recipe |

Plus the `finance_copilot.py` menu wrapper = 6 new Python files total.

---

## The 6 buckets — full synthesis from the research docs

### Bucket 1 — Video 4 ship list
See above. 5 new scripts + Copilot menu.

### Bucket 2 — Entire-project new code (beyond V4)
- **Revenue Recognition Engine (ASC 606)** — park as Video 5 candidate
- **Formula Integrity Fingerprinting** (VBA)
- **Dependency Impact Preview** (VBA)
- **Auto-Repair Suggestions** (VBA, complements existing Data Sanitizer)
- **Workbook Policy Validator** (VBA)
- **Intelligent Rollforward Assistant** (VBA)
- **Controlled Snapshot Sign-off** (VBA)
- **Narrative Variance Writer Python (deterministic)** — bigger sibling of the --talking-points flag we already shipped
- **Cohort Retention Analyzer** (Python)
- **CFO Pack Assembly Pipeline** (Python, L-effort)

### Bucket 3 — New code for Universal Toolkit
Already built (confirmed): Materiality Classifier, Exception Narrative Generator, Data Quality Scorecard, Header Row Auto-Detect, Quick Row Compare Count, Run Receipt Sheet, SHOW TOOLS button installer, Intelligence category pinned in Command Center, `word_report --talking-points`, 7 ZeroInstall Python scripts.

Genuinely new to add:
1. Dependency Impact Preview (VBA, M)
2. Workbook Policy Validator (VBA, M)
3. Auto-Repair Suggestions (VBA, M)
4. Formula Integrity Fingerprinting (VBA, M)
5. Intelligent Rollforward Assistant (VBA, L)
6. Fiscal Year Startup Check (VBA, S)
7. Quick Demo Mode Macro (VBA, S)
8. What's New Sheet (VBA, S)
9. Narrative Variance Writer Python (Python, M)
10. Workbook Dependency Scanner Python (Python, M) — reused from V4

### Bucket 4 — Cherry-picks to steal from the research
- `saas_arr_waterfall.py` from CODE_CATALOG.md — port into V4 hero
- `sox_evidence_collector.py` — port if SOX applies to Connor's team (unknown)
- `cohort_retention_analyzer.py` — could swap into V4b cookbook
- `revenue_recognition_engine.py` — park for Video 5
- Compare-and-Classify 4-status pattern (Identical / Modified / Removed / Added) — incorporate into v2 of existing `compare_workbooks.py`
- Formula Integrity Fingerprint pattern from CodexCodeIdeas.md VBA-02
- xlwings UDF pattern from report_extended.md — for the future Excel Button Edition

### Bucket 5 — New ideas not in the docs
1. **Copilot Prompt Generator** — Python preps data + outputs prompts you paste into M365 Copilot. Bridges to AI without violating constraints.
2. **Sample Data Generator** — `--demo` flag on any script auto-generates realistic fake data.
3. **Tool Usage Heat Tracker** — track which toolkit tools get used most.
4. **First-Run Onboarding Wizard** — 2-minute guided tour on first open.
5. **"Ask Copilot" Button in workbook** — clicks open copilot.microsoft.com with pre-drafted prompt + data on clipboard.
6. **Zip-and-Ship Button** — timestamped zip of workbook + outputs for sharing with auditors.
7. **Python Script Explorer** — list every Python script with 1-line descriptions.

### Bucket 6 — Future parking lot
- All AI API-dependent ideas (LLM contract parser, AI narrator, Instructor structured extraction)
- Warehouse SQL (Close Readiness Score View SQL, Allocation Drift, Forecast Backtest Warehouse, JE Duplicate Rings SQL, etc.)
- Infrastructure (Airflow, Flask/FastAPI, dbt, GitHub Actions CI)
- ML-dependent (Forecast Ensemble, Isolation Forest, SARIMA/Prophet, Splink)
- Platform misfits (Outlook, Slack/Teams webhooks, JIRA, Power Automate)
- Scope misfits (.NET add-in, VSTO, AWS cost, pyautogui RPA)

---

## Where the research docs disagree

| Question | Split | My recommendation |
|---|---|---|
| **V4 hero tool** | V3: Revenue Leakage Finder · V4/V6: ARR Waterfall · V5: xlwings buttons · V1: Exception Triage | **ARR Waterfall** — iPipeline is SaaS, 3 of 6 docs agree |
| **Combo structure** | V2: Menu · V3: Hero + Cookbook video + Menu tool · V4/V6: split 4a+4b · V5: xlwings primary | **Split 4a + 4b, menu as deliverable** |
| **Exception Triage role** | V1: flagship · V3/V4/V6: cookbook recipe | **Cookbook recipe** (V3 notes it's "visually flat" as hero) |
| **SOX Evidence Collector** | V3: finalist · V4: parked | **Park until SOX ownership confirmed** |
| **xlwings delivery** | V5: primary · Others: optional or v2 | **v2 for after V4 lands, confirm IT first** |

---

## Open questions for the user to confirm

1. Approve the 4a + 4b split? (Yes = proceed. No = we go back to single video.)
2. Approve ARR Waterfall as hero vs. Revenue Leakage Finder?
3. Does your team own SOX evidence work? (Affects whether SOX Evidence Collector is a finalist.)
4. OK to skip xlwings for now and park as "Excel Button Edition v2" post-V4?
5. Ship downloadable as Python CLI menu (simple) or both Python menu + Excel button version (bigger lift)?

---

## Effort estimate to execute the plan

| Phase | Work | Estimate |
|---|---|---|
| Build new Python scripts (5) | `saas_arr_waterfall`, `workbook_dependency_scanner`, `data_contract_checker`, `exception_triage_engine`, `control_evidence_pack` | 2–3 days |
| Build Finance Copilot menu wrapper | `finance_copilot.py` | 3–4 hours |
| Write video scripts (4a + 4b) | narration + on-screen callouts | 1 day |
| Generate demo input files for each recipe | realistic fake data | 4 hours |
| Record + edit | 4a + 4b | 1 day |
| **Total** | | **~5 days of focused work** |

---

## What's NOT in this proposal (intentional)

- **xlwings Excel Button Edition** — parked as v2 post-V4
- **AI-powered anything** — respects constraint
- **Revenue Recognition Engine** — too heavy for V4, future Video 5 material
- **SOX-specific tooling** — contingent on SOX ownership question
- **SQL-heavy ideas** — deferred; Python video shouldn't dilute with SQL

---

## Raw research files review — COMPLETE 2026-04-23

Reviewed all 14 raw files at `C:\Users\connor.atlee\RecTrial\Brainstorm\NewCodeResearch\ResearchFiles` via Explore subagent. **Finding: confidence HIGH that nothing new was missed.** The 6 compiled docs successfully captured every substantive idea that passes hard constraints and fits Finance-analyst audience. No new V4 ideas, no new toolkit additions, no factual corrections. Raw files add procedural detail + reinforce priorities but do not change the current plan.
