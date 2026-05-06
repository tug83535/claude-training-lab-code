# Report 1 — Claude Code Review Brief
## iPipeline Finance Automation Demo Project — Video 4 / Adoption Review

**Prepared for:** Claude Code  
**Prepared by:** ChatGPT review handoff  
**Source overview:** `2026-04-24_ipipeline-finance-automation-demo-master-overview.md`  
**Snapshot basis:** Overview file dated 2026-04-23 plus Connor's clarification after external review  
**Primary objective:** Re-review the project plan with corrected audience/scope assumptions and help Connor lock the Video 4 direction before more code is written.

---

## 0. Critical updated facts from Connor

The current overview file describes the project as a 4-video internal demo series for **2,000+ coworkers + CFO/CEO**. That framing is now too broad for the immediate plan.

Use these corrected assumptions:

1. **Real near-term audience:** approximately **50–150 coworkers**, not 2,000.
2. **Primary goal:** coworker training and showing what is possible, not enterprise-wide deployment.
3. **Video 4 objective:** practical Python automation for Finance users; still polished, but not mainly a CFO boardroom pitch.
4. **Revenue Leakage Finder feasibility:** Connor says he can build a convincing version.
5. **Python constraint:** Python is acceptable only if it is **100% safe** for coworkers.
6. **Support owner:** Connor owns support personally.
7. **Toolkit narrowing:** Connor is open to narrowing the public entry point from ~140 tools to 5–7 supported workflows.
8. **No external AI/API remains the right v1 constraint.**
9. **xlwings should remain parked for v2 unless IT/security makes it easy.**

These facts should override any stale implication in the overview that this is immediately becoming a company-wide product rollout.

---

## 1. Main review thesis

The project is technically strong, but the current risk is **over-scoping the finale and over-exposing the toolkit**.

The most important shift:

> Treat Video 4 as a targeted coworker-training finale, not a full enterprise product launch.

That means:
- one strong official Video 4, not a formal 4a/4b split;
- Revenue Leakage Finder should likely become the hero;
- the Python deliverable should be a safe local automation launcher, not a broad “Copilot” product;
- distribution should be a controlled SharePoint/package release for 50–150 people;
- public documentation should emphasize 5–7 supported workflows, not all 140 VBA tools and 28 Python scripts.

---

## 2. Claude Code task list

Please review the repository/project state and produce a decision memo. Do **not** start building new code until the plan is locked unless Connor explicitly asks you to proceed.

### Task A — Reconcile project framing

Find places where the project still implies:
- 2,000+ coworker distribution as immediate scope;
- CFO/CEO as primary Video 4 target;
- 4a/4b as the official finale structure;
- `finance_copilot.py` as the public-facing product name;
- 140 tools as the adoption surface.

Flag exact files/sections that should be updated.

Expected output:
- `RecTrial/Brainstorm/VIDEO_4_REVIEW_DECISION_MEMO.md`
- Include recommended changes and exact docs that need edits.

---

### Task B — Reassess Video 4 structure

Current plan in overview:
- Video 4a — “Python Shows You What Excel Can't”
- Video 4b — “Your Python Cookbook”
- Hero candidate: SaaS ARR/MRR Waterfall
- Alternate hero: Revenue Leakage Finder
- Deliverable: `finance_copilot.py`

Recommended revised plan:
- One official **Video 4 — Python Automation for Finance**
- 9–12 minutes, chaptered
- Optional recipe clips later, not branded as official 4b
- Hero: Revenue Leakage Finder
- Supporting demos:
  1. Python safety framing
  2. Revenue Leakage Finder
  3. Data Contract Checker
  4. Exception Triage Engine
  5. Control Evidence Pack
  6. Finance Automation Launcher
  7. Where to start / download

Review whether this is stronger than the current 4a/4b plan, given the corrected 50–150 coworker audience.

Expected output:
- `RecTrial/Brainstorm/VIDEO_4_REVISED_PLAN.md`
- Include final recommendation, video chapter outline, demo sequence, estimated build/recording effort, and known tradeoffs.

---

### Task C — Decide hero demo

Assess three possible hero structures:

1. **Revenue Leakage Finder**
   - Finds underbilling/overbilling or revenue-risk exceptions from synthetic but realistic files.
   - Produces ranked exception report, summary metrics, and recommended follow-ups.
   - Best for coworker training because the problem is concrete and practical.

2. **SaaS ARR/MRR Waterfall**
   - Polished executive-style reporting output.
   - Useful, but may look like “Python made a chart.”
   - Better as supporting output than primary hero.

3. **Revenue Control Tower**
   - Combines Data Contract Checker + Leakage Finder + Exception Triage + Evidence Pack + Executive Summary.
   - Strongest concept, but likely more build effort and recording complexity.

Recommendation from external review:
- Use **Revenue Leakage Finder** as the main hero.
- Optionally include a simple ARR movement chart or roll-forward table as a supporting artifact, not the main point.

Expected output:
- In the decision memo, include a ranked hero recommendation and explain what gets cut.

---

### Task D — Safety review for Python

Create or update a safety specification for the Python pack.

Minimum v1 rules:
1. No internet calls.
2. No external AI/API calls.
3. No credentials, tokens, secrets, or database connections.
4. Standard library only for new v1 scripts if feasible.
5. Input files are read-only.
6. Scripts never overwrite source files.
7. Outputs go to timestamped folders under `/outputs/`.
8. Every run writes a log.
9. Sample mode is available.
10. Clear failure messages are shown to the user.
11. Detailed stack traces, if needed, go to log files, not the main user experience.
12. Batch/destructive operations require explicit confirmation.
13. Logs should avoid storing sensitive raw data.
14. The launcher should include a visible safety disclaimer.

Expected output:
- `RecTrial/UniversalToolkit/python/PYTHON_SAFETY.md`
- or `RecTrial/Brainstorm/PYTHON_SAFETY_SPEC.md` if no code changes are desired yet.

---

### Task E — Reassess distribution plan for corrected audience size

For 50–150 coworkers, do **not** propose full IT deployment first. Recommend controlled SharePoint distribution.

Minimum viable release package:

```text
Finance Automation Toolkit v1.0
├── 00_START_HERE.pdf
├── Finance_Automation_Toolkit.xlsm
├── Python_Finance_Starter_Pack.zip
├── Sample_Files.zip
├── Quick_Reference_Card.pdf
├── Known_Limitations.pdf
├── Troubleshooting.pdf
└── Release_Notes.pdf
```

Expected output:
- `RecTrial/Brainstorm/MINIMUM_DISTRIBUTION_PLAN.md`
- Include SharePoint folder/page structure, launch message draft, support expectations, and pilot/release approach.

---

### Task F — Narrow the supported workflow surface

Do not lead with “140 tools.” Lead with 5–7 supported workflows.

Recommended public starter workflows:
1. Clean a messy Excel export.
2. Compare two files.
3. Consolidate sheets/files.
4. Find workbook issues and external links.
5. Generate an executive/workbook summary.
6. Run Python Revenue Leakage Finder on sample data.
7. Run Python Data Contract Checker on sample data.

Everything else should be “advanced / included for exploration.”

Expected output:
- `RecTrial/Brainstorm/SUPPORTED_WORKFLOWS_V1.md`
- Include which existing VBA/Python modules map to each workflow.

---

## 3. Video 4 recommended script outline

### Title
**Video 4 of 4 — Python Automation for Finance**

### Chapter 1 — Why Python after Excel/VBA? — 45 seconds
Message:
- Excel/VBA handles workbook-level automation well.
- Python helps with multi-file workflows, repeatable checks, larger data transformations, and safe report generation.

### Chapter 2 — Safety first — 60 seconds
Message:
- Local files only.
- No internet/API calls.
- Inputs are read-only.
- Outputs go to a separate folder.
- Logs are created so users know what happened.

### Chapter 3 — Hero: Revenue Leakage Finder — 2.5 to 3.5 minutes
Demo:
- Input: billing export, contract/expected revenue file, customer mapping.
- Process: compare expected vs actual, classify leakage risk, rank exceptions.
- Output:
  - summary metrics;
  - top exceptions;
  - recommended follow-ups;
  - output folder.

### Chapter 4 — Data Contract Checker — 90 seconds
Demo:
- Show red FAIL on bad file.
- Fix missing/wrong column.
- Re-run to green PASS.
- Emphasize that this prevents bad inputs before analysis starts.

### Chapter 5 — Exception Triage Engine — 90 seconds
Demo:
- Rank exceptions by impact, confidence, and recency.
- Show top 10 action list.

### Chapter 6 — Control Evidence Pack — 90 seconds
Demo:
- Generate manifest, hash list, log, and evidence folder.
- Emphasize repeatability and audit support.

### Chapter 7 — Launcher — 60 seconds
Demo:
- Show menu.
- Run sample mode.
- Show where outputs land.

### Chapter 8 — How to start — 30 seconds
Message:
- Start with sample files.
- Use supported workflows first.
- Contact Connor for issues.
- Do not use on sensitive production files until comfortable and aligned with team rules.

---

## 4. Recommended naming changes

Avoid making `finance_copilot.py` the main product name unless the project intentionally wants an AI-branded feel.

Better names:
- `finance_automation_launcher.py`
- `run_finance_tools.py`
- `python_finance_launcher.py`

Suggested user-facing label:
**Finance Automation Launcher**

Suggested package name:
**Python Finance Starter Pack**

Suggested SharePoint page:
**Finance Automation Toolkit v1.0**

---

## 5. Risks to explicitly flag

1. **Connor owns support personally.**
   - Therefore the public surface must be narrowed.
   - Do not encourage 50–150 people to explore 140 tools equally.

2. **Python safety must be inspectable.**
   - “Trust me” is not enough.
   - Write a visible `PYTHON_SAFETY.md`.

3. **Revenue Leakage Finder must not feel fake.**
   - Synthetic data is fine, but the business logic must be realistic.
   - Avoid toy examples.

4. **Command Center cleanup matters.**
   - If duplicate “Discovered” labels remain, fix before release.
   - A messy menu damages trust.

5. **Overbuilding Video 4 is the biggest schedule risk.**
   - A single strong 9–12 minute video is better than two okay videos.

---

## 6. Final recommendation to validate

Claude Code should validate or challenge this:

> Lock one official Video 4. Make Revenue Leakage Finder the hero. Use Data Contract Checker, Exception Triage, and Control Evidence Pack as supporting demos. Ship a safe local Python Automation Launcher and controlled SharePoint starter package. Narrow the public starting surface to 5–7 workflows. Park xlwings and enterprise IT deployment until v2.
