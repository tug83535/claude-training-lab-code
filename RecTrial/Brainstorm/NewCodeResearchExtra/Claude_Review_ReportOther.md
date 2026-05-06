# Review Plan for Claude Code

## Purpose

This report distills the critical feedback provided about the current iPipeline Finance Automation Demo project so that **Claude Code**—the large language model that constructed the overview and internal guides—can review the suggested adjustments. It highlights what is working, what needs improvement, and the reasoning behind each recommendation.

## High-Priority Feedback

### 1. Reconsider the Video 4 split

- The current plan splits Video 4 into **4a** (executive-focused) and **4b** (analyst-focused).
- Splitting can work only if each video stands alone and has a complete message.
- The current 4a/4b structure risks breaking the narrative: part A focuses on ARR waterfall and workbook dependency scanning, while part B shifts into control evidence and triage recipes.
- A single, cohesive Video 4 is likely stronger if it tells one business story from start to finish.

**Recommended direction:**

Use one consolidated Video 4 narrative centered on revenue leakage / ARR waterfall. Start with hidden workbook risk, move into recurring revenue movement, then end with how analysts can run the tools themselves.

### 2. Focus the hero demo on revenue leakage

- The ARR/MRR waterfall is a good CFO-facing visual.
- The stronger framing is not “look, a waterfall chart.”
- The stronger framing is “Python finds hidden revenue leakage and explains what changed.”
- A Revenue Leakage Finder may be more directly actionable, but the ARR waterfall is easier to explain visually.

**Recommended direction:**

Use the ARR waterfall as the visual hero, but frame it as a revenue leakage analysis.

### 3. Delivery interface matters

- The current `finance_copilot.py` menu is useful technically, but a CLI is not ideal for non-developer finance analysts.
- A command-line menu may still be acceptable for Video 4 if positioned as a v1 internal demo.
- The real adoption path should eventually be Excel button, simple GUI, or web app.
- Do not assume xlwings will be approved on locked-down corporate laptops.

**Recommended direction:**

Build the CLI for now, but document it as the simplest v1. Treat Excel button / GUI / web front end as the adoption version.

### 4. Simplify the universal toolkit

The toolkit currently has a very large surface area: about 140 VBA tools and 28 Python scripts. That is too much for a typical analyst to absorb.

**Recommended starter set:**

- Data cleaning tools
- Data sanitizer
- Highlight tools
- Comment inventory
- Tab organizer
- Column operations
- Sheet index / template clone
- Compare sheets
- Consolidate sheets
- Pivot refresh / pivot inventory
- Invoice duplicate detector
- GL validator
- Trial balance checker
- Workbook profiler / exec brief
- Materiality classifier
- Exception narratives
- Data quality scorecard

**Recommended direction:**

Create a Quick Start surface with the 15–20 most useful tools. Hide the long tail until users need it.

### 5. Distribution and adoption

The biggest project weakness is operational readiness. The code and videos are strong, but distribution is not solved.

**Minimum viable distribution plan:**

1. Put the package in one trusted SharePoint or Box location.
2. Include a Start Here guide.
3. Include a troubleshooting guide.
4. Create a signed Excel add-in if the VBA toolkit is meant to be widely adopted.
5. Package Python scripts so analysts do not need to understand terminal commands.
6. Assign an owner for support and maintenance.
7. Track adoption and issues.

### 6. Post-V4 roadmap

After V4 ships, avoid adding random new tools. Move toward governance and adoption.

**Move up:**

- Dual logging pattern
- `CONSTRAINTS.md`
- `BRAND.md`
- `RELEASE_READINESS_CHECKLIST.md`
- `TROUBLESHOOTING.md`
- Workbook Policy Validator
- Dependency Impact Preview
- Auto-Repair Suggestions

**Keep parked for now:**

- External AI APIs
- ML-heavy forecasting
- Airflow/dbt/infrastructure work
- Outlook/email automation unless there is a clear business case
- Third-party platform exploration unless leadership asks for it

## Clarifications Needed from Claude Code

1. Which existing files already contain partial Video 4 work?
2. Which scripts already exist and should be improved instead of duplicated?
3. Are there existing conventions for CLI arguments, output folders, and sample files?
4. Which docs are source-of-truth: `CLAUDE.md`, `Archive/tasks/todo.md`, or the `RecTrial/Brainstorm` docs?
5. Is the current repo clean enough for Codex and Claude Code to work on separate branches safely?
6. Which items from the Codex comparison are still unported and worth revisiting?

## Recommended Claude Code Task

Claude Code should perform a rigorous review before building more code.

### Step 1 — Inspect

Read:

- `CLAUDE.md`
- `Archive/tasks/todo.md`
- `Archive/tasks/lessons.md`
- `RecTrial/Brainstorm/VIDEO_4_CURRENT_PROPOSAL.md`
- `RecTrial/Brainstorm/VIDEO_4_DRAFT_IDEAS.md`
- `RecTrial/Brainstorm/FUTURE_AUTOMATION_IDEAS.md`
- `RecTrial/CodexCompare/COMPARISON_REPORT.md`
- `RecTrial/CodexCompare/CHERRY_PICK_TRACKER.md`
- `RecTrial/UniversalToolkit/python/`
- `RecTrial/Video4DemoFiles/`
- `FinalExport/`

### Step 2 — Decide

Make these calls explicitly:

- Single Video 4 vs split 4a/4b
- ARR Waterfall vs Revenue Leakage Finder
- CLI-only vs CLI plus future Excel/GUI path
- Which 15–20 tools are the first adoption set
- What gets parked until after V4

### Step 3 — Build only after deciding

Build only the files needed for the chosen V4 direction.

Likely targets:

- `saas_arr_waterfall.py`
- `workbook_dependency_scanner.py`
- `data_contract_checker.py`
- `exception_triage_engine.py`
- `control_evidence_pack.py`
- `finance_copilot.py`
- Video 4 runbook
- Start Here guide
- Troubleshooting guide
- Release checklist

## Conclusion

The project does not need more raw feature volume. It needs sharper product packaging, a clearer Video 4 story, and a distribution plan that real analysts can follow.

The strongest next move is to make Video 4 a single revenue-leakage story, build the minimum Python tools needed to support that story, and then package the output for adoption instead of continuing to expand the toolkit surface area.
