# Finance Toolkit Roadmap (Derived from RecTrial + External Build Brief)

## Objective

Build a practical, safe, local-first finance toolkit that follows this flow:

**Clean files → check workbook risk → reconcile → explain variances/exceptions → package evidence.**

---

## What to Adopt Immediately from the External Brief

1. Strong safety model (local-first, no external calls, non-destructive outputs).
2. Phased prioritization instead of trying to build everything at once.
3. v1 focus on 10 high-value tools built well.
4. Template/stub strategy for lower-priority items.
5. Consistent run logs + output folder discipline.

---

## What to Adjust for RecTrial Reality

1. RecTrial already contains many similar capabilities; avoid rebuilding duplicates blindly.
2. Prioritize integration/packaging over net-new tool sprawl.
3. Keep “finance analyst usability” ahead of architectural perfection.
4. Treat VBA/Python/PowerQuery/SQL/DAX as complementary lanes, not competing stacks.

---

## Phase Plan

## Phase 1 — Foundation (Must Be Solid)

1. Finance Automation Launcher
2. Control Evidence Pack Generator
3. Data Contract Checker
4. File/Data Cleaner (safe mode)
5. Workbook Health Check

**Exit criteria:**
- All 5 follow safety standard
- Sample mode works for each
- Output folder + logs consistent

## Phase 2 — Core Finance Workflows

6. Python Reconciliation Engine
7. Power Query Folder Consolidator Template
8. Power Query Unpivot Template
9. Budget vs Actual Variance Generator
10. Flux Analysis Generator

**Exit criteria:**
- At least 3 realistic finance sample scenarios run end-to-end
- Outputs understandable by non-developer analyst

## Phase 3 — High-Demo-Value Additions

11. Revenue Leakage Finder
12. Exception Triage Engine
13. Excel Cell-Level Diff Tool
14. Multi-File/Multi-Sheet Combiner
15. Journal Entry Formatter/Validator

**Exit criteria:**
- Strong before/after demo story
- Ranked output tables for actionability

## Phase 4 — Templates and Governance Enhancers

16. Formula Error Audit
17. External Link Finder/Breaker (detection-first)
18. Hardcoded Formula Detector
19. Named Range Cleaner
20. Sheet TOC Generator
21. Power Query Fuzzy Match Template
22. Three-Bucket Reconciler Template
23. SQL Trial Balance/GL Summary Template
24. SQL Variance Template
25. DAX Budget-vs-Actual Measure Pack

**Exit criteria:**
- Clean templates + simple adaptation instructions
- Explicit limitations documented

---

## Suggested v1 “Build Well” 10

1. Finance Automation Launcher
2. Workbook Health Check
3. File/Data Cleaner
4. Data Contract Checker
5. Python Reconciliation Engine
6. Revenue Leakage Finder
7. Exception Triage Engine
8. Control Evidence Pack
9. Budget vs Actual Variance Generator
10. Excel Cell-Level Diff Tool

---

## Decision Rules (to prevent overbuilding)

- If a similar tool already exists in RecTrial, prefer **hardening and packaging** over rebuilding.
- If a tool is heavy but low business frequency, ship as a **template/stub** first.
- If user cannot explain tool value in 20 seconds, deprioritize for v1.
- If setup friction is high, create a “light mode” first.

---

## Deliverable Shape Recommendation

- Keep current analysis docs in `CodexApril29/`.
- If implementation begins later, use:
  - `CodexApril29/src/`
  - `CodexApril29/tests/`
  - `CodexApril29/sample_data/`
  - `CodexApril29/outputs/`
  - `CodexApril29/logs/`

This keeps implementation isolated from RecTrial and aligns with your folder rule.
