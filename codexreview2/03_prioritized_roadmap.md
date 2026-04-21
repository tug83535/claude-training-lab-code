# Prioritized Roadmap for a Large Software Business

## Objective
Turn this branch from a strong demo + toolkit into a repeatable finance automation operating model for broad coworker use.

---

## Prioritization Framework
Each candidate was ranked by:
1. Business impact on close quality/speed.
2. Implementation effort in current stack.
3. Change-management complexity.
4. Audit/control value.

---

## Phase 1 (First 30–45 Days) — Highest ROI / Lowest Friction

1. **Close Readiness Score (SQL-02)**
2. **Exception Triage Engine (PY-01)**
3. **Exception Workbench Sheet (VBA-04)**
4. **Macro Runtime Telemetry Dashboard (VBA-08)**
5. **Control Evidence Pack Generator (PY-07)**

### Why this phase first
- Immediately improves daily workflow.
- Uses existing logs/validation outputs.
- Delivers visible value to both analysts and leadership.

### Suggested deliverables
- One SQL mart for readiness + exceptions.
- One Python triage script with scoring rules in config.
- One Excel tab for exception ownership workflow.
- One monthly evidence bundle zip output.

---

## Phase 2 (45–90 Days) — Governance + Forecast Maturity

6. **Allocation Drift Tracker (SQL-04)**
7. **Forecast Backtest Warehouse (SQL-08)**
8. **Forecast Ensemble Manager (PY-04)**
9. **Formula Integrity Fingerprinting (VBA-02)**
10. **Workbook Policy Validator (VBA-09)**

### Expected gains
- Measurable forecast accountability.
- Better confidence in workbook integrity.
- Cleaner governance and reduced manual review effort.

---

## Phase 3 (90+ Days) — Advanced Risk and Platformization

11. **Journal Entry Duplicate Ring Detection (SQL-01)**
12. **Close Calendar Risk Predictor (PY-06)**
13. **Workbook Dependency Scanner (PY-08)**
14. **Controlled Action Approvals (VBA-01)**
15. **Internal Exception Status API (OA-03)**

### Expected gains
- Early warning system for close delays and data risk.
- Stronger enterprise controls.
- Less reliance on single workbook mechanics.

---

## Implementation Notes (Practical)

- Keep **one source of truth** for outputs under `FinalExport/` or a dedicated delivery folder to avoid archive confusion.
- Introduce a **small metadata/config layer** before adding more scripts (rules, thresholds, ownership maps).
- Preserve the current non-technical onboarding advantage: every new feature should have a one-page “how to use” guide.
- Add a branch-level quality gate (tests + lint + basic static checks) before scaling tool count.

---

## Success Metrics (Track Monthly)

- Close cycle duration (hours).
- Count of high-severity exceptions unresolved > 48h.
- Forecast MAPE by product/entity.
- Number of post-close adjustments.
- Audit evidence prep time.
- Macro failure rate and median runtime.

---

## Recommended First Build Ticket Set

- Ticket 1: Build `close_readiness_score_vw` and exception severity table.
- Ticket 2: Implement `exception_triage.py` with config-driven weights.
- Ticket 3: Add `ExceptionWorkbench` Excel sheet + import macro.
- Ticket 4: Create `generate_evidence_pack.py` to package logs/results.
- Ticket 5: Publish `OPERATING_METRICS.md` with metric definitions and owners.

This ticket set is the fastest path from “excellent demo” to “repeatable business system.”
