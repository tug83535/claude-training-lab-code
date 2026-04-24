# CodexCodeIdeas — Full Code Ideas Reference

**Source:** `codexreview2/` folder (3 files: 01_branch_review.md, 02_new_automation_backlog.md, 03_prioritized_roadmap.md)

**Purpose:** This document captures every code idea, automation concept, and implementation suggestion found in the codexreview2 review. It is formatted for a Claude chat to read, understand, and act on.

---

## Project Context (What This Is)

This is a **Finance & Accounting enablement platform** built around one advanced Excel workbook, expanded with:

- **VBA macro system** — in-workbook command center, checks, dashboards, imports, exports
- **SQL pipeline** — staging, transformations, validations, allocation logic
- **Python toolkit** — orchestration, forecasting, simulation, reconciliation, CLI automation
- **Training/documentation layer** — guides, runbooks, QA logs for coworker adoption

The project is demo-ready. The next phase is adding **enterprise-grade automation with measurable business outcomes** — not more basic macros.

### What Already Exists (Do Not Duplicate)
- Dashboard/chart generation
- Scenario and sensitivity controls
- Reconciliations and validation checks
- Monte Carlo simulation
- AP matching and fuzzy matching
- Data sanitization and duplicate checks
- CLI orchestration and month-end automation
- Universal toolkit VBA modules (~140+ tools)

---

## Code Ideas — Full Backlog

---

## SECTION A: SQL Automations (Data Platform + Controls)

### SQL-01 — Journal Entry Duplicate Ring Detection
- **Business outcome:** Catches control circumvention patterns
- **Use case:** Detect near-duplicate journal entry patterns split across days, users, or entities
- **Implementation approach:** Graph-like grouping using similarity windows on amount / date / vendor / account
- **Stack:** SQL
- **Priority:** Phase 3 (90+ days)

### SQL-02 — Close Readiness Score View
- **Business outcome:** Prioritizes team effort early in close cycle
- **Use case:** Produce a single score (0–100) per entity per day indicating how ready that entity is to close
- **Implementation approach:** Weighted aggregation of failed validations + missing feeds + late postings; output as a SQL view (`close_readiness_score_vw`)
- **Stack:** SQL
- **Priority:** Phase 1 — build first (highest ROI)
- **Deliverable:** One SQL mart for readiness + exceptions

### SQL-03 — Vendor Payment Velocity Baselines
- **Business outcome:** Fraud/error early warning
- **Use case:** Detect unusual payment timing and amount shifts by vendor cohort
- **Implementation approach:** Rolling medians, MAD/z-score hybrid thresholds per vendor
- **Stack:** SQL
- **Priority:** Backlog (not yet phased)

### SQL-04 — Allocation Drift Tracker
- **Business outcome:** Governance and explainability for margin movement
- **Use case:** Detect silent drift in cost allocation percentages over time
- **Implementation approach:** Monthly delta view with tolerances and reason-code required flags when drift exceeds threshold
- **Stack:** SQL
- **Priority:** Phase 2 (45–90 days)

### SQL-05 — Workbook-to-Source Reconciliation Mart
- **Business outcome:** Faster confidence checks before executive reporting
- **Use case:** Compare workbook aggregate results against warehouse truth
- **Implementation approach:** Standardized reconciliation tables + variance reason taxonomy (categorized buckets)
- **Stack:** SQL
- **Priority:** Backlog

### SQL-06 — Subledger Completeness Control Matrix
- **Business outcome:** Prevents partial-close decisions based on incomplete data
- **Use case:** Ensure all required upstream feeds are present before close steps execute
- **Implementation approach:** Control table with expected feed times and row-count bounds; fail close gate if not met
- **Stack:** SQL
- **Priority:** Backlog

### SQL-07 — Policy-as-Code Rule Engine Tables
- **Business outcome:** Faster policy updates, lower maintenance risk
- **Use case:** Finance policy checks maintained without editing SQL code directly
- **Implementation approach:** Metadata-driven rule catalog table + dynamic execution procedure that reads rules at runtime
- **Stack:** SQL
- **Priority:** Backlog

### SQL-08 — Forecast Backtest Warehouse
- **Business outcome:** Objective model accountability
- **Use case:** Store every forecast run, its assumptions, and the realized actuals for comparison
- **Implementation approach:** Three tables: `forecast_run`, `forecast_assumption`, `forecast_actual`
- **Stack:** SQL
- **Priority:** Phase 2 (45–90 days)

### SQL-09 — Segregation-of-Duties Audit Query Pack
- **Business outcome:** Internal control coverage for audits
- **Use case:** Identify conflicting roles/actions in the transaction lifecycle
- **Implementation approach:** Role-action matrix joins + exception materialized views
- **Stack:** SQL
- **Priority:** Backlog

### SQL-10 — Close Bottleneck Heatmap Dataset
- **Business outcome:** Quantifiable process improvement roadmap
- **Use case:** Show which process steps most often create close cycle delays
- **Implementation approach:** Event timestamps and lag decomposition by step / entity / user
- **Stack:** SQL
- **Priority:** Backlog

---

## SECTION B: Python Automations (Orchestration + Intelligence)

### PY-01 — Exception Triage Engine
- **Business outcome:** Analysts work highest-value issues first
- **Use case:** Auto-rank exceptions by business impact and urgency
- **Implementation approach:** Rule + heuristic scoring formula: `impact * confidence * recency`; scoring weights stored in a config file for easy tuning
- **Stack:** Python
- **Priority:** Phase 1 — build second
- **Deliverable:** `exception_triage.py` with config-driven weights

### PY-02 — Narrative Variance Writer (Controlled)
- **Business outcome:** Faster executive pack drafting with auditable language
- **Use case:** Generate draft commentary using deterministic templates (no AI hallucination risk)
- **Implementation approach:** Template library + metric trigger thresholds + governance checks before output
- **Stack:** Python
- **Priority:** Backlog

### PY-03 — Finance Data Contract Checker
- **Business outcome:** Fewer downstream data breakages
- **Use case:** Schema and quality contract checks on every data handoff between systems
- **Implementation approach:** Declarative YAML contracts + Python validator runner that checks each feed on arrival
- **Stack:** Python
- **Priority:** Backlog

### PY-04 — Forecast Ensemble Manager
- **Business outcome:** Better forecast accuracy stability
- **Use case:** Combine multiple forecast models with weighted governance
- **Implementation approach:** Backtest-based weighting + champion/challenger model registry
- **Stack:** Python
- **Priority:** Phase 2 (45–90 days)

### PY-05 — Root Cause Reconciliation Assistant
- **Business outcome:** Shorter time-to-resolution for reconciliation breaks
- **Use case:** Propose likely cause categories when a reconciliation breaks
- **Implementation approach:** Deterministic rules + similarity matching to prior resolved issues (lookup against historical break/resolution log)
- **Stack:** Python
- **Priority:** Backlog

### PY-06 — Close Calendar Risk Predictor
- **Business outcome:** Proactive staffing and escalation before SLA misses happen
- **Use case:** Predict SLA miss probability for each close task based on current cycle state
- **Implementation approach:** Gradient boosting or logistic baseline model trained on historical cycle data
- **Stack:** Python (scikit-learn / lightweight ML)
- **Priority:** Phase 3 (90+ days)

### PY-07 — Control Evidence Pack Generator
- **Business outcome:** Audit prep hours reduced
- **Use case:** Auto-build audit evidence folders from run logs and validation results
- **Implementation approach:** Manifest builder + signed hash snapshots of outputs; package as zip with index
- **Stack:** Python
- **Priority:** Phase 1 — build fifth
- **Deliverable:** `generate_evidence_pack.py` that packages logs/results into a bundle

### PY-08 — Workbook Dependency Scanner
- **Business outcome:** Safer workbook structural changes
- **Use case:** Parse formulas, named ranges, and connections to map an impact dependency graph
- **Implementation approach:** `openpyxl` formula parser + graph export as JSON and/or interactive HTML
- **Stack:** Python + openpyxl
- **Priority:** Phase 3 (90+ days)

### PY-09 — CFO Pack Assembly Pipeline
- **Business outcome:** Consistent monthly deliverables every period
- **Use case:** Compile approved charts, tables, and commentary into one release artifact
- **Implementation approach:** Controlled templating + release tagging; locks content once approved
- **Stack:** Python
- **Priority:** Backlog

### PY-10 — Data Drift Monitor Service
- **Business outcome:** Avoids "quietly wrong" forecasts and reports going undetected
- **Use case:** Monitor distribution drift in critical metrics and trigger alerts when drift exceeds threshold
- **Implementation approach:** PSI (Population Stability Index) and KS tests + threshold alert output
- **Stack:** Python (scipy / statsmodels)
- **Priority:** Backlog

---

## SECTION C: VBA Automations (Excel Execution Layer)

### VBA-01 — Controlled Action Approvals
- **Business outcome:** Stronger governance in shared workbooks
- **Use case:** Require manager PIN or approval before high-impact macros execute
- **Implementation approach:** Approval gate lookup table + signed action log; action is blocked until approval record exists
- **Stack:** VBA
- **Priority:** Phase 3 (90+ days)

### VBA-02 — Formula Integrity Fingerprinting
- **Business outcome:** Catches silent formula breakage fast
- **Use case:** Detect unauthorized or accidental formula changes in protected/critical cell ranges
- **Implementation approach:** Hash formulas by range at baseline; compare against stored hash on demand or on open
- **Stack:** VBA
- **Priority:** Phase 2 (45–90 days)

### VBA-03 — Intelligent Rollforward Assistant
- **Business outcome:** Fewer period setup errors when rolling month tabs
- **Use case:** Roll month tabs with formula and mapping validation before committing the rollforward
- **Implementation approach:** Preflight checks + staged apply with undo capability if checks fail
- **Stack:** VBA
- **Priority:** Backlog

### VBA-04 — Exception Workbench Sheet
- **Business outcome:** One place for analyst actioning of all exceptions
- **Use case:** Unified triage sheet for all validation and reconciliation failures in one Excel tab
- **Implementation approach:** Import exception data from SQL/Python outputs; add workflow status columns (owner, due date, resolution); include import macro
- **Stack:** VBA + Excel sheet
- **Priority:** Phase 1 — build third
- **Deliverable:** `ExceptionWorkbench` sheet + import macro

### VBA-05 — Dependency Impact Preview
- **Business outcome:** Safer edits and clearer user confidence before making changes
- **Use case:** Show which downstream cells and charts will change before an action executes
- **Implementation approach:** Trace Excel precedents/dependents programmatically + surface in a summary UI popup
- **Stack:** VBA
- **Priority:** Backlog

### VBA-06 — Auto-Repair Suggestions (Not Auto-Apply)
- **Business outcome:** Guided remediation without hidden changes to the workbook
- **Use case:** Recommend fix options for detected data issues; user decides what to apply
- **Implementation approach:** Rules map issue type to proposed remediation options; display as menu choices
- **Stack:** VBA
- **Priority:** Backlog

### VBA-07 — Controlled Snapshot Sign-off
- **Business outcome:** Defensible, locked reporting snapshots for each close period
- **Use case:** Lock and document workbook state at monthly sign-off
- **Implementation approach:** Snapshot metadata capture + checksum of key ranges + approver name/timestamp stored in log sheet
- **Stack:** VBA
- **Priority:** Backlog

### VBA-08 — Macro Runtime Telemetry Dashboard
- **Business outcome:** Identifies slow and failing processes to optimize
- **Use case:** Show runtime, error rates, and usage frequency by Command Center action
- **Implementation approach:** Summarize existing VBA audit log data into a KPI dashboard sheet
- **Stack:** VBA (reads from existing VBA_AuditLog sheet)
- **Priority:** Phase 1 — build fourth

### VBA-09 — Workbook Policy Validator
- **Business outcome:** Standardization across all team workbooks
- **Use case:** Enforce naming standards (named ranges, tab order, required sheets present, font/color standards)
- **Implementation approach:** Policy definitions stored in a config section; compliance report output to a sheet
- **Stack:** VBA
- **Priority:** Phase 2 (45–90 days)

### VBA-10 — Data Entry Fraud Pattern Flags
- **Business outcome:** Additional detective control layer inside Excel
- **Use case:** Flag suspicious manual overrides based on timing + threshold combinations
- **Implementation approach:** Event log of manual cell edits + rule windows that score suspicious patterns
- **Stack:** VBA
- **Priority:** Backlog

---

## SECTION D: Other Automation Worth Considering

### OA-01 — Office Scripts + Power Automate Close Trigger
- **Use case:** Orchestration handoff trigger — not for basic native tasks
- **Approach:** Use Office Script only to trigger Python/SQL runs when files reach controlled states (e.g., file moved to "Ready for Close" folder)
- **Stack:** Office Scripts + Power Automate

### OA-02 — .NET Add-In for Signed Enterprise Deployment
- **Use case:** Organizations that block unsigned macros via Group Policy
- **Approach:** Move critical controls into a managed .NET add-in rather than workbook-level VBA; sign and deploy via IT
- **Stack:** .NET / VSTO

### OA-03 — Lightweight Internal API for Exception Status
- **Use case:** Avoid fragmented exception status scattered across files and emails
- **Approach:** Excel, VBA, and Python all read/write one lightweight exception service endpoint
- **Stack:** Python (Flask/FastAPI lightweight) + VBA HTTP calls
- **Priority:** Phase 3 (90+ days)

### OA-04 — dbt-style Model Layer for Finance SQL
- **Use case:** Versioned, tested, documented transformation DAG as dataset scale grows
- **Approach:** Adopt dbt or a dbt-inspired pattern for SQL transformations; improves maintainability
- **Stack:** dbt (or dbt-like structure)

### OA-05 — GitHub Actions Validation Bundle
- **Use case:** Ensure repo quality before user-facing delivery on every push
- **Approach:** Run lint / tests / data contract checks on every branch update via CI pipeline
- **Stack:** GitHub Actions + Python pytest + SQL linter

---

## Phased Implementation Roadmap

### Phase 1 — First 30–45 Days (Highest ROI / Lowest Friction)
| Order | Item | Type | Key Deliverable |
|-------|------|------|-----------------|
| 1 | Close Readiness Score (SQL-02) | SQL | `close_readiness_score_vw` + exception severity table |
| 2 | Exception Triage Engine (PY-01) | Python | `exception_triage.py` with config-driven weights |
| 3 | Exception Workbench Sheet (VBA-04) | VBA | `ExceptionWorkbench` sheet + import macro |
| 4 | Macro Runtime Telemetry Dashboard (VBA-08) | VBA | KPI dashboard from existing audit log |
| 5 | Control Evidence Pack Generator (PY-07) | Python | `generate_evidence_pack.py` |

**Why Phase 1 first:** Uses existing logs and validation outputs. Delivers visible value to analysts and leadership immediately.

### Phase 2 — 45–90 Days (Governance + Forecast Maturity)
| Order | Item | Type |
|-------|------|------|
| 6 | Allocation Drift Tracker (SQL-04) | SQL |
| 7 | Forecast Backtest Warehouse (SQL-08) | SQL |
| 8 | Forecast Ensemble Manager (PY-04) | Python |
| 9 | Formula Integrity Fingerprinting (VBA-02) | VBA |
| 10 | Workbook Policy Validator (VBA-09) | VBA |

### Phase 3 — 90+ Days (Advanced Risk + Platformization)
| Order | Item | Type |
|-------|------|------|
| 11 | Journal Entry Duplicate Ring Detection (SQL-01) | SQL |
| 12 | Close Calendar Risk Predictor (PY-06) | Python |
| 13 | Workbook Dependency Scanner (PY-08) | Python |
| 14 | Controlled Action Approvals (VBA-01) | VBA |
| 15 | Internal Exception Status API (OA-03) | Other |

---

## Recommended First 5 Build Tickets

These are the fastest path from "excellent demo" to "repeatable business system":

- **Ticket 1:** Build `close_readiness_score_vw` and exception severity table (SQL-02)
- **Ticket 2:** Implement `exception_triage.py` with config-driven weights (PY-01)
- **Ticket 3:** Add `ExceptionWorkbench` Excel sheet + import macro (VBA-04)
- **Ticket 4:** Create `generate_evidence_pack.py` to package logs/results (PY-07)
- **Ticket 5:** Publish `OPERATING_METRICS.md` with metric definitions and owners

---

## Success Metrics to Track Monthly

| Metric | What It Measures |
|--------|-----------------|
| Close cycle duration (hours) | Speed of financial close |
| High-severity exceptions unresolved > 48h | Exception management quality |
| Forecast MAPE by product/entity | Forecast accuracy |
| Number of post-close adjustments | Close quality |
| Audit evidence prep time | Audit readiness efficiency |
| Macro failure rate and median runtime | VBA system health |

---

## Implementation Guardrails (Practical Notes)

1. **One source of truth:** Keep all outputs under `FinalExport/` or a dedicated delivery folder — avoid archive confusion.
2. **Metadata/config layer first:** Before adding more scripts, introduce a small config layer for rules, thresholds, and ownership maps.
3. **One-page guide per feature:** Every new feature needs a plain-English "how to use" guide — preserve the non-technical onboarding advantage.
4. **Branch quality gate:** Add tests + lint + basic static checks before scaling the tool count further (see OA-05).

---

## What This Backlog Intentionally Excludes

These were explicitly ruled out as already covered by native Excel/OneDrive:
- Basic filters, slicers, sorts, pivots
- Autosave / version history / collaboration
- Simple import wizards
- Basic chart creation
- Basic conditional formatting

All ideas above focus on **controls, intelligence, auditability, and scalable operations** that Excel cannot provide natively.
