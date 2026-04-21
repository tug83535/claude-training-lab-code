# New Automation Backlog (SQL, Python, VBA, and Excel-Adjacent)

This backlog is intentionally focused on capabilities that **are not already solved by native modern Excel or OneDrive**.

Each idea includes business outcome, suggested stack, and why it matters.

---

## A) SQL Automations (Data Platform + Controls)

## SQL-01 — Journal Entry Duplicate Ring Detection
- **Use case:** detect near-duplicate JE patterns split across days, users, or entities.
- **Approach:** graph-like grouping using amount/date/vendor/account similarity windows.
- **Outcome:** catches control circumvention patterns.

## SQL-02 — Close Readiness Score View
- **Use case:** a single score (0–100) per entity/day indicating close readiness.
- **Approach:** weighted aggregation of failed validations + missing feeds + late postings.
- **Outcome:** prioritizes team effort early.

## SQL-03 — Vendor Payment Velocity Baselines
- **Use case:** detect unusual payment timing/amount shifts by vendor cohort.
- **Approach:** rolling medians, MAD/z-score hybrid thresholds.
- **Outcome:** fraud/error early warning.

## SQL-04 — Allocation Drift Tracker
- **Use case:** detect silent drift in allocation percentages over time.
- **Approach:** monthly delta view with tolerances and reason-code required flags.
- **Outcome:** governance and explainability for margin movement.

## SQL-05 — Workbook-to-Source Reconciliation Mart
- **Use case:** compare workbook aggregate results to warehouse truth.
- **Approach:** standardized reconciliation tables + variance reason taxonomy.
- **Outcome:** faster confidence checks before executive reporting.

## SQL-06 — Subledger Completeness Control Matrix
- **Use case:** ensure required upstream feeds are present before close steps run.
- **Approach:** control table with expected feed times and row-count bounds.
- **Outcome:** avoids partial-close decisions.

## SQL-07 — Policy-as-Code Rule Engine Tables
- **Use case:** finance policy checks maintained without editing SQL code.
- **Approach:** metadata-driven rule catalog + dynamic execution procedure.
- **Outcome:** faster policy updates, lower maintenance risk.

## SQL-08 — Forecast Backtest Warehouse
- **Use case:** store every forecast run, assumptions, and realized outcomes.
- **Approach:** forecast_run / forecast_assumption / forecast_actual tables.
- **Outcome:** objective model accountability.

## SQL-09 — Segregation-of-Duties Audit Query Pack
- **Use case:** identify conflicting roles/actions in transaction lifecycle.
- **Approach:** role-action matrix joins and exception materialized views.
- **Outcome:** internal control coverage for audits.

## SQL-10 — Close Bottleneck Heatmap Dataset
- **Use case:** show which process steps most often create close delays.
- **Approach:** event timestamps and lag decomposition by step/entity/user.
- **Outcome:** quantifiable process improvement roadmap.

---

## B) Python Automations (Orchestration + Intelligence)

## PY-01 — Exception Triage Engine
- **Use case:** auto-rank exceptions by business impact and urgency.
- **Approach:** rule + heuristic scoring (`impact * confidence * recency`).
- **Outcome:** analysts work highest-value issues first.

## PY-02 — Narrative Variance Writer (Controlled)
- **Use case:** generate draft commentary with deterministic templates.
- **Approach:** template library + metric triggers + governance checks.
- **Outcome:** faster executive pack drafting, auditable language.

## PY-03 — Finance Data Contract Checker
- **Use case:** schema and quality contract checks on every data handoff.
- **Approach:** declarative YAML contracts + Python validator runner.
- **Outcome:** fewer downstream breakages.

## PY-04 — Forecast Ensemble Manager
- **Use case:** combine multiple forecast models with weighted governance.
- **Approach:** backtest-based weighting + champion/challenger registry.
- **Outcome:** better accuracy stability.

## PY-05 — Root Cause Reconciliation Assistant
- **Use case:** propose likely cause categories for reconciliation breaks.
- **Approach:** deterministic rules + similarity matching to prior resolved issues.
- **Outcome:** shorter time-to-resolution.

## PY-06 — Close Calendar Risk Predictor
- **Use case:** predict SLA miss probability for each close task.
- **Approach:** gradient boosting / logistic baseline on historical cycle data.
- **Outcome:** proactive staffing and escalation.

## PY-07 — Control Evidence Pack Generator
- **Use case:** auto-build audit evidence folders from run logs/results.
- **Approach:** manifest builder + signed hash snapshots of outputs.
- **Outcome:** audit prep hours reduced.

## PY-08 — Workbook Dependency Scanner
- **Use case:** parse formulas/names/connections and map impact graph.
- **Approach:** `openpyxl` + graph export (JSON/HTML).
- **Outcome:** safer workbook changes.

## PY-09 — CFO Pack Assembly Pipeline
- **Use case:** compile approved charts/tables/commentary into one release artifact.
- **Approach:** controlled templating + release tagging.
- **Outcome:** consistent monthly deliverables.

## PY-10 — Data Drift Monitor Service
- **Use case:** monitor distribution drift in critical metrics and flags.
- **Approach:** PSI/KS tests + threshold alerts.
- **Outcome:** avoids “quietly wrong” forecasts/reports.

---

## C) VBA Automations (Excel Execution Where Users Work)

## VBA-01 — Controlled Action Approvals
- **Use case:** require manager PIN/approval for high-impact macros.
- **Approach:** approval gate table + signed action log.
- **Outcome:** stronger governance in shared workbooks.

## VBA-02 — Formula Integrity Fingerprinting
- **Use case:** detect unauthorized formula changes in protected zones.
- **Approach:** hash formulas by range and compare against baseline.
- **Outcome:** catches silent breakage fast.

## VBA-03 — Intelligent Rollforward Assistant
- **Use case:** roll month tabs with formula/map validation before commit.
- **Approach:** preflight checks + staged apply/undo.
- **Outcome:** fewer period setup errors.

## VBA-04 — Exception Workbench Sheet
- **Use case:** unified triage sheet for all validation/reconciliation failures.
- **Approach:** import from SQL/Python outputs and add workflow statuses.
- **Outcome:** one place for analyst actioning.

## VBA-05 — Dependency Impact Preview
- **Use case:** show downstream cells/charts that will change before action.
- **Approach:** trace precedents/dependents + summary UI.
- **Outcome:** safer edits and clearer user confidence.

## VBA-06 — Auto-Repair Suggestions (Not Auto-Apply)
- **Use case:** recommend fix options for detected data issues.
- **Approach:** rules map issue type to proposed remediations.
- **Outcome:** guided remediation without hidden changes.

## VBA-07 — Controlled Snapshot Sign-off
- **Use case:** lock/report workbook state at monthly sign-off.
- **Approach:** snapshot metadata + checksum + approver capture.
- **Outcome:** defensible reporting snapshots.

## VBA-08 — Macro Runtime Telemetry Dashboard
- **Use case:** show runtime, error rates, and usage by action.
- **Approach:** summarize existing audit logs into KPI dashboard.
- **Outcome:** identifies slow/failing processes to optimize.

## VBA-09 — Workbook Policy Validator
- **Use case:** enforce standards (named ranges, tab order, required sheets).
- **Approach:** policy definitions + compliance report output.
- **Outcome:** standardization across team workbooks.

## VBA-10 — Data Entry Fraud Pattern Flags
- **Use case:** flag suspicious manual overrides (timing + threshold combos).
- **Approach:** event log + rule windows.
- **Outcome:** additional detective control inside Excel.

---

## D) Other Coding/Automation Worth Considering

## OA-01 — Office Scripts + Power Automate Close Trigger
- Use script only for orchestration handoff, not basic native tasks.
- Trigger Python/SQL runs when files hit controlled states.

## OA-02 — .NET Add-In for Signed Enterprise Deployment
- For organizations blocking unsigned macros.
- Keep critical controls in managed add-in, not workbook-level VBA.

## OA-03 — Lightweight Internal API for Exception Status
- Excel/VBA/Python all read/write one exception service.
- Avoid fragmented status across files and emails.

## OA-04 — dbt-style Model Layer for Finance SQL
- Versioned, tested, documented transformation DAG.
- Improves maintainability once datasets scale.

## OA-05 — GitHub Actions Validation Bundle
- Run lint/tests/data contracts on every branch update.
- Ensures repo quality before user-facing delivery.

---

## Not Included On Purpose (Native/Commodity)
To respect your constraint, this backlog intentionally avoids:

- Basic filters/slicers/sorts/pivots
- Autosave/version history/collaboration features
- Simple import wizards already available in Excel UI
- Basic chart creation features users can do manually

The ideas above focus on **controls, intelligence, auditability, and scalable operations**.
