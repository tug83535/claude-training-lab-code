# Report 2 — Codex Build Specification
## iPipeline Finance Automation Demo Project — Video 4 Python Build

**Prepared for:** Codex  
**Purpose:** Build the revised Video 4 Python automation package safely and cleanly.  
**Source overview:** `2026-04-24_ipipeline-finance-automation-demo-master-overview.md`  
**Corrected audience:** 50–150 coworkers, not full 2,000-person deployment.  
**Primary goal:** coworker training and showing what is possible.  
**Owner/support:** Connor.

---

## 0. Build philosophy

This is not an enterprise platform build. This is a safe local demo/training package.

Priorities, in order:

1. **Safety**
2. **Clarity**
3. **Demo reliability**
4. **Useful outputs**
5. **Low dependency burden**
6. **Polish**
7. **Extensibility**

Do not optimize for cleverness. Optimize for a coworker being able to run a sample safely and understand what happened.

---

## 1. Non-negotiable safety requirements

All new v1 Python scripts should follow these rules:

1. **No internet calls**
   - No `requests`
   - No API calls
   - No telemetry
   - No package downloads
   - No external AI calls

2. **No credentials**
   - No passwords
   - No tokens
   - No database connection strings
   - No environment variable secrets

3. **No destructive source-file behavior**
   - Never edit input files in place.
   - Never overwrite source files.
   - Never delete files.
   - Never move user files.

4. **Outputs only in controlled output directories**
   - Use timestamped run folders:
     ```text
     outputs/YYYYMMDD_HHMMSS_<tool_name>/
     ```

5. **Logs for every run**
   - `run_log.json`
   - `run_summary.txt`
   - Include:
     - timestamp;
     - tool name;
     - input file path(s);
     - output file path(s);
     - status;
     - row counts;
     - warnings;
     - errors.

6. **Sample mode**
   - Every major tool must support `--sample`.

7. **Dry run where applicable**
   - Use `--dry-run` for tools that would create many outputs.

8. **Clear user-facing errors**
   - Main screen should show readable errors.
   - Detailed stack traces can go to logs if useful.

9. **Standard library first**
   - Prefer standard library only for new Video 4 scripts.
   - Use CSV/JSON/HTML/SVG outputs instead of requiring pandas/matplotlib.
   - If existing project scripts use dependencies, do not expand the dependency footprint for the new v1 demo unless necessary.

10. **No auto-opening files unless optional**
    - If implemented, make it opt-in via `--open-output`.
    - Default should simply print output path.

---

## 2. Recommended file locations

Use existing project structure where possible.

Preferred:

```text
RecTrial/
  UniversalToolkit/
    python/
      ZeroInstall/
        finance_automation_launcher.py
        revenue_leakage_finder.py
        data_contract_checker.py
        exception_triage_engine.py
        control_evidence_pack.py
        workbook_dependency_scanner.py
        common/
          safe_io.py
          logging_utils.py
          sample_data.py
          report_utils.py
        samples/
          revenue_leakage/
          data_contract/
          exception_triage/
          control_evidence/
          workbook_dependency/
        outputs/
        logs/
        PYTHON_SAFETY.md
        README_VIDEO4_PYTHON.md
```

If repository conventions require a different path, adapt but preserve:
- `ZeroInstall`
- `samples`
- `outputs`
- `common`
- visible safety/readme docs.

---

## 3. Scripts to build

### 3.1 `finance_automation_launcher.py`

User-facing name:
**Finance Automation Launcher**

Purpose:
- A simple menu that lets coworkers run safe sample workflows.
- It should launch the new Video 4 scripts.
- It may also list existing scripts, but the default menu should show only supported starter workflows.

Default menu:

```text
Finance Automation Launcher
Local-only. No internet. Inputs are read-only. Outputs go to /outputs.

1. Run sample Revenue Leakage Finder
2. Run sample Data Contract Checker
3. Run sample Exception Triage Engine
4. Run sample Control Evidence Pack
5. Run sample Workbook Dependency Scanner
6. Show safety rules
7. Show output folder
8. Exit
```

CLI options:
```text
python finance_automation_launcher.py
python finance_automation_launcher.py --sample revenue_leakage
python finance_automation_launcher.py --list-tools
python finance_automation_launcher.py --show-safety
```

Acceptance criteria:
- Works from Command Prompt.
- Does not require external packages.
- Does not crash if sample files are missing; instead instructs user how to regenerate them.
- Prints output folder after every run.
- Shows safety disclaimer on startup.

---

### 3.2 `revenue_leakage_finder.py`

Purpose:
- Hero demo for Video 4.
- Compare expected revenue vs actual billed revenue.
- Produce ranked exceptions and executive summary.

Inputs:
- `contracts.csv`
- `billing.csv`
- optional `customer_map.csv`
- optional `product_map.csv`

Synthetic sample data should be realistic enough to avoid toy-demo feel.

Recommended sample columns:

`contracts.csv`
```text
customer_id,customer_name,product,contract_start,contract_end,expected_mrr,price_increase_pct,billing_frequency,status
```

`billing.csv`
```text
invoice_id,customer_id,customer_name,product,billing_period,amount_billed,invoice_date,status
```

Core logic:
- Match expected MRR to billed amount by customer/product/period.
- Identify:
  - underbilling;
  - overbilling;
  - missing billing;
  - inactive customer billed;
  - product mismatch;
  - possible duplicate billing;
  - stale contract;
  - missing customer mapping.
- Calculate:
  - expected amount;
  - actual amount;
  - variance;
  - variance percent;
  - severity;
  - confidence;
  - recommended action.

Outputs:
```text
outputs/YYYYMMDD_HHMMSS_revenue_leakage/
  revenue_leakage_summary.html
  revenue_leakage_exceptions.csv
  revenue_leakage_summary.json
  top_10_action_list.csv
  run_log.json
  run_summary.txt
```

HTML report should include:
- total expected revenue;
- total billed revenue;
- net variance;
- estimated underbilling;
- estimated overbilling;
- exception count;
- top 10 exceptions;
- breakdown by exception type;
- plain-English interpretation.

CLI:
```text
python revenue_leakage_finder.py --sample
python revenue_leakage_finder.py --contracts inputs/contracts.csv --billing inputs/billing.csv
python revenue_leakage_finder.py --contracts ... --billing ... --output-dir outputs/
```

Acceptance criteria:
- Standard library only.
- Generates sample data if `--sample` is used and sample data does not exist.
- Does not modify input CSVs.
- Produces readable HTML report.
- Produces ranked CSV for analysts.
- Handles missing columns with clear error message.
- Handles empty files gracefully.
- Uses timestamped output folder.

---

### 3.3 `data_contract_checker.py`

Purpose:
- Validate whether input files have the expected schema before analysis.

Demo pattern:
- Show red FAIL.
- Fix bad file or run good sample.
- Show green PASS.

Inputs:
- CSV file.
- Schema JSON file.

Example schema:
```json
{
  "file_type": "billing_export",
  "required_columns": {
    "invoice_id": "string",
    "customer_id": "string",
    "billing_period": "date",
    "amount_billed": "number"
  },
  "optional_columns": {
    "product": "string",
    "status": "string"
  },
  "rules": [
    {"column": "amount_billed", "rule": "not_negative"},
    {"column": "billing_period", "rule": "not_blank"}
  ]
}
```

Outputs:
```text
outputs/YYYYMMDD_HHMMSS_data_contract/
  data_contract_report.html
  data_contract_results.csv
  data_contract_summary.json
  run_log.json
  run_summary.txt
```

Acceptance criteria:
- Clearly flags missing columns.
- Clearly flags type issues.
- Clearly flags blank required values.
- Exit code should be:
  - `0` for pass;
  - `1` for validation fail;
  - `2` for script/config error.
- Standard library only.

---

### 3.4 `exception_triage_engine.py`

Purpose:
- Rank exceptions by impact, confidence, recency, and type.

Inputs:
- CSV of exceptions, ideally compatible with Revenue Leakage Finder output.

Recommended columns:
```text
exception_id,customer_id,customer_name,exception_type,amount,confidence,last_seen_date,status,owner
```

Scoring:
```text
priority_score = impact_score * 0.45 + confidence_score * 0.30 + recency_score * 0.15 + repeat_score * 0.10
```

Outputs:
```text
outputs/YYYYMMDD_HHMMSS_exception_triage/
  exception_triage_report.html
  ranked_exceptions.csv
  top_10_action_list.csv
  run_log.json
  run_summary.txt
```

Acceptance criteria:
- Standard library only.
- Produces ranked output.
- Produces plain-English recommended action.
- Works with `--sample`.
- Does not modify source file.

---

### 3.5 `control_evidence_pack.py`

Purpose:
- Generate a simple audit/control evidence bundle from input files and run outputs.

Inputs:
- One or more files/folders.
- Optional control name.
- Optional owner.

Outputs:
```text
outputs/YYYYMMDD_HHMMSS_control_evidence/
  manifest.csv
  file_hashes.csv
  evidence_summary.html
  evidence_readme.txt
  run_log.json
  run_summary.txt
```

Core logic:
- Record file names, sizes, modified timestamps, SHA-256 hashes.
- Create manifest.
- Create summary report.
- Do not alter source files.

CLI:
```text
python control_evidence_pack.py --sample
python control_evidence_pack.py --input-dir outputs/some_previous_run --control-name "Revenue Leakage Review"
```

Acceptance criteria:
- Standard library only.
- Uses SHA-256 hashing.
- Handles folders recursively if requested.
- Does not copy sensitive source files by default unless explicitly told.
- Default should produce metadata/evidence summary, not duplicate raw data.

---

### 3.6 `workbook_dependency_scanner.py`

Purpose:
- Opener demo showing Python can inspect workbook structure differently than Excel.
- If standard library only, true `.xlsx` formula parsing is limited because `.xlsx` is zipped XML. That is still feasible with `zipfile` and `xml.etree.ElementTree`.
- If `.xlsm` parsing is too complex, scope to `.xlsx` sample workbooks for v1.

Inputs:
- `.xlsx` workbook.

Core logic:
- Open workbook as zip.
- Parse worksheets XML.
- Extract formula cells.
- Identify cross-sheet references using regex.
- Generate dependency list and simple HTML/SVG/CSV report.

Outputs:
```text
outputs/YYYYMMDD_HHMMSS_workbook_dependency/
  workbook_dependencies.html
  formula_inventory.csv
  sheet_dependency_edges.csv
  run_log.json
  run_summary.txt
```

Acceptance criteria:
- Standard library only.
- Works on provided sample workbook.
- If workbook structure is unsupported, fail gracefully.
- Do not edit workbook.
- Keep demo short; this is not the hero.

---

## 4. Common utility modules

Build shared utility modules to avoid copy/paste.

### `common/safe_io.py`

Responsibilities:
- create timestamped output directories;
- prevent source file overwrite;
- normalize relative/absolute paths;
- validate file existence;
- validate extension;
- safe CSV read helper;
- safe CSV write helper;
- safe JSON write helper;
- safe HTML write helper.

### `common/logging_utils.py`

Responsibilities:
- run start/end tracking;
- JSON log output;
- text summary output;
- warning/error capture;
- simple user-facing print helpers.

### `common/report_utils.py`

Responsibilities:
- minimal HTML report wrapper;
- simple tables;
- summary metric cards;
- escaping user-provided values with `html.escape`.

### `common/sample_data.py`

Responsibilities:
- generate realistic sample CSVs for each tool;
- use deterministic seed;
- avoid fake real customer names if that matters;
- create enough rows to look credible.

---

## 5. Demo data requirements

Synthetic data should feel like real Finance data but not expose actual company/customer details.

Use fake customer names:
- Northstar Insurance Group
- Harbor Life Systems
- Pioneer Benefits Co.
- Summit Policy Services
- Keystone Admin Solutions
- Atlas Brokerage Network

Use realistic fields:
- customer IDs
- product names
- billing periods
- invoice IDs
- expected MRR
- actual billed amounts
- contract status
- renewal dates
- exception types

Include intentional issues:
- underbilling;
- overbilling;
- missing invoice;
- duplicate invoice;
- inactive customer billed;
- product mismatch;
- missing customer ID;
- stale contract;
- amount formatting issue.

---

## 6. Documentation to create/update

Create:

```text
PYTHON_SAFETY.md
README_VIDEO4_PYTHON.md
SUPPORTED_WORKFLOWS_V1.md
```

Minimum content for `PYTHON_SAFETY.md`:
- local-only;
- no internet;
- no AI/API;
- read-only inputs;
- outputs folder;
- logs;
- no credentials;
- sample mode;
- limitations;
- who to contact.

Minimum content for `README_VIDEO4_PYTHON.md`:
- what this is;
- who it is for;
- how to run sample mode;
- how to find outputs;
- how to troubleshoot;
- what not to do.

Minimum content for `SUPPORTED_WORKFLOWS_V1.md`:
- 5–7 workflows;
- tool/script names;
- intended audience;
- risk level;
- support status.

---

## 7. Testing requirements

Add lightweight tests if project conventions allow.

At minimum:
- run each script with `--sample`;
- verify output folder created;
- verify expected files exist;
- verify no source files modified;
- verify logs exist;
- verify missing-column handling;
- verify launcher menu can call each sample workflow.

Suggested smoke test script:
```text
smoke_test_video4_python.py
```

Acceptance criteria:
```text
python smoke_test_video4_python.py
```

Should print:
```text
PASS revenue_leakage_finder
PASS data_contract_checker
PASS exception_triage_engine
PASS control_evidence_pack
PASS workbook_dependency_scanner
PASS finance_automation_launcher
```

---

## 8. Build order

Build in this order:

1. `common/` utilities.
2. sample data generator.
3. `data_contract_checker.py`.
4. `revenue_leakage_finder.py`.
5. `exception_triage_engine.py`.
6. `control_evidence_pack.py`.
7. `workbook_dependency_scanner.py`.
8. `finance_automation_launcher.py`.
9. docs.
10. smoke test.
11. final cleanup.

Reason:
- Data Contract Checker and safe IO establish core patterns.
- Revenue Leakage Finder is the hero and should reuse those patterns.
- Launcher should be built last so it calls stable tools.

---

## 9. Definition of done

The build is done only when:

1. Every new script runs in sample mode.
2. Every script is standard-library only or clearly documents any dependency.
3. No script calls the internet.
4. No script modifies input files.
5. Every run creates a timestamped output folder.
6. Every run creates logs.
7. The Revenue Leakage Finder produces a credible HTML report and CSV exception list.
8. The launcher can run all sample workflows.
9. `PYTHON_SAFETY.md` exists and is accurate.
10. `README_VIDEO4_PYTHON.md` explains how a non-developer starts.
11. Smoke test passes.
12. Any known limitation is documented instead of hidden.

---

## 10. Do not build these in v1

Do not build unless Connor explicitly overrides:

1. xlwings Excel Button Edition.
2. external AI/LLM API integrations.
3. database connections.
4. email automation.
5. scheduled automation.
6. Power Automate integration.
7. Streamlit/Dash app.
8. Flask/FastAPI service.
9. ML forecasting or anomaly detection.
10. anything requiring admin install.

These are v2/future, not Video 4 blockers.

---

## 11. Expected final response from Codex

When done, provide:

1. Files created/changed.
2. How to run each script.
3. How to run the launcher.
4. How to run the smoke test.
5. Safety guarantees implemented.
6. Known limitations.
7. Suggested Video 4 demo path using the built outputs.
