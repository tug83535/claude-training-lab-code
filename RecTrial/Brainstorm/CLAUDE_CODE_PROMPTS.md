# Claude Code Prompts — iPipeline Finance Automation Demo
## Ready-to-paste prompts for each remaining build task

**Branch this file lives on:** `copilot/codexreview-rectrial-folder-and-suggest-ideas-za5j`
*(This is the Copilot agent branch from the April 29 session. The April23CLD branch has an
earlier RecTrial snapshot. After merging or downloading this file, copy it wherever you need it.)*

**How to use:** Copy the prompt for the task you want to do, open a fresh Claude Code session,
and paste it in as your first message. Each prompt is self-contained — it includes all the
context Claude needs to produce the right output.

**Order matters:** Do Tier 1 first (planning docs), then Tier 2 (Python build), then Tier 3
(post-V4 items). Do not build Tier 2 code until the Tier 1 planning docs are approved by Connor.

---

## TIER 1 — PLANNING DOCS (Do first, before any code)

These five docs must be written and approved before any V4 Python scripts are touched.
Connor has already locked all 5 decisions (see `HANDOFF_2026-04-27.md` for details).

---

### PROMPT 1-A — VIDEO_4_REVIEW_DECISION_MEMO.md

```
I need you to write a planning document called VIDEO_4_REVIEW_DECISION_MEMO.md.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline (SaaS, insurance industry).
- I'm building a Finance automation package (Excel VBA + Python) for ~50–150 coworkers.
- The prior Video 4 plan (CMD-based 10-script demo, 4a+4b split) was pulled and a new
  direction is now locked.
- The five locked decisions are:
    1. Audience = 50–150 coworkers (not 2,000+ employees / CFO/CEO)
    2. Single Video 4, 9–12 min, chaptered (kills the 4a + 4b split)
    3. Hero = Revenue Leakage Finder as story; ARR Waterfall as closing visual artifact
    4. Public surface = 5–7 supported starter workflows (not "140 tools")
    5. Deliverable = finance_automation_launcher.py (renamed from finance_copilot.py
       because "Copilot" implies AI and the tool has none)

WHAT TO WRITE:
Write VIDEO_4_REVIEW_DECISION_MEMO.md. Save it to:
    RecTrial\Brainstorm\VIDEO_4_REVIEW_DECISION_MEMO.md

Required sections:
1. Header — date, version, author, status (LOCKED)
2. The five locked decisions — one paragraph each explaining WHAT was locked and WHY
3. What got cut — brief list of things that were in prior V4 plans but are now explicitly out:
   - 4a / 4b split
   - finance_copilot.py name
   - 140-tool public surface as the lead message
   - CFO/CEO primary audience
   - xlwings in V1
   - SOX Evidence Collector in V1 scope
4. Stale reference table — a table listing places in existing docs/code that still use old
   framing. Include columns: File | Section/Line | Stale text | What it should say instead.
   Check these locations for stale language:
   - CLAUDE.md (audience, CFO/CEO, 4a/4b references, finance_copilot.py, 140 tools)
   - tasks/todo.md (V4 replanning section)
   - RecTrial\Brainstorm\VIDEO_4_CURRENT_PROPOSAL.md (entire file is superseded)
   - RecTrial\PROJECT_OVERVIEW.md (audience section, Video 4 section)
   - Any references to finance_copilot.py in Python scripts or README files
5. What changed and why — a brief narrative (1–2 paragraphs) explaining the journey from
   the original V4 plan to the new direction, written for a future reader who wasn't there.

STYLE:
- Plain English. The audience for this doc is Connor + any future Claude session reading it.
- No jargon. Numbered sections. Tables where appropriate.
- Target length: 150–250 lines.
- Do NOT write any Python or VBA code as part of this task.
```

---

### PROMPT 1-B — VIDEO_4_REVISED_PLAN.md

```
I need you to write a production plan document called VIDEO_4_REVISED_PLAN.md.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- Video 4 is called "Python Automation for Finance." It is a single 9–12 min chaptered video.
- The hero is Revenue Leakage Finder ("Python found a possible billing problem").
- The ARR Waterfall chart is the closing visual artifact.
- The deliverable coworkers get is finance_automation_launcher.py (CLI menu, zero install).
- The audience is 50–150 coworkers in Finance, Accounting, and adjacent operations.
- All decisions are locked. Do not re-open them.
- The 6 Python scripts for V4 are:
    1. revenue_leakage_finder.py (hero — finds billing/contract gaps)
    2. data_contract_checker.py (validates file structure against a template)
    3. exception_triage_engine.py (flags transactions that need review)
    4. control_evidence_pack.py (generates an audit evidence folder)
    5. workbook_dependency_scanner.py (maps which cells/files reference what)
    6. finance_automation_launcher.py (CLI menu — entry point for coworkers)
- The 8 protected scripts (do NOT modify these until after Video 4 records):
    aging_report.py, bank_reconciler.py, compare_files.py, forecast_rollforward.py,
    fuzzy_lookup.py, pdf_extractor.py, variance_analysis.py, variance_decomposition.py

WHAT TO WRITE:
Write VIDEO_4_REVISED_PLAN.md. Save it to:
    RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md

Required sections:
1. Header — date, version, author, status (APPROVED)
2. Video overview — title, target length, format (chaptered), deliverable name
3. Chapter outline — 8 chapters with estimated time per chapter:
    Ch 1: Why Python after Excel/VBA — ~45 sec
    Ch 2: Safety first (show PYTHON_SAFETY.md) — ~60 sec
    Ch 3: Hero — Revenue Leakage Finder — ~2.5–3.5 min
    Ch 4: Data Contract Checker — ~90 sec
    Ch 5: Exception Triage Engine — ~90 sec
    Ch 6: Control Evidence Pack — ~90 sec
    Ch 7: The Launcher (finance_automation_launcher.py) — ~60 sec
    Ch 8: How to start — ~30 sec
4. Demo sequence — per chapter: what the coworker sees on screen, what Connor does,
   what the script outputs, what the narration says (1–2 sentence summary per chapter)
5. Sample data requirements — what fake/demo data is needed for each script demo,
   what makes the Revenue Leakage example feel real (realistic contract/billing structure,
   not toy numbers)
6. Build effort estimate — per script: already built or needs building, estimated hours
7. Recording effort estimate — total estimated time to record + edit + finalize
8. Known tradeoffs — what's good about this plan, what we're giving up, what risks remain
9. Optional recipe shorts (non-canonical, post-V4) — brief note on short follow-up videos
   that could walk through individual scripts in 2–3 min each

STYLE:
- Plain English. Numbered sections. Tables where appropriate.
- Target length: 200–300 lines.
- Do NOT write any Python or VBA code as part of this task.
```

---

### PROMPT 1-C — PYTHON_SAFETY.md

```
I need you to write a Python safety document called PYTHON_SAFETY.md.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- I'm distributing Python scripts to ~50–150 coworkers via a SharePoint zip package.
- Most coworkers are not technical. Some are. IT/security may review this doc.
- Connor owns support personally. If a script does something unexpected, Connor gets the call.
- The scripts are finance automation tools: file comparison, data cleaning, revenue leak
  detection, exception triage, audit evidence packaging, workbook scanning.
- All V4 scripts use Python standard library only (no third-party packages like pandas).
- The scripts run from a bundled Python 3.11 embeddable — no install required.
- These are the safety rules already designed into all V4 scripts:
    1. No internet calls
    2. No external AI/API calls
    3. No credentials, tokens, secrets, or database connections
    4. Standard library only for new V1 scripts where feasible
    5. Input files are read-only — scripts never modify the source file
    6. Scripts never overwrite source files
    7. Outputs go to timestamped folders under /outputs/
    8. Every run writes a log
    9. Sample/demo mode is available on all scripts
   10. Clear failure messages shown to user in plain English
   11. Detailed technical error info goes to log file, not main screen
   12. Batch or destructive operations require explicit confirmation before running
   13. Logs avoid storing raw sensitive data
   14. Launcher shows a visible safety disclaimer before running any script

WHAT TO WRITE:
Write PYTHON_SAFETY.md. Save it to:
    RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md

Required sections:
1. Header — title, version, date, audience (coworkers + IT/security reviewers)
2. In plain English — what these scripts do (2–3 sentences, non-technical)
3. What these scripts do NOT do — a clear, plain-English list of what is explicitly
   ruled out: no internet, no AI, no sending data anywhere, no overwriting your files, etc.
4. The 14 safety rules — one rule per line, plain English, readable by non-developers
5. What to do if something goes wrong — where the log file is, how to read it,
   who to contact (Connor Atlee, Finance — so coworkers know who to reach)
6. For IT / security reviewers — a brief section explaining the architecture:
   bundled Python, no registry changes, no internet access, outputs only go to local folders,
   scripts can be reviewed as plain .py text files
7. Version history (start with v1.0)

STYLE:
- Plain English throughout. No jargon. Write as if the reader is smart but not technical.
- Bold the key rules so they stand out.
- Target length: 80–120 lines.
- Do NOT write any Python code as part of this task.
```

---

### PROMPT 1-D — MINIMUM_DISTRIBUTION_PLAN.md

```
I need you to write a distribution plan document called MINIMUM_DISTRIBUTION_PLAN.md.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- I'm distributing a Finance automation package to ~50–150 coworkers via SharePoint.
- The package is a single zip file: FinanceTools_v1.0.zip
- Contents of the zip:
    FinanceTools.xlsm (Excel workbook with all VBA tools — ~140 tools in Command Center)
    python/ folder (bundled Python 3.11 embeddable + 6 V4 scripts + common/ utilities)
    samples/ folder (demo input files for each script)
    docs/ folder (PYTHON_SAFETY.md + quick start guide)
    README.txt (plain English — what's in here and how to start)
- Connor owns support personally. Must limit support load.
- The plan is: pilot with 10–20 users first, then broader 50–150 rollout.
- Pilot audience: 3–5 Finance + 3–5 Accounting + 3–5 Billing/RevOps + 1–2 managers
  + optional IT/security observer

WHAT TO WRITE:
Write MINIMUM_DISTRIBUTION_PLAN.md. Save it to:
    RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md

Required sections:
1. Header — date, version, author
2. Package contents — the 8-file/folder layout of FinanceTools_v1.0.zip with a brief
   description of each item
3. SharePoint setup — where to host the zip (Finance team SharePoint), folder/page structure,
   who gets the link, versioning approach (v1.0, v1.1, etc.)
4. Launch message — a ready-to-paste email: subject line + body. Written for coworkers.
   Friendly tone. Should explain what it is, why it's useful, how to get started in 3 steps.
5. Support expectations — what Connor will and won't support, how to report issues
   (email Connor directly), expected response time, what to do if a script crashes
6. Pilot plan — target audience (named roles, not names), how to invite them, what to ask
   them to try, how to collect feedback (simple: email Connor their top 3 questions)
7. Pilot success metrics:
   - 10 people open the package
   - 5 people run at least one script on a sample file
   - 3 people try it on a real file
   - Top 3 confusing points identified
   - Top 3 bugs found and fixed or documented
   - 2 concrete use cases captured (what did it help you do faster?)
8. Release gate — 11-checkpoint table. Columns: Checkpoint | Done? | Notes.
   Checkpoints:
   1. All 6 V4 scripts pass smoke tests
   2. finance_automation_launcher.py menu works end-to-end
   3. PYTHON_SAFETY.md written and reviewed by Connor
   4. ZIP package assembled and tested on a clean machine
   5. SharePoint page created and link tested
   6. Launch email drafted and reviewed by Connor
   7. README.txt written
   8. Pilot audience identified and invited
   9. Pilot feedback collected (min 3 users responded)
  10. Pilot bugs fixed or documented
  11. Final zip version tagged as v1.0 release

STYLE:
- Plain English. Numbered sections. Table for the release gate.
- Target length: 150–200 lines.
- Do NOT write any Python or VBA code as part of this task.
```

---

### PROMPT 1-E — SUPPORTED_WORKFLOWS_V1.md

```
I need you to write a supported workflows document called SUPPORTED_WORKFLOWS_V1.md.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- The Finance automation package supports ~140 VBA tools (all discoverable in the
  Command Center inside FinanceTools.xlsm) plus 6 Python scripts.
- For V1 adoption, we are publicly "leading with" only 5–7 starter workflows.
  The rest of the toolkit stays discoverable but is not part of the pitch.
- The 7 supported starter workflows for V1 are:
    1. Clean a messy Excel export
    2. Compare two files for differences
    3. Consolidate sheets or files into one
    4. Find workbook issues (broken links, errors, hidden sheets)
    5. Generate a workbook summary report
    6. Find possible revenue leakage (NEW — Python)
    7. Check file structure against a template (NEW — Python)
- The existing VBA modules that cover workflows 1–5:
    - Data Sanitizer / Data Cleaning: modUTL_DataSanitizer.bas
    - Sheet Compare: modUTL_Compare.bas
    - Quick Row Compare: modUTL_Compare.bas (UTL_QuickRowCompareCount)
    - Consolidate: modUTL_Consolidate.bas
    - Audit tools / external links: modUTL_Audit.bas + modAuditTools_v2.1.bas
    - Exec Brief: modUTL_ExecBrief.bas + modExecBrief_v2.1.bas
    - Profile Workbook: profile_workbook.py (Python, stdlib-only)
- The 2 NEW Python scripts for workflows 6–7:
    - revenue_leakage_finder.py
    - data_contract_checker.py

WHAT TO WRITE:
Write SUPPORTED_WORKFLOWS_V1.md. Save it to:
    RecTrial\Brainstorm\SUPPORTED_WORKFLOWS_V1.md

For each of the 7 workflows, write:
1. Workflow name and number
2. Plain-English description (1 paragraph — what problem does it solve, who would use it,
   what does it produce) — written for a non-technical Finance coworker
3. Which module(s) / script(s) deliver it (VBA module name, or Python script name)
4. Where to find it in the tool (Command Center category + action name, or Python launcher
   menu number)
5. Sample file used to demo it (filename from the samples/ folder, or note if TBD)
6. Key output — what does the coworker get when they run it? (sheet name, folder, file)
7. Known limitations for V1 (anything the tool does NOT handle that a coworker might expect)

Also add:
- An intro section explaining WHY we're leading with 7 workflows (not 140):
  Connor owns support personally. Narrowing the doorway limits support load. The full toolkit
  is always there inside the workbook — we just don't push it at onboarding time.
- A "what's in the full toolkit" note: 140+ tools across 23 VBA modules + 6 Python scripts.
  Coworkers can explore the Command Center once they're comfortable with the 7 starter flows.

STYLE:
- Plain English. One section per workflow. Use headers for each workflow.
- Target length: 150–200 lines.
- Do NOT write any Python or VBA code as part of this task.
```

---

## TIER 2 — V4 PYTHON BUILD (After Tier 1 docs are approved)

Do NOT start these prompts until all 5 Tier 1 planning docs are written and Connor has
approved VIDEO_4_REVISED_PLAN.md. The build spec from the 5th-pass research is at
`RecTrial\Brainstorm\NewCodeResearchExtra\02_codex_build_spec.md` — read it before building.

**⚠️ PROTECTED SCRIPTS — DO NOT EDIT UNTIL VIDEO 4 IS RECORDED:**
aging_report.py, bank_reconciler.py, compare_files.py, forecast_rollforward.py,
fuzzy_lookup.py, pdf_extractor.py, variance_analysis.py, variance_decomposition.py

---

### PROMPT 2-A — revenue_leakage_finder.py

```
I need you to build revenue_leakage_finder.py — the hero script for Video 4.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- This script is the HERO of Video 4: "Python Automation for Finance."
- The story is: "Python found a possible billing problem your Excel couldn't see."
- Audience: 50–150 non-technical Finance coworkers running this on their own files.
- This script must use Python STANDARD LIBRARY ONLY — no pandas, no openpyxl, no third-party.
- The script runs from a bundled Python 3.11 embeddable. Zero install for coworkers.

SAFETY RULES (all scripts must follow these):
- Input files are READ-ONLY. Never modify the source file.
- Outputs go to a timestamped folder: outputs/revenue_leakage_YYYYMMDD_HHMMSS/
- Every run writes a log file in that folder.
- Sample/demo mode must be available (--sample flag or menu option 0).
- Clear failure messages in plain English shown to user.
- Technical error details go to the log, not the main screen.
- No internet calls. No API calls. No credentials.

WHAT THE SCRIPT DOES:
Given a CSV or Excel-exported CSV of billing/contract/subscription data, find rows where:
1. A customer has an active contract but no corresponding invoice in the billing data
2. An invoiced amount is lower than the contracted amount (possible short-billing)
3. A contract end date has passed but the subscription is still being billed (possible overcharge
   OR the contract wasn't renewed — either needs review)
4. Duplicate invoice IDs (possible double-billing)

The script should:
- Accept a single input file (CSV) via command-line argument or interactive prompt
- Run all 4 checks
- Output a findings CSV with columns: check_type, row_number, customer_id, description,
  amount_expected, amount_found, severity (HIGH/MEDIUM/LOW)
- Output a plain-English summary to the screen showing how many issues were found per check
- Output a log file

SAMPLE DATA:
Create a sample input file at: samples/revenue_leakage_sample.csv
The sample should have ~30 rows and include at least 2–3 examples of each type of issue.
Make the data look realistic (not toy numbers) — use SaaS subscription amounts in the $1K–$50K
ARR range, realistic customer names (Company A, Company B, etc.), realistic contract dates.

WHERE TO SAVE:
- Script: RecTrial\UniversalToolkit\python\ZeroInstall\revenue_leakage_finder.py
- Sample: RecTrial\UniversalToolkit\python\ZeroInstall\samples\revenue_leakage_sample.csv

INTEGRATION:
After building, add this script as a menu option in finance_automation_launcher.py:
Option 1: Revenue Leakage Finder

Also add an entry to the smoke test file:
RecTrial\UniversalToolkit\python\ZeroInstall\smoke_test_video4_python.py
```

---

### PROMPT 2-B — data_contract_checker.py

```
I need you to build data_contract_checker.py.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- This script validates an input file's structure against a template/contract.
- Use case: coworker receives a monthly data export from another team. The contract says it
  should have 12 specific columns with specific data types. This script checks that.
- Audience: 50–150 non-technical Finance coworkers.
- Python STANDARD LIBRARY ONLY — no pandas, no openpyxl.
- Runs from bundled Python 3.11 embeddable.

SAFETY RULES (same as all V4 scripts):
- Input files read-only. Outputs to timestamped folder. Log every run.
- Sample/demo mode available. Plain-English messages. No internet/API calls.

WHAT THE SCRIPT DOES:
Given:
- An input CSV (the file to check)
- A contract/template (either a second CSV showing expected column names + types,
  OR a simple text/JSON spec that the user creates)

Check:
1. Are all required columns present?
2. Are any unexpected columns present (may be fine, but flag them)?
3. For each column, does the data match the expected type (number, date, text)?
4. Are there any blank required fields?
5. Are any numeric fields negative when they should be positive?
6. Does the row count match an expected range (optional — useful for "should be ~500 rows")?

Output:
- A findings CSV: check_type, column, row_number, issue_description, severity
- A summary to screen: N issues found, N warnings, file is PASS or FAIL
- A log file

SAMPLE DATA:
Create two sample files at:
samples/data_contract_sample_input.csv (the file to check — include ~5 intentional issues)
samples/data_contract_sample_template.csv (the expected column spec)

WHERE TO SAVE:
Script: RecTrial\UniversalToolkit\python\ZeroInstall\data_contract_checker.py
Samples: RecTrial\UniversalToolkit\python\ZeroInstall\samples\

INTEGRATION:
Add as menu option 2 in finance_automation_launcher.py: Data Contract Checker
Add to smoke test file.
```

---

### PROMPT 2-C — exception_triage_engine.py

```
I need you to build exception_triage_engine.py.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- This script reviews a transaction file and flags items that need human review.
- Use case: accounts payable or expense report review — find transactions that are
  unusually large, fall outside normal ranges, are round numbers (possible estimates),
  or match known exception patterns.
- Audience: 50–150 non-technical Finance coworkers.
- Python STANDARD LIBRARY ONLY. Runs from bundled Python 3.11 embeddable.

SAFETY RULES (same as all V4 scripts):
- Read-only inputs. Timestamped output folder. Log every run.
- Sample mode available. Plain-English messages. No internet/API calls.

WHAT THE SCRIPT DOES:
Given a CSV of transactions (date, amount, description, category, vendor, approver):

Run these exception checks:
1. Amount threshold — flag any transaction above $X (user sets threshold, default $10,000)
2. Round number check — flag transactions that are exactly round numbers ($5,000, $10,000)
   (round numbers often indicate estimates, not actuals)
3. Duplicate check — flag same amount + same vendor within 30 days (possible duplicate)
4. Missing required fields — flag rows with blank vendor, blank category, or blank approver
5. Weekend/holiday transactions — flag transactions dated on weekends (unusual for corp finance)
6. Consecutive sequence — flag vendors with 3+ transactions just under the approval threshold
   (classic split-transaction pattern to avoid approval requirements)

Output:
- Findings CSV: exception_type, row_number, amount, vendor, description, severity, reason
- Screen summary: N total exceptions, breakdown by type, overall risk level (LOW/MEDIUM/HIGH)
- Log file

SAMPLE DATA:
Create samples/exception_triage_sample.csv with ~50 rows including planted examples of each
exception type. Use realistic amounts and vendor names.

WHERE TO SAVE:
Script: RecTrial\UniversalToolkit\python\ZeroInstall\exception_triage_engine.py
Samples: RecTrial\UniversalToolkit\python\ZeroInstall\samples\

INTEGRATION:
Add as menu option 3 in finance_automation_launcher.py: Exception Triage Engine
Add to smoke test file.
```

---

### PROMPT 2-D — control_evidence_pack.py

```
I need you to build control_evidence_pack.py.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- This script generates a folder of audit-ready evidence files from a set of input files.
- Use case: auditor or manager asks "show me evidence that you reviewed the AP file this month."
  This script packages the inputs, outputs, and a run summary into a dated folder, ready to
  attach to an email or upload to audit software.
- Audience: 50–150 non-technical Finance coworkers.
- Python STANDARD LIBRARY ONLY. Runs from bundled Python 3.11 embeddable.

SAFETY RULES (same as all V4 scripts):
- Read-only inputs. Timestamped output folder. Log every run.
- Sample mode available. Plain-English messages. No internet/API calls.

WHAT THE SCRIPT DOES:
Given a folder of input files (CSVs, text files — whatever the coworker ran their analysis on):

1. Create a new timestamped evidence folder:
   outputs/control_evidence_YYYYMMDD_HHMMSS/
2. Copy the input files into an inputs/ subfolder (read-only copies — do not modify originals)
3. Generate a CONTROL_EVIDENCE_SUMMARY.txt file containing:
   - Date and time of package creation
   - Operator name (prompted from user)
   - Control name / review name (prompted from user)
   - List of input files included (filename, size, last modified date)
   - Any output files already in an outputs/ folder (if user points the script at a prior run)
   - A "reviewed by" signature line for manual sign-off
4. Generate a file hash log (SHA256 of each input file) to prove files haven't been altered
5. Print a summary to screen: "Evidence pack created at [path]. 3 input files included."

SAMPLE DATA:
Create a sample set of 3 small CSV files at:
samples/control_evidence_inputs/ (3 files: ap_export.csv, gl_extract.csv, bank_statement.csv)
Each file should have ~10 rows of realistic-looking finance data.

WHERE TO SAVE:
Script: RecTrial\UniversalToolkit\python\ZeroInstall\control_evidence_pack.py
Samples: RecTrial\UniversalToolkit\python\ZeroInstall\samples\

INTEGRATION:
Add as menu option 4 in finance_automation_launcher.py: Control Evidence Pack
Add to smoke test file.
```

---

### PROMPT 2-E — workbook_dependency_scanner.py

```
I need you to build workbook_dependency_scanner.py.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- This script scans Excel workbooks (as CSV exports or directly as .xlsx if possible with stdlib)
  and maps where formulas reference other files or sheets.
- Use case: Finance team has a nest of Excel files that feed each other. Before making a change,
  they want to know "which files depend on this one?" or "what does this formula pull from?"
- Audience: 50–150 non-technical Finance coworkers.
- Python STANDARD LIBRARY ONLY (use zipfile + xml.etree to read .xlsx natively — no openpyxl).
  Runs from bundled Python 3.11 embeddable.

SAFETY RULES (same as all V4 scripts):
- Read-only inputs. Timestamped output folder. Log every run.
- Sample mode available. Plain-English messages. No internet/API calls.

WHAT THE SCRIPT DOES:
Given one or more .xlsx files (or a folder of .xlsx files):

1. Open each .xlsx using zipfile + xml.etree (xlsx files are ZIP archives with XML inside)
2. Scan all cell formulas for:
   - External file references: =[OtherFile.xlsx]Sheet1!A1
   - Cross-sheet references: =Sheet2!A1
   - Named range references
3. Output a dependency map CSV:
   Columns: source_file, source_sheet, source_cell, reference_type (external/cross-sheet/named),
            target_file_or_sheet, target_cell, formula_preview
4. Output a plain-English summary:
   - N external file references found (files this workbook pulls from)
   - N cross-sheet references
   - List of external files referenced (so the user knows what else they need)
5. Flag any external references where the referenced file doesn't exist in the scanned folder

SAMPLE DATA:
Create sample .xlsx files at:
samples/workbook_dependency_sample/ — 2 .xlsx files where one references the other

WHERE TO SAVE:
Script: RecTrial\UniversalToolkit\python\ZeroInstall\workbook_dependency_scanner.py
Samples: RecTrial\UniversalToolkit\python\ZeroInstall\samples\

INTEGRATION:
Add as menu option 5 in finance_automation_launcher.py: Workbook Dependency Scanner
Add to smoke test file.
```

---

### PROMPT 2-F — finance_automation_launcher.py (update/finalize)

```
I need you to finalize finance_automation_launcher.py — the entry point coworkers use.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- This is the main CLI menu for the Video 4 Python distribution package.
- Name: finance_automation_launcher.py (NOT finance_copilot.py — "Copilot" implies AI)
- Coworkers launch this from an Excel button (VBA Shell() call) or by double-clicking.
- It must work inside a bundled Python 3.11 embeddable folder with zero install.
- Python STANDARD LIBRARY ONLY.

WHAT THE LAUNCHER DOES:
1. Shows a welcome banner with the tool name and version (v1.0)
2. Shows a safety disclaimer (1–2 sentences: "These scripts read your files but never modify them.
   All output goes to an outputs/ folder. See PYTHON_SAFETY.md for details.")
3. Shows a numbered menu:
    1. Revenue Leakage Finder
    2. Data Contract Checker
    3. Exception Triage Engine
    4. Control Evidence Pack
    5. Workbook Dependency Scanner
    0. Exit
4. Accepts the user's number input
5. Runs the chosen script (uses subprocess or import — whichever is cleaner for bundled Python)
6. After the script finishes, returns to the menu (don't exit — coworkers may want to run
   another tool)
7. Handles invalid input gracefully (not a number, number out of range)
8. Handles a KeyboardInterrupt (Ctrl+C) gracefully — shows "Exiting. Goodbye." instead of a
   traceback

UPDATES NEEDED (if a version already exists):
- Confirm menu items 1–5 match the 5 scripts listed above
- Confirm the safety disclaimer is present
- Confirm the banner shows the correct name (finance_automation_launcher.py, not finance_copilot.py)
- Confirm it returns to the menu after each run

WHERE TO SAVE:
RecTrial\UniversalToolkit\python\ZeroInstall\finance_automation_launcher.py

ALSO UPDATE:
smoke_test_video4_python.py — add a test that the launcher can be imported without errors
and that the menu string contains all 5 expected tool names.
```

---

## TIER 3 — POST-V4 SHIP (Do after Video 4 is recorded and distributed)

Do NOT start these until Video 4 is fully recorded and the pilot is in progress.

---

### PROMPT 3-A — Codex Batch 4: Dual-Logging Pattern

```
I need you to implement the Codex Batch 4 dual-logging pattern.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- The Codex parallel build (tug83535/AP_CodexVersion) introduced a dual-logging pattern
  where actions are logged both to the VBA_AuditLog sheet AND to a plain-text file.
- Batches 1–3 are already live. Batch 4 was deferred until after Video 4 ships because
  it touches the LogAction signature, which has a history of signature bugs.

CRITICAL WARNING — LogAction signature:
- LogAction takes 4 arguments: (Module As String, Procedure As String, Message As String, Status As String)
- The 4th argument MUST be a String like "OK" or "FAIL" — NEVER a Double (elapsed time).
- 13 instances of passing elapsed time as the 4th arg have been found and fixed historically.
- Before making ANY change to LogAction or any call site, grep the entire codebase for
  "LogAction" and review every call site for the correct signature. Do not break existing calls.

WHAT TO DO:
1. Read the CodexCompare/COMPARISON_REPORT.md and CodexCompare/CHERRY_PICK_TRACKER.md
   to understand what Batch 4 specifically proposes.
2. Read ALL existing LogAction call sites in vba/ (grep for "LogAction").
3. Implement the dual-logging pattern with zero changes to the existing 4-argument LogAction
   signature — the plain-text log output must be additive, not a replacement.
4. Test that all existing LogAction calls still work unchanged.
5. Update CodexCompare/CHERRY_PICK_TRACKER.md to mark Batch 4 complete.

DO NOT:
- Change the LogAction signature
- Break any existing call site
- Start this task until Video 4 is recorded and distributed
```

---

### PROMPT 3-B — Codex Batch 5: Top-Level Governance Docs

```
I need you to create the Codex Batch 5 governance documents.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- Batch 5 is four top-level governance docs recommended by the Codex parallel build.
- These were deferred until after Video 4 ships.

THE FOUR DOCS TO CREATE:
1. CONSTRAINTS.md — Banned features + threshold to justify any new feature
2. BRAND.md — Colors, fonts, formatting rules (non-negotiable). Source: docs/ipipeline-brand-styling.md
3. RELEASE_READINESS_CHECKLIST.md — Pre-release gate for any new version
4. TROUBLESHOOTING.md — Top 10 issues coworkers hit + how to fix them

WHERE TO SAVE:
All four files go in the repo root (same level as CLAUDE.md).

WHAT TO PUT IN EACH:
- CONSTRAINTS.md: Pull from the AGENTS.md "Code Quality Rules" + "banned patterns" sections.
  Add a "new feature threshold" — must have a concrete coworker use case, must not require
  IT/admin install, must work on any Finance workbook (not just the demo file).
- BRAND.md: Pull from docs/ipipeline-brand-styling.md. Add the VBA RGB values for reference.
  Note that modConfig color constants predate the brand guide and should NOT be edited — only
  NEW styling work uses the official brand colors.
- RELEASE_READINESS_CHECKLIST.md: Use the 11-checkpoint table from MINIMUM_DISTRIBUTION_PLAN.md
  as a starting point. Expand it to cover both VBA releases (new .xlsm version) and Python
  releases (new zip version).
- TROUBLESHOOTING.md: Anticipate the top 10 things coworkers will ask. Include:
  - "Script didn't run / nothing happened" — check that Python is in the right folder
  - "I got a permission error" — check that the file isn't open in another program
  - "The output folder is empty" — check the log file for error details
  - "The Excel Command Center doesn't open" — check that macros are enabled
  - "I can't find my output" — outputs go to [script folder]\outputs\

After creating all four files, update CodexCompare/CHERRY_PICK_TRACKER.md to mark Batch 5 complete.
```

---

### PROMPT 3-C — Video 5 "Getting Started" Micro-Video Plan

```
I need you to write a planning doc for Video 5 — the "Getting Started" micro-video.

CONTEXT:
- I am Connor Atlee, Finance & Accounting at iPipeline.
- After Video 4 ships, a short follow-up "Getting Started" video (3–5 min) will help
  coworkers who watched Video 4 but don't know how to actually download and start.
- This video does NOT replace Video 4 — it's a companion that bridges the gap between
  "I watched the demo" and "I'm running the tools."
- Audience: coworkers who want to use the tools but aren't sure where to start.

WHAT TO WRITE:
Write VIDEO_5_GETTING_STARTED_PLAN.md. Save it to:
    RecTrial\Brainstorm\VIDEO_5_GETTING_STARTED_PLAN.md

Required sections:
1. Video goal — one sentence: what should a coworker be able to do AFTER watching this?
2. Target length — 3–5 min (this is a utility video, not a feature showcase)
3. Step-by-step outline — what the viewer sees on screen for each step:
    Step 1: Go to [SharePoint link] and download FinanceTools_v1.0.zip
    Step 2: Unzip to a folder (show where — Desktop or Documents, not Downloads)
    Step 3: Open FinanceTools.xlsm. Enable macros when prompted.
    Step 4: Click the "Finance Tools" button (or press the keyboard shortcut) to open the
            Command Center. Run your first tool.
    Step 5: For Python tools — double-click finance_automation_launcher.py OR use the Excel
            Python button. Pick option 1 (Revenue Leakage Finder). Choose "Run sample."
4. What to show for each step (screenshots? live demo? narration only?)
5. Recording approach — will this use Director macro + ElevenLabs narration like Videos 1–3,
   or is it a simpler screen recording with Connor narrating live?
6. Distribution — where will this video live? Same SharePoint page as the zip download?
   Separate Teams post? Embedded in the README?

STYLE:
- Plain English. Short. This is a planning doc, not a script.
- Target length: 60–100 lines.
- Do NOT write any code as part of this task.
```

---

## REFERENCE — Key Files Claude Code Will Need

When starting any of the prompts above, these are the most important files to read first:

| File | What it contains | Tier relevance |
|------|-----------------|----------------|
| `RecTrial/HANDOFF_2026-04-27.md` | Full state of the world as of April 27 | All tiers |
| `RecTrial/OPEN_ITEMS_2026-04-27.md` | Detailed specs for the 5 planning docs | Tier 1 |
| `RecTrial/Brainstorm/VIDEO_4_REVIEW_DECISION_MEMO.md` | Locked decisions (after Tier 1) | Tier 2 |
| `RecTrial/Brainstorm/VIDEO_4_REVISED_PLAN.md` | Chapter outline + build effort (after Tier 1) | Tier 2 |
| `RecTrial/UniversalToolkit/python/PYTHON_SAFETY.md` | Safety rules (after Tier 1) | Tier 2 |
| `RecTrial/UniversalToolkit/python/ZeroInstall/` | Existing V4 Python scripts | Tier 2 |
| `RecTrial/UniversalToolkit/python/ZeroInstall/README_VIDEO4_PYTHON.md` | V4 Python README | Tier 2 |
| `RecTrial/UniversalToolkit/python/ZeroInstall/smoke_test_video4_python.py` | Smoke tests | Tier 2 |
| `CodexCompare/CHERRY_PICK_TRACKER.md` | Codex batch status | Tier 3 |
| `CLAUDE.md` (repo root) | Full project instructions and history | All tiers |
| `tasks/lessons.md` | Known bug patterns — READ BEFORE ANY VBA EDITS | All tiers |

---

## QUICK REMINDER — 8 Protected Python Scripts

**DO NOT EDIT these until Video 4 is fully recorded:**
- `aging_report.py`
- `bank_reconciler.py`
- `compare_files.py`
- `forecast_rollforward.py`
- `fuzzy_lookup.py`
- `pdf_extractor.py`
- `variance_analysis.py`
- `variance_decomposition.py`

Everything else is fair game.
