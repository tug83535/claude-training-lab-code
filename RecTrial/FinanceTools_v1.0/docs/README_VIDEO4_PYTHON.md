# Finance Automation Toolkit v1.0 — Python Pack
## Video 4: Python Automation for Finance — iPipeline

---

## What is this?

A set of Python scripts that run directly on your Windows machine — no installation required.
Each script reads your CSV or Excel files, runs an analysis, and writes a report to an `/outputs/` folder.
**Your original files are never modified.**

---

## Quick Start

### Step 1 — Run the launcher
```
python finance_automation_launcher.py
```
A numbered menu appears. Pick a number and press Enter.

### Step 2 — Run in sample mode first
Every tool has a sample mode that uses built-in demo data. Run it in sample mode before using your own files. This shows you what the output looks like without touching anything sensitive.

### Step 3 — Run against your own files
Once you understand the output, point the tool at your own file:
```
python revenue_leakage_finder.py contracts.csv billing.csv
python data_contract_checker.py my_export.csv
```

### Step 4 — Find your results
All output goes to the `/outputs/` folder. Each run creates a new timestamped subfolder so previous results are never overwritten.

---

## The Five Tools

### 1. Revenue Leakage Finder
```
python revenue_leakage_finder.py contracts.csv billing.csv
python revenue_leakage_finder.py --sample
```
Finds five types of billing problems: customers billing without a contract, stale contracts still billing, base quantity anomalies, invoice amount drift, and customer name mismatches. Produces a branded HTML report, a ranked exceptions CSV, and an ARR gap waterfall chart.

**Input files needed:**
- `contracts.csv` — one row per contract: customer_id, customer_name, status, term_end, base_fee, base_quantity, billing_basis, expected_annual_revenue
- `billing.csv` — one row per invoice: invoice_id, customer_id, customer_name, billing_period, amount_billed

---

### 2. Data Contract Checker
```
python data_contract_checker.py myfile.csv
python data_contract_checker.py --sample
```
Validates a CSV file before you run analysis on it. Checks for blank rows, non-numeric amounts, unparseable dates, duplicate rows, and business-rule violations (like base_quantity = 0). Produces a PASS/WARN/FAIL report so you know whether the file is ready for analysis.

**Input files needed:**
- Any CSV file you want to validate

---

### 3. Exception Triage Engine
```
python exception_triage_engine.py exceptions_ranked.csv
python exception_triage_engine.py --sample
```
Takes the exceptions CSV from Revenue Leakage Finder and scores each exception on four dimensions: dollar impact, confidence, recency, and whether the same customer appears repeatedly. Produces a ranked report and a top-10 action list with plain-English recommended actions.

**Note:** Run Revenue Leakage Finder first. The `--sample` flag auto-finds the most recent Leakage Finder output.

---

### 4. Control Evidence Pack
```
python control_evidence_pack.py --sample
python control_evidence_pack.py --input-dir outputs/20260428_123456_revenue_leakage_finder
python control_evidence_pack.py --input-dir outputs/20260428_123456_revenue_leakage_finder --control-name "Q2 2026 Revenue Review"
```
Creates a tamper-evident evidence bundle from any analysis output folder. Records each file's name, size, last-modified date, and SHA-256 hash. If someone asks "what files were analyzed and when?" — this folder answers that question precisely.

**When to use it:** After any significant analysis that may go to audit, leadership, or a review ticket.

---

### 5. Workbook Dependency Scanner
```
python workbook_dependency_scanner.py myworkbook.xlsx
python workbook_dependency_scanner.py --sample
```
Maps every cross-sheet formula reference inside an Excel file. Shows which sheets reference other sheets, which cells contain the formulas, and flags any hidden sheets that are referenced (because deleting those would break formulas). No Excel installation needed — reads the file directly.

**Input files needed:**
- Any `.xlsx` or `.xlsm` file

---

## Where do outputs go?

```
ZeroInstall/
  outputs/
    20260428_123456_revenue_leakage_finder/
      leakage_report.html
      exceptions_ranked.csv
      arr_waterfall.html
      run_log.json
      run_summary.txt
    20260428_123501_data_contract_checker/
      contract_check_report.html
      issues_detail.csv
      ...
```

Each run creates a new folder with a timestamp. Previous runs are never touched.

---

## Troubleshooting

**"Python is not recognized as a command"**
Python is not installed or not in your PATH. Open a browser and go to python.org/downloads. Install Python 3.10 or later. During installation, check "Add Python to PATH".

**"FileNotFoundError: Input file not found"**
The file path is wrong. Try dragging the file into the Command Prompt window instead of typing the path — it auto-fills the full path for you.

**"The file is not a valid .xlsx"**
The file may be open in Excel. Close it in Excel first, then run the scan.

**"ModuleNotFoundError: No module named 'common'"**
You must run scripts from the `ZeroInstall/` folder, not from another directory. Run:
```
cd C:\path\to\ZeroInstall
python script_name.py --sample
```

**Something else went wrong**
Contact Connor Atlee — Finance & Accounting.

---

## Safety reminder

1. Input files are opened read-only — never modified.
2. All outputs go to `/outputs/` only.
3. No data leaves your machine — no network connections.
4. Start with sample mode before using real files.
5. Do not run on files with SSNs, passwords, or payment card numbers.
6. Questions? Contact Connor.

---

*Finance Automation Toolkit v1.0 | iPipeline | Connor Atlee — Finance & Accounting*
