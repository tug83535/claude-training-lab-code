# Python Safety — Finance Automation Toolkit v1

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date:** 2026-04-28
**Audience:** Finance & Accounting coworkers (non-developers) + IT/security reviewers
**Purpose:** A plain-English, inspectable record of the safety rules this toolkit follows. Every Python script in this package was built to satisfy all 14 rules below.

---

## How Python runs in this toolkit

You do not need to open a terminal or install anything. The toolkit uses Excel buttons.

When you click a button in `FinanceTools.xlsm`, Excel internally calls Python using a copy of Python that is bundled inside the toolkit folder — in `python\python-embedded\`. This bundled Python belongs to the toolkit. It never modifies your system, never installs anything on your machine, and is removed simply by deleting the toolkit folder if you ever want to clean up.

The scripts can also be run from Command Prompt by advanced users who prefer it — the exact same code works both ways.

No Python installation on your laptop is required. No administrator rights are required. No IT involvement is required for normal use.

---

## The 14 safety rules

### 1. No internet calls
Scripts do not connect to the internet under any circumstances. No web requests, no API calls, no telemetry, no package downloads, no external AI calls. Everything runs locally on your machine with your local files.

### 2. No external AI or API calls
Scripts do not call OpenAI, Anthropic, or any other AI service. There is no "AI" inside the toolkit — the name "Finance Automation" is intentional. What you see is what runs: Python logic applied to your files.

### 3. No credentials, passwords, or secrets
Scripts do not ask for passwords, tokens, database connection strings, or any kind of secret. If a script ever prompts you for a password, something has gone wrong — stop and contact Connor.

### 4. Your input files are never changed
Scripts open your files as read-only. They do not edit, overwrite, move, or delete any file you provide as input. Your original files remain exactly as they were before you ran the script.

### 5. Scripts never write to the folder where your input lives
Output always goes to a separate `outputs\` folder inside the toolkit directory. Scripts cannot write files back to the folder where your input came from.

### 6. All outputs go to a timestamped folder
Every time you run a script, a new folder is created inside `outputs\` with a timestamp:
```
outputs\YYYYMMDD_HHMMSS_toolname\
```
Running the same script twice creates two separate folders. Nothing is overwritten. You always have the full history of every run.

### 7. Every run creates a log
Each run produces two log files in the output folder:
- `run_log.json` — machine-readable record of what was analyzed, row counts, any warnings, any errors
- `run_summary.txt` — plain English version of the same

If something goes wrong, the log tells you what happened. If you need to show someone "what did this script actually do?" — share the log.

### 8. Sample mode is available
Every major script supports a sample mode. In sample mode, the script uses pre-built synthetic data files (included in the toolkit) instead of your own files. Always run sample mode first to understand what a script does before pointing it at your own data.

From Excel: there is a "Run Sample" button and a "Run on Your File" button for each tool.
From Command Prompt: add `--sample` to any command.

### 9. Clear error messages — no cryptic code on your screen
If a script encounters a problem (missing column, wrong file format, empty file), it shows a plain message explaining what went wrong. You will not see a wall of Python code. Detailed technical information, if needed for troubleshooting, goes to the log file — not the main output.

### 10. Sensitive data is not stored in logs
Run logs record file names, row counts, column names, exception types, and statistics. They do not copy the contents of your data files into the log. If your billing file has customer amounts in it, the log records "processed 312 rows, found 14 exceptions" — not the actual dollar amounts.

### 11. No destructive operations without explicit confirmation
Scripts that create or overwrite anything ask before proceeding, except for creating the timestamped output folder (which always gets a new name, so nothing is ever overwritten).

### 12. No files are deleted
Scripts never delete files — yours or their own. If a script creates an output folder, it stays there until you manually clean it up. Output folders can be deleted safely once you've reviewed and saved what you need.

### 13. Relative paths only — the toolkit works from any folder
All paths inside scripts are relative to the toolkit folder. Scripts use `Path(__file__).parent` (Python) and `ThisWorkbook.Path` (VBA Excel buttons) to locate their own files. This means the toolkit works correctly regardless of where you unzipped it — `C:\Users\yourname\Desktop\FinanceTools\`, a network drive, or anywhere else.

There is one exception: when you point a script at a file on your own machine (like a billing export from your OneDrive folder), you supply that path. The script reads from it but never writes back to it.

### 14. The Excel workbook shows a safety reminder on startup
When you open `FinanceTools.xlsm`, a brief notice confirms: local-only, no internet, inputs are read-only, outputs go to the outputs folder. This is not a clickthrough warning — it is a one-line status bar message so you can always confirm the toolkit is running in safe mode.

---

## What the bundled Python contains

The `python\python-embedded\` folder contains Python 3.11 (Windows embeddable distribution — the official version from python.org). It also includes the following pre-bundled packages:

| Package | What it does |
|---|---|
| pandas | Data analysis and CSV processing |
| openpyxl | Reading and writing Excel files (.xlsx) |
| matplotlib | Charts and graphs (used for the ARR waterfall output) |

These packages are included in the bundle. No pip install is required for any of them. If a script needs a package that is not on this list, it will be flagged in the script's header comment and added to the bundle before the toolkit is released.

Scripts do not use any packages outside the standard library and the four above without explicit documentation.

---

## Running on your own real files — what's safe, what to avoid

**Safe to run against:**
- Billing exports or contract lists from internal systems (e.g., a CSV exported from CustVol or a billing platform)
- Excel files from your own drive or a shared Finance folder
- Any file you have legitimate access to and that does not contain information restricted to specific individuals

**Do not run against:**
- HR files, payroll data, or personnel records
- Files containing SSNs, banking credentials, or personal financial information
- Any file labeled Confidential or Restricted that you haven't confirmed is in scope for your analysis
- Production system files that are live and being actively written by another process

**Before running on a real file for the first time:**
1. Run sample mode first so you understand what the output looks like.
2. Verify your input file is in the `inputs\` folder or note its path.
3. Check the output folder after the run — review it before sharing anything with a colleague or manager.
4. If the output looks wrong or unexpected, check `run_summary.txt` in the output folder before re-running.

**OneDrive paths:** If your files are on a OneDrive-redirected Desktop (common at iPipeline — the path looks like `C:\Users\yourname\OneDrive - iPipeline\Desktop\`), scripts handle this correctly using standard path resolution. You can supply the full path when prompted and it will work.

---

## IT and security reviewer notes

**Bundled Python:** The toolkit includes a copy of the official Python 3.11 Windows embeddable distribution (from python.org). This is a portable, self-contained interpreter with no installer. It writes no registry entries and modifies no system settings. It can be removed by deleting the `python\python-embedded\` folder.

**Network activity:** None. Scripts do not make outbound connections. Verify with a network monitor on first run if needed.

**File system writes:** Scripts write only to the `outputs\` subfolder of the toolkit directory. They do not write to system folders, registry, AppData, or any location outside their own folder.

**Macro security:** `FinanceTools.xlsm` contains VBA macros that call Python via `Shell()`. The macros are open to inspection in the VBA editor (Alt+F11). No macros are password-protected.

**Endpoint scanner note:** The toolkit ships with `python.exe` inside the folder structure. Some endpoint scanning tools may flag `.exe` files in SharePoint packages. If the IT endpoint scanner flags this package, please review the above and contact Connor for clarification or an exception request. This is a known launch dependency that Connor is coordinating with IT separately.

---

## Who to contact

Connor Atlee — Finance & Accounting
**For:** questions about what a script does, unexpected results, errors you can't resolve from the log, permission to use the toolkit on a specific file type.

**Not for:** IT security concerns about the package — those go to the IT helpdesk with Connor cc'd.

---

**End of Python Safety doc.**
