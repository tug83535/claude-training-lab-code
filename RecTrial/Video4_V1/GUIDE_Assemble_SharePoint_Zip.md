# Step-by-Step Guide — Assemble the SharePoint Zip Package
## Building the FinanceTools_v1.0 distribution folder

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date:** 2026-04-29
**Time required:** ~20 minutes
**Skill level:** No coding needed — this is copy/paste and folder creation only

---

## What you are building

A single folder called `FinanceTools_v1.0` that contains everything a coworker needs
to run the Finance Tools. When you zip this folder and put it on SharePoint, coworkers
download it, unzip it, open Excel, and click the button. Nothing else required.

**Final folder structure:**
```
FinanceTools_v1.0\
├── FinanceTools.xlsm              ← the Excel workbook with the button
├── python\
│   └── python-embedded\
│       ├── python.exe             ← bundled Python 3.11 (no install needed)
│       └── [~20 other Python files]
├── scripts\
│   ├── finance_automation_launcher.py
│   ├── revenue_leakage_finder.py
│   ├── data_contract_checker.py
│   ├── exception_triage_engine.py
│   ├── control_evidence_pack.py
│   ├── workbook_dependency_scanner.py
│   └── common\
│       ├── safe_io.py
│       ├── logging_utils.py
│       ├── report_utils.py
│       └── sample_data.py
├── samples\
│   ├── contracts_sample.csv
│   └── billing_sample.csv
├── outputs\                       ← empty folder, results appear here after each run
└── docs\
    ├── README_VIDEO4_PYTHON.md
    └── PYTHON_SAFETY.md
```

**⚠ One step depends on the Excel guide:** Placing `FinanceTools.xlsm` into the folder
(Step 14) requires the workbook to be built first. Every other step is independent.
You can complete Steps 1–13 and 15–20 in any order.

---

## PART 1 — Create the folder structure

### Step 1 — Choose where to build the package
You need a location to build the package before zipping it. Use this path:

`C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\`

You will create this folder in the next step.

### Step 2 — Create the root folder
1. Open **File Explorer**
2. Navigate to `C:\Users\connor.atlee\RecTrial\`
3. Right-click in the empty white space of the folder → **New** → **Folder**
4. Name it exactly: `FinanceTools_v1.0`
5. Press Enter

### Step 3 — Create the subfolders
Open the `FinanceTools_v1.0` folder you just created. Inside it, create these folders
one at a time (right-click → New → Folder):

- `python`
- `scripts`
- `samples`
- `outputs`
- `docs`

After this step your folder should look like:
```
FinanceTools_v1.0\
├── python\
├── scripts\
├── samples\
├── outputs\
└── docs\
```

### Step 4 — Create the python-embedded subfolder
1. Open the `python` folder you just created
2. Inside it, create one more folder: `python-embedded`

Your python folder should now look like:
```
FinanceTools_v1.0\
└── python\
    └── python-embedded\     ← empty for now, filled in Part 2
```

### Step 5 — Create the common subfolder inside scripts
1. Open the `scripts` folder
2. Inside it, create one more folder: `common`

Your scripts folder should now look like:
```
FinanceTools_v1.0\
└── scripts\
    └── common\              ← empty for now, filled in Part 3
```

---

## PART 2 — Download and place bundled Python 3.11

This is the "zero install" piece. You are downloading a self-contained version of Python
that ships inside the zip. Coworkers never need to install Python themselves.

### Step 6 — Download the Python 3.11 embeddable package
1. Open your browser and go to:
   **https://www.python.org/downloads/release/python-3119/**
2. Scroll down to the section called **"Files"**
3. Find the row that says: **Windows embeddable package (64-bit)**
   - The filename will be: `python-3.11.9-embed-amd64.zip`
   - The size is approximately 8–9 MB
4. Click the filename to download it
5. Save it somewhere easy to find — your Downloads folder is fine

**Why 64-bit:** iPipeline laptops run 64-bit Windows. If you are ever unsure, 64-bit is
the safe choice for any modern Windows machine made in the last 10 years.

### Step 7 — Extract the Python embeddable package
1. Find the downloaded file `python-3.11.9-embed-amd64.zip` in your Downloads folder
2. Right-click it → **Extract All...**
3. When asked where to extract, click **Browse** and navigate to:
   `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\python\python-embedded\`
4. Click **Select Folder** → then click **Extract**

**Important:** Make sure you extract the FILES into `python-embedded\` — not a subfolder
inside `python-embedded\`. After extraction, `python.exe` should be directly inside
`python-embedded\`, not inside another folder nested within it.

### Step 8 — Verify the extraction
Open `FinanceTools_v1.0\python\python-embedded\` in File Explorer.

You should see approximately 20 files including:
- `python.exe` ← this is the key one
- `python311.dll`
- `python311.zip`
- Several `.pyd` files

If you see a single folder inside `python-embedded\` instead of these files — you
extracted into an extra subfolder. Cut all the files from that inner folder and paste
them directly into `python-embedded\`.

---

## PART 3 — Copy the Python scripts

All scripts are already built and ready at:
`C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\ZeroInstall\`

### Step 9 — Copy the main scripts
1. Open File Explorer and navigate to:
   `C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\ZeroInstall\`
2. Select these 6 files (hold Ctrl and click each one):
   - `finance_automation_launcher.py`
   - `revenue_leakage_finder.py`
   - `data_contract_checker.py`
   - `exception_triage_engine.py`
   - `control_evidence_pack.py`
   - `workbook_dependency_scanner.py`
3. Right-click → **Copy**
4. Navigate to: `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\scripts\`
5. Right-click in the empty space → **Paste**

### Step 10 — Copy the common utilities
1. Still in the ZeroInstall folder, open the `common` subfolder
2. Select all 4 files inside it (Ctrl + A):
   - `safe_io.py`
   - `logging_utils.py`
   - `report_utils.py`
   - `sample_data.py`
3. Right-click → **Copy**
4. Navigate to: `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\scripts\common\`
5. Right-click → **Paste**

---

## PART 4 — Copy the sample data files

### Step 11 — Copy the sample CSV files
1. Navigate to:
   `C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\ZeroInstall\samples\`
2. Select both files:
   - `contracts_sample.csv`
   - `billing_sample.csv`
3. Right-click → **Copy**
4. Navigate to: `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\samples\`
5. Right-click → **Paste**

---

## PART 5 — Copy the documentation files

### Step 12 — Copy the README
1. Navigate to:
   `C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\ZeroInstall\`
2. Find `README_VIDEO4_PYTHON.md` and copy it
3. Paste it into: `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\docs\`

### Step 13 — Copy the safety rules file
1. Navigate to:
   `C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\`
2. Find `PYTHON_SAFETY.md` and copy it
3. Paste it into: `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\docs\`

---

## PART 6 — Place FinanceTools.xlsm

**⚠ This step requires the Excel workbook to be built first.**
Complete `GUIDE_Build_FinanceTools_xlsm.md` before doing this step.

### Step 14 — Copy FinanceTools.xlsm into the package root
1. Find your completed `FinanceTools.xlsm` file (wherever you saved it during the Excel guide)
2. Copy it
3. Paste it directly into: `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\`
   *(Root of the folder — not inside any subfolder)*

After this step, `FinanceTools.xlsm` should sit at the same level as the `python\`,
`scripts\`, `samples\`, `outputs\`, and `docs\` folders.

---

## PART 7 — Verify the complete folder structure

### Step 15 — Do a final folder check
Open `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\` and confirm it looks like this:

```
FinanceTools_v1.0\
├── FinanceTools.xlsm                    ✓
├── python\
│   └── python-embedded\
│       ├── python.exe                   ✓
│       └── [~20 other files]            ✓
├── scripts\
│   ├── finance_automation_launcher.py   ✓
│   ├── revenue_leakage_finder.py        ✓
│   ├── data_contract_checker.py         ✓
│   ├── exception_triage_engine.py       ✓
│   ├── control_evidence_pack.py         ✓
│   ├── workbook_dependency_scanner.py   ✓
│   └── common\
│       ├── safe_io.py                   ✓
│       ├── logging_utils.py             ✓
│       ├── report_utils.py              ✓
│       └── sample_data.py               ✓
├── samples\
│   ├── contracts_sample.csv             ✓
│   └── billing_sample.csv              ✓
├── outputs\                             ✓ (empty — that's correct)
└── docs\
    ├── README_VIDEO4_PYTHON.md          ✓
    └── PYTHON_SAFETY.md                 ✓
```

If anything is missing, go back to the relevant Part and copy the missing file.

---

## PART 8 — Test the full zero-install path

This is the moment of truth. You are testing that the button finds bundled Python
and the menu opens — with no system Python required.

### Step 16 — Open FinanceTools.xlsm from the package folder
1. Navigate to `C:\Users\connor.atlee\RecTrial\FinanceTools_v1.0\`
2. Double-click `FinanceTools.xlsm` to open it
3. If the yellow "Macros have been disabled" bar appears → click **Enable Content**

**Important:** Open it from inside `FinanceTools_v1.0\` — not from wherever you
originally saved it. The VBA code finds Python relative to the workbook's current location.

### Step 17 — Click the Finance Tools button
Click the **Finance Tools** button on the sheet.

**What you should see:**
A command-line window opens with the Finance Tools menu:
```
============================================================
        Finance Tools — Finance Automation Launcher
============================================================

 1. Revenue Leakage Finder
 2. Data Contract Checker
 3. Exception Triage Engine
 4. Control Evidence Pack
 5. Workbook Dependency Scanner
 6. Show Safety Rules
 7. Open Outputs Folder
 8. Exit

Select an option (1-8):
```

If you see this menu — the zero-install path works. ✓

**If you still see "Python not found" error:**
- Confirm `python.exe` is directly inside `python-embedded\` (not in a subfolder)
- Confirm you opened FinanceTools.xlsm from inside `FinanceTools_v1.0\`
- Check the path shown in the error message — it tells you exactly where it looked

### Step 18 — Run a quick test of each tool
In the menu, type each number and press Enter to confirm each tool runs in sample mode:

| Option | Expected result |
|---|---|
| 1 | Revenue Leakage Finder runs, outputs folder created, HTML report generated |
| 2 | Data Contract Checker runs, shows PASS on sample data |
| 3 | Exception Triage Engine runs, top_10_action_list.csv created |
| 4 | Control Evidence Pack runs, evidence_summary.html created |
| 7 | File Explorer opens showing the outputs\ folder |
| 8 | Menu closes |

After running options 1–4, check `FinanceTools_v1.0\outputs\` in File Explorer.
You should see timestamped subfolders with the results inside.

---

## PART 9 — Zip the folder for SharePoint

### Step 19 — Zip the FinanceTools_v1.0 folder
Once testing passes:
1. Navigate to `C:\Users\connor.atlee\RecTrial\`
2. Right-click the `FinanceTools_v1.0` folder
3. Click **Send to** → **Compressed (zipped) folder**
   *(On Windows 11 you may need to click "Show more options" first)*
4. A file called `FinanceTools_v1.0.zip` is created in the same location

### Step 20 — Rename the zip if needed
The zip should be named exactly: `FinanceTools_v1.0.zip`
If Windows named it differently, right-click → Rename → correct it.

This zip is what goes on SharePoint. Coworkers download it, unzip it, and open Excel.

---

## Troubleshooting

**"Python not found" error after placing all files:**
Open the error message and read the exact path it tried. Then open File Explorer and
confirm `python.exe` is at that exact path. The most common cause is an extra subfolder
created during extraction (e.g., `python-embedded\python-3.11.9-embed-amd64\python.exe`
instead of `python-embedded\python.exe`).

**Menu opens but a tool crashes with an error:**
Type the error message exactly and send it to Claude. Include which option number you ran.

**The outputs folder doesn't appear after running a tool:**
The scripts create it automatically. If it is missing, the script likely errored. Check
the command-line window — there will be a plain-English error message explaining what
went wrong.

**Excel says "Macros have been disabled" every time you open it:**
This is a Windows security setting. Click Enable Content each time, or add the
`FinanceTools_v1.0` folder to Excel's Trusted Locations:
File → Options → Trust Center → Trust Center Settings → Trusted Locations → Add new location.

---

*End of guide. Version 1.0 — 2026-04-29.*
*Once the zip tests cleanly, let Claude know — next step is uploading to SharePoint for the pilot.*
