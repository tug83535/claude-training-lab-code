# TESTRUN — Download & Setup Guide
**KBT P&L Toolkit v2.1 — Live Testing Preparation**
Last updated: 2026-03-01

---

## What You Need to Download

You need **three things** from GitHub:

| # | What | Where in the repo | Why |
|---|------|--------------------|-----|
| 1 | The Excel file | `excel/` folder | This is the P&L workbook you will test in |
| 2 | All 31 VBA module files (`.bas`) | `vba/` folder | These get imported into Excel one time before testing |
| 3 | The test plan | `qa/TEST_PLAN.md` | Your step-by-step checklist while testing |

> **Note:** The repo documentation says 32 modules, but 31 `.bas` files are currently
> in the `vba/` folder. Import all 31. This will be reconciled after testing.

---

## Step 1 — Download Everything as a ZIP

This is the easiest way. It downloads the entire repo in one click.

1. Go to your GitHub repo in a browser:
   `https://github.com/tug83535/claude-training-lab-code`

2. Make sure you are on the correct branch. Look for a button near the top-left
   of the page that shows a branch name. Click it and select:
   `claude/review-code-testing-s4dsQ`

3. Click the green **`< > Code`** button (top right of the file list)

4. Click **Download ZIP**

5. Save the ZIP somewhere easy to find (Desktop is fine)

6. Right-click the ZIP → **Extract All** → choose a destination folder
   (for example: `C:\Users\YourName\Desktop\KBT_TestRun`)

---

## Step 2 — Locate the Three Things You Need

After extracting, navigate into the folder. You will see:

```
KBT_TestRun\
  excel\
    KeystoneBenefitTech_PL_Model.xlsx   ← YOUR TEST FILE
  vba\
    modAdmin_v2.1.bas
    modAllocation_v2.1.bas
    modAuditTools_v2.1.bas
    modAWSRecompute_v2.1.bas
    modConfig_v2.1.bas
    modConsolidation_v2.1.bas
    modDashboard_v2.1.bas
    modDataGuards_v2.1.bas
    modDataQuality_v2.1.bas
    modDataSanitizer_v2.1.bas
    modDemoTools_v2.1.bas
    modDrillDown_v2.1.bas
    modETLBridge_v2.1.bas
    modForecast_v2.1.bas
    modFormBuilder_v2.1.bas
    modImport_v2.1.bas
    modIntegrationTest_v2.1.bas
    modLogger_v2.1.bas
    modMasterMenu_v2.1.bas
    modMonthlyTabGenerator_v2.1.bas
    modNavigation_v2.1.bas
    modPDFExport_v2.1.bas
    modPerformance_v2.1.bas
    modReconciliation_v2.1.bas
    modScenario_v2.1.bas
    modSearch_v2.1.bas
    modSensitivity_v2.1.bas
    modTrendReports_v2.1.bas
    modUtilities_v2.1.bas
    modVarianceAnalysis_v2.1.bas
    modVersionControl_v2.1.bas
    (31 files total)
  qa\
    TEST_PLAN.md                         ← YOUR TEST CHECKLIST
```

---

## Step 3 — Prepare the Excel File

1. Copy `KeystoneBenefitTech_PL_Model.xlsx` somewhere safe
   (example: `C:\Users\YourName\Desktop\KBT_TestRun\`)

2. Rename the copy to `KeystoneBenefitTech_PL_Model_TEST.xlsm`
   - Important: change the extension from `.xlsx` to `.xlsm`
   - The `.xlsm` extension is required for Excel to save VBA macros inside the file

3. Open the renamed file in Excel

4. If Excel shows a yellow bar saying **"Macros have been disabled"** — click
   **Enable Content**

---

## Step 4 — Enable Trust Access to VBA (One-Time Setup)

You only need to do this once. Without it, importing the VBA modules will fail.

1. In Excel, click **File** → **Options**
2. Click **Trust Center** (bottom of the left sidebar)
3. Click **Trust Center Settings...**
4. Click **Macro Settings** on the left
5. Select **Enable all macros**
6. Check the box that says **Trust access to the VBA project object model**
7. Click **OK** → **OK**
8. **Close and reopen the Excel file**

---

## Step 5 — Import All 31 VBA Modules

This is the most important step. You are loading all the macro code into Excel.

1. Open the Excel file (the `.xlsm` file from Step 3)

2. Press **Alt + F11** on your keyboard
   - This opens the Visual Basic Editor (a separate window)

3. In the VBA Editor, click **File** in the top menu → **Import File...**

4. Navigate to your `vba\` folder

5. Select the first `.bas` file and click **Open**

6. Repeat steps 3–5 for every `.bas` file — all 31 of them
   - You must import each one individually
   - It takes about 2–3 minutes total

7. When done, click **File** → **Save** in the VBA Editor, then close it
   (press **Alt + F11** again or click the X on the VBA Editor window)

8. Back in Excel, press **Ctrl + S** to save the workbook

> **Tip:** After importing all modules, press **Alt + F11** again and look at
> the left panel (Project Explorer). You should see 31 module names listed
> under your workbook. If you see fewer, check which ones are missing.

---

## Step 6 — Verify Everything Loaded (Quick Check)

Before running any tests:

1. In the VBA Editor, click **Debug** → **Compile VBAProject**
   - If nothing happens, that is correct — it means zero errors
   - If a red error window appears, stop and report it

2. Click the **Immediate Window** at the bottom
   (if you don't see it: View → Immediate Window, or press Ctrl+G)

3. Type this exactly and press Enter:
   ```
   ?APP_VERSION
   ```
4. It should print: `2.1.0`
   - If it does, you are ready to test
   - If you get an error, stop and report it

---

## Step 7 — Open the Test Plan

Open `qa\TEST_PLAN.md` in any text editor (Notepad, Word, or a Markdown viewer).

Work through it in order:
- **Start with T1** (Compilation tests) — these must all pass before continuing
- Log each test as PASS, FAIL, or SKIP as you go
- Note any error messages exactly as they appear

---

## Quick Reference — Files Used During Testing

| File | Purpose |
|------|---------|
| `KeystoneBenefitTech_PL_Model_TEST.xlsm` | The workbook you run all tests in |
| `vba/*.bas` (all 31 files) | Already imported — no need to touch these again |
| `qa/TEST_PLAN.md` | Your test checklist |
| `qa/INTEGRATION_TEST_GUIDE.md` | Detailed walkthrough for the integration tests (T7) |
| `qa/VALIDATION_REPORT.md` | Reference for the 6 known acceptable FAIL checks |

---

## If Something Goes Wrong

- **"Macro not found" error:** The module that contains that macro may not have
  imported correctly. Re-import that specific `.bas` file.
- **"Compile error" on a specific line:** Write down the module name, line number,
  and error message exactly, and report back.
- **Excel crashes or freezes:** Save and reopen. If it keeps happening, check that
  you enabled macros and Trust Access (Step 4).
- **Any other error:** Write down the exact error message and which test you were
  running, and report it.
