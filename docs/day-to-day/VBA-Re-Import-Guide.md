# VBA Module Re-Import Guide — Step by Step

**Purpose:** Import the 10 updated `.bas` files from the repo into your Excel workbook.
This replaces the old versions of these modules with the bug-fixed, updated versions.

**Time needed:** About 5-10 minutes.

---

## Before You Start

1. **Save your Excel file** — File > Save (or Ctrl+S)
2. **Close any other Excel files** — only your demo file should be open
3. Make sure you have the latest repo files pulled (the `.bas` files in the `vba/` folder)

---

## How to Open the VBA Editor

1. Open your demo Excel file (`ExcelDemoFile_adv.xlsm`)
2. Press **Alt + F11** — the VBA Editor window opens
3. On the left side you'll see the **Project Explorer** panel
4. Look for your project — it will say something like `VBAProject (ExcelDemoFile_adv.xlsm)`
5. Click the **+** next to "Modules" to expand the module list

---

## Step-by-Step: Remove Old Module, Import New One

**You will repeat these 3 steps for each of the 10 files below.**

### Step A — Delete the Old Module

1. In the Project Explorer, find the module (example: `modConfig`)
2. **Right-click** on the module name
3. Click **Remove modConfig...**
4. A dialog asks "Do you want to export the module before removing it?"
5. Click **No** — we already have the updated version in the repo

### Step B — Import the New Module

1. In the VBA Editor menu bar, click **File**
2. Click **Import File...**
3. Navigate to your repo's `vba/` folder
4. Select the `.bas` file (example: `modConfig_v2.1.bas`)
5. Click **Open**
6. The module appears in the Project Explorer under "Modules"

### Step C — Verify It Imported

1. Double-click the newly imported module in Project Explorer
2. You should see the code appear in the main code window
3. Scroll to the top — confirm you see `Attribute VB_Name = "modConfig"` (or whatever module it is)
4. Move on to the next file

---

## The 10 Files to Re-Import (Do Them in This Order)

| # | File to Import | Module Name | What Changed |
|---|----------------|-------------|--------------|
| 1 | `modConfig_v2.1.bas` | modConfig | Color constant fixes (CLR_NAVY, CLR_ALT_ROW) |
| 2 | `modReconciliation_v2.1.bas` | modReconciliation | dateCol/amtCol constant fixes + LogAction fix |
| 3 | `modVarianceAnalysis_v2.1.bas` | modVarianceAnalysis | GenerateCommentary row fix + YoY Variance Analysis |
| 4 | `modDashboard_v2.1.bas` | modDashboard | Split — this is the base module only (charts) |
| 5 | `modDashboardAdvanced_v2.1.bas` | modDashboardAdvanced | **NEW module** — ExecDashboard, Waterfall, ProductComp |
| 6 | `modDemoTools_v2.1.bas` | modDemoTools | LogAction fix + new CreateDisclaimerSheet |
| 7 | `modTrendReports_v2.1.bas` | modTrendReports | LogAction fix |
| 8 | `modMonthlyTabGenerator_v2.1.bas` | modMonthlyTabGenerator | LogAction fixes + TestUpdateHeaderText wrapper |
| 9 | `modDataQuality_v2.1.bas` | modDataQuality | Letter Grade (A-F) feature + LogAction fix |
| 10 | `modPDFExport_v2.1.bas` | modPDFExport | LogAction fix |

### Important Notes for Specific Files

**File #4 (modDashboard):** This module was split in two. The old `modDashboard` had ~1,400 lines. The new one has ~530 lines. The rest moved to `modDashboardAdvanced`. So when you delete the old `modDashboard` and import the new one, it will look smaller — that's correct.

**File #5 (modDashboardAdvanced):** This is a **brand new module** — it won't exist in the workbook yet. Skip Step A (delete) and go straight to Step B (import).

---

## After All 10 Are Imported

### 1. Compile Check

1. In the VBA Editor, click **Debug** in the menu bar
2. Click **Compile VBAProject**
3. **What you want to see:** Nothing happens — that means it compiled with zero errors
4. **If you get an error:** Write down the exact error message and which module/line it points to

### 2. Quick Smoke Test

1. Close the VBA Editor (click the X or press Alt+Q)
2. In Excel, run the Command Center (click the button or run `LaunchCommandCenter`)
3. Try one simple action — like **Action 1 (Run Reconciliation)** — to confirm macros work
4. If it runs without errors, you're good

### 3. Save

1. **File > Save** (Ctrl+S)
2. Make sure it saves as `.xlsm` (macros-enabled)

---

## Troubleshooting

**"Sub or Function not defined" error after compile:**
- A module is referencing a sub that doesn't exist yet. Make sure all 10 files are imported.
- Check that `modDashboardAdvanced` was imported (File #5) — it's new and easy to miss.

**Module imported with wrong name:**
- Sometimes VBA names the module based on the `Attribute VB_Name` line inside the file.
- If a module shows up with the wrong name, right-click it > Remove > No > then re-import.

**"Name conflicts with existing module" error:**
- You forgot to delete the old version first (Step A). Delete it, then import again.

**Compile error on a specific line:**
- Write down the module name, line number, and error text. We'll fix it in the next session.

---

## After Re-Import: Run the Disclaimer Sheet

Once all 10 modules are imported and the compile passes:

1. Press **Alt + F8** (or go to Developer > Macros)
2. Find `CreateDisclaimerSheet` in the list
3. Click **Run**
4. A new "Disclaimer" sheet appears at the end of your workbook
5. It clearly states all financial data is fictional

---

*Guide created: 2026-03-05 | Branch: claude/resume-ipipeline-demo-qKRHn*
