# Getting Started — First Time Setup Guide

## iPipeline P&L Automation Toolkit — Setup from Scratch

**Estimated Time:** 10–15 minutes

---

## Table of Contents

1. [Before You Begin — What You Need](#1-before-you-begin--what-you-need)
2. [Step 1: Save the File to Your Computer](#2-step-1-save-the-file-to-your-computer)
3. [Step 2: Open the File in Excel](#3-step-2-open-the-file-in-excel)
4. [Step 3: Enable Macros (Critical)](#4-step-3-enable-macros-critical)
5. [Step 4: Configure Trust Center Settings (One Time Only)](#5-step-4-configure-trust-center-settings-one-time-only)
6. [Step 5: Verify the Toolkit Is Working](#6-step-5-verify-the-toolkit-is-working)
7. [Step 6: Run Your First 5 Actions](#7-step-6-run-your-first-5-actions)
8. [Step 7: Understand the Workbook Layout](#8-step-7-understand-the-workbook-layout)
9. [What to Expect Going Forward](#9-what-to-expect-going-forward)
10. [Troubleshooting — Common Setup Issues](#10-troubleshooting--common-setup-issues)
11. [Getting Help](#11-getting-help)

---

## 1. Before You Begin — What You Need

### Required

| Item | Details |
|---|---|
| **Computer** | Windows PC (Windows 10 or later) |
| **Excel Version** | Microsoft Excel Desktop — Excel 2019, Excel 2021, or Microsoft 365 (any plan that includes the desktop app) |
| **The P&L File** | The file named something like `ExcelDemoFile_adv.xlsm` (shared by the Finance Automation Team via SharePoint, email, or Teams) |
| **File Format** | The file MUST be `.xlsm` (macro-enabled workbook). If you see `.xlsx`, the macros are missing — contact the team for the correct file. |

### NOT Compatible

| Platform | Why Not |
|---|---|
| **Excel Online (browser)** | Does not support VBA macros at all |
| **Mac Excel** | VBA UserForms have known issues on Mac — some features may not work |
| **Google Sheets** | Does not support VBA |
| **LibreOffice Calc** | VBA compatibility is incomplete |
| **Mobile Excel (iPhone/Android/iPad)** | Does not support macros |

> **Bottom line:** You need **Windows Excel Desktop** (2019, 2021, or 365). If you are unsure which version you have, open Excel, click **File > Account** and look under "Product Information."

---

## 2. Step 1: Save the File to Your Computer

### If the File Was Shared via SharePoint or OneDrive

1. Open the SharePoint or OneDrive link in your web browser
2. Find the file (it will be named something like `ExcelDemoFile_adv.xlsm`)
3. Click the **three dots (...)** next to the file name
4. Click **"Download"**
5. The file will download to your **Downloads** folder
6. **Move the file** from Downloads to a permanent location on your computer. We recommend:
   - `C:\Users\YourName\Documents\iPipeline P&L\` (create this folder if it doesn't exist)
   - Or any folder on your local drive that you will remember

### If the File Was Shared via Email

1. Open the email containing the attachment
2. Click on the attached file name to download it
3. Save it to a permanent location (not your Downloads folder — it's easy to lose files there)
4. We recommend: `C:\Users\YourName\Documents\iPipeline P&L\`

### If the File Was Shared via Teams

1. Open the Teams chat or channel where the file was shared
2. Click on the file name
3. Click the **"..."** menu in the top right corner of the file preview
4. Click **"Download"**
5. Move the file to your permanent folder

### IMPORTANT: Do NOT Open the File Directly from SharePoint/OneDrive

If you open the file directly from a SharePoint or OneDrive link (without downloading first), it may open in "read-only" mode or in the browser, which does not support macros. **Always download the file first, then open it from your local computer.**

> **What you should see at this point:** The `.xlsm` file saved in a folder on your local computer (e.g., `C:\Users\YourName\Documents\iPipeline P&L\ExcelDemoFile_adv.xlsm`).

---

## 3. Step 2: Open the File in Excel

1. Navigate to the folder where you saved the file
2. **Double-click** the `.xlsm` file to open it in Excel
3. Wait for Excel to fully load the file (this may take 5–10 seconds for a large workbook)

### What You Might See

When Excel opens a file with macros, you will typically see one or more of these:

**Scenario A: Yellow Security Warning Bar**
- You see a yellow bar at the top of Excel that says: *"SECURITY WARNING: Macros have been disabled."*
- There is an **"Enable Content"** button on the right side of this bar
- **Action:** Click **"Enable Content"** — this is required for the toolkit to work
- This is normal and expected — Excel is protecting you from unknown macros

**Scenario B: Protected View Bar**
- You see a yellow or orange bar that says: *"PROTECTED VIEW: Be careful — files from the Internet can contain viruses."*
- There is an **"Enable Editing"** button
- **Action:** Click **"Enable Editing"** first, then look for the macro security warning (Scenario A) and click "Enable Content"

**Scenario C: No Warning at All**
- If your Trust Center is already configured (Step 4), the file may open without any warnings
- This is also normal — it means your settings are already correct

**Scenario D: "Macros in this workbook are disabled" Dialog**
- A dialog box pops up saying macros are disabled
- **Action:** You need to adjust your Trust Center settings (go to Step 4)

> **What you should see at this point:** The P&L workbook is open in Excel, and you have clicked "Enable Content" (if prompted). The workbook should show multiple sheet tabs at the bottom.

---

## 4. Step 3: Enable Macros (Critical)

**This is the most important step.** If macros are not enabled, none of the 62 Command Center actions will work.

### Quick Enable (Every Time You Open the File)

Every time you open the `.xlsm` file, Excel may show the yellow "Security Warning" bar. Simply click **"Enable Content"** each time.

### Why Does Excel Ask Every Time?

Excel's default security settings require you to confirm that you trust macros in each file, each time you open it. This is a security feature. You can change this behavior in Step 4 (Trust Center settings) so it only asks once.

---

## 5. Step 4: Configure Trust Center Settings (One Time Only)

This step changes your Excel settings so the P&L file's macros are automatically trusted. **You only need to do this once.** After this, the file will open without the yellow warning bar.

### Option A: Add a Trusted Location (Recommended)

This tells Excel: "Any file in this specific folder is trusted." This is the safest approach because it only trusts files in one folder.

**Step-by-step instructions:**

1. **Open Excel** (any workbook, or a blank workbook)

2. **Click "File"** in the top-left corner of Excel
   - You will see the green sidebar (Backstage view)

3. **Click "Options"** at the bottom of the left sidebar
   - The "Excel Options" dialog box will open

4. **Click "Trust Center"** in the left sidebar of the Options dialog
   - You will see Trust Center information on the right side

5. **Click the "Trust Center Settings..."** button on the right side
   - The Trust Center dialog will open

6. **Click "Trusted Locations"** in the left sidebar of the Trust Center
   - You will see a list of folders that Excel already trusts

7. **Click "Add new location..."** button at the bottom
   - A small dialog box will appear

8. **Click "Browse..."** and navigate to the folder where you saved the P&L file
   - Example: `C:\Users\YourName\Documents\iPipeline P&L\`

9. **Select the folder** and click **"OK"**

10. **Check the box** that says "Subfolders of this location are also trusted" (optional but recommended)

11. **Click "OK"** to close the Trusted Location dialog

12. **Click "OK"** to close the Trust Center

13. **Click "OK"** to close Excel Options

14. **Close and reopen the P&L file.** It should now open without any security warnings.

> **What you should see:** When you open the P&L file after adding the Trusted Location, there should be NO yellow security warning bar. The file opens clean and the macros are ready to use.

### Option B: Change Macro Security Settings (Alternative)

If you don't want to set up a Trusted Location, you can change the overall macro security setting. This is less targeted but also works.

**Step-by-step instructions:**

1. **Open Excel** (any workbook, or a blank workbook)

2. **Click "File"** > **"Options"** > **"Trust Center"** > **"Trust Center Settings..."**
   - (Same as Steps 1–5 in Option A)

3. **Click "Macro Settings"** in the left sidebar of the Trust Center

4. **Select "Disable VBA macros with notification"** (this is the default and recommended setting)
   - This means Excel will show the yellow bar and let you click "Enable Content" for each file

5. **Alternatively, select "Enable VBA macros"** (not recommended for general use, but acceptable if you only work with trusted files from your team)

6. **Click "OK"** to close the Trust Center

7. **Click "OK"** to close Excel Options

> **Recommendation:** Option A (Trusted Location) is safer and easier in the long run. Set it up once and forget about it.

### What About "Trust access to the VBA project object model"?

You may have heard about this setting. Here is the deal:

- This setting is **only needed if you want to automatically build the Command Center form using Mode A (programmatic creation)**
- For **normal daily use**, you do NOT need this setting enabled
- The Command Center form is already built into the file when you receive it
- **Leave this unchecked** unless specifically instructed otherwise

---

## 6. Step 5: Verify the Toolkit Is Working

Now let's make sure everything is set up correctly. Follow these steps exactly:

### Test 1: Open the Command Center

1. Make sure the P&L workbook is open and is the active window
2. Press **Ctrl + Shift + M** on your keyboard
3. **Expected result:** The "AUTOMATION COMMAND CENTER" window appears
4. If it appears — **the toolkit is working**
5. Click **"Close"** to close the Command Center for now

**If the Command Center does NOT appear:**
- Did you click "Enable Content" in the yellow security bar? (Go back to Step 3)
- Is the P&L workbook the active window? (Click on it first, then try again)
- Try pressing **Alt + F8** to see if "LaunchCommandCenter" is listed. If it is not listed, the VBA modules may not be properly loaded in the file — contact the Finance Automation Team

### Test 2: Check the Version Number

1. Press **Ctrl + Shift + M** to open the Command Center
2. Look at the top of the Command Center window
3. You should see **"Version 2.1.0"** (or a later version number)
4. If you see the version number, the core modules are loaded correctly
5. Click **"Close"**

### Test 3: Run a Safe Action

1. Press **Ctrl + Shift + M** to open the Command Center
2. In the **Search** box, type **"health"**
3. You should see **"Action 45: Quick Health Check"** appear in the list
4. Click on it to select it
5. Click **"Run & Close"**
6. **Expected result:** A message box appears showing 5 test results (all should say PASS)
7. Click **OK** to close the message box

**If all 5 checks show PASS:** Congratulations — the toolkit is fully operational.

**If any check shows FAIL:** Note which check failed and contact the Finance Automation Team with the details.

### Test 4: Navigate to the Home Sheet

1. Press **Ctrl + Shift + H** on your keyboard
2. **Expected result:** Excel navigates to the main "Report-->" sheet
3. This confirms the keyboard shortcuts are working

> **What you should see at this point:** The Command Center opens properly, shows version 2.1.0, and the Quick Health Check passes all 5 tests. You are ready to use the toolkit.

---

## 7. Step 6: Run Your First 5 Actions

Now that you have confirmed the toolkit is working, let's run 5 actions to familiarize yourself with how it works. These are safe, read-only actions that will not change your data.

### Your First 5 Actions (In This Order)

#### First Action: Scan for Data Quality Issues (Action 7)

1. Press **Ctrl + Shift + M** to open the Command Center
2. Click on **"Data Quality"** in the left category panel
3. Click on **"Scan for Data Quality Issues"** in the right action panel
4. Click **"Run & Close"**
5. **What happens:** A new sheet called "Data Quality Report" will be created showing any data issues found. You will also see a **Letter Grade** (A through F) at the top of the report.
6. **Look at the results.** Don't worry about fixing anything yet — this was just to show you what a data scan looks like.

#### Second Action: Run Reconciliation Checks (Action 3)

1. Press **Ctrl + Shift + M** to open the Command Center
2. Search for **"reconciliation"** in the search box
3. Click on **"Run Reconciliation Checks"**
4. Click **"Run & Close"**
5. **What happens:** The "Checks" sheet will be updated with PASS/FAIL results for each reconciliation check. Green cells are good, red cells need attention.

#### Third Action: Run Variance Analysis (Action 6)

1. Press **Ctrl + Shift + M** to open the Command Center
2. Click on **"Analysis"** in the left category panel
3. Click on **"Run Variance Analysis"**
4. Click **"Run & Close"**
5. **What happens:** A "Variance Analysis" sheet will be created showing month-over-month changes. Significant variances (over 15%) are highlighted.

#### Fourth Action: Build Dashboard Charts (Action 12)

1. Press **Ctrl + Shift + M** to open the Command Center
2. Search for **"dashboard"** in the search box
3. Click on **"Build Dashboard Charts"**
4. Click **"Run & Close"**
5. **What happens:** An "Executive Dashboard" sheet will be created with professional charts showing revenue trends, expense breakdowns, and product comparisons.

#### Fifth Action: Export Report Package as PDF (Action 10)

1. Press **Ctrl + Shift + M** to open the Command Center
2. Click on **"Reporting"** in the left category panel
3. Click on **"Export Report Package (PDF)"**
4. Click **"Run & Close"**
5. **What happens:** A file save dialog will appear. Choose a location (like your Desktop) and click Save. A polished multi-sheet PDF will be created.
6. **Open the PDF** to see the finished product — this is what you can send to leadership.

> **Congratulations!** You have just run 5 professional-grade financial analysis actions in under 5 minutes. This is the power of the Command Center.

---

## 8. Step 7: Understand the Workbook Layout

The P&L workbook contains multiple sheets organized by purpose. Here is what each sheet does:

### Core Reporting Sheets (These Are Your Main Working Sheets)

| Sheet Name | What It Contains | How Often You Use It |
|---|---|---|
| **Report-->** | The main landing page / report view. Start here. | Every time you open the file |
| **P&L - Monthly Trend** | Revenue and expense data by month, all on one sheet. This is the big picture view. | Monthly — review trends |
| **Functional P&L - Monthly Trend** | Same data as above, but organized by functional area (department) rather than natural account. | Monthly — departmental view |
| **Product Line Summary** | Revenue and expense breakdown by product line (iGO, Affirm, InsureSight, DocFast). | Monthly — product analysis |
| **Checks** | Reconciliation check results (PASS/FAIL for each check). | Monthly — after running Action 3 |

### Monthly Detail Sheets

| Sheet Name | What It Contains |
|---|---|
| **Functional P&L Summary - Jan 25** | Detailed P&L for January 2025 |
| **Functional P&L Summary - Feb 25** | Detailed P&L for February 2025 |
| **Functional P&L Summary - Mar 25** | Detailed P&L for March 2025 |
| *(Additional months are generated by Action 1)* | Apr through Dec added on demand |

### Input / Reference Sheets

| Sheet Name | What It Contains | Who Edits It |
|---|---|---|
| **Assumptions** | All the key drivers: growth rates, allocation percentages, product revenue shares. This is the ONLY sheet you should manually edit during normal use. | Finance team |
| **Data Dictionary** | Definitions of every data field used in the workbook. | Reference only — do not edit |
| **AWS Allocation** | AWS cost allocation details and methodology. | Finance team (quarterly) |

### Behind-the-Scenes Sheets (Usually Hidden)

| Sheet Name | What It Contains | Notes |
|---|---|---|
| **CrossfireHiddenWorksheet** | The raw GL (General Ledger) transaction detail. This is the source data that feeds everything else. | Normally hidden — do not edit |
| **VBA_AuditLog** | Automatic log of every action you run. | Normally hidden — view with Action 41 |
| **Scenarios** | Saved scenario snapshots. | Normally hidden — managed by Actions 20–23 |
| **Version History** | Saved version snapshots. | Normally hidden — managed by Actions 31–35 |

### Generated Sheets (Created by Actions)

These sheets are created when you run specific actions. They will not exist until you run the corresponding action for the first time:

| Sheet Name | Created By | What It Shows |
|---|---|---|
| **Data Quality Report** | Action 7 | Data scan results with letter grade |
| **Variance Analysis** | Action 6 | Month-over-month variance table |
| **Sensitivity Analysis** | Action 5 | What-if scenario results |
| **Executive Dashboard** | Action 12 | Professional charts and visuals |
| **Variance Commentary** | Action 46 | Auto-generated variance explanations |
| **Rolling Forecast** | Action 18 | Projected values for remaining months |
| **Search Results** | Action 15 (search) | Cross-sheet search results |
| **Validation Report** | Action 47 | Cross-sheet validation results |
| **Integration Test Report** | Action 44 | Test suite results |
| **Allocation Output** | Action 24 | Allocation calculation results |
| **Tech Documentation** | Action 36 | Auto-generated workbook documentation |
| **Change Management Log** | Action 37 | Change request tracking |

---

## 9. What to Expect Going Forward

### Daily Use

- Open the P&L file
- Press **Ctrl + Shift + M** to open the Command Center
- Run the actions you need
- Close the file when done (save if you made changes you want to keep)

### Monthly Close Process

The recommended workflow for month-end close is documented in the "How to Use the Command Center" guide (see the Monthly Close Workflow section). The short version:

1. Import new GL data (Action 17)
2. Scan for data quality issues (Action 7)
3. Fix any issues found (Action 8)
4. Run reconciliation checks (Action 3)
5. Run variance analysis (Action 6)
6. Build dashboard (Action 12)
7. Export PDF package (Action 10)

### Updates and New Versions

When the Finance Automation Team releases an updated version of the P&L file:

1. Download the new file
2. Save it to your iPipeline P&L folder (replacing the old file, or keeping both)
3. Open the new file — it will already have the latest macros
4. You do NOT need to re-do the Trust Center settings (if you set up a Trusted Location, any file in that folder is automatically trusted)

---

## 10. Troubleshooting — Common Setup Issues

### Issue 1: "Macros have been disabled" and there is no "Enable Content" button

**Cause:** Your macro security is set to "Disable all macros without notification."

**Fix:**
1. Close the P&L file
2. Open Excel (any blank workbook)
3. Go to **File > Options > Trust Center > Trust Center Settings > Macro Settings**
4. Select **"Disable VBA macros with notification"**
5. Click OK, OK
6. Reopen the P&L file — you should now see the "Enable Content" button

### Issue 2: I clicked "Enable Content" but the Command Center still won't open

**Cause:** The macro modules may not be loaded in the file.

**Fix:**
1. Press **Alt + F8** to open the Macro dialog
2. Look for "LaunchCommandCenter" in the list
3. If it is NOT in the list, the VBA modules are missing from the file — contact the Finance Automation Team for the correct `.xlsm` file
4. If it IS in the list, click it and click **Run**

### Issue 3: The file opens in "Protected View" and I can't do anything

**Cause:** The file came from the internet (email attachment, SharePoint download) and Excel is in Protected View.

**Fix:**
1. Look for the yellow or orange bar at the top that says "PROTECTED VIEW"
2. Click **"Enable Editing"** on the right side of that bar
3. After clicking Enable Editing, you may see a second yellow bar about macros — click **"Enable Content"** on that bar
4. The file should now be fully functional

### Issue 4: I see "Run-time error" when I try to run an action

**Cause:** Something unexpected happened during the action.

**Fix:**
1. Click **"End"** on the error dialog (do NOT click "Debug" unless you know VBA)
2. Try running the action again
3. If the same error appears, note the **error number** and **error description** (e.g., "Run-time error 1004: Method 'Range' of object '_Worksheet' failed")
4. Contact the Finance Automation Team with the error details

### Issue 5: The file is very slow to open or run actions

**Cause:** Excel may be recalculating the entire workbook on open, or the file is on a network drive.

**Fix:**
1. **Move the file to a local drive** (not a network drive or OneDrive synced folder). Local drives are much faster for large Excel files.
2. **Disable automatic calculation temporarily:** Go to **Formulas > Calculation Options > Manual**. Then press **F9** when you want to recalculate.
3. **Close other large Excel files** before opening the P&L file.
4. **Check your RAM:** If your computer has less than 8 GB of RAM, large workbooks will be slow. Close other programs to free up memory.

### Issue 6: The file extension is .xlsx instead of .xlsm

**Cause:** The file was saved as a regular workbook (without macros) instead of a macro-enabled workbook.

**Fix:**
- You have the wrong file. The correct file has the `.xlsm` extension. Contact the Finance Automation Team for the correct file.
- **Do NOT try to rename .xlsx to .xlsm** — this does not add the macros, it just changes the file name and will cause errors.

### Issue 7: I get a message about "Trusted Access to the VBA project object model"

**Cause:** You are trying to run the automatic Command Center builder (Mode A), which requires special permissions.

**Fix:**
- For normal daily use, you do NOT need this. The Command Center form is already built into the file.
- If you specifically need Mode A (told to by the Finance Automation Team), go to: **File > Options > Trust Center > Trust Center Settings > Macro Settings** and check the box **"Trust access to the VBA project object model"**. Then close and reopen the file.

### Issue 8: Some sheet tabs at the bottom are missing

**Cause:** Sheets may be hidden (which is normal for behind-the-scenes sheets like VBA_AuditLog and Scenarios).

**Fix:**
- Run **Action 52 (Unhide All Worksheets)** from the Command Center to make all hidden sheets visible
- Or right-click any sheet tab at the bottom and click **"Unhide..."** to see a list of hidden sheets

### Issue 9: I accidentally deleted a sheet

**Cause:** Sheet deletion cannot be undone in Excel.

**Fix:**
1. **If you haven't saved:** Close the file WITHOUT saving (click "Don't Save" when prompted). Reopen the file — the deleted sheet will be back.
2. **If you already saved:** Check if you have a saved version (Action 35 to list versions). If so, you may be able to restore from a version (Action 34).
3. **If no version exists:** Contact the Finance Automation Team. They can provide a fresh copy of the file.

### Issue 10: "Compile error" appears when I try to open the Command Center

**Cause:** A VBA module has a coding error.

**Fix:**
1. Click **"OK"** on the error dialog
2. If the VBA Editor opens, close it (**Alt + Q**)
3. Do NOT try to fix the code yourself
4. Contact the Finance Automation Team with a screenshot of the error

---

## 11. Getting Help

### Self-Service Resources

| Resource | What It Covers | Where to Find It |
|---|---|---|
| **How to Use the Command Center** guide | All 62 actions explained in detail | Ask the Finance Automation Team |
| **Quick Reference Card** | 1-page cheat sheet of all actions | Ask the Finance Automation Team |
| **Action 50: About This Toolkit** | Version and build info | Run from Command Center |
| **Action 45: Quick Health Check** | 5-point workbook health test | Run from Command Center |
| **Action 44: Full Integration Test** | 18-test comprehensive check | Run from Command Center |

### Contact the Finance Automation Team

If you need help with setup, encounter an error, or have questions:

1. **Take a screenshot** of the error or issue
2. **Note which action** you were trying to run (if applicable)
3. **Note your Excel version** (File > Account > Product Information)
4. **Send all of this** to the Finance Automation Team via email or Teams

### Common Questions

**Q: Can I use this on a shared network drive?**
A: It will work, but performance will be slower. We recommend keeping a copy on your local drive for best performance.

**Q: Can two people have the file open at the same time?**
A: No. Excel macro-enabled workbooks should only be opened by one person at a time. If you need to share results, use the PDF export (Action 10).

**Q: Will this work if I upgrade my Excel version?**
A: Yes. The toolkit is compatible with Excel 2019, 2021, and Microsoft 365 on Windows. Future Excel updates will not break the macros.

**Q: I'm on a Mac. Can I use this?**
A: Not reliably. The Command Center uses VBA UserForms which have known issues on Mac Excel. We recommend using a Windows PC.

---

## Setup Checklist

Use this checklist to confirm you have completed all setup steps:

- [ ] Downloaded the `.xlsm` file to a local folder on my computer
- [ ] Opened the file in Windows Excel Desktop (2019, 2021, or 365)
- [ ] Clicked "Enable Content" when prompted (or set up Trusted Location)
- [ ] Pressed Ctrl+Shift+M and the Command Center opened successfully
- [ ] Verified version number shows 2.1.0 (or later)
- [ ] Ran Action 45 (Quick Health Check) and all 5 tests passed
- [ ] Ran my first 5 actions (Steps 7.1 through 7.5) successfully
- [ ] I know where to find help if I need it

**If all boxes are checked — you are ready to use the iPipeline P&L Automation Toolkit.**

---

## Document Information

| Field | Value |
|---|---|
| **Document Title** | Getting Started — First Time Setup Guide |
| **Version** | 1.0 |
| **Last Updated** | March 5, 2026 |
| **Author** | Finance Automation Team |
| **Audience** | All iPipeline Employees (First-Time Setup) |
| **Estimated Setup Time** | 10–15 minutes |
| **Prerequisites** | Windows PC + Excel Desktop (2019/2021/365) |

---

*This document is part of the iPipeline P&L Automation Toolkit documentation suite. After completing setup, see "How to Use the Command Center" for a complete guide to all 62 actions.*
