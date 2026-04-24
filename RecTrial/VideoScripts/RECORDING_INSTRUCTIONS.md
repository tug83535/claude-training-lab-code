# VIDEO RECORDING INSTRUCTIONS — Step-by-Step Clickthrough Guide

**Purpose:** Keep this document open on your second monitor while recording in Camtasia. Silent recording — pre-made audio plays in headphones, you click through at the right pace.

**Recording Settings:** 1920x1080, 30fps, H.264 MP4. Excel zoom 110%. Windows scaling 100%. All notifications OFF (Teams, Outlook, Windows, OneDrive). Taskbar auto-hidden.

---

# PRE-RECORDING CHECKLIST (Do This Before ANY Video)

- [ ] Close all apps except Excel (and Command Prompt for Video 3 Python clips)
- [ ] Desktop wallpaper: solid dark color (no icons visible)
- [ ] Excel: File > Options > General > uncheck "Show Start screen"
- [ ] Excel: View > uncheck Gridlines (cleaner look on camera)
- [ ] Camtasia: recording region locked to Excel window
- [ ] Second monitor: this document open, scrolled to the right video section
- [ ] Headphones: audio script loaded, tested at comfortable volume
- [ ] Do a 10-second test recording and play it back to check quality

---
---

# VIDEO 1 — "WHAT'S POSSIBLE" (~5 min, 7 clips)

## Pre-Setup for Video 1

**File state:** Demo file open, on the "Report-->" landing page. No output sheets should exist yet (no Data Quality Report, no Variance Analysis, no Dashboard, etc.). If they exist from a previous run, delete them first.

**How to get clean state:**
1. Open the demo file
2. Right-click and delete these sheets if they exist: Data Quality Report, Variance Analysis, Variance Commentary, Executive Dashboard, Sensitivity Analysis, YoY Variance Analysis, Time Saved Analysis, Executive Brief, Integration Test Report
3. Save
4. Close and reopen (so splash screen fires fresh — dismiss it for now)
5. Navigate to the "Report-->" tab

**Command Center:** Should be CLOSED when recording starts. You will open it on camera.

---

### CLIP 1: Title Card (5 sec)

**What to record:** Static title card. You can add this in Camtasia post-production or hold on a branded slide.

**Action:** Record 5 seconds of nothing (or skip — add in editing).

---

### CLIP 2: Opening Hook — Landing Page (30 sec)

**Screen before:** Excel open, "Report-->" sheet visible. This is the styled landing page with the iPipeline branding.

**Actions:**
1. Start recording
2. Slowly scroll down through the Report--> sheet so the viewer sees the full layout
3. Pause on any branded header or summary section for 2-3 seconds
4. The audio will mention "65 automated actions" — just hold steady on the landing page

**Expected screen:** Clean branded landing page with iPipeline colors (navy headers, blue accents).

**Gotcha:** Make sure no random cell is selected with a blinking cursor in a weird spot. Click on cell A1 before starting.

---

### CLIP 3: Command Center Overview (40 sec)

**Screen before:** Still on "Report-->" sheet.

**Actions:**
1. Press **Ctrl+Shift+M** to open the Command Center
2. Wait 1-2 seconds for the form to appear
3. Slowly scroll through the Command Center list so all 65 actions are visible (the viewer needs to see how many there are)
4. In the search box at the top, type **variance** — this filters the list to show only variance-related actions
5. Pause 2-3 seconds on the filtered results
6. Clear the search box
7. Close the Command Center (click X or Cancel)

**Expected screen:** frmCommandCenter UserForm pops up with a list of all 65 actions organized by category. Search filters in real time.

**Gotcha:** If frmCommandCenter doesn't exist in the workbook, it will fall back to an InputBox menu — this looks much worse on camera. Make sure the UserForm is built and working before recording. Test Ctrl+Shift+M first.

**Gotcha #2:** If "Trust access to the VBA project object model" is not enabled, the form may not build. Verify beforehand: File > Options > Trust Center > Trust Center Settings > Macro Settings > check "Trust access to the VBA project object model."

---

### CLIP 4: Data Quality Scan (40 sec)

**Screen before:** Command Center closed, on any sheet.

**Actions:**
1. Press **Ctrl+Shift+M** to open Command Center
2. Type **7** in the action number box (or click Action 7: "Data Quality Scan")
3. Click Run/OK
4. Wait for the scan to complete (takes 2-5 seconds depending on file size)
5. Excel will auto-navigate to the new "Data Quality Report" sheet
6. Slowly scroll down through the report — the viewer needs to see:
   - The letter grade badge at the top (A, B, C, D, or F)
   - Category breakdown: Text-Stored Numbers, Blank Rows, Duplicates, Formula Errors, etc.
   - Color-coded cells (green = good, red = issues found)
7. Pause 2-3 seconds on the letter grade

**Expected output:** A new sheet called "Data Quality Report" with a styled header, letter grade (likely A or B on the clean demo file), and detailed breakdown by category.

**Gotcha:** If the file is perfectly clean, the letter grade will be A and all categories will show 0 issues. This is fine — it shows the tool works. If you want a more dramatic demo, you can intentionally break something first (paste a text "123" in a number column), but this adds complexity.

---

### CLIP 5: Variance Commentary (45 sec)

**Screen before:** On the Data Quality Report sheet (from previous clip).

**Actions:**
1. Press **Ctrl+Shift+M** to open Command Center
2. Type **46** in the action number box (Action 46: "Generate Variance Commentary")
3. Click Run/OK
4. Wait for generation (3-5 seconds)
5. Excel will navigate to the new "Variance Commentary" sheet
6. **PAUSE 2-3 seconds** — this is the "jaw drop" moment. Let the viewer read.
7. Slowly scroll down through the auto-generated narratives
8. Each row should have: line item name, variance amount, percentage, and a plain-English narrative like "Revenue decreased 12.5% month-over-month, declining from $X to $Y..."
9. Pause on one particularly good narrative for 3 seconds

**Expected output:** "Variance Commentary" sheet with styled headers and auto-generated plain-English commentary for every P&L line item with a material variance.

**Gotcha:** This reads from the P&L Monthly Trend sheet. If that sheet has no data or only one month, the commentary will be empty or minimal. Make sure at least 2 months of data exist.

---

### CLIP 6: Executive Dashboard (40 sec)

**Screen before:** On the Variance Commentary sheet.

**Actions:**
1. Press **Ctrl+Shift+M** to open Command Center
2. Type **12** in the action number box (Action 12: "Build Dashboard")
3. Click Run/OK
4. Wait for chart generation (5-10 seconds — this builds 8 charts)
5. Excel navigates to the dashboard sheet
6. Slowly scroll/pan across the dashboard to show all charts:
   - Revenue trend line
   - Expense breakdown
   - Product comparison bars
   - Waterfall chart
   - KPI summary cards
7. Pause on the waterfall chart for 2-3 seconds (most visually impressive)

**Expected output:** "Executive Dashboard" sheet with 8 branded charts in iPipeline colors (navy, blue, lime green accents).

**Gotcha:** Chart generation can sometimes throw an error if expected data columns are missing. Do a test run before recording. If any chart fails, the others should still build — but it looks bad on camera.

**Gotcha #2:** If an old Executive Dashboard sheet already exists, the macro will delete and rebuild it. This is fine, but you'll see a brief flash. If this bothers you, delete the old one manually first.

---

### CLIP 7: Bridge to Universal Tools + Closing (30 sec + 5 sec card)

**Screen before:** On the Executive Dashboard sheet showing the charts.

**Actions:**
1. No macro to run here — this is a talking-head moment (audio only)
2. While the audio mentions "140+ universal tools that work on ANY Excel file," slowly scroll back to the landing page (Report--> tab)
3. Hold on the landing page for the closing card audio
4. End recording

**Expected screen:** Landing page, static. The audio does the work here.

---
---

# VIDEO 2 — "FULL DEMO WALKTHROUGH" (~18 min, 19 clips)

## Pre-Setup for Video 2

**CRITICAL — Clean file state required.** The file must be in a "fresh" state as if no macros have been run yet.

**How to get clean state:**
1. Open the demo file
2. Delete ALL output sheets: Data Quality Report, Variance Analysis, Variance Commentary, Executive Dashboard, Sensitivity Analysis, YoY Variance Analysis, Time Saved Analysis, Executive Brief, Integration Test Report, What-If Impact
3. Delete hidden sheets if they exist: WhatIf_Baseline
4. Clear the Checks sheet (select all data below headers, delete)
5. Clear the VBA_AuditLog sheet (unhide it: right-click any tab > Unhide > VBA_AuditLog, then clear all data below headers, then re-hide it: right-click > Hide)
6. Make sure no version snapshots exist (check for sheets starting with "VER_")
7. Save the file
8. **Close Excel completely**
9. Reopen the file — the splash screen will fire on open (this is Clip 1)
10. **DO NOT dismiss the splash screen yet — start recording first**

Wait — actually, you need Camtasia already recording when you open the file. So:
1. Do steps 1-7 above in advance
2. Save and close Excel
3. Start Camtasia recording
4. Double-click the .xlsm file to open it
5. The splash screen fires — that's your first clip

---

### CLIP 1: File Opens + Splash Screen (15 sec)

**Screen before:** Desktop or Excel loading screen. Camtasia is already recording.

**Actions:**
1. Double-click the .xlsm file on your desktop (or from File Explorer)
2. Excel opens → the splash screen fires automatically (from modSplashScreen)
3. The splash shows:
   - "KEYSTONE BENEFITECH"
   - "P&L Reporting & Allocation Model"
   - Version 2.1.0, 34 VBA Modules, 62 Command Center Actions
   - "Press Ctrl+Shift+M to open the Command Center"
4. Read it for 3-4 seconds
5. Click OK
6. A second dialog asks "Would you like to open the Command Center?" — click **No** (you'll open it later)

**Expected screen:** MsgBox splash with branding info, then the file opens to the Report--> landing page.

**Gotcha:** If macros are disabled, the splash won't fire. Make sure you click "Enable Content" on the security bar if it appears. Better yet, add the file location to Trusted Locations beforehand: File > Options > Trust Center > Trust Center Settings > Trusted Locations > Add new location.

**Gotcha #2:** If frmSplash UserForm exists, it shows a nicer form that auto-dismisses after 5 seconds. If it doesn't exist, the MsgBox fallback fires. Both work, but the UserForm looks better.

---

### CLIP 2: Workbook Tour — All Sheets (45 sec)

**Screen before:** File is open, on the Report--> landing page.

**Actions:**
1. Click on each sheet tab slowly, left to right, pausing 2-3 seconds on each:
   - **Report-->** (landing page — already here)
   - **P&L - Monthly Trend** (the main data — pause here, scroll right to show months)
   - **Product Line Summary** (product breakdown)
   - **Functional P&L - Monthly Trend** (functional expense view)
   - **Assumptions** (driver sheet — pause here, this is important for What-If later)
   - **Checks** (reconciliation — should be empty/blank at this point)
   - **Charts & Visuals** (if it exists — pre-built charts)
   - Any monthly tabs (Mar 25, etc.)
2. Navigate back to Report--> when done

**Expected screen:** Each sheet shows structured, branded data. The viewer sees the scope of the workbook.

**Gotcha:** Don't click too fast. The audio will be describing each sheet — give it time to match. If you have 13+ tabs, you might need to use the tab scroll arrows at the bottom left.

---

### CLIP 3: Command Center Overview (45 sec)

**Screen before:** On Report--> landing page.

**Actions:**
1. Press **Ctrl+Shift+M** to open the Command Center
2. Slowly scroll through all 65 actions
3. Point out the categories visually (the list is grouped):
   - Monthly Operations (1-5)
   - Analysis & Data Quality (6-12)
   - Navigation & Import (13-17)
   - Forecast & Scenario (18-25)
   - Consolidation & Version Control (26-35)
   - Admin & Testing (36-50)
   - Utilities (51-62)
   - What-If Demo (63-65)
4. Demo the search: type **"reconciliation"** — shows filtered results
5. Clear search
6. Close the Command Center

**Expected screen:** Same as Video 1 Clip 3 but you're spending a bit more time here.

---

### CLIP 4: GL Import (45 sec)

**Screen before:** Command Center closed, on any sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 17: Import Data Pipeline**
3. Click Run/OK
4. A file picker dialog appears — navigate to your test CSV/Excel file
   - **SETUP REQUIRED:** Have a sample GL data file ready on your desktop. Use a .csv or .xlsx with GL columns (Date, Account, Description, Amount, etc.)
5. Select the file and click Open
6. The import runs — status bar shows progress
7. When complete, the imported data appears on the GL detail sheet

**Expected output:** Data imported into the workbook. A success MsgBox appears.

**Gotcha:** If you don't have a sample import file ready, this will just show a file picker and you'll have nothing to pick. **Prepare the file in advance.**

**Gotcha #2:** If the import file has unexpected column headers, the import may fail or map incorrectly. Use a file with known headers that match what modImport expects.

**Alternative:** If the import is too risky for a live recording, you can skip this clip and mention it verbally. The GL data may already be in the file.

---

### CLIP 5: Data Quality Scan (40 sec)

**Screen before:** On any sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 7: Data Quality Scan** (type 7)
3. Click Run/OK
4. Wait 2-5 seconds for scan
5. Auto-navigates to "Data Quality Report" sheet
6. Scroll down slowly through the full report:
   - Letter grade at top (A/B/C/D/F)
   - Each category: Text-Stored Numbers, Blank Rows, Duplicates, Formula Errors, Missing Values
   - Issue counts and locations
7. Pause on the letter grade for 2-3 seconds

**Expected output:** "Data Quality Report" sheet with styled header and category breakdown.

**Gotcha:** Same as Video 1 — on a clean file you'll get grade A. That's fine for this demo.

---

### CLIP 6: Reconciliation (45 sec)

**Screen before:** On Data Quality Report sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 3: Reconciliation Checks** (type 3)
3. Click Run/OK
4. Wait 2-5 seconds
5. Auto-navigates to "Checks" sheet
6. Scroll through the results:
   - Each check has a description and PASS/FAIL status
   - Color-coded: green cells = PASS, red cells = FAIL
   - Cross-sheet validations show which sheets were compared
7. Pause on the PASS/FAIL column for 2-3 seconds

**Expected output:** "Checks" sheet populated with reconciliation results. Most should be PASS on the demo file.

**Gotcha:** If some checks fail, that's actually fine — it shows the tool catches real issues. But if you want all-PASS for the demo, make sure the data is internally consistent.

---

### CLIP 7: Variance Analysis — Month over Month (40 sec)

**Screen before:** On Checks sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 6: Run Variance Analysis** (type 6)
3. Click Run/OK
4. Wait 2-5 seconds
5. Auto-navigates to "Variance Analysis" sheet
6. Scroll through:
   - Line items with dollar variances and percentage variances
   - Highlighted cells for material variances (>15% threshold by default)
   - Positive variances (favorable) vs negative (unfavorable)
7. Pause on a highlighted material variance for 2-3 seconds

**Expected output:** "Variance Analysis" sheet comparing the two most recent months.

**Gotcha:** Needs at least 2 months of data in P&L Monthly Trend. If only 1 month exists, the analysis will be empty.

---

### CLIP 8: Variance Commentary (45 sec)

**Screen before:** On Variance Analysis sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 46: Generate Variance Commentary** (type 46)
3. Click Run/OK
4. Wait 3-5 seconds
5. Auto-navigates to "Variance Commentary" sheet
6. **PAUSE 2-3 seconds** — let the viewer absorb the auto-generated English narratives
7. Slowly scroll down, reading one or two narratives yourself
8. Each narrative explains what changed, by how much, and what it means

**Expected output:** "Variance Commentary" sheet with plain-English paragraphs for each material variance.

**THIS IS THE JAW-DROP MOMENT.** The audio will emphasize this. Let the screen breathe — don't rush past it.

---

### CLIP 9: Year-over-Year Variance (30 sec)

**Screen before:** On Variance Commentary sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. You need to run the YoY Variance action. Check your Command Center for this — it may be part of Action 6 or a separate action in modVarianceAnalysis.
   - Look for "YoY Variance" or "Year over Year" in the Command Center search
3. Click Run/OK
4. Wait for results
5. Navigate to the "YoY Variance Analysis" sheet (if created)
6. Scroll through — similar format to MoM but comparing same month across years

**Expected output:** "YoY Variance Analysis" sheet comparing current period to same period last year.

**Gotcha:** This requires data from two different fiscal years. If the demo file only has one year of data, this may produce empty results or an error. **Test this beforehand.**

---

### CLIP 10: Dashboard Charts (45 sec)

**Screen before:** On YoY Variance sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 12: Build Dashboard** (type 12)
3. Click Run/OK
4. Wait 5-10 seconds (builds 8 charts)
5. Auto-navigates to dashboard sheet
6. Slowly scroll/pan across all 8 charts:
   - Revenue trend (line chart)
   - Expense breakdown (pie or bar)
   - Product comparison (grouped bars)
   - Monthly P&L bars
   - Waterfall chart (revenue to net income)
   - Any other charts built by the module
7. Pause on each chart for 2-3 seconds
8. End on the waterfall chart (most visually impressive)

**Expected output:** Full dashboard with 8 branded charts in iPipeline colors.

**Gotcha:** Same as Video 1 — test first. Chart errors mid-recording look bad.

---

### CLIP 11: Executive Dashboard (30 sec)

**Screen before:** On the dashboard charts.

**Actions:**
1. The Executive Dashboard may already be built as part of Action 12, or it may be a separate step
2. If separate: Press **Ctrl+Shift+M**, look for "Executive Dashboard" or "Create Executive Dashboard"
3. Navigate to the "Executive Dashboard" sheet
4. This shows:
   - KPI summary cards (Revenue, Expenses, Net Income, Margin)
   - Key metrics at a glance
   - Color-coded status indicators
5. Pause on the KPI cards for 2-3 seconds

**Expected output:** Clean executive summary with KPI cards and key metrics.

---

### CLIP 12: PDF Export (30 sec)

**Screen before:** On Executive Dashboard.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 10: Export Report Package** (type 10)
3. Click Run/OK
4. A save dialog appears — choose your Desktop for easy access
5. The export runs (5-10 seconds — exports multiple sheets to PDF)
6. A success MsgBox appears showing the file path
7. **Optional:** Minimize Excel briefly, open the PDF on your desktop to show the output (professional headers/footers, page numbers, iPipeline branding)

**Expected output:** A multi-page PDF file saved to the location you chose. Contains all report sheets with print headers/footers.

**Gotcha:** The PDF includes 7+ sheets. If any sheet doesn't exist yet (e.g., you skipped a step), the export might skip it or error. Make sure all report sheets exist from previous clips.

**Gotcha #2:** If no default PDF printer is configured, the export may fail. Test beforehand.

---

### CLIP 13: Executive Brief (40 sec)

**Screen before:** Back in Excel after PDF export.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Search for **"Executive Brief"** or find the action for `GenerateExecBrief`
   - This may not have a Command Center action number — check if it's wired up. If not, run from VBA editor: Alt+F11 > Immediate Window > type `modExecBrief.GenerateExecBrief` and press Enter
   - **Better approach:** If it IS wired to an action number, use that
3. Wait 3-5 seconds for the scan
4. Auto-navigates to "Executive Brief" sheet
5. Scroll through the 5 sections:
   - **Revenue & P&L Highlights** — top-line numbers
   - **Reconciliation Status** — summary of checks
   - **Key Assumptions & Drivers** — from the Assumptions sheet
   - **Product Line Overview** — product performance
   - **Workbook Health** — error count, formula count, hidden sheets, external links
6. Each section has color-coded status indicators (green/yellow/red)
7. Pause on Workbook Health for 2-3 seconds

**Expected output:** "Executive Brief" sheet with 5 styled sections, Navy headers, wrapped text, color-coded indicators.

**Gotcha:** Make sure the Assumptions sheet and Checks sheet have data (from earlier clips). If they're empty, those sections of the brief will be thin.

---

### CLIP 14: Executive Mode Toggle (20 sec)

**Screen before:** On Executive Brief sheet. Multiple tabs visible at the bottom.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 48: Toggle Executive Mode** (type 48)
3. Click Run/OK
4. Watch the tabs at the bottom — several sheets will HIDE, leaving only the key executive-level sheets visible
5. Pause 2 seconds to show the reduced tab list
6. Press **Ctrl+Shift+M** again
7. Select **Action 48** again to TOGGLE BACK
8. All sheets reappear

**Expected output:** Sheets toggle between full view (all tabs) and executive view (only key summary sheets).

**Gotcha:** This uses xlSheetHidden, so sheets can still be manually unhidden. The toggle should work cleanly back and forth. Test it twice beforehand.

---

### CLIP 15: Version Control — Save Snapshot (30 sec)

**Screen before:** All sheets visible again after toggling Executive Mode back.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 32: Save Version** (type 32)
3. Click Run/OK
4. An InputBox appears asking for a version name
5. Type: **"March Close Draft 1"**
6. Click OK
7. A success MsgBox confirms the snapshot was saved
8. **Optional:** Show the version by pressing Ctrl+Shift+M > Action 35 (List Versions) to display a list of saved versions

**Expected output:** Version snapshot saved internally. If you list versions, you'll see "March Close Draft 1" with a timestamp.

**Gotcha:** This creates a hidden sheet (prefixed "VER_"). If the workbook already has many versions, the file size grows. For the demo, one version is enough.

---

### CLIP 16: What-If Scenario Demo — THE WOW MOMENT (90 sec)

**Screen before:** On any sheet. This is the Tier 1 wow moment — take your time.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 63: Run What-If Demo** (type 63)
3. Click Run/OK
4. A menu appears with 9 options:
   ```
   1. Revenue drops 15%
   2. Revenue increases 10%
   3. AWS costs increase 25%
   4. Headcount grows 20%
   5. All expenses cut 10%
   6. Best case: Revenue +15%, Expenses -5%
   7. Worst case: Revenue -20%, Expenses +15%
   8. Custom (pick your own driver & %)
   9. Restore original values
   ```
5. Type **1** and click OK (Revenue drops 15%)
6. Wait 3-5 seconds — the macro:
   - Saves the current Assumptions as a baseline (first run only, to hidden WhatIf_Baseline sheet)
   - Modifies the Assumptions sheet to reflect -15% revenue
   - Recalculates the workbook
   - Creates a "What-If Impact" sheet showing before/after comparison
7. Auto-navigates to "What-If Impact" sheet
8. **PAUSE 3-4 seconds** — let the viewer see the impact numbers
9. Scroll through the impact report:
   - Shows which drivers changed
   - Before value → After value → Dollar impact → Percentage change
   - Ripple effects across P&L line items
10. Now restore: Press **Ctrl+Shift+M**
11. Select **Action 65: Restore Baseline** (type 65)
12. Click Run/OK
13. A confirmation dialog appears — click Yes
14. The Assumptions sheet restores to original values
15. The What-If Impact sheet is cleaned up
16. A success message confirms restoration

**Expected output:** "What-If Impact" sheet showing the full ripple effect of revenue dropping 15%. Then clean restoration to original state.

**Gotcha (CRITICAL):** The baseline is saved on FIRST RUN ONLY. If you ran What-If before and didn't restore, the baseline may already be modified. **Always restore before the demo recording session, or delete the WhatIf_Baseline hidden sheet so it saves a fresh baseline.**

**Gotcha #2:** After running What-If, the Assumptions sheet has modified values. If you DON'T restore before continuing the demo, all subsequent clips will show the modified numbers. **Always restore before moving to Clip 17.**

---

### CLIP 17: Integration Test (30 sec)

**Screen before:** Back to normal state after restoring What-If baseline.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 44: Run Full Integration Test** (type 44)
3. Click Run/OK
4. Wait 5-10 seconds (runs 18 tests)
5. Auto-navigates to "Integration Test Report" sheet
6. Scroll through:
   - Each test has a name, description, and PASS/FAIL status
   - Look for the summary at the top: "18/18 PASS" (or similar)
   - Green rows = PASS, Red rows = FAIL
7. Pause on the summary line for 2-3 seconds

**Expected output:** "Integration Test Report" with 18/18 tests passing (all green).

**Gotcha:** If any test fails, it will show red. On a properly set up demo file, all should pass. If something fails, it's usually because a required sheet is missing or data was modified. **Run this test BEFORE recording to verify all 18 pass.**

---

### CLIP 18: Audit Log (25 sec)

**Screen before:** On Integration Test Report.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Select **Action 41: View Audit Log** (type 41)
3. Click Run/OK
4. The hidden "VBA_AuditLog" sheet becomes visible and activates
5. Scroll through the log entries:
   - Timestamp | Module | Action | Status | Details
   - You should see entries from every macro you ran during the demo
   - This proves every action was logged automatically
6. Pause for 2-3 seconds — the viewer sees a full audit trail from the entire demo
7. **Optional:** Right-click the VBA_AuditLog tab > Hide (to re-hide it)

**Expected output:** "VBA_AuditLog" sheet with timestamped entries for every action run during this recording session.

**Gotcha:** If the log is empty, the LogAction calls aren't working. This should not happen on a properly imported file, but test beforehand.

---

### CLIP 19: Time Saved Calculator — Closing Moment (45 sec)

**Screen before:** On VBA_AuditLog or any sheet.

**Actions:**
1. Press **Ctrl+Shift+M**
2. Search for **"Time Saved"** or find the action for `ShowTimeSavedReport`
   - Check if this is wired to a Command Center action number. If not, run from: Alt+F11 > Immediate Window > `modTimeSaved.ShowTimeSavedReport`
3. Wait 3-5 seconds
4. Auto-navigates to "Time Saved Analysis" sheet
5. Scroll through the report:
   - Table with all 62 actions listed
   - Columns: Action Name, Category, Manual Time (min), Automated Time (min), Time Saved (min), Savings %
   - Each row shows the comparison
6. Scroll to the bottom for the **Executive Summary box**:
   - "Manual: X hours per monthly close"
   - "Automated: Y hours per monthly close"
   - "Saved: Z hours per monthly close"
   - **"Annual: (Z x 12) hours per year"** ← THIS is the closing number
7. **PAUSE 3-4 seconds on the annual savings number** — this is the mic-drop closing moment
8. End recording

**Expected output:** "Time Saved Analysis" sheet with full ROI breakdown and annual savings number.

**Gotcha:** The time estimates are hardcoded in the module. They should be reasonable but review them before recording to make sure the annual number looks impressive and believable.

---
---

# VIDEO 3 — "UNIVERSAL TOOLS" (~10 min, 13 clips)

## Pre-Setup for Video 3

**CRITICAL — This video uses a DIFFERENT file.** NOT the Finance Automation demo file.

**Sample file:** Use `Sample_Quarterly_Report.xlsx` (already built and in the repo under `videodraft/`). If you need to rebuild it, run `python/build_sample_file.py`.

**What the sample file should contain (intentional mess):**
- 4-6 sheets with different data (Sales, Expenses, Inventory, etc.)
- Text-stored numbers (numbers formatted as text — green triangle in corner)
- Floating-point noise (e.g., 99.99999999997 instead of 100)
- Blank rows scattered in the data
- Leading/trailing spaces in text cells
- Merged cells in a column (e.g., department names)
- Mixed date formats (1/15/2025, Jan 15, 2025, 2025-01-15)
- Some cells with comments/notes
- At least 1 hidden sheet
- No iPipeline branding (raw/unstyled)
- Duplicate values in at least one column

**Universal toolkit modules must be imported into this sample file.** Before recording:
1. Open Sample_Quarterly_Report.xlsx
2. Alt+F11 > File > Import > import all 23 modUTL_*.bas files from `UniversalToolsForAllFiles/vba/`
3. Save as .xlsm (macro-enabled)
4. Close and reopen
5. Test each macro you plan to demo (see clips below)

**Also needed for Python clips:** Command Prompt open and ready. Have two test Excel files on your desktop for the compare script. Have a sample PDF with tables for the PDF extractor.

---

## CHAPTER 1 — Data Cleanup (3 clips)

### CLIP 1: Data Sanitizer (60 sec)

**Screen before:** Sample file open, on a sheet with messy data (text-stored numbers, FP tails visible).

**Setup:** Select a range or sheet that has visible text-stored numbers (green triangles in cell corners) and floating-point tails (numbers with long decimal strings).

**Actions:**
1. First, show the PREVIEW: Alt+F8 > type `PreviewSanitizeChanges` > Run
   - Or use the Universal Command Center if imported: run `LaunchUTLCommandCenter` > select "Preview Sanitize Changes"
2. Wait 2-3 seconds
3. A new sheet "UTL_Sanitizer_Preview" appears showing what WOULD change (no edits yet)
4. Scroll through — each row shows: Sheet, Cell, Current Value, Issue Type, Proposed Fix
5. Pause 2 seconds — the viewer sees the dry-run report
6. Now run the FULL sanitize: Alt+F8 > type `RunFullSanitize` > Run
7. A dialog may ask to confirm — click Yes
8. Wait 3-5 seconds (backs up sheets first, then fixes)
9. Success MsgBox shows count of fixes applied
10. Navigate back to the data sheet — the green triangles are gone, FP tails are fixed

**Expected output:** "UTL_Sanitizer_Preview" sheet (from preview), then the actual data is cleaned in place.

**Gotcha:** The sanitizer skips columns with header keywords like "Date", "Name", "ID", "Customer", etc. This is by design — it won't corrupt your dates or names. But if your test data has numeric columns with those headers, they'll be skipped. Use headers like "Amount", "Quantity", "Total" for columns you want fixed.

**Gotcha #2:** `RunFullSanitize` creates backup sheets (named "BKP_SheetName"). These are safety copies. They'll clutter the tab bar — that's expected. Mention it's a safety feature.

---

### CLIP 2: Highlights — Threshold + Duplicates (60 sec)

**Screen before:** On a data sheet with numeric values.

**Actions — Part A: Threshold Highlighting (30 sec):**
1. Select a range of numeric cells (e.g., a column of dollar amounts)
2. Alt+F8 > type `HighlightByThreshold` > Run
3. An InputBox asks for the threshold value — type **10000**
4. A second InputBox asks direction: "above", "below", or "equal" — type **above**
5. Wait 1 second
6. All cells above $10,000 are highlighted (yellow or custom color)
7. Pause 2 seconds on the result

**Actions — Part B: Duplicate Highlighting (30 sec):**
1. Select a column with duplicate values (e.g., customer names or product codes)
2. Alt+F8 > type `HighlightDuplicateValues` > Run
3. Wait 1 second
4. All duplicate values are highlighted in orange
5. Pause 2 seconds
6. Clear highlights: Alt+F8 > `ClearHighlights` > Run > choose "Active Sheet"

**Expected output:** Cells highlighted by threshold (Part A), then duplicates highlighted in orange (Part B), then cleared.

**Gotcha:** The 500K cell safety cap means if you select the entire sheet (Ctrl+A), it may refuse on very large files. Select a specific range instead.

---

### CLIP 3: Comments — Extract and Count (40 sec)

**Screen before:** On a sheet that has cell comments/notes.

**Setup:** Make sure the sample file has at least 5-10 comments across 2-3 sheets. If not, add some manually before recording (right-click cell > Insert Comment).

**Actions:**
1. Alt+F8 > type `CountComments` > Run
2. A MsgBox appears showing comment count per sheet:
   - "Sales: 4 comments"
   - "Expenses: 2 comments"
   - "Total: 6 comments"
3. Click OK
4. Alt+F8 > type `ExtractAllComments` > Run
5. Wait 1-2 seconds
6. A new "UTL_CommentReport" sheet appears with:
   - Sheet name, Cell address, Author, Comment text, Date (if available)
   - Styled header row
7. Scroll through — all comments extracted to one place
8. Pause 2 seconds

**Expected output:** Comment count summary, then "UTL_CommentReport" sheet with all comments extracted.

---

## CHAPTER 2 — Sheet & Column Tools (3 clips)

### CLIP 4: Tab Organizer (50 sec)

**Screen before:** Sample file with 4-6 tabs visible at the bottom.

**Actions — Part A: Color Tabs by Keyword (25 sec):**
1. Alt+F8 > type `ColorTabsByKeyword` > Run
2. InputBox asks for keyword — type **"Sales"**
3. A color picker dialog appears (7 options) — pick a number (e.g., 3 for green)
4. All tabs with "Sales" in the name turn green
5. Pause 2 seconds

**Actions — Part B: Reorder Tabs (25 sec):**
1. Alt+F8 > type `ReorderTabs` > Run
2. A numbered list of all sheets appears — pick which one(s) to move
3. Choose position: front, back, or after a specific sheet
4. Tabs reorder on screen
5. Pause 2 seconds — visible tab order changed

**Expected output:** Colored tabs and reordered tab bar.

**Gotcha:** Tab coloring only works on sheet tabs — some Excel versions show the color more subtly than others. Zoom in on the tab bar if needed.

---

### CLIP 5: Column Ops — Split and Merge (50 sec)

**Screen before:** On a sheet with a column that has combined data (e.g., "Smith, John" or "New York, NY 10001").

**Setup:** Make sure one column has data that can be split (comma-separated, or space-separated names).

**Actions — Part A: Split Column (25 sec):**
1. Select the column header of the column to split
2. Alt+F8 > type `SplitColumn` > Run
3. A dialog asks for the delimiter — choose "Comma" (or type custom)
4. Wait 1-2 seconds
5. The column splits into 2+ columns, data redistributed
6. Pause 2 seconds

**Actions — Part B: Combine Columns (25 sec):**
1. Select 2 columns to merge (e.g., First Name + Last Name)
2. Alt+F8 > type `CombineColumns` > Run
3. A dialog asks for separator — choose "Space" or "Comma + Space"
4. Wait 1 second
5. A new column appears with combined values
6. Pause 2 seconds

**Expected output:** Column split into multiple columns (Part A), then columns merged into one (Part B).

**Gotcha:** SplitColumn inserts new columns to the right. Make sure there's room (no data in adjacent columns) or it may overwrite. Test on the specific sheet beforehand.

---

### CLIP 6: Sheet Tools — Sheet Index + Template Cloner (50 sec)

**Screen before:** On any sheet in the sample file.

**Actions — Part A: Sheet Index (20 sec):**
1. Alt+F8 > type `ListAllSheetsWithLinks` > Run
2. Wait 1-2 seconds
3. A new "UTL_SheetIndex" sheet appears with:
   - Column A: Sheet names
   - Column B: Clickable hyperlinks to each sheet
   - Column C: Visibility status (Visible/Hidden/Very Hidden)
4. Click one hyperlink to show it works — jumps to that sheet
5. Navigate back to UTL_SheetIndex

**Actions — Part B: Template Cloner (30 sec):**
1. Alt+F8 > type `TemplateCloner` > Run
2. A numbered list of sheets appears — pick a sheet to clone (e.g., pick the Sales sheet)
3. An InputBox asks how many copies — type **3**
4. Wait 2-3 seconds
5. Three new tabs appear: "Sales (1)", "Sales (2)", "Sales (3)" — exact copies
6. Pause 2 seconds on the tab bar showing the clones

**Expected output:** Sheet index with hyperlinks (Part A), then 3 cloned sheets (Part B).

**Gotcha:** Cloned sheet names must be ≤31 characters (Excel limit). If the original name is long, clones may get truncated. Use a short-named sheet for the demo.

---

## CHAPTER 3 — Analysis & Building (4 clips)

### CLIP 7: Compare Sheets (50 sec)

**Screen before:** Sample file with at least 2 sheets that have similar structure but different values (e.g., "Sales" and "Sales (1)" — the clone you just made, after modifying a few cells).

**Setup:** After cloning in Clip 6, go into "Sales (1)" and change 3-5 cell values. This gives the compare tool something to find.

**Actions:**
1. Alt+F8 > type `CompareSheets` > Run
2. A numbered list of sheets appears — pick Sheet A (e.g., "Sales")
3. Pick Sheet B (e.g., "Sales (1)")
4. A dialog asks "Highlight differences on source sheets?" — click Yes
5. Wait 2-5 seconds
6. A new "UTL_CompareReport" sheet appears with:
   - Cell address | Sheet A Value | Sheet B Value | Match/Mismatch
   - Summary: X cells compared, Y differences found
7. Navigate to "Sales" sheet — the changed cells are highlighted in red
8. Pause 3 seconds on the diff report

**Expected output:** "UTL_CompareReport" with cell-by-cell comparison, plus red highlighting on source sheets.

---

### CLIP 8: Consolidate Sheets (40 sec)

**Screen before:** Sample file with multiple data sheets (Sales, Expenses, Inventory, or the clones).

**Actions:**
1. Alt+F8 > type `ConsolidateSheets` > Run
2. A numbered list of sheets appears — select 2-3 sheets to combine (type their numbers comma-separated)
3. A dialog asks "Skip headers on sheets 2+?" — click Yes
4. A dialog asks "Add Source Sheet column?" — click Yes
5. Wait 2-3 seconds
6. A new "UTL_Consolidated" sheet appears with:
   - All data from selected sheets stacked vertically
   - A "Source Sheet" column on the right showing which sheet each row came from
   - Headers from the first sheet applied at top
7. Scroll down to show data from different source sheets
8. Pause 2 seconds

**Expected output:** "UTL_Consolidated" sheet with merged data and source tracking.

**Gotcha:** Sheets must have the same column structure for clean consolidation. If columns don't match, the result may look misaligned. Use sheets with identical headers.

---

### CLIP 9: PivotTools + LookupBuilder + ValidationBuilder (60 sec)

**Screen before:** On the consolidated sheet (or any data sheet).

**Actions — Part A: Pivot Tools — List All Pivots (15 sec):**
1. Alt+F8 > type `ListAllPivots` > Run
2. If the file has pivot tables: a "UTL_PivotReport" sheet appears listing them all
3. If no pivots exist: a MsgBox says "No pivot tables found in this workbook" — that's fine, acknowledge and move on
4. (If pivots exist, also show `RefreshAllPivots` for the one-click refresh)

**Actions — Part B: Lookup Builder (25 sec):**
1. Alt+F8 > type `BuildVLOOKUP` > Run
2. Step 1: Select the lookup keys (e.g., a column of product IDs)
3. Step 2: Select the source table range (the lookup table)
4. Step 3: Enter the column number to return (e.g., 3 for price)
5. Step 4: Select where to put the results
6. Wait 1 second — VLOOKUP formulas are written, wrapped in IFERROR
7. Scroll to show the formulas populated with real values
8. Pause 2 seconds

**Actions — Part C: Validation Builder (20 sec):**
1. Select a range of cells (e.g., a "Status" column)
2. Alt+F8 > type `CreateDropdownList` > Run
3. Type the list values: **"Open, Closed, Pending, Cancelled"**
4. Click OK
5. Click on one of the cells — a dropdown arrow appears with the 4 options
6. Pause 2 seconds

**Expected output:** Pivot inventory (Part A), VLOOKUP formulas (Part B), dropdown validations (Part C).

**Gotcha (Lookup Builder):** The source table must have the lookup column as the FIRST column for VLOOKUP. If it's not first, use `BuildINDEXMATCH` instead. For the demo, set up the data so VLOOKUP works cleanly.

---

## CHAPTER 4 — Universal Command Center (1 clip)

### CLIP 10: Universal Command Center (50 sec)

**Screen before:** On any sheet in the sample file.

**Actions:**
1. Alt+F8 > type `LaunchCommandCenter` > Run
   - Or if you've set up a shortcut, use that
2. The Universal Command Center menu appears with all tool categories:
   - Data Sanitize (4 tools)
   - Highlights (5 tools)
   - Comments (4 tools)
   - Tab Organizer (6 tools)
   - Column Ops (4 tools)
   - Sheet Tools (4 tools)
   - Compare (3 tools)
   - Consolidate (2 tools)
   - Pivot Tools (4 tools)
   - Lookup Builder (4 tools)
   - Validation (6 tools)
3. Scroll through the full menu slowly — the viewer sees ALL categories
4. Demo the search: type **"duplicate"** — filters to relevant tools
5. Clear search, type **"pivot"** — shows pivot tools
6. Clear search
7. Pick any tool from the menu and run it (e.g., "Count Comments" for a quick result)
8. Close the Command Center

**Expected output:** Full categorized menu of 140+ tools with real-time keyword search.

**Gotcha:** The Universal Command Center (`modUTL_CommandCenter.LaunchCommandCenter`) is DIFFERENT from the demo file's Command Center (`modFormBuilder.LaunchCommandCenter`). Make sure you're running the right one. In the sample file, only the UTL version should be imported.

---

## PYTHON CLIPS (2 clips)

### CLIP 11: File Comparison Script (60 sec)

**Screen before:** Command Prompt open (not PowerShell — use `cmd` for simpler visuals). Two Excel files on your desktop ready to compare.

**Setup:**
- Have two Excel files ready: `Budget_v1.xlsx` and `Budget_v2.xlsx` (or similar names)
- They should be similar but with 5-10 differences (changed numbers, added rows)
- Have Python installed with `pandas` and `openpyxl` (`pip install pandas openpyxl`)
- Know the exact path to `compare_files.py`

**Actions:**
1. In Command Prompt, type the command (have this pre-typed in Notepad and paste it):
   ```
   python compare_files.py "C:\Users\Connor\Desktop\Budget_v1.xlsx" "C:\Users\Connor\Desktop\Budget_v2.xlsx"
   ```
2. Press Enter
3. Wait 3-5 seconds — terminal shows progress messages
4. Output: "COMPARISON_REPORT.xlsx saved to C:\Users\Connor\Desktop\"
5. Switch to Excel (Alt+Tab or click)
6. Open COMPARISON_REPORT.xlsx from your desktop
7. Show the SUMMARY sheet: Added/Removed/Changed counts
8. Click into a detail sheet showing cell-by-cell differences
   - Columns: Location, Column, File1 Value, File2 Value, Change Type
   - Color coding: Green = Added, Red = Removed, Yellow = Changed
9. Pause 3 seconds on the diff report

**Expected output:** Terminal shows success, then Excel report with color-coded differences.

**Gotcha:** If Python isn't in your PATH, use the full path: `"C:\Users\Connor\AppData\Local\Programs\Python\Python39\python.exe"`. Test the exact command beforehand.

**Gotcha #2:** If the two files have different sheet names, the script compares matching sheets only. Make sure both files have at least one sheet with the same name.

---

### CLIP 12: PDF Extractor Script (60 sec)

**Screen before:** Command Prompt open. A sample PDF with tables on your desktop.

**Setup:**
- Have a PDF with at least 1-2 visible data tables (financial statement, report, etc.)
- **The PDF must have selectable text** — scanned image PDFs will not work
- Install `pdfplumber`: `pip install pdfplumber pandas openpyxl`
- Know the exact path to `pdf_extractor.py`

**Actions:**
1. In Command Prompt, type:
   ```
   python pdf_extractor.py "C:\Users\Connor\Desktop\SampleReport.pdf"
   ```
2. Press Enter
3. Wait 3-10 seconds (depends on PDF size)
4. Terminal shows: "Found X tables across Y pages" and "Saved to PDF_EXTRACTED_TABLES.xlsx"
5. Switch to Excel
6. Open PDF_EXTRACTED_TABLES.xlsx
7. Show the extracted table(s):
   - Each table gets its own sheet: "Page1_Table1", "Page2_Table1", etc.
   - Headers styled with dark blue
   - Data aligned in proper columns
8. Pause 3 seconds — the viewer sees a PDF table now in Excel

**Expected output:** Terminal success, then Excel file with extracted tables.

**Gotcha:** If the PDF has no tables, the script will say "No tables found." Use a PDF you've verified has extractable tables. Financial statements, invoices, and structured reports work best.

**Gotcha #2:** If `pdfplumber` isn't installed, the script will fail with ImportError. Install it beforehand: `pip install pdfplumber`.

---

### CLIP 13: Closing + Universal Command Center Recap (30 sec)

**Screen before:** Back in Excel with the sample file.

**Actions:**
1. No macro to run — this is audio only
2. While the closing audio plays, slowly scroll through the tab bar showing all the output sheets created during the demo:
   - UTL_Sanitizer_Preview
   - UTL_CommentReport
   - UTL_SheetIndex
   - UTL_CompareReport
   - UTL_Consolidated
   - (any others created during clips)
3. End on the main data sheet, clean and organized
4. End recording

---
---

# POST-RECORDING CHECKLIST

After recording all 3 videos:

- [ ] Review each recording in Camtasia — check for:
  - [ ] No error dialogs visible that shouldn't be there
  - [ ] No personal information visible (file paths with your username are OK if expected)
  - [ ] No notification pop-ups that crept in
  - [ ] Mouse movements are smooth and deliberate (no frantic clicking)
  - [ ] Each pause is long enough for the audio to match (2-3 seconds minimum on key moments)
- [ ] Trim dead time at start/end of each clip
- [ ] Add chapter markers in Camtasia for Video 2 (7 chapters) and Video 3 (4 chapters)
- [ ] Sync audio tracks to video clips
- [ ] Export final versions
- [ ] Upload to SharePoint

---

# QUICK REFERENCE — ACTION NUMBERS

| # | Action | Module | Used In |
|---|--------|--------|---------|
| 3 | Reconciliation Checks | modReconciliation | V2 Clip 6 |
| 6 | Variance Analysis | modVarianceAnalysis | V2 Clip 7 |
| 7 | Data Quality Scan | modDataQuality | V1 Clip 4, V2 Clip 5 |
| 10 | PDF Export | modPDFExport | V2 Clip 12 |
| 12 | Build Dashboard | modDashboard | V1 Clip 6, V2 Clip 10 |
| 17 | Import Data Pipeline | modImport | V2 Clip 4 |
| 32 | Save Version | modVersionControl | V2 Clip 15 |
| 35 | List Versions | modVersionControl | V2 Clip 15 (optional) |
| 41 | View Audit Log | modLogger | V2 Clip 18 |
| 44 | Run Full Integration Test | modIntegrationTest | V2 Clip 17 |
| 46 | Generate Commentary | modVarianceAnalysis | V1 Clip 5, V2 Clip 8 |
| 48 | Toggle Executive Mode | modNavigation | V2 Clip 14 |
| 63 | Run What-If Demo | modWhatIf | V2 Clip 16 |
| 65 | Restore Baseline | modWhatIf | V2 Clip 16 |
| — | GenerateExecBrief | modExecBrief | V2 Clip 13 |
| — | ShowTimeSavedReport | modTimeSaved | V2 Clip 19 |

---

# EMERGENCY RECOVERY — If Something Goes Wrong During Recording

**Macro errors out with a VBA error dialog:**
- Click "End" (not Debug)
- The action failed but Excel is fine
- Re-run the action. If it fails again, skip it and note it for re-recording later

**Excel freezes or hangs:**
- Wait 30 seconds — some macros are just slow
- If still frozen after 60 seconds, press Ctrl+Break to interrupt VBA
- If totally unresponsive, end task via Task Manager and restart from last good state

**Wrong sheet is showing:**
- Just click the correct tab. You can edit this out in Camtasia

**Command Center doesn't open (Ctrl+Shift+M not working):**
- The keyboard shortcut may not be set up. Use Alt+F8 > type `LaunchCommandCenter` > Run instead
- To set up the shortcut: Alt+F8 > type `SetupKeyboardShortcuts` > Run (if it exists in modNavigation)

**What-If won't restore baseline:**
- Run Action 65 (Restore Baseline) manually
- If that fails: right-click the "WhatIf_Baseline" tab > Unhide, then manually copy the values back to Assumptions
- Nuclear option: close without saving and reopen from the last saved version

**PDF export fails:**
- Check that a PDF printer is configured (Microsoft Print to PDF or similar)
- Try exporting a single sheet instead: Action 11 (Export Single Sheet)
