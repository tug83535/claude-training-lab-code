# Demo Script Plan — KBT P&L Automation Toolkit
**Draft as of 2026-03-01 | For internal iPipeline presentation**

---

## Context

This is a draft demo script for recording a walkthrough of the Keystone BenefitTech P&L Automation Toolkit — a full Excel + VBA + Python + SQL system built to replace a 15+ hour manual monthly close process with a sub-2-hour, one-click automated workflow. The audience is 2,000+ iPipeline employees plus the CFO and CEO. The demo needs to be polished, professional, and accessible to non-technical Finance & Accounting staff.

---

## Recommended Video Breakdown

Rather than one 45-minute recording, break it into **5 shorter videos**. Shorter videos are easier to re-record if something goes wrong, easier for coworkers to reference later (they can jump to just the part they need), and feel more professional.

| # | Video Title | Audience | Length |
|---|-------------|----------|--------|
| 1 | **The Big Picture** — What This Is & Why It Matters | Everyone | ~6 min |
| 2 | **Live Demo Part 1** — The Command Center & Monthly Workflow | Everyone | ~14 min |
| 3 | **Live Demo Part 2** — Dashboards, Analysis & PDF Reports | Everyone | ~10 min |
| 4 | **Python & Advanced Analytics** (Optional Layer) | Technical / Power Users | ~8 min |
| 5 | **How to Set It Up** — Installation Guide for Coworkers | IT / Setup Users | ~7 min |

**Total runtime: ~45 minutes**

---

## Pre-Recording Checklist (Do Before You Hit Record)

- [ ] Close all other applications — clean desktop
- [ ] Set screen resolution to 1920x1080 (full HD)
- [ ] Open Excel with the P&L demo file already loaded
- [ ] Enable all macros (trust center already set up)
- [ ] Make sure all VBA modules are imported and Command Center works
- [ ] Have a browser tab open with the Streamlit dashboard running (for Video 4)
- [ ] Clear any leftover generated sheets from previous test runs (Variance, DQ Report, etc.)
- [ ] Turn off notifications (Slack, Outlook, Teams, Windows popups)
- [ ] Use a clean user profile / hide personal bookmarks if screen sharing
- [ ] Do a dry run of each video once before recording for real

---

---

# VIDEO 1: The Big Picture (~6 minutes)

**[SCENE: Slide or clean Word/PowerPoint document with project title. Or just speak over the Excel file with it minimized.]**

---

### INTRO (0:00–0:45)

> "Hi everyone. My name is [Your Name], and I work in Finance & Accounting here at iPipeline.
>
> Over the past several months, I've been building something that I'm really excited to share with you — a fully automated P&L reporting and analysis toolkit built right inside Excel.
>
> This is not a third-party tool. It's not software you need to buy. It runs entirely inside our existing Excel environment, and I built every piece of it from the ground up."

---

### THE PROBLEM (0:45–2:00)

**[SCENE: Show a blank Excel file or just talk — describe the before state in plain language]**

> "Before I show you what it does, let me tell you what problem it solves.
>
> Every month, our team would spend 15 or more hours manually building and updating P&L reports. That meant:
>
> - Copying and pasting data between sheets by hand
> - Manually checking formulas for errors
> - Building charts from scratch every single month
> - Sending back and forth emails to find out what the numbers meant
> - And then doing it all over again the next month
>
> It was slow, it was error-prone, and it took time away from actual analysis — the things that matter to the business.
>
> I knew there had to be a better way."

---

### THE SOLUTION (2:00–4:00)

**[SCENE: Switch to the Excel file. Show the Report--> table of contents sheet.]**

> "So I built this.
>
> This is the Keystone BenefitTech P&L Automation Toolkit — version 2.1.
>
> What you're looking at is our P&L model. It has 13 core sheets, covers 4 product lines — iGO, Affirm, InsureSight, and DocFast — and tracks every dollar across 7 departments.
>
> But what makes this different from any regular Excel file is this:"

**[Press Ctrl+Shift+M to open Command Center]**

> "One keystroke brings up the Command Center — 50 automated actions that anyone on the team can run in one click.
>
> Data quality scanning. Reconciliation checks. Variance analysis. Dashboard charts. PDF exports. Automated commentary. All of it — one click.
>
> What used to take 15 hours now takes under 2."

---

### WHAT YOU'LL SEE (4:00–6:00)

> "Over the next few videos, I'm going to walk you through exactly what this does.
>
> In the next video, I'll show you the full monthly workflow — how a real month-end close runs from start to finish using this system.
>
> After that, I'll show you the dashboards and reports we can generate automatically.
>
> Then I'll show you the Python analytics layer — which goes even further with forecasting and scenario modeling.
>
> And finally, if you're someone who needs to set this up, I have a step-by-step installation guide video as well.
>
> Let's get into it."

---

---

# VIDEO 2: Live Demo Part 1 — Command Center & Monthly Workflow (~14 minutes)

**[SCENE: Excel file open, on the Report--> sheet]**

---

### INTRO (0:00–0:30)

> "In this video I'm going to walk through what a real month-end close looks like using this toolkit — from raw data all the way to a finished reconciliation.
>
> I'll show you 6 of the most important actions, and you'll see how they chain together into a clean, repeatable process."

---

### STEP 1: OPEN THE COMMAND CENTER (0:30–2:00)

**[Press Ctrl+Shift+M]**

> "The Command Center is the heart of this system.
>
> I press Ctrl+Shift+M from anywhere in the workbook and this menu appears.
>
> You can see 50 actions here, organized across four pages. Each one does something specific. You don't need to know VBA, you don't need to touch a formula — you just click.
>
> Today I'll walk through the core monthly workflow."

---

### STEP 2: SCAN DATA QUALITY (2:00–4:30)

**[Click Action 7 — Scan Data Quality]**

> "The first thing I always do at the start of a new month is run the data quality scan.
>
> This checks all 510 rows of our GL data for 8 types of problems:
> - Duplicate transactions
> - Text-stored numbers — where a cell looks like a number but Excel treats it as text
> - Blank cells that should have data
> - Misspelled category names
> - Mixed date formats
> - And more.
>
> Watch what happens when I click it."

**[Let the scan run. Show DQ Report sheet that gets generated.]**

> "In just a few seconds, it scanned every row and generated this Data Quality Report sheet.
>
> Each issue is flagged with a severity level — High, Medium, Low — and tells you exactly what cell, what sheet, and what the problem is.
>
> Before I had this, finding data issues was a manual, manual process. Now it's one click."

---

### STEP 3: FIX TEXT NUMBERS (4:30–6:00)

**[Click Action 8 — Fix Text Numbers]**

> "Now I'll run the Fix Text Numbers command.
>
> This only touches the cells that were already flagged by the scan — it won't accidentally convert GL IDs or other text fields. That was an important safeguard we built in.
>
> Watch — it runs instantly, and those cells are now real numbers that Excel can calculate with."

---

### STEP 4: GENERATE MONTHLY TABS (6:00–8:30)

**[Click Action 1 — Generate Monthly Tabs, OR Action 42 for single-month]**

> "Next — generating the monthly tabs.
>
> Our workbook has January, February, and March already set up as Functional P&L Summary sheets.
>
> When I click Generate Monthly Tabs, it clones the March template and automatically creates April through December — updating all the formula references, applying the right formatting, and color-coding each tab.
>
> This used to take 30-45 minutes to do by hand for every month. Watch."

**[Let it run. Show the new tabs appearing.]**

> "Done. Every month is now ready for data entry, with formulas already in place pointing to the right source data."

---

### STEP 5: RUN RECONCILIATION CHECKS (8:30–11:00)

**[Click Action 3 — Run Reconciliation Checks]**

> "Now I run reconciliation.
>
> This is where the system acts like an auditor. It evaluates 9 pre-built checks on our Checks sheet — things like:
> - Do the GL totals match the P&L trend?
> - Do all product line allocations add up to 100%?
> - Do the functional summaries match the natural P&L?
>
> Every check returns either PASS or FAIL — highlighted in green or red — and a reconciliation report gets exported automatically."

**[Show the Checks sheet with results highlighted]**

> "This is the kind of thing that used to require a senior analyst to manually trace through three different sheets. Now it takes 3 seconds.
>
> And because every run is logged in the audit trail, we have a full history of when checks were run and what the results were."

---

### STEP 6: VARIANCE ANALYSIS (11:00–13:00)

**[Click Action 6 — Variance Analysis]**

> "The last step I'll show in this video is variance analysis.
>
> I select January and February as my comparison months, click Run, and the system compares every line item between the two months — flagging anything that changed by more than 15%.
>
> But here's the part I'm really proud of:"

**[Click Action 46 — Variance Commentary]**

> "After the variance analysis runs, I can click Variance Commentary — and it automatically writes plain-English sentences describing the top 5 variances.
>
> Things like: 'AWS Compute for InsureSight increased $48,000 MoM, representing a 22% increase driven by compute expansion.'
>
> That's ready to paste into an executive briefing. No one had to write it."

---

### WRAP UP (13:00–14:00)

> "So that's the core monthly workflow — data quality scan, fix numbers, generate tabs, reconcile, analyze variances, write commentary.
>
> Six steps. Maybe 10 minutes of actual work.
>
> In the next video, I'll show you the dashboards and PDF reports."

---

---

# VIDEO 3: Live Demo Part 2 — Dashboards, Analysis & PDF Reports (~10 minutes)

**[SCENE: Excel file open, Command Center ready]**

---

### INTRO (0:00–0:30)

> "In this video, I'll show you three things: the automated dashboard charts, the sensitivity analysis tool, and the professional PDF export.
>
> These are the outputs that go to leadership — the things that make the numbers visible and actionable."

---

### DASHBOARDS (0:30–4:00)

**[Click Action 12 — Build Dashboard]**

> "I'll start with the dashboard.
>
> I click Build Dashboard, and the system creates three charts automatically — it detects how many months of data we have and sizes everything dynamically. I don't touch a single chart manually.
>
> The first chart is a Revenue Trend line — 12 months of revenue across all four product lines."

**[Scroll to show each chart as you describe it]**

> "The second is Contribution Margin Trend — this is the story of how our margins are moving month to month.
>
> The third is Product Mix — a pie chart showing what percentage of revenue each product line contributed this month.
>
> Now watch this:"

**[Click Action 43 — Executive Dashboard (KPI cards)]**

> "The Executive Dashboard takes it a step further. It creates a summary view with KPI cards showing Total Revenue, Net Income, Average Margin, and the Top Product — plus a summary table underneath.
>
> This is what I'd put on slide one of any CFO presentation."

**[Click Action 44 — Waterfall Chart]**

> "And finally — the Waterfall Chart.
>
> This is the revenue-to-net-income bridge. It shows how we started with revenue and ended at net income, step by step — what the expenses were, where margin was gained or lost.
>
> Waterfall charts are notoriously painful to build manually in Excel. This one builds itself in one click."

---

### SEARCH (4:00–5:30)

**[Click Action 15 — Cross-Sheet Search]**

> "Here's a feature that saves a ton of time during reviews.
>
> I type a keyword — let's say 'DocFast' — and it searches every visible sheet in the entire workbook and returns every cell that matches, with the sheet name, cell address, and surrounding context.
>
> No more Ctrl+F and clicking through sheet by sheet. One search, all results."

---

### PDF EXPORT (5:30–8:00)

**[Click Action 10 — Export Report Package]**

> "Now — the PDF export.
>
> When I click Export Report Package, it takes all the report sheets — the P&L trend, functional summaries, checks, dashboard, variance analysis — and exports them as a single professional PDF.
>
> Each page has a header with the company name and report title, a footer with the page number and date, and everything is formatted for landscape printing.
>
> The file is automatically named with a timestamp so you always know exactly when it was generated.
>
> This is the package you'd attach to an email to leadership at month end."

**[Show the PDF file or describe it opening]**

> "Before this, someone would manually set print areas, format each sheet, export them one at a time, and merge them. That's 30-45 minutes of clicking. Now it's 15 seconds."

---

### AUDIT TRAIL (8:00–9:30)

**[Navigate to VBA_AuditLog sheet]**

> "Last thing I want to show you in this video — the audit trail.
>
> Every single command that anyone runs gets logged here automatically. You can see the date, the time, the module name, what action was taken, and whether it succeeded.
>
> This is our internal controls story. If anyone ever asks 'when was this reconciliation run?' or 'who generated this report?' — the answer is right here."

---

### WRAP UP (9:30–10:00)

> "That's the reporting layer. Dashboards, waterfall charts, cross-sheet search, professional PDF, full audit trail.
>
> In the next video, I'll show you the Python analytics layer — which goes even further."

---

---

# VIDEO 4: Python & Advanced Analytics (~8 minutes)

**[SCENE: Split view or switch between Excel and terminal. Have Streamlit dashboard already running in browser.]**

**[NOTE: This video is optional — most relevant for power users, FP&A leads, and anyone technical. Regular staff don't need to run Python to use the toolkit — this is an add-on layer.]**

---

### INTRO (0:00–0:45)

> "Everything I've shown you so far lives inside Excel — no special software needed.
>
> But for teams that want to go even further, I built a Python analytics layer on top of the Excel file.
>
> Python is a programming language that can do things Excel can't — statistical forecasting, interactive web dashboards, fuzzy matching, automated email reports.
>
> I want to show you what's possible, even if you never use this yourself."

---

### STREAMLIT DASHBOARD (0:45–2:30)

**[Switch to browser showing the Streamlit dashboard]**

> "This is the Streamlit dashboard.
>
> It reads our P&L data directly from the Excel file and builds an interactive web app — right in your browser.
>
> You can filter by product line, by month, by department. The charts update in real time. There's no clicking through Excel sheets — it's a clean, modern interface.
>
> To launch it, you run one line: python pnl_runner.py dashboard. That's it."

---

### MONTH-END CLOSE AUTOMATION (2:30–4:00)

**[Switch to terminal. Run: python pnl_runner.py month-end --month 1]**

> "The month-end close command runs 6 categories of checks against our data — GL completeness, allocation balance, reconciliation, variance flags, items needing review, and a close status summary.
>
> Each check returns PASS, FAIL, WARN, or SKIP.
>
> This runs in about 5 seconds and gives you a clean status report you could share with a manager."

---

### FORECASTING (4:00–5:30)

**[Run: python pnl_runner.py forecast --months 3]**

> "The forecasting module takes our historical data and projects the next several months.
>
> It uses three methods — a simple moving average, exponential smoothing, and a trend line — and you can compare all three to see which fits your data best.
>
> The output is a table you can paste into a board report or planning document.
>
> This would normally require a financial analyst with modeling experience. Now it's one command."

---

### ALLOCATION SIMULATOR (5:30–7:00)

**[Run: python pnl_runner.py allocate or describe it]**

> "Finally — the What-If Allocation Simulator.
>
> This lets you ask questions like: 'What happens to our product margins if we shift 5% of AWS costs from InsureSight to DocFast?'
>
> You set the override, run the simulation, and it shows you the before and after — side by side — for every product line.
>
> There are 3 preset scenarios built in, or you can set your own custom allocations."

---

### WRAP UP (7:00–8:00)

> "That's the Python layer — interactive dashboard, automated close checks, forecasting, scenario modeling.
>
> None of this requires any coding knowledge to use. It's all set up and ready to run.
>
> In the final video, I'll show you exactly how to install and set this up from scratch."

---

---

# VIDEO 5: How to Set It Up — Installation Guide (~7 minutes)

**[SCENE: Fresh Excel file, or walk through as if setting up for the first time.]**

---

### INTRO (0:00–0:30)

> "In this last video, I'm going to walk you through the setup — how to get this working on your own computer.
>
> I'll go step by step. You don't need any technical experience. If you can open Excel, you can do this."

---

### STEP 1: ENABLE MACROS / TRUST CENTER (0:30–2:00)

> "Before we do anything, we need to tell Excel to allow macros to run.
>
> Open Excel. Go to File > Options > Trust Center > Trust Center Settings > Macro Settings.
>
> Select 'Enable all macros' — or, if your IT department requires it, 'Disable all macros with notification.'
>
> Also go to the Trusted Locations tab and add the folder where you saved the Excel file.
>
> Click OK on everything to save.
>
> This only needs to be done once."

---

### STEP 2: OPEN THE FILE AND ENABLE EDITING (2:00–2:45)

> "Open the file KeystoneBenefitTech_PL_Model.xlsm — make sure it's the .xlsm version, which is the macro-enabled version.
>
> If you see a yellow bar at the top that says 'Enable Editing' — click it.
>
> If you see another bar about macros — click 'Enable Content.'
>
> You should now see the workbook normally."

---

### STEP 3: IMPORT VBA MODULES (2:45–4:30)

> "Now we need to import the VBA code.
>
> Press Alt+F11 to open the VBA editor.
>
> In the left panel, right-click on the workbook name and choose 'Import File.'
>
> Navigate to the folder called 'vba' in the project folder.
>
> Select all 13 module files — they end in .bas — and import them one by one.
>
> Once all modules are imported, close the VBA editor with Alt+F11 again."

---

### STEP 4: TEST THE COMMAND CENTER (4:30–5:30)

> "Now let's make sure everything works.
>
> Press Ctrl+Shift+M. If the Command Center opens — you're done. Everything is working.
>
> If you get an error, the most common fix is to go back to the Trust Center and make sure macros are fully enabled.
>
> Try clicking Action 3 — Run Reconciliation Checks — just to confirm a command runs successfully."

---

### STEP 5: PYTHON SETUP (5:30–6:30) [Optional]

> "If you also want to run the Python analytics — this part is optional — you need Python installed.
>
> Open Command Prompt. Type: pip install -r requirements.txt and press Enter.
>
> That installs everything Python needs automatically.
>
> Then you can run any command with: python pnl_runner.py followed by the command name.
>
> The full list of commands is in the USER_TRAINING_GUIDE.md file in the docs folder."

---

### WRAP UP (6:30–7:00)

> "That's it. You're set up.
>
> If you run into any issues, check the QUICK_START guide — it has a troubleshooting section at the end.
>
> And of course, feel free to reach out to me directly.
>
> Thanks for watching."

---

---

## What's Left to Do BEFORE Recording

### On the Project / Code
- [ ] Verify all 13 VBA modules import cleanly with no errors on a fresh machine
- [ ] Do a full end-to-end dry run of every command shown in Videos 2 and 3
- [ ] Confirm Streamlit dashboard runs without errors (`python pnl_runner.py dashboard`)
- [ ] Confirm `python pnl_runner.py month-end --month 1` runs without errors
- [ ] Confirm `python pnl_runner.py forecast --months 3` runs without errors
- [ ] Clear any leftover test sheets before recording (Variance Analysis, DQ Report, etc.)
- [ ] Resolve ISSUE-013, ISSUE-014, ISSUE-015 if they affect anything shown in the demo

### On the Repo
- [ ] Upload existing files from local APCLDmerge_ALL folder into GitHub
- [ ] Create .gitignore (exclude .xlsm, __pycache__, .pyc, .db)
- [ ] Populate /training/ folder with coworker training materials
- [ ] Rewrite README.md with professional overview
- [ ] Fill CompletePackageStorage/production/ with the final demo file

### On the Recording Setup
- [ ] Decide which computer you will record on
- [ ] Choose recording software (OBS Studio is free; Loom and Camtasia are also good options)
- [ ] Set screen resolution to 1920x1080 before recording
- [ ] Do at least one full dry run of each video before recording for real
- [ ] Fill in your name and title in the Video 1 script intro
- [ ] Turn off all notifications before recording

---

## What's Left to Do on the Project Overall

### Immediate / Next Session

| Priority | Task |
|----------|------|
| High | Upload local APCLDmerge_ALL files to repo |
| High | Populate /training/ folder with coworker guides |
| High | Complete final README.md rewrite |
| Medium | Create .gitignore |
| Medium | Fill CompletePackageStorage/production/ |
| Medium | Resolve ISSUE-013, 014, 015 (deferred from v2.1 QA) |

### Before the Presentation

| Priority | Task |
|----------|------|
| High | Full end-to-end dry run on a clean machine |
| High | Verify all modules import cleanly on a fresh Excel install |
| High | Record and edit all 5 demo videos |
| Medium | Create a one-page printed Quick Reference card for coworkers |
| Medium | Create the PowerPoint slide deck (for live CFO/CEO presentation if needed) |

### Nice-to-Have (Future Phase)

| Task | Why It Matters |
|------|----------------|
| Outlook email integration | Auto-send PDF report to leadership at close |
| One-Click P&L Generator from Raw GL | Full end-to-end from trial balance to finished P&L |
| Export All Charts to PowerPoint | Huge time saver for board presentations |
| Data Entry UserForm | Guided data entry, reduces input errors |
| Timestamp Audit Trail on cell changes | Strong compliance story for CFO/CEO |

---

## Estimated Overall Project Completion

| Area | Status |
|------|--------|
| VBA Code (50 actions) | 100% — Production ready |
| Python Scripts (11 scripts) | 100% — Production ready |
| SQL Scripts | 100% — Production ready |
| Documentation (10 docs) | 100% — Production ready |
| QA / Testing | ~95% — 3 issues deferred |
| Repo Cleanup | ~40% — Session tasks pending |
| Training Materials | ~10% — Folder is empty |
| Demo Videos | 0% — Not yet recorded |
| Presentation Deck | 0% — Not yet created |
| Final Package (production/) | 0% — Files not yet uploaded |

---

## Recommended Recording Order

Record in this order — not necessarily video number order. This minimizes re-work if something goes wrong.

1. **Video 5 first** (Setup Guide) — easiest, mostly menu clicks, low pressure, no live macros
2. **Video 1 second** (Big Picture) — no clicking, just talking, easy to re-record in one take
3. **Video 2 third** (Monthly Workflow) — most important video, do it when you're freshest
4. **Video 3 fourth** (Dashboards & PDF) — builds on Video 2 being set up already
5. **Video 4 last** (Python) — if Streamlit causes any issues, this is the easiest video to skip or cut short
