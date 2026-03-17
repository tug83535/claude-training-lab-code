# Connor's SharePoint Folder Map (Updated)

**Personal Reference -- What Goes Where**
**Last Updated:** March 17, 2026
**Location:** SharePoint > Documents > Finance Automation

---

## Full Folder Tree

```
Finance Automation/
|
|--- START HERE/
|    |--- 00-Start-Here-Welcome.pdf
|    |--- 02-Getting-Started-First-Time-Setup.pdf
|    |--- 03-What-This-File-Does-Overview.pdf
|    |--- Source-Code-vs-Universal-Toolkit.pdf
|
|--- Demo File/
|    |--- ExcelDemoFile_adv.xlsm
|
|--- Quick Reference/
|    |--- 01-How-to-Use-the-Command-Center.pdf
|    |--- 04-Quick-Reference-Card.pdf
|    |--- AP-Copilot-Prompt-Guide.pdf
|
|--- Training Guides/
|    |--- 05-User-Training-Guide.pdf
|    |--- 07-Operations-Runbook.pdf
|    |--- 08-WhatIf-Scenario-Guide.pdf
|    |--- 09-Universal-CommandCenter-Guide.pdf
|    |--- 10-VBA-Module-Reference-List.pdf
|    |--- Dynamic-Chart-Filter-Setup-Guide.pdf
|
|--- Demo Videos/
|    |--- Video 1 - Whats Possible.mp4
|    |--- Video 2 - Full Walkthrough.mp4
|    |--- Video 3 - Universal Tools.mp4
|
|--- Source Code/
|    |--- VBA Modules/       (39 .bas files)
|    |--- Python Scripts/    (13 .py + requirements.txt)
|    |--- SQL Scripts/       (4 .sql files)
|
|--- Universal Toolkit/
     |--- 06-Universal-Toolkit-Guide.pdf
     |--- VBA Modules/       (23 .bas files)
     |--- Python Scripts/    (22 .py + requirements.txt)
```

---

## Folder-by-Folder Breakdown

### START HERE/
**Who it's for:** Every coworker. Read these first.

| File | Purpose |
|------|---------|
| 00-Start-Here-Welcome.pdf | Welcome page -- what this is, why it exists, where to begin |
| 02-Getting-Started-First-Time-Setup.pdf | How to enable macros, open the file, first steps |
| 03-What-This-File-Does-Overview.pdf | What the 65 actions do, how the file is organized |
| Source-Code-vs-Universal-Toolkit.pdf | Explains the difference between Source Code and Universal Toolkit folders |

---

### Demo File/
**Who it's for:** Everyone. This is the main deliverable.

| File | Purpose |
|------|---------|
| ExcelDemoFile_adv.xlsm | The P&L file with all 39 VBA modules and Command Center built in. Download this, enable macros, press Ctrl+Shift+M. |

**Maintenance rule:** When you update this file, KEEP ONLY THE LATEST VERSION HERE. Save old versions on your own machine, not on SharePoint. Coworkers should never wonder "which one do I download?"

---

### Quick Reference/
**Who it's for:** People who already set up the file and need quick answers.

| File | Purpose |
|------|---------|
| 01-How-to-Use-the-Command-Center.pdf | How to open, search, and run actions |
| 04-Quick-Reference-Card.pdf | One-page cheat sheet of all 65 actions |
| AP-Copilot-Prompt-Guide.pdf | How to use AI (CoPilot/Claude) to adapt the code for your own files |

---

### Training Guides/
**Who it's for:** People who want to go deeper.

| File | Purpose |
|------|---------|
| 05-User-Training-Guide.pdf | Full walkthrough of every feature |
| 07-Operations-Runbook.pdf | Day-to-day operations and maintenance |
| 08-WhatIf-Scenario-Guide.pdf | How to use the What-If scenario tool |
| 09-Universal-CommandCenter-Guide.pdf | How to use the Universal Command Center |
| 10-VBA-Module-Reference-List.pdf | All 39 modules listed with what each one does |
| Dynamic-Chart-Filter-Setup-Guide.pdf | How to add dropdown chart filters to your own files |

---

### Demo Videos/
**Who it's for:** Everyone. Watch before or after downloading the file.

| File | Purpose |
|------|---------|
| Video 1 - Whats Possible.mp4 | 5-min highlight reel -- what the toolkit can do |
| Video 2 - Full Walkthrough.mp4 | 18-min full demo of every major feature |
| Video 3 - Universal Tools.mp4 | 10-min demo of tools that work on ANY Excel file |

---

### Source Code/
**Who it's for:** Coworkers who want to READ the code, learn from it, or adapt pieces for their own projects using the CoPilot Prompt Guide.

| Subfolder | What's In It | Count |
|-----------|-------------|-------|
| VBA Modules/ | All 39 .bas files that power the demo Excel file | 39 files |
| Python Scripts/ | The 13 Python scripts + requirements.txt | 14 files |
| SQL Scripts/ | The 4 SQL scripts | 4 files |

**Key point:** This code is ALREADY inside the .xlsm file (for VBA). You do NOT need to download these to use the demo file. These are here so you can read the code, copy snippets, or use the CoPilot Prompt Guide to adapt them for your own files.

---

### Universal Toolkit/
**Who it's for:** Coworkers who want to add automation tools to their OWN Excel files (not the demo file).

| Subfolder | What's In It | Count |
|-----------|-------------|-------|
| 06-Universal-Toolkit-Guide.pdf | Step-by-step guide for importing and using these tools | 1 file |
| VBA Modules/ | 23 universal .bas modules (140+ tools) | 23 files |
| Python Scripts/ | 22 universal Python scripts + requirements.txt | 23 files |

**Key point:** These are DIFFERENT from Source Code. Source Code only works with the demo P&L file. Universal Toolkit works on ANY Excel file. See the "Source Code vs Universal Toolkit" doc in START HERE for a full explanation.

---

## Quick Decision Chart

| If you have... | Put it in... |
|----------------|-------------|
| Updated .xlsm demo file | Demo File/ (replace the old one) |
| New or updated training guide PDF | The right folder (START HERE, Quick Reference, or Training Guides) |
| Recorded demo video | Demo Videos/ |
| Updated .bas, .py, or .sql from the demo | Source Code/ > right subfolder |
| Updated universal tool (.bas or .py) | Universal Toolkit/ > right subfolder |

---

## Checklist -- When You Publish a New Version

Use this every time you update anything on SharePoint:

- [ ] Replace the .xlsm in Demo File/ (save old version on your machine first)
- [ ] Update any training guide PDFs that changed
- [ ] Update .bas/.py/.sql files in Source Code/ if demo code changed
- [ ] Update Universal Toolkit/ files if universal tools changed
- [ ] Spot-check: open the SharePoint page in a browser and make sure nothing looks broken

---

*Excel Automation Toolkit -- iPipeline Finance & Accounting -- 2026*
