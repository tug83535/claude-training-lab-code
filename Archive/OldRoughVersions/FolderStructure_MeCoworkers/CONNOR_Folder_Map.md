# Connor's SharePoint Folder Map — What Goes Where

**Last Updated:** March 10, 2026
**Location:** SharePoint > iPipeline Finance & Accounting > Documents > Automation Project

This is your personal reference for exactly what goes in every folder. When you update a file, add a new guide, or publish a new version — check this doc so nothing ends up in the wrong spot.

---

## Full Folder Tree

```
Automation Project/
│
├── 1. Demo File/
│   ├── Backups/
│   ├── Excel File - Clean/
│   └── Excel File - Fully Built/
│
├── 2. Training Guides/
│   ├── 1. Getting Started/
│   ├── 2. Quick Reference/
│   ├── 3. Python, VBA, & SQL/
│   └── 4. CoPilot Prompt Guide (code self-help)/
│
├── 3. Demo Video Walkthrough/
│
├── 4. Source Code/
│   ├── Python Modules/
│   ├── SQL Modules/
│   └── VBA Modules/
│
└── 5. Universal Code Library/
    ├── Python Kit/
    └── VBA Kit/
```

---

## Folder-by-Folder Breakdown

### 1. Demo File

This is the main event — the P&L Excel file that you demo to leadership and share with coworkers.

| Subfolder | What Goes Here | File Types | Notes |
|-----------|---------------|------------|-------|
| **Backups/** | Previous versions of both Excel files | .xlsm, .xlsx | Before you update either Excel file, drop the old version here with a date in the filename (e.g., `ExcelDemoFile_adv_2026-03-10.xlsm`). This is your safety net. |
| **Excel File - Clean/** | The P&L file with NO macros/VBA code | .xlsx only | This is a regular `.xlsx` — no macros, no Command Center. Just the raw P&L data, sheets, and formatting. For coworkers who want to explore the data without enabling macros. |
| **Excel File - Fully Built/** | The P&L file WITH all 32 VBA modules and Command Center | .xlsm only | This is the full toolkit. All 62 Command Center actions work from this file. This is what you demo to leadership and what power users will use daily. |

**Maintenance rules:**
- When you update the .xlsm (re-import VBA modules, fix bugs, add features), always drop the old version in Backups first
- When you update the .xlsm, also update the .xlsx (save a copy as .xlsx to strip the macros)
- The .xlsm in "Fully Built" should ALWAYS be the latest working version

---

### 2. Training Guides

All the guides that teach coworkers how to use the toolkit. Organized by topic so people can find what they need fast.

| Subfolder | What Goes Here | Specific Files |
|-----------|---------------|----------------|
| **1. Getting Started/** | Everything a brand-new user reads first | Guide 00 — Start Here Welcome.pdf |
| | | Guide 02 — Getting Started First Time Setup.pdf |
| | | Guide 03 — What This File Does Overview.pdf |
| **2. Quick Reference/** | Guides they come back to repeatedly for quick answers | Guide 01 — How to Use the Command Center.pdf |
| | | Guide 04 — Quick Reference Card.pdf |
| | | Dynamic Chart Filter Setup Guide.pdf |
| **3. Python, VBA, & SQL/** | Technical/code-related guides for people who want to go deeper | Guide 06 — Universal Toolkit Guide.pdf |
| **4. CoPilot Prompt Guide (code self-help)/** | The AI self-help prompt library | AP Copilot Prompt Guide.pdf |

**Why this split:**
- **Getting Started** = "I just got this file, what do I do?" — new users start here and read in order (00 → 02 → 03)
- **Quick Reference** = "I've been using this for a week, how do I do X?" — the guides people bookmark and revisit
- **Python, VBA, & SQL** = "I want to use the standalone code tools on my own files" — the Universal Toolkit guide lives here because it's about importing VBA/Python modules separately
- **CoPilot Prompt Guide** = "I want to use AI to help me with Excel/VBA/Python" — self-contained, its own category

**Where does Guide 05 go?**
- Guide 05 (Video Demo Script & Storyboard) goes in **3. Demo Video Walkthrough/** — NOT in Training Guides. It's the companion doc to the video itself.

---

### 3. Demo Video Walkthrough

Your demo videos for the CFO/CEO presentation and coworker training.

| What Goes Here | File Types | Notes |
|---------------|------------|-------|
| Demo video files | .mp4, .mov, etc. | The recorded walkthrough videos |
| Guide 05 — Video Demo Script & Storyboard | .pdf | The script/storyboard that goes with the videos |

**Future additions:** If you record multiple videos (e.g., one for leadership, one for coworkers, one for the Universal Toolkit), you can add subfolders later. For now, keep it flat.

---

### 4. Source Code

The raw code files that live inside the Excel file (VBA) and the standalone scripts (Python, SQL). This is for people who want to see the code outside of Excel, or for your own reference.

| Subfolder | What Goes Here | File Types | Source in GitHub Repo |
|-----------|---------------|------------|----------------------|
| **Python Modules/** | All 14 Python scripts | .py | `python/` folder |
| **SQL Modules/** | All SQL scripts | .sql | `sql/` folder |
| **VBA Modules/** | All 34 VBA module files | .bas | `vba/` folder |

**Important:** These are the SAME files that are in the GitHub repo. When you update code in GitHub, also update the copies here so they stay in sync.

**Who uses this folder:** Mostly you (Connor). Coworkers who are curious about the code can browse here, but they don't need to touch these files — the VBA is already baked into the .xlsm file.

---

### 5. Universal Code Library

The standalone tools that work on ANY Excel file — not just the P&L demo file. This is for coworkers who want to use specific tools (Data Sanitizer, Branding, Sheet Tools, etc.) on their own workbooks.

| Subfolder | What Goes Here | File Types | Source in GitHub Repo |
|-----------|---------------|------------|----------------------|
| **Python Kit/** | Universal Python scripts (standalone tools) | .py | `UniversalToolsForAllFiles/python/` |
| **VBA Kit/** | Universal VBA modules (import into any workbook) | .bas | `UniversalToolsForAllFiles/vba/` |

**Important:** These are DIFFERENT from the Source Code folder. Source Code (folder 4) has the P&L-specific code. Universal Code Library (folder 5) has tools that work on any file.

**Who uses this folder:** Coworkers who want to add specific tools to their own Excel files. They would import the .bas files into their workbook via VBA Editor (Alt+F11 > File > Import). Guide 06 (Universal Toolkit Guide) walks them through exactly how to do this.

---

## Quick Decision Chart — "Where Does This File Go?"

| If you have... | Put it in... |
|----------------|-------------|
| A new version of the P&L .xlsm file | 1. Demo File > Excel File - Fully Built/ (and backup the old one first) |
| A new version of the clean .xlsx file | 1. Demo File > Excel File - Clean/ |
| A new or updated training guide PDF | 2. Training Guides > the right subfolder (see table above) |
| A recorded demo video | 3. Demo Video Walkthrough/ |
| The video script/storyboard | 3. Demo Video Walkthrough/ |
| An updated .bas, .py, or .sql file | 4. Source Code > the right subfolder |
| A universal tool (.bas or .py) | 5. Universal Code Library > the right kit |
| An old version of any file you're replacing | 1. Demo File > Backups/ (for Excel files) |

---

## Checklist — When You Publish a New Version

Use this every time you update the toolkit and push it to SharePoint:

- [ ] Back up the current .xlsm and .xlsx to 1. Demo File > Backups/
- [ ] Replace the .xlsm in Excel File - Fully Built/
- [ ] Save a fresh .xlsx copy (File > Save As > .xlsx) and replace Excel File - Clean/
- [ ] Update any training guide PDFs that changed → put in the right subfolder in 2. Training Guides/
- [ ] Update .bas/.py/.sql files in 4. Source Code/ if code changed
- [ ] Update Universal Code Library files in 5. if universal tools changed
- [ ] Open the coworker README and update the "Last Updated" date if anything visible to coworkers changed
