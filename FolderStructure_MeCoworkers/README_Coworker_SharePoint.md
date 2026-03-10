# Welcome to the P&L Automation Toolkit

**iPipeline Finance & Accounting**
**Created by:** Connor Atlee (Connor.Atlee@ipipeline.com)
**Last Updated:** March 2026

---

## What Is This?

This is a collection of Excel automation tools built for iPipeline's Finance and Accounting team. The centerpiece is a P&L workbook with 62 one-click actions that automate tasks like reconciliation, variance analysis, data quality checks, PDF exports, dashboards, and more.

Everything you need is organized into 5 folders. Here is what each one contains and when you would use it.

---

## Folder Guide

### 1. Demo File

**This is where you get the Excel file.**

| Subfolder | What's Inside | When to Use It |
|-----------|--------------|----------------|
| **Excel File - Fully Built** | The P&L workbook with all 62 automation actions built in (.xlsm) | **Start here.** This is the main file. Download it, open it in Excel, enable macros, and press Ctrl+Shift+M to open the Command Center. |
| **Excel File - Clean** | The same P&L workbook but without any macros or code (.xlsx) | Use this if you just want to look at the data and sheets without enabling macros. No automation — just the raw workbook. |
| **Backups** | Previous versions of the Excel files | You do not need this folder. It is for version history only. |

**Which file should I download?**
- If you want to use the automation tools → download from **Excel File - Fully Built**
- If you just want to browse the data → download from **Excel File - Clean**

---

### 2. Training Guides

**Step-by-step guides that show you how to use everything.** Start with folder 1 and work your way through.

| Subfolder | What's Inside | Who It's For |
|-----------|--------------|-------------|
| **1. Getting Started** | Welcome guide, first-time setup instructions, and a full overview of what the toolkit does | **Everyone.** Read these first if you are new. They walk you through downloading the file, enabling macros, and running your first actions. |
| **2. Quick Reference** | The complete Command Center guide (all 62 actions explained), a printable quick reference card, and a guide for adding dropdown chart filters | **Everyone.** These are the guides you will come back to most often. Bookmark them. |
| **3. Python, VBA, & SQL** | The Universal Toolkit guide — how to use standalone code tools on your own Excel files | **Optional.** Only read this if you want to add specific tools (like the Data Sanitizer or Branding module) to your own separate workbooks. |
| **4. CoPilot Prompt Guide (code self-help)** | Pre-built prompts you can paste into Microsoft CoPilot or ChatGPT to troubleshoot VBA and Python code on your own | **Optional.** Read this if you use AI tools and want ready-made prompts for debugging, improving, or understanding Excel code. |

**Recommended reading order for new users:**
1. Guide 00 — Start Here (Welcome) — 2 minutes
2. Guide 02 — Getting Started (First Time Setup) — 10 minutes
3. Guide 03 — What This File Does (Overview) — 5 minutes
4. Guide 01 — How to Use the Command Center — reference as needed
5. Guide 04 — Quick Reference Card — print it and keep it handy

---

### 3. Demo Video Walkthrough

**Video recordings showing the toolkit in action.** Watch these to see what the tools look like before you try them yourself, or to learn specific workflows.

---

### 4. Source Code

**The raw code files behind the toolkit.** This folder contains the VBA, Python, and SQL source code that powers everything.

| Subfolder | What's Inside |
|-----------|--------------|
| **Python Modules** | 14 Python scripts for data analysis, forecasting, ETL, and reporting |
| **SQL Modules** | SQL scripts for database queries and data extraction |
| **VBA Modules** | 34 VBA module files — the code that runs inside the Excel file |

**Do I need this folder?** Probably not. The VBA code is already built into the .xlsm file — you do not need to download or import anything from here to use the toolkit. This folder exists for transparency and for anyone who wants to review or learn from the code.

---

### 5. Universal Code Library

**Standalone tools you can add to ANY Excel file** — not just the P&L demo file.

| Subfolder | What's Inside |
|-----------|--------------|
| **Python Kit** | Python scripts that work on any Excel workbook |
| **VBA Kit** | VBA modules you can import into any Excel workbook |

**Examples of what's in here:**
- **Data Sanitizer** — Fixes text-stored numbers, floating-point noise, and formatting issues on any sheet
- **iPipeline Branding** — Applies iPipeline brand colors, fonts, and formatting to any workbook
- **Sheet Tools** — Creates a clickable sheet index, clones templates, generates unique IDs
- **Audit Tools** — Finds external links, audits hidden sheets, creates masked copies

**How to use these:** See the Universal Toolkit Guide in Training Guides > 3. Python, VBA, & SQL. It has step-by-step instructions for importing these modules into your own files.

---

## Quick Start (5 Minutes)

If you just want to get up and running as fast as possible:

1. **Download** the .xlsm file from **1. Demo File > Excel File - Fully Built**
2. **Save** it to a folder on your local computer (not a network drive)
3. **Open** it in Excel (Windows desktop — Excel 2019, 2021, or 365)
4. **Click** "Enable Content" when the yellow security bar appears
5. **Press** Ctrl+Shift+M to open the Command Center
6. **Run** Action 45 (Quick Health Check) to verify everything works

That's it. You are ready to use the toolkit. For the full setup walkthrough, see Training Guides > 1. Getting Started.

---

## Need Help?

| What You Need | What to Do |
|--------------|-----------|
| Setup help or first-time questions | Read Guide 02 (Getting Started) in Training Guides > 1. Getting Started |
| How to run a specific action | Read Guide 01 (Command Center) in Training Guides > 2. Quick Reference |
| Quick answer about what an action does | Read Guide 04 (Quick Reference Card) in Training Guides > 2. Quick Reference |
| Self-service troubleshooting | Run Action 45 (Quick Health Check) from the Command Center |
| AI-powered code help | See the CoPilot Prompt Guide in Training Guides > 4. CoPilot Prompt Guide |
| Still stuck | Contact Connor Atlee — Connor.Atlee@ipipeline.com |

---

*iPipeline Finance & Accounting — Automation Project*
*Confidential — Internal Use Only*
