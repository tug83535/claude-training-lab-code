# FinalExport — Everything You Need in One Place

This folder contains everything for the iPipeline P&L Demo — the demo file, all code, all guides, and all video scripts. Nothing else in the repo matters once this folder is complete.

---

## What's In Each Folder

### DemoFile/
**What:** The final Excel demo file (.xlsm) with all 39 VBA modules imported.
**Status:** Empty until you complete the final re-import. Drop the finished file here.
**Goes on SharePoint:** YES — this is the main file coworkers will download.

### DemoVBA/ (39 .bas files + 1 UserForm code file)
**What:** All 39 VBA modules for the demo Excel file. These are the source code files that get imported into the .xlsm via Alt+F11 > File > Import.
- 34 core modules (Config, FormBuilder, Dashboard, Reconciliation, etc.)
- 5 optional add-ins (TimeSaved, SplashScreen, ProgressBar, WhatIf, ExecBrief)
- 1 UserForm code file (frmCommandCenter_code.txt)

**Goes on SharePoint:** NO — these are source code. The code lives inside the .xlsm file.

### DemoPython/ (13 Python scripts + 4 SQL scripts)
**What:** The 13 demo Python scripts that complement the Excel file, plus 4 SQL scripts.
- Python: pnl_runner, pnl_forecast, pnl_dashboard, pnl_monte_carlo, etc.
- SQL: staging, transformations, validations, enhancements
- Includes requirements.txt for pip install

**Goes on SharePoint:** YES — for coworkers who want to run the Python tools.

### UniversalToolkit/ (27 VBA modules + 22 Python scripts)
**What:** The universal tools that work on ANY Excel file (not just the demo).
- vba/ — 23 core modules (140+ tools) + 4 NewTools modules
- python/ — 18 core scripts + 4 NewTools scripts + requirements.txt
- These are what become the future .xlam add-in (Scenario 2 in the sharing plan)

**Goes on SharePoint:** LATER — after demo is done and you build the .xlam add-in.

### Guides/ (15 PDFs)
**What:** All finalized training guides for coworkers. Ready to post.
- 00 - Start Here Welcome
- 01 - How to Use the Command Center
- 02 - Getting Started First Time Setup
- 03 - What This File Does Overview
- 04 - Quick Reference Card
- 05 - User Training Guide
- 06 - Universal Toolkit Guide
- 07 - Operations Runbook
- 08 - WhatIf Scenario Guide
- 09 - Universal CommandCenter Guide
- 10 - VBA Module Reference List
- AP - CoPilot Prompt Guide
- Dynamic Chart Filter Setup Guide
- START HERE
- SharePoint Welcome README

**Goes on SharePoint:** YES — upload all of these.

### VideoRecording/ (5 files)
**What:** Everything you need when you sit down to record the demo videos.
- Video_Demo_Master_Plan.md — overall strategy and structure
- Video_1_Script_Whats_Possible.md — Script for Video 1 (highlight reel)
- Video_2_Script_Full_Demo_Walkthrough.md — Script for Video 2 (full walkthrough)
- Video_3_Script_Universal_Tools.md — Script for Video 3 (universal tools)
- COMPILED_VIDEO_PACKAGE.md — all scripts compiled into one reference doc

**Goes on SharePoint:** The finished VIDEOS go on SharePoint, not these scripts.

### ForConnor/ (2 PDFs)
**What:** Your personal reference docs.
- Connors_Project_Wrap_Up.pdf — full project summary and what's left
- Connors_SharePoint_Folder_Map.pdf — how to organize SharePoint folders

**Goes on SharePoint:** NO — these are just for you.

---

## Quick Checklist — What To Do

1. Re-import all 39 VBA modules into the Excel file (use DemoVBA/ files)
2. Save the final .xlsm and drop it into DemoFile/
3. Upload to SharePoint: DemoFile/ + Guides/ + DemoPython/
4. Sit down with VideoRecording/ open and record the 3 demo videos
5. Upload finished videos to SharePoint
6. Later: build .xlam add-in from UniversalToolkit/ and post that too

---

## File Counts

| Folder | Files | Description |
|--------|-------|-------------|
| DemoVBA | 40 | 39 .bas modules + 1 UserForm code |
| DemoPython | 14 | 13 .py scripts + requirements.txt |
| DemoPython/sql | 4 | SQL scripts |
| UniversalToolkit/vba | 27 | 23 core + 4 NewTools .bas modules |
| UniversalToolkit/python | 23 | 18 core + 4 NewTools .py + requirements.txt |
| Guides | 15 | Training guide PDFs |
| VideoRecording | 5 | Video scripts + master plan |
| ForConnor | 2 | Personal reference docs |
| **Total** | **130** | Everything in one place |
