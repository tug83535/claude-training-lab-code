# Guide Update & PDF Conversion — Instructions for Claude

## Background
Connor has a demo P&L Excel file with VBA macros for iPipeline (Finance & Accounting).
There are two sets of training guides that need work:

1. **3 existing PDFs** that are now outdated because new features were added
2. **8 brand-new guides** that need iPipeline branding applied and PDF conversion

All finished PDFs go in: `training/FinalGuidesPDFs/`

---

## Brand Styling Reference

**READ THIS FIRST:** `docs/ipipeline-brand-styling.md`

Quick summary:
- Primary: iPipeline Blue `#0B4779` | Navy `#112E51` | Innovation Blue `#4B9BCB`
- Accents: Lime Green `#BFF18C` | Aqua `#2BCCD3`
- Neutrals: Arctic White `#F9F9F9` | Charcoal `#161616`
- Fonts: Arial Bold (headings), Arial Narrow (subheadings), Arial Regular (body)
- Match the same look and feel as the 9 existing PDFs in `training/FinalGuidesPDFs/`

---

## PART 1 — Update 3 Outdated PDFs

These PDFs were made when the project had 62 Command Center actions and 14 universal
toolkit modules. Since then, new features were added. The source markdown needs to be
updated, re-branded, and re-converted to PDF.

### 1A. Quick Reference Card
- **Source:** `OldRoughVersions/FinalRoughGuides/04-Quick-Reference-Card.md`
- **Replace:** `training/FinalGuidesPDFs/04-Quick-Reference-Card.pdf`
- **What changed:**
  - Was 62 actions — now **65 actions**
  - Add these 3 new actions to the card:
    - Action 63: Run What-If Scenario Demo (7 presets + custom + restore)
    - Action 64: Load Pre-Built Scenario
    - Action 65: Restore to Original Values
  - These belong under a new category: **"What-If Demo"**

### 1B. Universal Toolkit Guide
- **Source:** `OldRoughVersions/FinalRoughGuides/06-Universal-Toolkit-Guide.md`
- **Replace:** `training/FinalGuidesPDFs/06-Universal-Toolkit-Guide.pdf`
- **What changed:**
  - Was 14 modules / ~100 tools — now **23 modules / ~140+ tools**
  - Add these 9 new modules:

| Module | Tools | What It Does |
|--------|-------|-------------|
| modUTL_ColumnOps | 7 | Column insert/delete/move/split/merge/fill/swap |
| modUTL_Compare | 1 | Sheet comparison with color-coded diff report |
| modUTL_Consolidate | 1 | Multi-sheet data consolidation with source tracking |
| modUTL_Highlights | 3 | Threshold, top/bottom N, duplicate highlighting |
| modUTL_PivotTools | 4 | PivotTable create/refresh/style/drill-down |
| modUTL_TabOrganizer | 6 | Sort/color/group/reorder/rename tabs in bulk |
| modUTL_Comments | 3 | Extract/clear/convert comments and notes |
| modUTL_ValidationBuilder | 5 | Data validation builder: lists, numbers, dates, custom |
| modUTL_LookupBuilder | 2 | VLOOKUP/INDEX-MATCH formula builder with preview |

  - Also add mention of:
    - modUTL_CommandCenter — master menu for all toolkit tools
    - modUTL_WhatIf — universal What-If scenario tool
  - Update all counts: "14 modules" → "23 modules", "~100 tools" → "~140+ tools"

### 1C. How to Use the Command Center
- **Source:** `OldRoughVersions/FinalRoughGuides/01-How-to-Use-the-Command-Center.md`
- **Replace:** `training/FinalGuidesPDFs/01-How-to-Use-the-Command-Center.pdf`
- **What changed:**
  - Add a brief mention that the Command Center now has 65 actions (was 62)
  - Page 4 of the menu (or wherever the action list ends) should include:
    - Action 63: Run What-If Scenario Demo
    - Action 64: Load Pre-Built Scenario
    - Action 65: Restore to Original Values

---

## PART 2 — Brand & Convert 8 New Guides to PDF

These guides are finished markdown files that have never been branded or converted to PDF.
Apply the same iPipeline branding as the existing PDFs, then save to `training/FinalGuidesPDFs/`.

| # | Source File | Suggested PDF Name |
|---|------------|-------------------|
| 1 | `training/LastGuidesReview/USER_TRAINING_GUIDE.md` | `05-User-Training-Guide.pdf` |
| 2 | `training/LastGuidesReview/OPERATIONS_RUNBOOK.md` | `07-Operations-Runbook.pdf` |
| 3 | `training/LastGuidesReview/WhatIf-Demo-Guide.md` | `08-WhatIf-Demo-Guide.pdf` |
| 4 | `training/LastGuidesReview/WhatIf-Universal-Guide.md` | `09-WhatIf-Universal-Guide.pdf` |
| 5 | `training/LastGuidesReview/Universal-CommandCenter-Guide.md` | `10-Universal-CommandCenter-Guide.pdf` |
| 6 | `training/LastGuidesReview/Universal-NewTools-Guide.md` | `11-Universal-NewTools-Guide.pdf` |
| 7 | `training/LastGuidesReview/CoPilot-Quick-Start-Card.md` | `12-CoPilot-Quick-Start-Card.pdf` |
| 8 | `training/LastGuidesReview/VBA-Module-Reference-List.md` | `13-VBA-Module-Reference-List.pdf` |

### Notes on each guide:

1. **User Training Guide** — Complete reference for all 65 Command Center actions across 16 categories. This is the big one — step-by-step for every action. Written for non-technical Finance staff.

2. **Operations Runbook** — Monthly P&L close cycle procedures: month-open, mid-month, month-close. Shows exact command sequences for each phase.

3. **What-If Demo Guide** — How to use Actions 63-65 in the demo file. 7 preset scenarios, custom mode, restore baseline. Written for the CFO/CEO demo.

4. **What-If Universal Guide** — How to use modUTL_WhatIf on ANY Excel file. Presets, custom percentages, baseline save/restore, impact report.

5. **Universal Command Center Guide** — The master menu for all ~140+ universal toolkit tools. 13 categories, keyword search, tool inventory, auto-discovery.

6. **Universal New Tools Guide** — Covers the 9 newest VBA modules (38 tools): Compare, Consolidate, PivotTools, Comments, ColumnOps, Highlights, TabOrganizer, ValidationBuilder, LookupBuilder.

7. **CoPilot Quick Start Card** — One-page cheat sheet: 5 common scenarios pointing users to the right CoPilot prompt section.

8. **VBA Module Reference List** — Catalog of all 38 demo file VBA modules + frmCommandCenter. Grouped by category with "easiest to adapt" section.

---

## PART 3 — PDFs That Do NOT Need Changes

These 6 are still accurate. Do not modify:

- `00-Start-Here-Welcome.pdf`
- `02-Getting-Started-First-Time-Setup.pdf`
- `03-What-This-File-Does-Overview.pdf`
- `AP-Copilot-Prompt-Guide.pdf`
- `Dynamic-Chart-Filter-Setup-Guide.pdf`
- `Welcome_README.pdf`

---

## Final Checklist

When done, `training/FinalGuidesPDFs/` should contain:

- [ ] `00-Start-Here-Welcome.pdf` (unchanged)
- [ ] `01-How-to-Use-the-Command-Center.pdf` (UPDATED — 65 actions)
- [ ] `02-Getting-Started-First-Time-Setup.pdf` (unchanged)
- [ ] `03-What-This-File-Does-Overview.pdf` (unchanged)
- [ ] `04-Quick-Reference-Card.pdf` (UPDATED — 65 actions)
- [ ] `05-User-Training-Guide.pdf` (NEW)
- [ ] `06-Universal-Toolkit-Guide.pdf` (UPDATED — 23 modules)
- [ ] `07-Operations-Runbook.pdf` (NEW)
- [ ] `08-WhatIf-Demo-Guide.pdf` (NEW)
- [ ] `09-WhatIf-Universal-Guide.pdf` (NEW)
- [ ] `10-Universal-CommandCenter-Guide.pdf` (NEW)
- [ ] `11-Universal-NewTools-Guide.pdf` (NEW)
- [ ] `12-CoPilot-Quick-Start-Card.pdf` (NEW)
- [ ] `13-VBA-Module-Reference-List.pdf` (NEW)
- [ ] `AP-Copilot-Prompt-Guide.pdf` (unchanged)
- [ ] `Dynamic-Chart-Filter-Setup-Guide.pdf` (unchanged)
- [ ] `Welcome_README.pdf` (unchanged)

**Total: 17 PDFs (6 unchanged + 3 updated + 8 new)**

---

## Current Project Stats (for accuracy in guides)

- Demo file VBA modules: **39 total** (34 core + 5 optional add-ins)
- Demo file Command Center actions: **65 total**
- Universal Toolkit VBA modules: **23 total**
- Universal Toolkit tools: **~140+ total**
- Python scripts: **22 total** (18 base + 4 in NewTools/)
- Total bugs found and fixed to date: **35+**
- Branch: `claude/resume-ipipeline-demo-qKRHn`
