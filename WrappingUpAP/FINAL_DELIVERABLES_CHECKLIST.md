# Final Deliverables Checklist — What You Need to Finish

**Branch:** `claude/resume-ipipeline-demo-qKRHn`
**Last Updated:** 2026-03-09

This document lists exactly what files and actions are needed to get to a finished product — the recording, the SharePoint upload, and the coworker package. Nothing else.

---

## The 3 Things You're Delivering

1. **The Demo Video** — screen recording showing the P&L file + Command Center in action
2. **The SharePoint Package** — the Excel file + training PDFs uploaded for coworkers to use
3. **The Universal Toolkit** (later, after demo) — the ~100 tools for any Excel file

---

## Deliverable 1: The Demo Video

### What You Need From the Branch

| What | Where in Branch | Status | Notes |
|------|----------------|--------|-------|
| Video script + storyboard | `FinalRoughGuides/05-Video-Demo-Script-and-Storyboard.md` | Draft — needs your review | 18-22 min, 3 parts, word-for-word narration |
| Video package / build checklist | `videodraft/COMPILED_VIDEO_PACKAGE.md` | Ready | Master reference — tool counts, demo stats, structure |
| Demo plan | `videodraft/VIDEO_DEMO_PLAN.md` | Ready | Flow, tips, open questions (recording software, webcam, etc.) |
| The Excel file itself | Your local copy (with all 34 VBA modules imported) | Ready | You already imported and tested all 62 actions |

### What You Still Need to Do

- [ ] **Review the video script** (`05-Video-Demo-Script-and-Storyboard.md`) — make sure the order, narration, and talking points are what you want
- [ ] **Answer the open questions** in `VIDEO_DEMO_PLAN.md` — recording software, webcam yes/no, background music, where to host the video
- [ ] **Do a dry run** — walk through the full demo once without recording to practice timing
- [ ] **Record the video** — screen + narration following the script
- [ ] **Save the video** to `CompletePackageStorage/production/`

---

## Deliverable 2: The SharePoint Package

### What You Need From the Branch

| What | Where in Branch | Status | Action Needed |
|------|----------------|--------|---------------|
| The Excel demo file | Your local `.xlsm` (with VBA imported) | Ready | Save a clean copy |
| Guide: How to Use the Command Center | `FinalRoughGuides/01-How-to-Use-the-Command-Center.md` | Draft — needs review | Review → approve → convert to PDF |
| Guide: Getting Started / First Time Setup | `FinalRoughGuides/02-Getting-Started-First-Time-Setup.md` | Draft — needs review | Review → approve → convert to PDF |
| Guide: What This File Does (Leadership) | `FinalRoughGuides/03-What-This-File-Does-Leadership-Overview.md` | Draft — needs review | Review → approve → convert to PDF |
| Guide: Quick Reference Card | `FinalRoughGuides/04-Quick-Reference-Card.md` | Draft — needs review | Review → approve → convert to PDF |
| CoPilot Prompt Guide | `FinalRoughGuides/CoPilotPromptGuide/AP_Copilot_PromptGuideHelpV2.md` (+ .docx) | Ready | Convert to PDF if not already |

### What You Still Need to Do

- [ ] **Review guides 01-04** — read each one, flag anything that's wrong or missing
- [ ] **Fix the minor outdated items** found during guide review (see "Known Guide Fixes" below)
- [ ] **Convert approved guides to PDF** — coworkers get PDFs, not markdown
- [ ] **Lock down the Excel file:**
  - [ ] Open it fresh on a different machine or clean Excel session — confirm it works out of the box
  - [ ] Check that no personal file paths, test data, or debug code is left in the macros
  - [ ] Save final copy as `.xlsm`
- [ ] **Save final files to a dedicated output folder on your computer:**
  - [ ] The `.xlsm` file
  - [ ] All approved PDF guides
  - [ ] A dated backup copy in a separate backups folder
- [ ] **Upload to SharePoint:**
  - [ ] Create `Finance Automation/` folder
  - [ ] Create subfolders: `Demo File/`, `Training/`, `Video/`
  - [ ] Upload the `.xlsm` to `Demo File/`
  - [ ] Upload the training PDFs to `Training/`
  - [ ] Upload the video to `Video/`
  - [ ] Set permissions — one group for the whole folder
  - [ ] Pin the folder or add to team Quick Links

---

## Deliverable 3: Universal Toolkit (Later — After Demo)

This is Scenario 2. Don't worry about this until after the demo and SharePoint upload are done.

### What You Need From the Branch

| What | Where in Branch | Status |
|------|----------------|--------|
| VBA modules (13 files) | `UniversalToolsForAllFiles/vba/` + `vba/NewTools/` | Ready |
| Python scripts (22 files) | `UniversalToolsForAllFiles/python/` + `python/NewTools/` | Ready |
| How-to guide | `UniversalToolsForAllFiles/UNIVERSAL_TOOLS_HOW_TO_GUIDE.md` | Ready |
| Full toolkit training guide | `FinalRoughGuides/06-Universal-Toolkit-Guide.md` | Draft — needs review |

### What You'd Need to Do (Later)

- [ ] Convert Python scripts to `.exe` files (PyInstaller) so coworkers don't need Python installed
- [ ] Convert the how-to guide and toolkit guide to PDF
- [ ] Upload `.bas` files, `.exe` files, and PDFs to `Universal Tools (Optional)/` on SharePoint
- [ ] (Eventually) Package VBA tools into `KBT_UniversalTools.xlam` add-in

---

## Known Guide Fixes Needed (Minor)

These were found during review on 2026-03-09. All are small text updates — no structural changes.

| Guide | What to Fix |
|-------|-------------|
| Guide 01 | "7-sheet PDF" → "multi-sheet PDF" (4 places) — PDF export now dynamically discovers all monthly tabs |
| Guide 03 | "~99 tools" → "~100 tools" (4 places) |
| Guide 03 | "21 bugs found" → update to actual total (30+) |
| Guide 03 | "7-sheet report" → "multi-sheet report" (2 places) |
| Guide 05 | "~99 tools" → "~100 tools" (1 place) |
| Guide 06 | "Which Modules to Import" table missing 4 NewTools modules (modUTL_DataCleaningPlus, modUTL_AuditPlus, modUTL_DuplicateDetection, modUTL_NumberFormat) |
| VIDEO_DEMO_PLAN.md | "99 tools" → "~100 tools" (1 place) |

---

## Files You Do NOT Need for Deliverables

These folders are project internals — useful during development but not part of what you deliver to coworkers:

| Folder | Why It's Not Needed |
|--------|-------------------|
| `CLAUDE.md` | Instructions for Claude AI sessions — internal only |
| `tasks/` | Session management (todo.md, lessons.md) — internal only |
| `qa/` | Test plans and QA tracking — internal only |
| `_internal/` | All dev-only folders (NewTesting, ProjectRefresh, review, Testing_Issues, TESTRUN, DemofileChartBuild) — internal only |
| `docs/` | Developer docs and setup guides — internal only (coworkers get the training PDFs instead) |

---

## Summary — Your Finish Line

**To be done with everything, you need:**

1. Review and approve 4 training guides (01-04)
2. Fix the minor guide text updates listed above
3. Convert approved guides to PDF
4. Lock down the Excel file (clean test on fresh machine)
5. Copy everything to `CompletePackageStorage/production/`
6. Record the demo video
7. Upload to SharePoint

**That's it. Everything else is already built.**
