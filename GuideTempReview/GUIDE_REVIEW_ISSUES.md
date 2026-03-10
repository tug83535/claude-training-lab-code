# Guide Review — Issues Found

**Reviewer:** Claude (Opus 4.6 — Connor's other Claude account)
**Date:** 2026-03-10
**Scope:** All 8 PDFs in GuideTempReview/

---

## GLOBAL ISSUE (Affects ALL 8 Guides)

### ISSUE G-1: "P&L;" Rendering Bug — Every Guide
- **Severity:** HIGH — Visible to all readers
- **Where:** Every single guide, dozens of occurrences
- **Problem:** "P&L" renders as "P&L;" throughout all 8 PDFs. The semicolon is an HTML/markdown entity artifact (`&amp;` → `&` conversion failure during PDF generation).
- **Examples:**
  - Guide 00 page 1: "P&L; Automation Toolkit"
  - Guide 01 page 1: "P&L; Automation Toolkit — Complete User Guide"
  - Guide 02 page 2: "C:\Users\YourName\Documents\P&L; Toolkit\"
  - Guide 03 page 1: "P&L; Automation Toolkit — Overview"
  - Guide 04 page 1: "P&L; - Monthly Trend"
  - Guide 06 page 1: "P&L; Command Center"
  - CoPilot Guide: "P&L; file" in table of contents
  - Dynamic Chart Guide: appears in Need Help footer
- **Fix:** In the markdown source files, find every `P&L` and make sure the `&` is not being HTML-encoded. If using a markdown-to-PDF tool, check its HTML entity handling. A global find/replace of `P&L;` → `P&L` on the final output would also work.

### ISSUE G-2: "FP&A;" Same Rendering Bug
- **Severity:** MEDIUM
- **Where:** Guide 03 page 9 ("FP&A; Leadership"), Guide 06 page 19 ("FP&A; methodology")
- **Fix:** Same root cause as G-1. Fix the `&` encoding.

### ISSUE G-3: "P&L-specific;" Same Bug, Hyphenated Form
- **Severity:** LOW
- **Where:** Guide 06 page 24 ("For the P&L-specific; Command Center")
- **Fix:** Same root cause.

---

## MISSING GUIDE

### ISSUE M-1: Guide 05 — Video Demo Script & Storyboard Is Missing
- **Severity:** HIGH
- **Where:** The GuideTempReview folder
- **Problem:** All 8 guides consistently reference "Guide 05 — Video Demo Script & Storyboard" in their Complete Guide Set boxes and in the Guide 00 routing table. But there is no `05-Video-Demo-Script-and-Storyboard.pdf` in the folder. Only 8 PDFs exist (00, 01, 02, 03, 04, 06, CoPilot, Dynamic Chart).
- **Fix:** Either add the Guide 05 PDF to the folder, or remove all references to Guide 05 from every guide's Complete Guide Set box and routing table. The source markdown exists at `FinalRoughGuides/05-Video-Demo-Script-and-Storyboard.md` — it just needs to be converted to PDF and included.

---

## GUIDE 00 — Start Here (4 pages)

### ISSUE 00-1: Duplicate "Need Help?" Section
- **Severity:** MEDIUM
- **Where:** Pages 3 and 4
- **Problem:** Two separate "Need Help?" boxes appear. Page 3 has a general help list (Setup problems / Don't know which action / VBA errors / Questions). Page 4 has a different help box (Action 45 / Action 44 / CoPilot / Contact). These should be one section.
- **Fix:** Merge into a single "Need Help?" section. The page 4 version with Action 44/45 self-help steps is the better one — combine the page 3 bullet points into it.

### ISSUE 00-2: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1, Complete Guide Set box
- **Problem:** The box on page 1 lists guides but does NOT include Guide 05. However, the table on page 2 correctly lists Guide 05. These are inconsistent.
- **Fix:** Add "Guide 05 Video Demo Script & Storyboard" to the Complete Guide Set box on page 1 (assuming Guide 05 PDF is added per M-1).

---

## GUIDE 01 — How to Use the Command Center (23 pages)

### ISSUE 01-1: No Issues Found (Content)
- **Severity:** N/A
- **Notes:** This guide is excellent. All 62 actions documented with What/When/Expect format. The ASCII Command Center layout diagram is a great touch. Monthly Close Workflow is correct. Troubleshooting and FAQ are thorough. Keyboard shortcuts are accurate.

### ISSUE 01-2: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1, Complete Guide Set box
- **Fix:** Same as 00-2 — add Guide 05 to the box.

---

## GUIDE 02 — Getting Started / First Time Setup (15 pages)

### ISSUE 02-1: No Issues Found (Content)
- **Severity:** N/A
- **Notes:** This guide is outstanding. Every single click is written out. The 4 test verification sequence (Command Center → Version Number → Health Check → Navigate Home) is smart. The "First 5 Actions" walkthrough is well-chosen. The 10 troubleshooting issues cover every common problem. Setup Checklist at the end is a great touch.

### ISSUE 02-2: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1, Complete Guide Set box
- **Fix:** Same as 00-2.

---

## GUIDE 03 — What This File Does / Overview (14 pages)

### ISSUE 03-1: Test Status Table Is Incomplete/Placeholder
- **Severity:** MEDIUM
- **Where:** Page 8, Section 8 "Quality and Reliability" — Testing table
- **Problem:** The text says "The toolkit has been tested through a structured 8-category test plan:" and shows a table that starts with "T1: Compilation & Load | 8 tests" but the table appears to be cut off or incomplete. The Status column values are not visible in the extracted text.
- **Fix:** Complete the testing status table with all 8 categories and their current pass/fail status. Or simplify it to a summary statement like "69 tests across 8 categories — 15 passed, 54 remaining" since the exact numbers will change.

### ISSUE 03-2: "Toolkit Time" Column Missing Manual Comparison
- **Severity:** LOW — Nice to have
- **Where:** Page 7, "What the Toolkit Delivers Per Close Cycle" table
- **Problem:** The table shows toolkit time (e.g., "Reconciliation checks: 10 seconds") but does NOT show the manual comparison time. This is a leadership-facing document — the before/after contrast is the most compelling data point.
- **Fix:** Add a "Manual Time" column. Example: "Reconciliation checks — Manual: 2+ hours → Toolkit: 10 seconds". This directly supports the Time Saved Calculator idea from the Last Optional Adds list and is exactly what CFO/CEO want to see.

### ISSUE 03-3: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1
- **Fix:** Same as 00-2.

---

## GUIDE 04 — Quick Reference Card (9 pages)

### ISSUE 04-1: Duplicate "Need Help?" Section
- **Severity:** MEDIUM
- **Where:** Pages 8-9
- **Problem:** Same issue as Guide 00. Two "Need Help?" sections appear — one embedded in the content (page 8, with Action 45/44/50 and contact info), and a second as the standard footer box (page 9). They have slightly different content.
- **Fix:** Merge into one. Keep the more detailed version with Action 45/44/50.

### ISSUE 04-2: Should This Be Printable on Fewer Pages?
- **Severity:** LOW — Design preference
- **Where:** Entire guide
- **Problem:** The guide title says "1-page printable cheat sheet" but it's actually 9 pages. This is fine for a comprehensive reference card, but the description is misleading.
- **Fix:** Either change the description from "1-page printable cheat sheet" to "printable quick reference card" everywhere it appears (Guide 00, Guide 02, Guide 04 itself), OR create a true 1-page condensed version as a separate companion sheet. The 9-page version is great as-is — just fix the description.

### ISSUE 04-3: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1
- **Fix:** Same as 00-2.

---

## GUIDE 06 — Universal Toolkit Guide (26 pages)

### ISSUE 06-1: No Issues Found (Content)
- **Severity:** N/A
- **Notes:** This is the strongest guide in the set. The "Which Modules to Import" table is perfect for the audience. The Top 20 list is well-chosen. The 6 Use Case Playbooks are practical and specific. The VBA vs Python troubleshooting split is clear. The FAQ answers the right questions. The dependency on modUTL_Core is called out correctly and prominently.

### ISSUE 06-2: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1
- **Fix:** Same as 00-2.

---

## COPILOT PROMPT GUIDE (18 pages)

### ISSUE CP-1: No Issues Found (Content)
- **Severity:** N/A
- **Notes:** Excellent guide. The Quick Reference table at the top is well-organized (A through H categories + Mega Prompt). The All-in-One prompt structure (5 parts) is thorough. The "Tips for Getting the Best Results" chapter at the end adds real value. The privacy/safe sharing note is responsible and appropriate.

### ISSUE CP-2: Guide Number Label Says "Guide 08 of 08"
- **Severity:** LOW
- **Where:** Page 1, header
- **Problem:** The header says "Guide 08 of 08" but this is the CoPilot Prompt Guide, which isn't numbered in the official guide set (it's listed as an "Additional Resource" in Guide 00, not a numbered guide). Meanwhile, Guide 05 and Guide 07 (Dynamic Chart) exist as numbered items. The numbering is inconsistent across guides.
- **Fix:** Decide on a consistent numbering scheme. Suggestion:
  - 00: Start Here
  - 01: Command Center
  - 02: Getting Started
  - 03: Overview
  - 04: Quick Reference
  - 05: Video Demo Script (missing — add it)
  - 06: Universal Toolkit
  - 07: Dynamic Chart Filter (currently unnumbered)
  - 08: CoPilot Prompt Library
  Then update the "Guide XX of 08" headers on every guide to match.

### ISSUE CP-3: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1
- **Fix:** Same as 00-2.

---

## DYNAMIC CHART FILTER SETUP GUIDE (6 pages)

### ISSUE DC-1: "Need Help?" Footer References Command Center Actions
- **Severity:** LOW
- **Where:** Page 6
- **Problem:** The "Need Help?" footer mentions "Run Action 45 (Quick Health Check)" and "Run Action 44 (Full Integration Test)". But this guide is about adding dropdown filters to ANY Excel file — not specifically the P&L demo file. A reader working on their own file won't have those actions available.
- **Fix:** Either remove the Action 44/45 references from this specific guide's footer (keep just the CoPilot and Contact lines), or add a note: "Actions 44 and 45 are only available in the P&L Automation Toolkit file."

### ISSUE DC-2: Guide Number Label Says "Guide 07 of 08"
- **Severity:** LOW
- **Where:** Page 1, header
- **Problem:** Same numbering inconsistency as CP-2. This guide is called "Guide 07 of 08" in its header but is listed as unnumbered ("--") in the Guide 00 table.
- **Fix:** Assign it number 07 officially and update Guide 00's table to show "07" instead of "--".

### ISSUE DC-3: Guide Set Box Skips Guide 05
- **Severity:** LOW
- **Where:** Page 1
- **Fix:** Same as 00-2.

---

## SUMMARY — All Issues by Priority

### HIGH Priority (Fix Before Publishing)
| # | Issue | Guides Affected |
|---|-------|----------------|
| G-1 | "P&L;" semicolon rendering bug | ALL 8 guides |
| M-1 | Guide 05 PDF is missing from folder | Folder + all guide references |

### MEDIUM Priority (Should Fix)
| # | Issue | Guides Affected |
|---|-------|----------------|
| G-2 | "FP&A;" same rendering bug | Guide 03, Guide 06 |
| 00-1 | Duplicate "Need Help?" section | Guide 00 |
| 03-1 | Test status table incomplete | Guide 03 |
| 04-1 | Duplicate "Need Help?" section | Guide 04 |

### LOW Priority (Nice to Fix)
| # | Issue | Guides Affected |
|---|-------|----------------|
| G-3 | "P&L-specific;" hyphenated rendering bug | Guide 06 |
| 00-2 | Guide Set box skips Guide 05 | ALL 8 guides |
| 03-2 | Missing "Manual Time" comparison column | Guide 03 |
| 04-2 | "1-page cheat sheet" is actually 9 pages | Guide 04 + references |
| CP-2 | Guide numbering inconsistency (08 vs unnumbered) | CoPilot Guide |
| DC-1 | Need Help footer references P&L-only actions | Dynamic Chart Guide |
| DC-2 | Guide numbering inconsistency (07 vs "--") | Dynamic Chart Guide |

---

## WHAT'S GOOD — Strengths Worth Noting

These guides are genuinely world-class. Specific strengths:

1. **Consistency** — Every guide follows the same structure: branded header, Complete Guide Set box, Table of Contents, section-by-section content, Need Help footer, Confidential footer. Professional and cohesive.

2. **Audience awareness** — Written exactly right for non-technical Finance & Accounting staff. No jargon. Every click spelled out. Screenshots described in words where images aren't possible.

3. **Guide 01** (Command Center) — The ASCII Command Center layout diagram is genius. The "What It Replaces" before/after table on page 1 immediately sells the value. All 62 actions documented consistently.

4. **Guide 02** (Getting Started) — The 4-test verification sequence is exactly what a first-time user needs. The 10 troubleshooting issues cover every realistic scenario. The Setup Checklist at the end is a nice capstone.

5. **Guide 03** (Overview) — The cost comparison table ($0 vs $50K-$200K/year) is the single most persuasive element in the entire package. Keep this front and center for the CFO/CEO.

6. **Guide 04** (Quick Reference) — Clean, scannable, well-organized. The Monthly Close Workflow and Top 10 tables are exactly what people will pin to their desk.

7. **Guide 06** (Universal Toolkit) — The "Which Modules to Import" table is perfect. The 6 Use Case Playbooks are practical and immediately actionable. Best guide in the set.

8. **CoPilot Guide** — The All-in-One prompt structure (5 parts) is thorough and practical. The A-H category system is well-organized. The privacy note shows maturity.

9. **Dynamic Chart Guide** — Concise, focused, exactly the right level of detail. Two methods (Helper Table vs PivotTable) gives users options based on comfort level.

**Bottom line:** Fix the "P&L;" rendering bug globally, add Guide 05, merge the duplicate Need Help sections, and these are ready to publish.
