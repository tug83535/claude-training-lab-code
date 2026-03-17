# iPipeline Finance Automation — Complete Video & Delivery Package

**Compiled:** 2026-03-06
**Purpose:** This document contains EVERYTHING needed to produce and deliver the iPipeline Finance Automation demo videos and SharePoint package. It is designed to be handed to another AI (Claude Code or similar) for review, execution, and production support.

**What's in this document:**
1. Master Plan — all decisions, strategy, folder structure
2. Video 1 Script — "What's Possible" (4–5 min overview)
3. Video 2 Script — "Full Demo Walkthrough" (15–18 min chaptered)
4. Video 3 Script — "Universal Tools" (8–10 min chaptered)
5. START HERE Document — content spec for the SharePoint landing PDF
6. Sample File Spec — requirements for the Video 3 demo file

**Important context:** The universal code library contains 12 VBA modules (~78 tools) and 22 Python scripts. Every script in this document was written after a full review of every code file.

---
---



================================================================================
================================================================================
# SECTION 1: MASTER PLAN
================================================================================
================================================================================

# iPipeline Finance Automation — Video & Delivery Master Plan

**Created:** 2026-03-06
**Status:** Final Plan — Ready for Production

---

## Executive Summary

This document captures every decision made during the planning review for the iPipeline Finance Automation demo videos and SharePoint delivery package. It serves as the single source of truth for producing and delivering the final product.

**What we're delivering:**
- 3 videos (overview, full demo, universal tools)
- A demo Excel file for exploration
- A universal code library (VBA, Python, SQL) for coworkers to use on their own files
- Step-by-step training guides (PDF format)
- A master Copilot prompts document
- A sample Excel file used in Video 3
- A START HERE document that orients anyone landing in the SharePoint folder

---

## The Goal

Show iPipeline employees — primarily Finance & Accounting (65%) and the broader company (35%) — what can be done with VBA, Python, and SQL to improve daily workflows and cut down on manual work. The demo file demonstrates what's possible. The universal code library gives people tools they can grab and apply to their own files immediately.

This is a practical enablement project, not just a showcase.

---

## Context & Background

- **Who built it:** Connor (Finance & Accounting team member, not a developer) with AI assistance (Claude)
- **What was built:** A single .xlsm file with 62 automated VBA macros organized into a Command Center, plus 14 Python scripts, ~99 universal tools, and 6 training guides
- **CFO involvement:** The CFO tasked Connor with this project but has not yet seen the output
- **Testing:** All code is done, demo file built and tested. No external users have tested it yet — the video will be the first exposure
- **Timeline:** ASAP — everything is built, just needs packaging and recording

---

## Video Strategy — Three Videos

### Why Three Videos (Not One)

Different audiences need different things:
- The CFO and leadership need to see impact in under 5 minutes
- The Finance team needs a detailed walkthrough to understand and trust the toolkit
- Coworkers in other departments want to know what tools they can grab for their own files

One video can't serve all three well. Three focused videos let each audience watch only what's relevant to them.

### Video 1: "What's Possible" — Overview (4–5 minutes)

**Audience:** Everyone at iPipeline (2,000+ employees)
**Purpose:** Build awareness. Show a few high-impact features. Point people to the right next step.
**Tone:** Helpful colleague sharing something useful. Lead with the positive — "here's what's now possible" — not the pain of the old process.

**Flow:**
1. Opening hook (30 sec) — What if reports, checks, and charts could be done in seconds? 62 actions, one click each, nothing to install.
2. Command Center intro (30 sec) — Open it, show search, show categories. Quick and visual.
3. Feature demos (2.5 min total):
   - Data Quality Scan + Letter Grade (~40 sec)
   - Variance Commentary — auto-generated English narratives (~40 sec)
   - Dashboard / Executive Dashboard (~40 sec)
   - PDF Export (~20 sec)
4. Bridge to the universal library (30 sec) — The code behind this works on any Excel file. There's a library of tools you can use on your own work.
5. Closing + CTA (30 sec) — Where to find the files, guides, and other videos.

### Video 2: "Full Demo Walkthrough" (15–18 minutes, chaptered)

**Audience:** Finance & Accounting team, interested power users
**Purpose:** Detailed tour of the P&L demo file. Shows how things actually work, step by step.
**Scope:** Purely the demo file. No mention of the universal code library — that's Video 3's job.

**Chapters:**
1. The Workbook & Command Center (2 min) — Sheet orientation, Command Center walkthrough, search and launch
2. Data Import & Quality (2.5 min) — GL Import → Data Quality Scan + Letter Grade → Reconciliation Checks (PASS/FAIL)
3. Analysis (3 min) — Variance Analysis → Variance Commentary (auto narratives) → YoY Variance
4. Reporting & Visuals (2.5 min) — Dashboard Charts → Executive Dashboard → PDF Export
5. Enterprise Features (2.5 min) — Executive Mode → Version Control → Scenario Management → Sensitivity Analysis
6. Under the Hood (1.5 min) — Integration Test (18/18 PASS) → Audit Log → brief architecture mention
7. Closing & Next Steps (1 min) — Where to find the file, training guides, how to ask questions

### Video 3: "Universal Tools" (8–10 minutes, chaptered)

**Audience:** Anyone at iPipeline who uses Excel and wants to automate their own work
**Purpose:** Show examples of the universal tools running on a plain sample file, then point people to SharePoint to get the code and guides
**Key Decision:** Uses a separate, simple sample Excel file (not the demo file) to make it clear these tools work on ANY file

**Flow:**
1. Opening (30 sec) — Context: these tools work on any Excel file, here's what they do
2. Chapter 1: Sheet Tools (2.5 min) — AutoFit, Sort, Protect, Find & Replace across sheets
3. Chapter 2: Data Tools (2.5 min) — Data cleanup, formatting, restructuring tools
4. Chapter 3: Python & SQL Tools (2 min) — Button-click demos showing output (audience doesn't need to know code)
5. Closing (30 sec) — Where to find the full library, guides, and Copilot prompts on SharePoint

**Demo pattern for each tool:** Here's the file before → here's the click → here's the file after.

---

## Production Decisions

| Decision | Choice |
|----------|--------|
| Webcam overlay | No — screen recording only |
| Title cards between sections | Yes — Videos 2 and 3 (chaptered videos). Simple cards with chapter name + one-line description. Not needed for Video 1. |
| Time savings overlay | Text overlay after key actions: "Manual: 2 hours → Automated: 10 seconds". Use 3–4 times in Video 2, not on every action. |
| Background music | Intro and outro only on all three videos. Silence during demos. |
| Script approach | Full word-for-word scripts for all three videos. Practice enough to deliver naturally. Script visible on second monitor during recording. |
| Recording approach | Record each chapter/section as a separate clip, stitch together in editing. Allows re-recording individual sections without redoing the whole video. |
| Resolution / format | 1920×1080, 30fps, MP4 (H.264) |
| Tone | Professional but approachable. Helpful colleague, not salesperson. Confident, practical, focused on "here's what this does for you." |

---

## SharePoint Folder Structure

```
iPipeline Finance Automation/
│
├── START HERE.pdf
│
├── Videos/
│   ├── 1 - What's Possible (Overview).mp4
│   ├── 2 - Full Demo Walkthrough.mp4
│   └── 3 - Universal Tools.mp4
│
├── Demo File/
│   ├── iPipeline_PnL_Demo.xlsm
│   └── Sample_File_For_Universal_Tools.xlsx
│
├── Universal Code Library/
│   ├── VBA/
│   ├── Python/
│   └── SQL/
│
├── Training Guides/
│   ├── 01 - Getting Started.pdf
│   ├── 02 - Command Center Guide.pdf
│   ├── 03 - VBA Tools Guide.pdf
│   ├── 04 - Python Tools Guide.pdf
│   └── (additional guides as needed)
│
└── Copilot Prompts.pdf
```

**START HERE.pdf** — A one-page document at the top level that tells anyone landing in this folder exactly what they're looking at, which video to watch first, and where to go based on what they need.

**Training guides** — Delivered as PDFs. Written for a complete beginner audience (step-by-step, assume they know nothing) while still being useful for advanced users.

**Copilot Prompts** — One master PDF document with pre-built prompts for coworkers who need help using the code or making minor edits.

---

## Key Principles

1. **The demo file shows what's possible.** It's an example, not the deliverable.
2. **The universal code library is the deliverable.** That's what people grab and use on their own files.
3. **Every guide assumes the reader knows nothing.** Detailed step-by-step instructions regardless of skill level.
4. **Each video serves a different audience.** Nobody should have to sit through content that isn't for them.
5. **Lead with the positive.** "Here's what's now possible" — not "here's what was broken."
6. **Copilot prompts fill the gap.** For anyone who needs help applying the code, pre-built prompts get them there without needing Connor.

---

## Build Checklist

| Step | Item | Status |
|------|------|--------|
| 1 | Master Plan Document (this file) | ✅ Complete |
| 2 | Video 1 outline + full script | ⬜ Not started |
| 3 | Video 2 outline + full script | ⬜ Not started |
| 4 | Video 3 outline + full script | ⬜ Not started |
| 5 | START HERE document | ⬜ Not started |
| 6 | Sample Excel file for Video 3 | ⬜ Not started |

---

*Document created: 2026-03-06 | Source: AI Briefing Video Review conversation*


================================================================================
================================================================================
# SECTION 2: VIDEO 1 SCRIPT
================================================================================
================================================================================

# Video 1 Script — "What's Possible"

**Runtime Target:** 4:00–5:00
**Format:** Screen recording, no webcam, voice-over narration
**Audience:** All iPipeline employees (2,000+)
**Purpose:** Build awareness, show high-impact features, point to next steps
**Music:** Subtle corporate/tech track during title card and closing card only

---

## Pre-Recording Checklist

Before hitting record, make sure:

- [ ] Excel is the only application open
- [ ] Desktop is clean — no icons, no taskbar if possible (auto-hide taskbar)
- [ ] Excel is maximized to full screen
- [ ] Zoom level set to 100% or 110% (pick one and don't change it during recording)
- [ ] Windows display scaling set to 100% (not 125% or 150%)
- [ ] All notifications silenced (Teams, Outlook, Windows notifications OFF)
- [ ] Demo file is open and on the Report--> landing page
- [ ] Command Center is closed (you'll open it on camera)
- [ ] No previous macro outputs visible — start clean (no leftover Variance Analysis sheets, dashboards, etc.)
- [ ] Script visible on second monitor or printed nearby
- [ ] Audio test done — 30-second recording, listen back with headphones
- [ ] Screen recording software running, resolution confirmed at 1920×1080, 30fps

---

## Title Card (5 seconds)

**On screen:** Branded title card (iPipeline colors)

```
iPipeline Finance Automation
What's Possible
```

**Audio:** Brief music sting (2-3 seconds), then fade to silence

**Production note:** Create this as a simple image or slide. iPipeline Blue (#0B4779) background, white text, Arial font. Can add the iPipeline logo if permitted.

---

## Section 1: Opening Hook

**Duration:** 0:05–0:35 (30 seconds)
**On screen:** The demo file is open, showing the Report--> landing page

### Script:

> "This is a single Excel file. Nothing to install, nothing to configure — you just open it and go.
>
> Inside are 62 automated actions that handle reporting, analysis, data quality checks, charts, exports, and more — each one triggered with a single click.
>
> In the next few minutes, I'm going to show you what that looks like."

### Screen Actions:
- File is already open on the Report--> page
- Slowly scroll down the Report--> page as you narrate, so the viewer sees it's a real, populated workbook
- No clicking yet — just a smooth, slow scroll

### Production Notes:
- Pace yourself. This is the first thing people hear — speak clearly and confidently
- Don't rush the "62 automated actions" line. Let that number land.
- The scroll should be slow and deliberate, not frantic

---

## Section 2: Command Center Introduction

**Duration:** 0:35–1:15 (40 seconds)
**On screen:** Opening and browsing the Command Center

### Script:

> "Everything runs from one place — the Command Center.
>
> You can open it with Ctrl+Shift+M, or from the button on the landing page.
>
> [open Command Center]
>
> Every action is organized by category — Monthly Operations, Analysis, Reporting, Enterprise Features, and more. You can scroll through or just search.
>
> [type a search term, e.g., 'variance']
>
> Type what you're looking for and it filters instantly. Find the action, click Run, and it handles the rest.
>
> Let me show you a few examples."

### Screen Actions:
1. Click the Command Center button (or use Ctrl+Shift+M) — pause for a beat after it opens so the viewer can take it in
2. Slowly scroll through the categories — don't rush, let people read a few action names
3. Click into the search bar, type "variance" — show the filtered results
4. Clear the search, leave the Command Center open

### Production Notes:
- When the Command Center opens, pause for 1–2 seconds before speaking. Let the visual register.
- Mouse movements should be smooth and deliberate — hover over category headers so people can read them
- Don't explain every category. The scroll does the work.

---

## Section 3: Feature Demo — Data Quality Scan + Letter Grade

**Duration:** 1:15–1:55 (40 seconds)
**On screen:** Running the Data Quality Scan from the Command Center

### Script:

> "First — data quality. Before you do anything with your numbers, you want to know if the data is clean.
>
> [click Run on Data Quality Scan]
>
> One click, and it scans your entire workbook across six categories — completeness, accuracy, consistency, formatting, outliers, and cross-references.
>
> [sheet appears with results]
>
> It gives you a letter grade — right there at the top. In this case, [read the grade]. You get a full breakdown underneath showing exactly where issues are, if any.
>
> Fifteen seconds, start to finish."

### Screen Actions:
1. In the Command Center, find "Data Quality Scan" (search or scroll)
2. Click Run
3. Wait for it to complete — the macro runs and creates/navigates to the Data Quality Report sheet
4. Pause on the letter grade badge (28pt colored badge) — hold for 2–3 seconds so the viewer sees it
5. Slowly scroll down to show the category breakdown

### Production Notes:
- The letter grade badge is your visual anchor here. Make sure the viewer sees it clearly.
- Read the actual grade that appears — don't script a specific grade in advance since it may vary
- After "Fifteen seconds, start to finish" — pause briefly before moving on. Let the speed sink in.

### Time Savings Overlay (optional):
After the action completes, display text overlay:
```
⏱ This scan has never been done manually — now it takes 15 seconds
```

---

## Section 4: Feature Demo — Variance Commentary

**Duration:** 1:55–2:40 (45 seconds)
**On screen:** Running Variance Commentary from the Command Center

### Script:

> "Next — one of the most useful features in the whole file.
>
> After running a variance analysis, the system can automatically generate written commentary for the top five variances.
>
> [click Run on Variance Commentary]
>
> [sheet appears with narratives]
>
> These are plain English narratives — ready to drop into an email, a report, or a presentation. It identifies the line item, the dollar and percentage change, and describes what happened.
>
> No copying numbers into a paragraph. No writing it yourself. One click."

### Screen Actions:
1. Navigate back to Command Center (Ctrl+Shift+M)
2. Find "Variance Commentary" — click Run
3. Wait for it to complete — navigates to the Variance Commentary sheet
4. Slowly scroll through the generated narratives — pause on at least two so the viewer can read a few words
5. Highlight (with your mouse cursor) one narrative to draw the eye

### Production Notes:
- This is your jaw-drop feature. Build it up with "one of the most useful features in the whole file."
- Let the viewer READ the narratives. Don't talk over them immediately. Give 2–3 seconds of silence while the text is visible.
- Move your cursor near (not on top of) the text so the viewer knows where to look
- This feature is what makes a CFO lean forward. Give it room to breathe.

---

## Section 5: Feature Demo — Executive Dashboard

**Duration:** 2:40–3:20 (40 seconds)
**On screen:** Running the Executive Dashboard from the Command Center

### Script:

> "When it's time to present to leadership, you need visuals — not spreadsheets.
>
> [click Run on Executive Dashboard]
>
> [dashboard appears]
>
> One click builds a full executive dashboard — KPI summary cards at the top, a waterfall chart showing how you get from budget to actual, and a product line comparison. All branded, all formatted, all ready to present.
>
> You can also build a full set of eight charts on a separate sheet, or export everything to a clean, formatted PDF — headers, footers, page numbers included."

### Screen Actions:
1. Navigate back to Command Center
2. Find "Executive Dashboard" — click Run
3. Wait for it to build — the Executive Dashboard sheet appears
4. Slowly scroll through the dashboard — pause on the KPI cards, then the waterfall chart, then the product comparison
5. Brief pause at the end

### Production Notes:
- The dashboard is visually rich — let it fill the screen and give the viewer time to absorb it
- Don't describe every chart element. The visual speaks for itself. Your narration just frames it.
- The mention of PDF export and the eight-chart sheet is a quick verbal teaser — you're NOT navigating to those. Just planting the idea that there's more.

---

## Section 6: Bridge to Universal Code Library

**Duration:** 3:20–3:50 (30 seconds)
**On screen:** Back on the Command Center or the Report--> landing page

### Script:

> "That's a sample of what this file can do for a P&L close process. But the code behind it — the VBA, the Python, the SQL — isn't locked to this one file.
>
> There's a library of universal tools that work on any Excel spreadsheet. Formatting, cleanup, sorting, searching across sheets — all reusable, all documented, all available on SharePoint.
>
> There's a separate video walking through those tools if you want to see what's available for your own work."

### Screen Actions:
- Navigate back to the Command Center or the Report--> page — something visually neutral
- No new actions being run. This is narration over a static screen.
- Optionally, scroll slowly through the Command Center categories showing the universal tools section if one is visible

### Production Notes:
- This is a transition moment, not a demo. Keep the energy up but don't try to demo anything new.
- The key phrase is "work on any Excel spreadsheet" — emphasize that
- Don't linger. 30 seconds, then move to closing.

---

## Section 7: Closing + Call to Action

**Duration:** 3:50–4:20 (30 seconds)
**On screen:** Report--> landing page or a closing title card

### Script:

> "Everything you just saw runs from this one Excel file — nothing to install, no cost, no IT involvement.
>
> If you want to explore the file yourself, watch the full demo walkthrough, or grab tools from the code library — it's all on SharePoint. There are step-by-step guides for everything, and if you need help, there are pre-built Copilot prompts to walk you through it.
>
> [PAUSE]
>
> Thanks for watching."

### Screen Actions:
- Stay on the Report--> landing page during the first part
- On "it's all on SharePoint" — optionally cut to a closing title card showing:

```
Find everything on SharePoint:
[SharePoint folder link or path]

Videos | Demo File | Code Library | Training Guides
```

### Production Notes:
- "Nothing to install, no cost, no IT involvement" — say this clearly. It's a key selling point.
- The pause before "Thanks for watching" lets the CTA land. Don't rush to the end.
- The closing title card should stay on screen for 5+ seconds so people can note the SharePoint location
- Brief music sting on the closing card (same track as the opening)

---

## Closing Title Card (5 seconds)

**On screen:** Branded closing card (same style as opening)

```
iPipeline Finance Automation

Videos | Demo File | Code Library | Guides
[SharePoint location]

Questions? Contact Connor [last name / email]
```

**Audio:** Brief music, fade out

---

## Total Runtime Breakdown

| Section | Duration | Cumulative |
|---------|----------|------------|
| Title Card | 0:05 | 0:05 |
| Opening Hook | 0:30 | 0:35 |
| Command Center | 0:40 | 1:15 |
| Data Quality Scan | 0:40 | 1:55 |
| Variance Commentary | 0:45 | 2:40 |
| Executive Dashboard | 0:40 | 3:20 |
| Bridge to Universal Library | 0:30 | 3:50 |
| Closing + CTA | 0:30 | 4:20 |
| Closing Card | 0:05 | 4:25 |
| **TOTAL** | | **~4:25** |

Buffer for natural pauses, macro run times, and breathing room: expect **4:30–5:00** in practice.

---

## Recording Tips Specific to This Video

1. **Record each section as a separate clip.** If you flub the Variance Commentary narration, you only re-record that 45-second section — not the whole video.

2. **Do a full dry run first.** Run every macro you plan to demo, in order, before recording. Make sure nothing errors out and all outputs look clean.

3. **Reset the file between takes.** If you need to re-record a section, make sure previous macro outputs are cleared so you're starting from the same state.

4. **Watch your mouse.** During narration-only moments (bridge section, closing), keep the mouse still or move it very slowly. Jittery mouse movement is distracting.

5. **Leave 2 seconds of silence at the start and end of each clip.** Makes editing and stitching much easier.

6. **Speak slightly slower than you think you should.** Screen recording narration almost always sounds rushed on playback. Slow down 10%.

---

*Script created: 2026-03-06 | Part of Video Demo Master Plan*


================================================================================
================================================================================
# SECTION 3: VIDEO 2 SCRIPT
================================================================================
================================================================================

# Video 2 Script — "Full Demo Walkthrough"

**Runtime Target:** 15:00–18:00
**Format:** Screen recording, no webcam, voice-over narration, chaptered
**Audience:** Finance & Accounting team (primary), interested power users (secondary)
**Purpose:** Detailed tour of the P&L demo file — show how things work, step by step
**Scope:** Demo file only. No mention of universal code library (that's Video 3).
**Music:** Subtle corporate/tech track during title card, chapter cards, and closing card only

---

## Pre-Recording Checklist

Before hitting record, make sure:

- [ ] Excel is the only application open
- [ ] Desktop is clean — no icons, taskbar auto-hidden
- [ ] Excel is maximized to full screen
- [ ] Zoom level set to 100% or 110% (same as Video 1 — be consistent across all videos)
- [ ] Windows display scaling set to 100%
- [ ] All notifications silenced (Teams, Outlook, Windows notifications OFF)
- [ ] Demo file is open and on the Report--> landing page
- [ ] File is in a CLEAN state — no leftover macro outputs from previous runs:
  - [ ] No Variance Analysis sheet
  - [ ] No Variance Commentary sheet
  - [ ] No Data Quality Report sheet
  - [ ] No Executive Dashboard sheet
  - [ ] No YoY Variance Analysis sheet
  - [ ] No Sensitivity Analysis sheet
  - [ ] Dashboard/Charts sheet is blank or default state
  - [ ] Checks sheet is blank or default state
  - [ ] Version Control has no saved snapshots (or a clean starting state)
  - [ ] Audit Log is empty or has minimal entries
- [ ] Command Center is closed
- [ ] Script visible on second monitor or printed
- [ ] Audio test done — 30 seconds, listen back with headphones
- [ ] Screen recording running, 1920×1080, 30fps confirmed
- [ ] You have run through the ENTIRE demo sequence at least once today to confirm all macros execute without errors

**IMPORTANT:** Record each chapter as a separate clip. This lets you re-record any chapter without redoing the entire video. Leave 2 seconds of silence at the start and end of each clip.

---

## Title Card (5 seconds)

**On screen:**

```
iPipeline Finance Automation
Full Demo Walkthrough
```

**Audio:** Brief music sting (2–3 seconds), fade to silence
**Style:** iPipeline Blue (#0B4779) background, white text, Arial font

---

## Opening (Before Chapter 1)

**Duration:** 0:05–0:45 (40 seconds)
**On screen:** Report--> landing page

### Script:

> "Welcome to the full walkthrough of the iPipeline Finance Automation file.
>
> This is a single Excel workbook that automates the monthly P&L close process — from importing data, to running quality checks, to generating analysis, building dashboards, and producing final deliverables. Sixty-two actions, all accessible from one control panel.
>
> I'm going to walk you through how it works, chapter by chapter. Everything you see is running live — no slides, no mockups. Let's get into it."

### Screen Actions:
- File is open on Report--> page
- Slow scroll down the landing page as you narrate
- No clicking yet

### Production Notes:
- This opening sets expectations: it's live, it's real, it's organized by chapter
- Speak at a measured pace — this is a longer video, you don't need to rush
- The "sixty-two actions" line should be clear and confident, same energy as Video 1

---

## CHAPTER CARD: Chapter 1

**On screen (3 seconds):**

```
Chapter 1
The Workbook & Command Center
Your home base for everything
```

**Audio:** Brief music accent or silence

---

## Chapter 1: The Workbook & Command Center

**Duration:** 2:00
**On screen:** Navigating the workbook sheets and Command Center

### Script:

> "Let's start with what's inside this file.
>
> The landing page — Report — gives you a summary of the workbook and quick navigation to any section. Think of it as your table of contents.
>
> [click through a few sheet tabs at the bottom]
>
> The file has over a dozen sheets. You've got the main P&L Monthly Trend sheet — this is your core financial data, revenue and expenses by month, with a full-year total and budget column.
>
> [click to P&L - Monthly Trend, pause]
>
> There are Functional P&L Summary sheets — one for each month — that break things down by department.
>
> [click to one monthly tab, pause briefly]
>
> A Product Line Summary sheet showing revenue by product — iGO, Affirm, InsureSight, DocFast.
>
> [click to Product Line Summary, pause]
>
> An Assumptions sheet with the key financial drivers — growth rates, allocation percentages, revenue shares.
>
> [click to Assumptions, pause]
>
> And a General Ledger sheet with the raw transaction data.
>
> [click to General Ledger, pause]
>
> You don't need to memorize any of this — because everything runs from one place.
>
> [open Command Center with Ctrl+Shift+M]
>
> This is the Command Center. Every automated action in this file is listed here, organized by category. You can scroll through to browse, or use the search bar to find what you need.
>
> [scroll through categories slowly]
>
> Monthly Operations, Analysis and Reporting, Enterprise Features, Utilities — it's all here. Pick an action, click Run, and it handles the rest.
>
> [type 'reconciliation' in search bar, show filtered results, then clear]
>
> That's your home base. Every demo from here on out starts from this screen."

### Screen Actions (detailed):
1. Start on Report--> page
2. Click P&L - Monthly Trend tab — pause 2 seconds, let viewer see the data layout
3. Click one Functional P&L Summary tab (e.g., Jan) — pause 1 second
4. Click Product Line Summary tab — pause 2 seconds
5. Click Assumptions tab — pause 1 second
6. Click General Ledger tab — pause 1 second
7. Press Ctrl+Shift+M to open Command Center — pause 2 seconds after it opens
8. Slowly scroll through categories (5–6 seconds of scrolling)
9. Click search bar, type "reconciliation" — show results — clear search
10. Leave Command Center open for transition to Chapter 2

### Production Notes:
- Don't linger on any sheet too long. This is orientation, not analysis. 1–2 seconds per tab is enough.
- The Command Center opening is the payoff of this chapter. Give it room.
- When scrolling through categories, move slowly enough that a viewer could pause and read action names
- End with the Command Center open — it creates a natural bridge to Chapter 2

---

## CHAPTER CARD: Chapter 2

**On screen (3 seconds):**

```
Chapter 2
Data Import & Quality
Getting your data in — and making sure it's clean
```

---

## Chapter 2: Data Import & Quality

**Duration:** 2:30
**On screen:** Running GL Import, Data Quality Scan, and Reconciliation Checks

### Section 2A: GL Import

**Duration:** 0:50

### Script:

> "Before you can do any analysis, you need your data. The General Ledger Import pulls in GL data from a CSV or Excel file with format validation built in.
>
> [navigate to GL Import in Command Center, click Run]
>
> It reads the source file, validates the format, maps the columns, and loads the transactions into the workbook. If something doesn't match the expected structure, it tells you.
>
> [import completes — show the General Ledger sheet with data]
>
> What used to take around 45 minutes of manual copying, pasting, and reformatting — done in about 30 seconds."

### Screen Actions:
1. In Command Center, find "Import GL Data" (search or scroll)
2. Click Run
3. If a file dialog appears, navigate to the source file (have this ready in a known location)
4. Wait for import to complete
5. Navigate to General Ledger sheet to show the loaded data
6. Slow scroll through a few rows so viewer can see real transaction data

### Time Savings Overlay:
```
⏱ Manual: ~45 minutes → Automated: ~30 seconds
```

### Production Notes:
- Have the source file ready in an easy-to-find location so the file dialog doesn't fumble
- If the import runs fast enough that there's dead air, fill with: "It's validating the format as it goes"
- This is a setup feature, not a wow feature. Keep it brisk.

---

### Section 2B: Data Quality Scan

**Duration:** 0:50

### Script:

> "Now that the data is loaded, the first thing you want to know is — how clean is it?
>
> [navigate to Data Quality Scan in Command Center, click Run]
>
> The Data Quality Scan checks your entire workbook across six categories: completeness, accuracy, consistency, formatting, outliers, and cross-references.
>
> [scan completes — Data Quality Report sheet appears]
>
> Right at the top — a letter grade. [Read the grade]. That tells you at a glance whether your data is ready to work with.
>
> [scroll down slowly]
>
> Below that, each category gets its own score and detail. If there are issues, it tells you exactly where — which sheet, which column, what the problem is.
>
> This scan has never been done manually — there was no practical way to do it. Now it takes about fifteen seconds."

### Screen Actions:
1. Back to Command Center (Ctrl+Shift+M)
2. Find "Data Quality Scan" — click Run
3. Wait for completion — Data Quality Report sheet appears
4. Pause 2–3 seconds on the letter grade badge (28pt colored)
5. Slowly scroll down through the category breakdown
6. Hover cursor near specific findings if visible

### Time Savings Overlay:
```
⏱ Previously: Never done (no practical method) → Now: ~15 seconds
```

### Production Notes:
- The letter grade badge is a strong visual. Hold on it. Don't talk for 1–2 seconds while it's on screen.
- If specific issues are flagged, briefly mention one: "For example, it found [X] in [sheet]." This makes it feel real.
- This feature appeared in Video 1 as well — that's intentional. Repetition reinforces. But in Video 2, you go deeper into the category breakdown.

---

### Section 2C: Reconciliation Checks

**Duration:** 0:50

### Script:

> "Next step — make sure all the numbers tie out.
>
> [navigate to Reconciliation Checks in Command Center, click Run]
>
> The reconciliation engine runs a series of validation checks across every sheet — verifying that cross-sheet totals match, that revenue and expense lines balance, and that formulas are intact.
>
> [Checks sheet appears with PASS/FAIL results]
>
> Each check gets a clear PASS or FAIL. Green means it ties. Red means something needs attention.
>
> [scroll through the results]
>
> In this case — [describe what you see: e.g., 'all checks passing' or 'one item flagged']. Either way, you know exactly where you stand in ten seconds instead of two hours."

### Screen Actions:
1. Back to Command Center
2. Find "Run Reconciliation Checks" — click Run
3. Wait for completion — Checks sheet appears
4. Pause on the PASS/FAIL scorecard — the green/red visual is immediately readable
5. Slowly scroll through all checks
6. If any FAIL items exist, hover cursor near them

### Time Savings Overlay:
```
⏱ Manual: ~2 hours → Automated: ~10 seconds
```

### Production Notes:
- PASS/FAIL with color coding is visually satisfying. Let the viewer see the full list.
- Read the actual results — don't script specific outcomes since they depend on the demo data state
- The "two hours to ten seconds" comparison is one of your strongest. Let the overlay stay visible for 3–4 seconds.

---

## CHAPTER CARD: Chapter 3

**On screen (3 seconds):**

```
Chapter 3
Analysis
Making sense of your numbers
```

---

## Chapter 3: Analysis

**Duration:** 3:00
**On screen:** Variance Analysis, Variance Commentary, YoY Variance

### Section 3A: Variance Analysis

**Duration:** 1:00

### Script:

> "Your data is in, it's clean, and it reconciles. Now — what's actually happening in the numbers?
>
> [navigate to Variance Analysis in Command Center, click Run]
>
> The Variance Analysis compares each line item month over month and flags anything that moved more than fifteen percent. Revenue, expenses, margins — it checks everything.
>
> [Variance Analysis sheet appears]
>
> Items over the threshold are highlighted automatically. You can see the dollar change, the percentage change, and whether it's favorable or unfavorable. For expense items, the favorable/unfavorable logic is automatically reversed — a decrease in costs is flagged as favorable, not unfavorable.
>
> [scroll through, pausing on highlighted items]
>
> Instead of scanning hundreds of rows yourself, you get a filtered view of what actually needs your attention."

### Screen Actions:
1. Command Center → find "Variance Analysis" → click Run
2. Wait for the Variance Analysis sheet to appear
3. Pause on the header row — let viewer see the column structure
4. Scroll through slowly, pausing on highlighted/flagged items
5. Hover cursor near a flagged item to draw attention

### Production Notes:
- The highlighted items are the visual hook here. Make sure they're visible.
- The cost-line reversal is a subtle but important detail for Finance people. Mention it but don't over-explain.
- Keep the energy up — you're building toward the Variance Commentary, which is the payoff.

---

### Section 3B: Variance Commentary

**Duration:** 1:00

### Script:

> "This is one of the features I'm most excited about.
>
> You've got your flagged variances. Now the system can write the commentary for you.
>
> [navigate to Variance Commentary in Command Center, click Run]
>
> [Variance Commentary sheet appears with written narratives]
>
> These are plain English narratives for the top five variances. Each one identifies the line item, states the dollar and percentage change, and describes what happened — in complete sentences, ready to paste into an email, a report, or a board deck.
>
> [slowly scroll through the narratives — give the viewer time to read]
>
> [pause for 2–3 seconds of silence while text is visible]
>
> Writing these manually — pulling the numbers, doing the comparison, putting it into words — that's typically an hour of work. This takes about five seconds."

### Screen Actions:
1. Command Center → find "Variance Commentary" → click Run
2. Wait for the Variance Commentary sheet to appear
3. Pause for 2–3 seconds — let the viewer take in the full page before narrating
4. Slowly scroll through each narrative (there should be ~5)
5. Hover cursor near one narrative to draw the eye
6. Pause again after reading "five seconds" — let the impact sit

### Time Savings Overlay:
```
⏱ Manual: ~1 hour → Automated: ~5 seconds
```

### Production Notes:
- THIS IS YOUR JAW-DROP MOMENT in this video. Treat it accordingly.
- The 2–3 seconds of silence while narratives are on screen is critical. Resist the urge to keep talking. Let people read.
- After "five seconds" — hold for a beat. Don't immediately rush to the next feature.
- If the generated narratives look particularly good, consider reading one aloud: "For example — [read one sentence of a narrative]." This makes it even more concrete.

---

### Section 3C: YoY Variance

**Duration:** 1:00

### Script:

> "Variance Analysis gives you month over month. But leadership often wants year over year — how does this year compare to last year, and how are we tracking against budget?
>
> [navigate to YoY Variance in Command Center, click Run]
>
> [YoY Variance Analysis sheet appears]
>
> This builds a full Year-over-Year comparison. Full-year total versus prior year, full-year total versus budget, with dollar and percentage variances for every line.
>
> [scroll through the sheet]
>
> Same idea — items beyond the threshold are flagged. You get a complete picture of where you're ahead, where you're behind, and by how much.
>
> What would normally take a couple of hours of pulling data from two different periods and building the comparison — done in about ten seconds."

### Screen Actions:
1. Command Center → find "YoY Variance" → click Run
2. Wait for YoY Variance Analysis sheet to appear
3. Pause on the header — let viewer see the column structure (FY Total, Prior Year, Budget, $ Variance, % Variance)
4. Scroll through slowly, pausing on flagged items
5. Brief pause at the end

### Production Notes:
- This feature is important for the Finance audience but less dramatic than Variance Commentary. Keep it solid but don't over-dwell.
- Emphasize the "leadership wants year over year" framing — it connects the feature to a real request they get regularly
- Smooth transition to Chapter 4

---

## CHAPTER CARD: Chapter 4

**On screen (3 seconds):**

```
Chapter 4
Reporting & Visuals
Turning analysis into deliverables
```

---

## Chapter 4: Reporting & Visuals

**Duration:** 2:30
**On screen:** Dashboard Charts, Executive Dashboard, PDF Export

### Section 4A: Dashboard Charts

**Duration:** 0:50

### Script:

> "You've done the analysis. Now you need to present it.
>
> [navigate to Build Dashboard in Command Center, click Run]
>
> [Charts & Visuals sheet appears with 8 charts]
>
> One click builds eight branded charts in a grid layout — revenue trends, expense breakdowns, margin analysis, product mix, and more. All formatted in iPipeline colors, all properly labeled.
>
> [slowly scroll through the chart grid]
>
> These are the visuals you'd normally build one at a time in a separate PowerPoint or chart tool. Here, they're generated directly from your data in about fifteen seconds."

### Screen Actions:
1. Command Center → find "Build Dashboard" → click Run
2. Wait for the Charts & Visuals sheet to appear
3. Slowly scroll through the chart grid — pause briefly on each chart (1–2 seconds per chart)
4. Let the branded colors and formatting speak for themselves

### Production Notes:
- The chart grid is visually impressive at a glance. Give the viewer a full-screen moment when it first appears.
- Don't describe every chart in detail. "Revenue trends, expense breakdowns, margin analysis, product mix, and more" covers it. The visual does the work.
- Mention "iPipeline colors" to signal this is branded and presentation-ready, not generic Excel charts.

---

### Section 4B: Executive Dashboard

**Duration:** 0:50

### Script:

> "For a more focused leadership view, there's the Executive Dashboard.
>
> [navigate to Executive Dashboard in Command Center, click Run]
>
> [Executive Dashboard sheet appears]
>
> This puts everything on one sheet — KPI summary cards across the top, a waterfall chart showing how you get from budget to actual, and a product line comparison at the bottom.
>
> [scroll slowly from KPI cards → waterfall → product comparison]
>
> This is designed to be the one sheet you pull up when the CFO asks 'how are we doing this month.' One click, one sheet, full picture."

### Screen Actions:
1. Command Center → find "Executive Dashboard" → click Run
2. Wait for Executive Dashboard sheet to appear
3. Pause at the top — KPI cards visible — hold 2 seconds
4. Scroll down to waterfall chart — hold 2 seconds
5. Scroll down to product comparison — hold 2 seconds

### Production Notes:
- "The one sheet you pull up when the CFO asks how are we doing" — this is a relatable hook for your Finance audience. They've all been in that moment.
- The KPI cards, waterfall, and product comparison are three distinct visual elements. Give each one a moment.
- Don't say "the CFO" in a way that feels like name-dropping. It's a scenario, not a reference to your actual CFO.

---

### Section 4C: PDF Export

**Duration:** 0:50

### Script:

> "When you need a final deliverable — something you can email, save to a shared drive, or print — the PDF Export handles it.
>
> [navigate to PDF Export in Command Center, click Run]
>
> [export runs — PDF is generated]
>
> It takes seven key sheets from the workbook and compiles them into a single, clean PDF. Each page has proper headers and footers — the report title, the date, page numbers. Formatted for printing or sharing.
>
> [open the generated PDF briefly if possible, or show the output location]
>
> Manually formatting and exporting seven sheets to a clean PDF — that's easily thirty minutes of adjusting print areas, fixing page breaks, and hoping nothing shifts. This takes about ten seconds."

### Screen Actions:
1. Command Center → find "PDF Export" → click Run
2. Wait for the export to complete
3. If the PDF opens automatically, pause on the first page — let viewer see the formatting
4. Scroll to page 2 to show headers/footers and clean formatting
5. If it doesn't auto-open, navigate to the output file and open it briefly

### Time Savings Overlay:
```
⏱ Manual: ~30 minutes → Automated: ~10 seconds
```

### Production Notes:
- The PDF is the tangible deliverable — the thing someone actually sends. That makes this feature feel real and practical.
- If the PDF looks crisp and professional on screen, hold on it for a few seconds. It sells itself.
- "Hoping nothing shifts" — small moment of relatability for anyone who's fought with Excel print formatting.

---

## CHAPTER CARD: Chapter 5

**On screen (3 seconds):**

```
Chapter 5
Enterprise Features
Power tools for control and flexibility
```

---

## Chapter 5: Enterprise Features

**Duration:** 2:30
**On screen:** Executive Mode, Version Control, Scenario Management, Sensitivity Analysis

### Section 5A: Executive Mode

**Duration:** 0:30

### Script:

> "When leadership needs to review the file, they don't need to see every technical sheet. Executive Mode cleans it up.
>
> [navigate to Executive Mode in Command Center — or use Ctrl+Shift+R — click Run/toggle]
>
> One click hides all the working sheets and leaves only the presentation-ready views. Toggle it off and everything comes back.
>
> [toggle off to show sheets returning]
>
> Simple, but it makes a big difference when you're sharing the file with someone who just wants the highlights."

### Screen Actions:
1. Show the full tab bar at the bottom — many sheet tabs visible
2. Toggle Executive Mode ON — watch tabs disappear, leaving only key sheets
3. Pause 2 seconds — let viewer see the clean state
4. Toggle Executive Mode OFF — tabs return
5. Brief pause

### Production Notes:
- This is a quick hit — don't over-explain. The visual toggle is self-explanatory.
- The disappearing/reappearing tabs is a satisfying visual. Let it happen without talking over it.
- 30 seconds max. In, out, move on.

---

### Section 5B: Version Control

**Duration:** 0:40

### Script:

> "Version Control lets you save a snapshot of the entire workbook at any point — and compare or restore it later.
>
> [navigate to Version Control — save a snapshot]
>
> You give it a name — 'Pre-Close' or 'March Draft 1' — and it saves the full state. If something goes wrong, or if someone overwrites your work, you go back to Version Control, pick a snapshot, and restore it.
>
> [show the save confirmation or snapshot list]
>
> Every snapshot is timestamped and logged. You always know what changed and when."

### Screen Actions:
1. Command Center → find "Version Control" area
2. Run Save Snapshot — enter a name when prompted (e.g., "March Draft 1")
3. Show the confirmation or the snapshot list
4. Briefly show the Restore option (don't actually restore — just show it exists)

### Production Notes:
- The use case "someone overwrites your work" is universally relatable. It earns a knowing nod.
- Don't actually demonstrate a restore — it would require setup and adds time. Just showing that the option exists is enough.
- Keep it tight. 40 seconds.

---

### Section 5C: Scenario Management

**Duration:** 0:30

### Script:

> "Scenario Management lets you save and load different sets of assumptions. You can set up a Base Case, an Optimistic case, a Conservative case — each with different growth rates, allocation percentages, whatever drivers matter.
>
> [show Scenario Management — save or load a scenario]
>
> Switch between them with one click and the entire workbook recalculates. You can also compare scenarios side by side to see the impact of different assumptions."

### Screen Actions:
1. Command Center → find "Scenario Management"
2. Show the scenario list (if scenarios are pre-saved, show Base Case, Optimistic, etc.)
3. Load one scenario — show that it updates the Assumptions sheet
4. Brief pause

### Production Notes:
- This is a powerful feature but hard to demo visually in 30 seconds. Focus on the concept and the one-click switch.
- If you can show the Assumptions sheet values change when you load a different scenario, that's the visual proof. Even a few cells changing is enough.

---

### Section 5D: Sensitivity Analysis

**Duration:** 0:50

### Script:

> "Sensitivity Analysis takes this a step further. Instead of switching between preset scenarios, you can run what-if analysis on any key assumption and see how it ripples through the entire P&L.
>
> [navigate to Sensitivity Analysis in Command Center, click Run]
>
> [Sensitivity Analysis sheet appears]
>
> What happens to total revenue if growth is two percent higher? What happens to margins if our allocation changes by five points? This answers those questions instantly.
>
> [scroll through the results]
>
> Doing this manually — changing an assumption, recalculating, recording the result, changing it back, trying the next one — that's four or more hours of tedious work. This runs all of them in about twenty seconds."

### Screen Actions:
1. Command Center → find "Sensitivity Analysis" → click Run
2. Wait for the Sensitivity Analysis sheet to appear
3. Pause on the output — let viewer see the structure
4. Scroll through the results slowly
5. Hover near key data points

### Time Savings Overlay:
```
⏱ Manual: 4+ hours → Automated: ~20 seconds
```

### Production Notes:
- Frame the what-if questions conversationally: "What happens if..." — this makes it relatable to how leadership actually asks these questions
- The four-hours-to-twenty-seconds comparison is dramatic. Let the overlay stay visible.
- This is the last feature demo before the "Under the Hood" chapter. End with energy.

---

## CHAPTER CARD: Chapter 6

**On screen (3 seconds):**

```
Chapter 6
Under the Hood
Built to be trusted
```

---

## Chapter 6: Under the Hood

**Duration:** 1:30
**On screen:** Integration Test, Audit Log

### Section 6A: Integration Test

**Duration:** 0:50

### Script:

> "With this many automated actions, you need to know the system is working correctly. The Integration Test runs eighteen automated checks across the entire workbook — sheet existence, data integrity, formula health, macro functionality.
>
> [navigate to Integration Test in Command Center, click Run]
>
> [test runs — results appear]
>
> Eighteen out of eighteen — all passing.
>
> [pause on the results]
>
> This runs every time you want to verify the file is in a good state. Before a close, after making changes, anytime you want peace of mind — one click."

### Screen Actions:
1. Command Center → find "Integration Test" → click Run
2. Wait for test to complete — results appear
3. Pause on the 18/18 PASS result — hold 3 seconds
4. If results are listed individually, slowly scroll through them

### Production Notes:
- 18/18 PASS is a confidence builder. It says "this isn't fragile, it's tested."
- Don't explain what each individual test does — "sheet existence, data integrity, formula health, macro functionality" covers the categories. The 18/18 result is what matters.
- The phrase "peace of mind" resonates with anyone who's worked with complex spreadsheets.

---

### Section 6B: Audit Log

**Duration:** 0:40

### Script:

> "Every action you run is logged automatically.
>
> [navigate to the Audit Log — this may be a hidden sheet, so show how to access it]
>
> The Audit Log records a timestamp, the module that ran, and the result for every single action. If you need to know who ran what, and when — it's all here.
>
> [scroll through the log showing the entries from today's demo]
>
> You'll see entries from everything we just did — every import, every scan, every export. Full traceability."

### Screen Actions:
1. Navigate to the Audit Log sheet (unhide if hidden, or access through Command Center)
2. Show the log entries — they should include timestamps from the demo you just ran
3. Scroll through slowly, hovering near a few entries
4. Brief pause

### Production Notes:
- The audit log filled with entries from this very demo is a nice moment — "you'll see entries from everything we just did" makes it tangible
- Don't spend too long here. The point is: it exists, it's automatic, it's complete. 40 seconds.
- If the Audit Log is on a hidden sheet, briefly show the unhide action — that itself is a feature (hidden from casual users, available when needed)

---

## CHAPTER CARD: Chapter 7

**On screen (3 seconds):**

```
Chapter 7
Next Steps
Where to go from here
```

---

## Chapter 7: Closing & Next Steps

**Duration:** 1:00
**On screen:** Report--> landing page, then closing card

### Script:

> "That's the full walkthrough. Let me recap what we covered.
>
> We imported GL data, checked data quality, ran reconciliation, analyzed variances month over month and year over year, generated written commentary, built dashboards and an executive view, exported a clean PDF, managed scenarios and sensitivity analysis, ran a full integration test, and reviewed the audit trail.
>
> All from one Excel file, all through the Command Center, all in a matter of minutes.
>
> [brief pause]
>
> If you want to explore the file yourself, it's available on SharePoint along with step-by-step training guides for everything you just saw. If you run into any questions, reach out — I'm happy to help.
>
> Thanks for watching."

### Screen Actions:
- Navigate back to the Report--> landing page for the recap section
- During the recap list, optionally show a slow scroll or keep the landing page static — don't click around
- On "it's available on SharePoint" — cut to closing card

### Production Notes:
- The recap should be spoken at a slightly faster pace than the rest of the video — it's a summary, not a re-explanation
- Don't re-demo anything during the recap. Just list it.
- "Reach out — I'm happy to help" is the right CTA for this audience. They're the Finance team. They'll have questions. Make it easy.
- End confidently. No "so yeah, that's it" energy. Clear, professional close.

---

## Closing Title Card (5 seconds)

**On screen:**

```
iPipeline Finance Automation

Demo File & Training Guides
Available on SharePoint
[SharePoint location]

Questions? Contact Connor [last name / email]
```

**Audio:** Brief music, fade out

---

## Total Runtime Breakdown

| Section | Duration | Cumulative |
|---------|----------|------------|
| Title Card | 0:05 | 0:05 |
| Opening | 0:40 | 0:45 |
| Ch 1 Card | 0:03 | 0:48 |
| Ch 1: Workbook & Command Center | 2:00 | 2:48 |
| Ch 2 Card | 0:03 | 2:51 |
| Ch 2A: GL Import | 0:50 | 3:41 |
| Ch 2B: Data Quality Scan | 0:50 | 4:31 |
| Ch 2C: Reconciliation Checks | 0:50 | 5:21 |
| Ch 3 Card | 0:03 | 5:24 |
| Ch 3A: Variance Analysis | 1:00 | 6:24 |
| Ch 3B: Variance Commentary | 1:00 | 7:24 |
| Ch 3C: YoY Variance | 1:00 | 8:24 |
| Ch 4 Card | 0:03 | 8:27 |
| Ch 4A: Dashboard Charts | 0:50 | 9:17 |
| Ch 4B: Executive Dashboard | 0:50 | 10:07 |
| Ch 4C: PDF Export | 0:50 | 10:57 |
| Ch 5 Card | 0:03 | 11:00 |
| Ch 5A: Executive Mode | 0:30 | 11:30 |
| Ch 5B: Version Control | 0:40 | 12:10 |
| Ch 5C: Scenario Management | 0:30 | 12:40 |
| Ch 5D: Sensitivity Analysis | 0:50 | 13:30 |
| Ch 6 Card | 0:03 | 13:33 |
| Ch 6A: Integration Test | 0:50 | 14:23 |
| Ch 6B: Audit Log | 0:40 | 15:03 |
| Ch 7 Card | 0:03 | 15:06 |
| Ch 7: Closing | 1:00 | 16:06 |
| Closing Card | 0:05 | 16:11 |
| **TOTAL** | | **~16:11** |

Buffer for natural pauses, macro run times, and breathing room: expect **16:30–18:00** in practice.

---

## Recording Tips Specific to This Video

1. **Record each chapter as a separate clip.** This is essential for a 16+ minute video. Don't try to do it in one take.

2. **Reset the file between chapter takes if needed.** If you need to re-record Chapter 3, make sure the Chapter 2 outputs are still present (since Chapter 3 builds on them) but Chapter 3's outputs are cleared.

3. **Watch your energy level.** Sixteen minutes is a long narration. If you feel your energy dropping in the later chapters, take a break and come back. Record Chapter 6 tomorrow if needed — nobody will know.

4. **The Chapter Cards serve as natural break points.** In editing, you'll stitch clips together at these cards. They hide any discontinuity.

5. **Keep a consistent mouse style throughout.** Smooth, deliberate movements. Same speed. If you're a fast clicker in Chapter 1 and slow in Chapter 6, it'll feel inconsistent.

6. **Note the cumulative runtime as you edit.** If the video is trending over 18 minutes, look for sections where you paused too long or narrated too slowly. Tighten those first before cutting content.

7. **Remember: you can always do a second take of just one chapter.** That's the whole point of recording in sections. Don't settle for a mediocre Chapter 4 because Chapters 1–3 were great.

---

*Script created: 2026-03-06 | Part of Video Demo Master Plan*


================================================================================
================================================================================
# SECTION 4: VIDEO 3 SCRIPT
================================================================================
================================================================================

# Video 3 Script — "Universal Tools"

**Runtime Target:** 8:00–10:00
**Format:** Screen recording, no webcam, voice-over narration, chaptered
**Audience:** Anyone at iPipeline who uses Excel and wants to automate their own work
**Purpose:** Show examples of the universal tools running on a plain sample file, then point people to SharePoint
**Key principle:** Uses a separate, simple sample Excel file — NOT the demo file
**Music:** Subtle corporate/tech track during title card, chapter cards, and closing card only

---

## Pre-Recording Checklist

Before hitting record, make sure:

- [ ] Excel is the only application open
- [ ] Desktop is clean — no icons, taskbar auto-hidden
- [ ] Excel is maximized to full screen
- [ ] Zoom level set to 100% or 110% (SAME as Videos 1 and 2 — consistency across all three)
- [ ] Windows display scaling set to 100%
- [ ] All notifications silenced
- [ ] Sample file (Sample_Quarterly_Report.xlsx) is open and ready
- [ ] The sample file has some intentional "mess" baked in:
  - [ ] A few merged cells in column A
  - [ ] Some text-stored numbers (numbers formatted as text)
  - [ ] A few blank rows scattered in the data
  - [ ] Some extra spaces in text cells
  - [ ] At least one column of dates in mixed formats
  - [ ] A few error values (#N/A, #REF!) in formulas
  - [ ] At least one hidden sheet
  - [ ] Some unstyled headers (no formatting)
- [ ] Script visible on second monitor or printed
- [ ] Audio test done
- [ ] Screen recording running, 1920×1080, 30fps

**IMPORTANT:** The sample file should look generic and relatable — something any iPipeline employee might have on their desktop. Name it something like "Sample_Quarterly_Report.xlsx" or "Team_Data_Export.xlsx". Use realistic-looking but fictional data (employee names, departments, amounts, dates).

---

## Title Card (5 seconds)

**On screen:**

```
iPipeline Finance Automation
Universal Tools — For Any Excel File
```

**Audio:** Brief music sting (2–3 seconds), fade to silence
**Style:** iPipeline Blue (#0B4779) background, white text, Arial font

---

## Opening

**Duration:** 0:05–0:50 (45 seconds)
**On screen:** The sample file is open, showing a typical messy spreadsheet

### Script:

> "The P&L demo file showed what automation can do for a specific workflow. But the code library behind it includes over 75 VBA tools and 22 Python scripts that work on any Excel file — not just that one.
>
> In this video, I'm going to show you a handful of those tools running on a regular spreadsheet. Nothing special about this file — it's just a typical data export with some common problems. Messy formatting, text stored as numbers, blank rows, inconsistent dates — the kind of thing you deal with every day.
>
> Every tool you see here is available on SharePoint. Grab what you need, use the step-by-step guides, and if you get stuck, there are pre-built Copilot prompts to help."

### Screen Actions:
- File is open showing the sample data
- Slow scroll through the file so the viewer can see the "mess" — blank rows, mixed formatting, errors visible
- No clicking yet — just visual context

### Production Notes:
- "Over 75 VBA tools and 22 Python scripts" — say this clearly. The scale is impressive.
- The scroll should make the viewer feel the familiar pain of a messy spreadsheet. Don't rush it.
- Keep this under 45 seconds. You're setting context, not explaining.

---

## CHAPTER CARD: Chapter 1

**On screen (3 seconds):**

```
Chapter 1
Data Cleanup
Fixing the mess in seconds
```

---

## Chapter 1: Data Cleanup

**Duration:** 2:30
**These tools come from: modUTL_DataCleaning, modUTL_DataCleaningPlus, modUTL_DataSanitizer**

This chapter shows 4 cleanup tools, each one solving a specific, common Excel problem.

### Demo 1A: Delete Blank Rows

**Duration:** 0:30

### Script:

> "First problem — blank rows scattered through the data. Every filter, every formula, every sort breaks when you have random empty rows in the middle of your table.
>
> [run Delete Blank Rows]
>
> One click — it finds and removes every completely empty row. It creates a backup copy of the sheet first, just in case. [Read the confirmation: X blank rows deleted.]"

### Screen Actions:
1. Show the data with visible blank rows (scroll briefly to point them out)
2. Run DeleteBlankRows from the macro menu (Alt+F8 or developer tab)
3. Confirm the dialog
4. Show the result — clean data, no gaps

### Production Notes:
- The backup-before-destructive-action detail builds trust. Mention it.
- Quick and snappy. This is an appetizer.

---

### Demo 1B: Remove Leading/Trailing Spaces

**Duration:** 0:30

### Script:

> "Next — invisible spaces hiding in your text cells. You can't see them, but they break VLOOKUP matches, they mess up filters, and they make duplicates that aren't really duplicates.
>
> [select a column, run Remove Leading/Trailing Spaces]
>
> [Read confirmation: X cells cleaned.]
>
> Those phantom spaces are gone. Every VLOOKUP that was failing because of a trailing space — fixed."

### Screen Actions:
1. Select a column of text data
2. Run RemoveLeadingTrailingSpaces
3. Show the confirmation message

### Production Notes:
- "Breaks VLOOKUP matches" — this hits home for anyone who's fought with VLOOKUP. Name the pain specifically.

---

### Demo 1C: Convert Text to Numbers

**Duration:** 0:30

### Script:

> "One of the most common data import problems — numbers stored as text. The cell shows 1,250 but it's actually a text string. Your SUM formula returns zero. Your chart is blank.
>
> [select a range, run ConvertTextToNumbers]
>
> [Read confirmation: X text-stored numbers converted.]
>
> Now they're real numbers. SUM works. Charts work. Everything downstream works."

### Screen Actions:
1. Show a column where numbers have the green triangle (text-stored number indicator) or demonstrate that SUM returns 0
2. Select the range
3. Run ConvertTextToNumbers
4. Show the confirmation

### Production Notes:
- If you can show a SUM formula going from 0 to the correct total after conversion, that's a powerful visual. Consider setting this up in the sample file.

---

### Demo 1D: Unmerge Cells & Fill Down

**Duration:** 0:30

### Script:

> "Merged cells. Every data person's least favorite thing. They break sorting, filtering, copy-paste — essentially everything.
>
> [select a range with merged cells, run Unmerge And Fill Down]
>
> It unmerges every cell in the selection and fills the value down into the blanks that are left behind. Your data is flat and filterable now."

### Screen Actions:
1. Show the merged cells in column A (department names merged across rows)
2. Select the range
3. Run UnmergeAndFillDown
4. Show the result — flat, clean data with values filled down

### Production Notes:
- "Every data person's least favorite thing" — a moment of shared frustration. The audience will nod.

---

## CHAPTER CARD: Chapter 2

**On screen (3 seconds):**

```
Chapter 2
Formatting & Standardization
Making every file look professional
```

---

## Chapter 2: Formatting & Standardization

**Duration:** 2:00
**These tools come from: modUTL_Formatting, modUTL_Branding**

### Demo 2A: AutoFit All Columns & Rows

**Duration:** 0:20

### Script:

> "Quick one — AutoFit across the entire workbook. Every column, every row, every sheet — properly sized in one click.
>
> [run AutoFit All Columns & Rows — choose "Yes" for all sheets]
>
> Done. No more scrolling sideways to read truncated headers."

### Screen Actions:
1. Show some columns that are too narrow or too wide
2. Run AutoFitAllColumnsRows
3. Show the immediate visual improvement

### Production Notes:
- Fast. 20 seconds max. The visual before/after speaks for itself.

---

### Demo 2B: Apply iPipeline Branding

**Duration:** 0:40

### Script:

> "This one is my favorite for making files look professional fast.
>
> [run Apply iPipeline Branding]
>
> It automatically detects your header row, applies the official iPipeline Blue background with white text, sets alternating row colors for readability, and styles any total or summary rows in Navy Blue. All in official brand fonts and colors.
>
> [scroll through the formatted sheet]
>
> Five seconds ago this was a plain spreadsheet. Now it looks like it came from the corporate template library."

### Screen Actions:
1. Show the sheet with plain, unstyled data (default Excel look)
2. Run ApplyiPipelineBranding
3. Confirm the dialog
4. Pause on the result — the transformation should be visually dramatic
5. Slowly scroll through to show headers, alternating rows, and total rows

### Production Notes:
- THIS IS YOUR VISUAL WOW MOMENT in Video 3. The before/after contrast of plain Excel → branded professional table is instantly impressive.
- Pause for 2–3 seconds after the branding applies. Let the visual register.
- "Came from the corporate template library" — that's the reaction you want from viewers.

---

### Demo 2C: Date Format Standardizer

**Duration:** 0:30

### Script:

> "Mixed date formats — some cells show MM/DD/YYYY, others show DD-MMM-YY, and a few are just serial numbers from an import. It's a mess.
>
> [run DateFormatStandardizer]
>
> [Read confirmation: X date cells standardized to MM/DD/YYYY.]
>
> Every date in the workbook is now in the same format. No more guessing whether 03/05 is March 5th or May 3rd."

### Screen Actions:
1. Show a column with visibly mixed date formats
2. Run DateFormatStandardizer
3. Show the confirmation
4. Scroll through the column — all dates now consistent

---

### Demo 2D: Highlight Negatives in Red

**Duration:** 0:20

### Script:

> "Standard Finance formatting — every negative number should be red and bold. One click applies this across every sheet in the workbook.
>
> [run Highlight Negatives Red — choose "Yes" for all sheets]
>
> Instant visual scan — you can see where the losses and shortfalls are without reading a single number."

### Screen Actions:
1. Run HighlightNegativesRed
2. Show the result — negative numbers now jump out in red

---

## CHAPTER CARD: Chapter 3

**On screen (3 seconds):**

```
Chapter 3
Audit & Investigation
Finding problems before they find you
```

---

## Chapter 3: Audit & Investigation

**Duration:** 1:30
**These tools come from: modUTL_Audit, modUTL_AuditPlus, modUTL_WorkbookMgmt**

### Demo 3A: Workbook Health Check

**Duration:** 0:40

### Script:

> "Before you start working with any file you've received, run the Workbook Health Check.
>
> [run WorkbookHealthCheck]
>
> It scans the entire workbook and gives you a diagnostic report — how many sheets, how many formulas, how many errors, how many external links, how many blank cells. If anything needs attention, it flags it.
>
> [read a few key lines from the report]
>
> Think of it as a checkup for your spreadsheet. Ten seconds and you know exactly what you're working with."

### Screen Actions:
1. Run WorkbookHealthCheck
2. Read the message box report — pause so the viewer can see the stats
3. Brief pause at the end

### Production Notes:
- This is a credibility builder. It says "these tools are thorough and professional."

---

### Demo 3B: External Link Finder

**Duration:** 0:25

### Script:

> "If a file has formulas pointing to other workbooks — external links — you want to know about it before those links break.
>
> [run ExternalLinkFinder]
>
> It creates a report listing every cell that references an external file, with the exact sheet, cell address, and linked file path. If there are none, it tells you the workbook is self-contained."

### Screen Actions:
1. Run ExternalLinkFinder
2. If links are found, show the report sheet. If not, show the clean message.

---

### Demo 3C: Unhide All Sheets, Rows & Columns

**Duration:** 0:25

### Script:

> "Ever received a file and suspected there were hidden sheets or rows? One click reveals everything.
>
> [run UnhideAllSheetsRowsColumns]
>
> [Read confirmation: X hidden sheets revealed, all hidden rows and columns shown.]
>
> No more right-clicking and unhiding one sheet at a time."

### Screen Actions:
1. Show the tab bar — the sample file should have at least one hidden sheet
2. Run UnhideAllSheetsRowsColumns
3. Show the hidden sheet appearing in the tab bar

---

## CHAPTER CARD: Chapter 4

**On screen (3 seconds):**

```
Chapter 4
Python & SQL Power Tools
Command-line tools for bigger jobs
```

---

## Chapter 4: Python & SQL Power Tools

**Duration:** 1:30
**These tools come from the 22 Python scripts**

### Script (intro — 15 seconds):

> "Beyond VBA, the library includes 22 Python scripts for heavier-duty work — data consolidation, bank reconciliation, fuzzy matching, PDF extraction, even running SQL queries against your Excel files.
>
> These are command-line tools, but you don't need to be a programmer. Each one has a step-by-step guide, and the Copilot prompts can walk you through it. Let me show you what a couple of these look like."

### Demo 4A: Universal Data Cleaner (clean_data.py)

**Duration:** 0:30

### Script:

> "The Python Data Cleaner does in one command what would take five or six VBA macros — it removes empty rows and columns, trims spaces, converts text-stored numbers, standardizes dates, removes duplicates, and gives you a before-and-after summary.
>
> [show the terminal command and output]
>
> You point it at a file, it creates a cleaned copy, and it tells you exactly what it changed. No manual steps."

### Screen Actions:
1. Open a command prompt or terminal alongside Excel
2. Run: `python clean_data.py "Sample_Quarterly_Report.xlsx"`
3. Show the output summary (rows removed, cells cleaned, etc.)
4. Briefly open the cleaned output file to show the result

### Production Notes:
- The terminal might feel intimidating to non-technical viewers. Keep the narration reassuring: "one command, one file path, done."
- The step-by-step guide covers exactly how to open a terminal and type this command. Mention that.

---

### Demo 4B: SQL Query Tool (sql_query_tool.py)

**Duration:** 0:30

### Script:

> "This one is especially powerful — it lets you run SQL queries directly on your Excel or CSV files. No database needed. Your spreadsheet becomes a queryable table.
>
> [show a SQL query running against the sample file]
>
> Filter, aggregate, join two files together — anything SQL can do, you can do on your spreadsheets. If you know SQL, this will change how you work with Excel data."

### Screen Actions:
1. In the terminal, run a SQL query against the sample file
2. Example: `python sql_query_tool.py "Sample_Quarterly_Report.xlsx" --query "SELECT Department, SUM(Amount) FROM data GROUP BY Department ORDER BY SUM(Amount) DESC"`
3. Show the output

### Production Notes:
- This is a power-user feature. Not everyone will use it — but the people who do will love it.
- "If you know SQL, this will change how you work with Excel data" — that's a strong statement for the right audience. Let it land.

---

### Brief mention (15 seconds — narration only, no live demo):

> "The library also includes tools for bank reconciliation with fuzzy matching, PDF table extraction, budget vs. actual consolidation, forecast roll-forwards, variance decomposition, and more. Twenty-two scripts in total. Every one documented, every one available on SharePoint."

### Production Notes:
- This is a verbal inventory, not a demo. You're painting the picture of scale.
- Show the SharePoint folder or a list of script names on screen while narrating this.

---

## Chapter 5: Closing + Call to Action

**Duration:** 0:45
**On screen:** Sample file (now cleaned and formatted) or closing card

### Script:

> "That's a sample of what's in the universal tools library. We showed maybe ten percent of what's available.
>
> Here's how to get started:
>
> Everything is on SharePoint — the VBA code, the Python scripts, the SQL tools, and the sample files. Each tool has a step-by-step guide written for someone who's never touched code before. And if you need help, there's a master document with pre-built Copilot prompts that will walk you through using any tool, step by step.
>
> Start with the guide. Pick one tool that solves a problem you deal with every week. Try it. And if you have questions or ideas for new tools, reach out — I'm happy to help.
>
> Thanks for watching."

### Screen Actions:
- During "Everything is on SharePoint" — cut to closing card showing:

```
SharePoint: [folder path]

VBA Tools (75+) | Python Scripts (22) | SQL Tools
Step-by-Step Guides | Copilot Prompts

Questions? Contact Connor [last name / email]
```

### Production Notes:
- "Pick one tool that solves a problem you deal with every week" — this is a great CTA because it's specific and low-commitment. One tool. One problem. Try it.
- The closing card should stay on screen for 5+ seconds
- Brief music sting on the closing card

---

## Closing Title Card (5 seconds)

**On screen:**

```
iPipeline Finance Automation

Universal Tools — For Any Excel File
Available on SharePoint
[SharePoint location]

Questions? Contact Connor [last name / email]
```

**Audio:** Brief music, fade out

---

## Total Runtime Breakdown

| Section | Duration | Cumulative |
|---------|----------|------------|
| Title Card | 0:05 | 0:05 |
| Opening | 0:45 | 0:50 |
| Ch 1 Card | 0:03 | 0:53 |
| Ch 1: Data Cleanup (4 demos) | 2:30 | 3:23 |
| Ch 2 Card | 0:03 | 3:26 |
| Ch 2: Formatting (4 demos) | 2:00 | 5:26 |
| Ch 3 Card | 0:03 | 5:29 |
| Ch 3: Audit (3 demos) | 1:30 | 6:59 |
| Ch 4 Card | 0:03 | 7:02 |
| Ch 4: Python/SQL (2 demos + verbal list) | 1:30 | 8:32 |
| Closing + CTA | 0:45 | 9:17 |
| Closing Card | 0:05 | 9:22 |
| **TOTAL** | | **~9:22** |

Buffer for natural pauses and command execution time: expect **9:30–10:30** in practice.

---

## Tools Selected for Demo (and Why)

### Chapter 1 — Data Cleanup (problems everyone has):
| Tool | Source Module | Why This One |
|------|-------------|--------------|
| Delete Blank Rows | modUTL_DataCleaning | Universal pain point, visual before/after |
| Remove Spaces | modUTL_DataCleaning | Invisible problem with visible consequences (broken VLOOKUPs) |
| Text to Numbers | modUTL_DataCleaning | Extremely common import problem, satisfying fix |
| Unmerge & Fill Down | modUTL_DataCleaning | Everyone hates merged cells — instant relatability |

### Chapter 2 — Formatting (visual transformation):
| Tool | Source Module | Why This One |
|------|-------------|--------------|
| AutoFit All | modUTL_Formatting | Quick win, visible improvement |
| iPipeline Branding | modUTL_Branding | VISUAL WOW MOMENT — plain sheet → branded professional |
| Date Standardizer | modUTL_Formatting | Common problem, clean fix |
| Highlight Negatives | modUTL_Formatting | Finance standard, instant visual value |

### Chapter 3 — Audit (trust and investigation):
| Tool | Source Module | Why This One |
|------|-------------|--------------|
| Workbook Health Check | modUTL_WorkbookMgmt | Comprehensive diagnostic, builds confidence |
| External Link Finder | modUTL_Audit | Solves a real problem (broken links), generates a report |
| Unhide All | modUTL_WorkbookMgmt | Simple, dramatic, universally useful |

### Chapter 4 — Python/SQL (power tools):
| Tool | Source Script | Why This One |
|------|-------------|--------------|
| Universal Data Cleaner | clean_data.py | Shows Python's power — 6 operations in one command |
| SQL Query Tool | sql_query_tool.py | Game-changer for SQL users, impressive for non-SQL viewers |

### Not Demoed (but mentioned or available):
The following are available on SharePoint but not shown in the video:
- 14 Finance-specific VBA tools (Duplicate Invoice Detector, GL Validator, Trial Balance Checker, Ratio Dashboard, etc.)
- Fuzzy Match / Fuzzy Lookup
- Bank Reconciler
- PDF Table Extractor
- Budget vs. Actual Consolidator
- Forecast Roll-Forward
- Variance Decomposition
- Word Report Generator
- 40+ additional VBA tools across all modules

---

## Sample File Requirements

The sample file (Sample_Quarterly_Report.xlsx) needs the following "problems" baked in for the demos to work:

| Problem | Where | For Which Demo |
|---------|-------|---------------|
| 5-10 blank rows scattered in data | Rows 15, 28, 43, etc. | Delete Blank Rows |
| Leading/trailing spaces in text cells | Column B (names or descriptions) | Remove Spaces |
| Numbers stored as text (green triangles) | Column D or E (amounts) | Text to Numbers |
| Merged cells | Column A (department names) | Unmerge & Fill Down |
| Narrow/wide columns | Various | AutoFit All |
| No header formatting (plain default look) | Row 1 | iPipeline Branding |
| Mixed date formats | Column C | Date Standardizer |
| Negative numbers (plain, not red) | Column D or E | Highlight Negatives |
| At least 1 external link formula | Any cell | External Link Finder |
| At least 1 hidden sheet | A sheet named "Notes" or "Archive" | Unhide All |
| A few #N/A or #REF! errors | Scattered | Workbook Health Check |

**Suggested data structure for the sample file:**

| Column A | Column B | Column C | Column D | Column E | Column F |
|----------|----------|----------|----------|----------|----------|
| Department | Employee Name | Date | Amount | Budget | Variance |

~100-150 rows of data. Fictional names, realistic departments (Engineering, Sales, Marketing, Finance, Operations). Amounts in the $500-$50,000 range. Some negative variances.

---

## Recording Tips Specific to This Video

1. **Record each chapter as a separate clip.** Same as Videos 1 and 2.

2. **Practice the VBA macro runs.** Open the sample file, run each macro in order, confirm they all work correctly on the specific data in the sample file. Do this BEFORE recording.

3. **Practice the Python demos.** Open the terminal, run the commands, make sure the output looks clean and the scripts execute without errors. Have the exact commands ready to paste.

4. **The iPipeline Branding demo is your key visual moment.** If any single demo needs to be perfect, it's this one. Practice it twice.

5. **For the Python chapter, consider pre-recording the terminal output** and splicing it in during editing. Terminal text can be hard to read in a screen recording — you may want to zoom in or increase font size in the terminal.

6. **Keep the demo pattern consistent:** state the problem → run the tool → show the result. Every demo follows this exact rhythm. The viewer learns the pattern and starts anticipating the payoff.

---

*Script created: 2026-03-06 | Part of Video Demo Master Plan*
*Based on full review of: 12 VBA modules (~78 tools) + 22 Python scripts*


================================================================================
================================================================================
# SECTION 5: START HERE DOCUMENT (Content Spec)
================================================================================
================================================================================

# START HERE — Content Specification

**Format:** Branded PDF (iPipeline Blue #0B4779 header, Arial/Helvetica fonts)
**Pages:** 2
**Location in SharePoint:** Top level of the iPipeline Finance Automation folder
**Status:** PDF already built and ready to use

## Content Structure

### Banner Header
- Title: "iPipeline Finance Automation"
- Subtitle: "Your guide to everything in this folder"
- iPipeline Blue background, white text

### Section: What Is This?
- One paragraph explaining the toolkit: complete automation for Finance workflows plus universal code library for any Excel file
- Stats line: "75+ VBA tools | 22 Python scripts | Step-by-step guides | Pre-built Copilot prompts"

### Section: Where to Start (Three Audience Paths)
Three cards side by side:

**Card 1 — "I want to see what this can do"**
Watch Video 1: What's Possible (5 min). Shows the highlights. If you only have time for one thing, watch this.

**Card 2 — "I'm on the Finance team"**
Watch Video 2: Full Demo Walkthrough (16 min). Covers every feature step by step. Then explore the demo file yourself.

**Card 3 — "I want tools for my own files"**
Watch Video 3: Universal Tools (10 min). Shows tools that work on any Excel file. Then go to the Universal Code Library folder and grab what you need.

### Section: What's in This Folder
Table listing each folder/file and what it contains:
- Videos/ — 3 videos (5 min, 16 min, 10 min)
- Demo File/ — The P&L automation file (.xlsm) plus the sample file used in Video 3
- Universal Code Library/ — VBA, Python, SQL tools organized by type
- Training Guides/ — Step-by-step PDFs for every tool, written for beginners
- Copilot Prompts.pdf — Pre-built prompts for help using any tool

### Section: Important Notes
Yellow callout box: "All financial data in the demo file is fictional. Product names are real iPipeline products, but every number, vendor, GL account, and transaction is fabricated for demonstration purposes. The tools and automation logic are production-ready and can be applied to real data."

### Section: Quick Start — 5 Minutes to Your First Automation
1. Watch Video 1: What's Possible (5 minutes)
2. Open the Training Guides folder and read the Getting Started guide
3. Pick one tool that solves a problem you deal with every week
4. Follow the step-by-step guide for that tool. If you get stuck, use the Copilot prompts.
5. Questions or ideas? Reach out to Connor.

### Footer
- "Questions or feedback? Contact Connor | Built with AI assistance (Claude) | Zero cost, zero IT involvement"
- "iPipeline Finance Automation • 2026"



================================================================================
================================================================================
# SECTION 6: SAMPLE FILE SPECIFICATION (For Video 3)
================================================================================
================================================================================

# Sample File Requirements — Sample_Quarterly_Report.xlsx

**Purpose:** This file is used in Video 3 to demonstrate the universal tools. It should look like a typical, slightly messy data export that any iPipeline employee might have on their desktop.

**Important:** Connor will have Claude Code build this file to match the quality standard of the P&L demo file. This spec defines what problems must be baked in for each demo to work.

## Data Structure

| Column A | Column B | Column C | Column D | Column E | Column F | Column G | Column H |
|----------|----------|----------|----------|----------|----------|----------|----------|
| Department | Employee Name | Transaction Date | Amount | Budget | Variance | Category | Status |

- ~120-150 rows of data
- Fictional names, realistic departments (Engineering, Sales, Marketing, Finance, Operations, Customer Success, Product, HR)
- Amounts in the $500–$48,000 range
- Some negative variances
- Categories: Software, Consulting, Travel, Equipment, Training, Subscriptions, Facilities, Marketing Spend
- Statuses: Approved, Pending, Reviewed, Flagged, Complete

## Required Problems (Intentional Mess for Demos)

| Problem | Where | For Which Demo | Details |
|---------|-------|---------------|---------|
| Merged cells | Column A (department names) | Unmerge & Fill Down | 3 merge groups, each spanning 12-18 rows |
| Leading/trailing spaces | Column B (employee names) | Remove Spaces | ~15% of name cells have invisible extra spaces |
| Numbers stored as text | Column D (amounts), Column E (budgets) | Text to Numbers | ~20% of Amount and ~12% of Budget stored as text strings |
| Blank rows scattered in data | Rows ~16, 33, 51, 68, 84, 102 | Delete Blank Rows | 6 completely empty rows inserted at irregular intervals |
| Mixed date formats | Column C (dates) | Date Standardizer | 5 different formats: MM/DD/YYYY, DD-MMM-YYYY, YYYY-MM-DD, MM-DD-YYYY, Month DD, YYYY — all stored as text |
| Negative values not formatted | Column F (variance) | Highlight Negatives | Negative numbers appear plain, not red |
| No header formatting | Row 1 | iPipeline Branding + AutoFit | Plain default Excel look, no colors, no bold |
| Poor column widths | All columns | AutoFit All | Some too narrow (truncated), some too wide |
| Error formulas | Below data area | Workbook Health Check | At least 1 #DIV/0! and 1 #N/A error |
| External link formula | Below data area | External Link Finder | 1 formula referencing [OtherWorkbook.xlsx] |
| Hidden sheets | Sheet tabs | Unhide All | 2 hidden sheets: "Archive Notes" and "Legacy Data" with minimal content |
| Total/Summary row | Bottom of data | Branding (total row detection) | Row with "Total" in Column A and SUM formulas in D, E, F |

## Style Notes
- The file should look generic and relatable — NOT styled, NOT branded, NOT polished
- It should feel like "something someone exported from a system and dropped on their desktop"
- The filename "Sample_Quarterly_Report.xlsx" reinforces this generic feel

