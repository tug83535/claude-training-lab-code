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
