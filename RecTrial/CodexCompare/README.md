# Finance & Accounting Automation Demo — Codex Build

## Your Mission

You are Codex. You are building a brand-new, world-class Finance & Accounting automation demo project **from the ground up**. The project will be presented to **2,000+ employees** and the **CFO and CEO of iPipeline**. It must be polished, professional, and represent "what's possible" when a Finance team combines **Excel + VBA + SQL + Python** in creative ways.

You have no prior code to reference. You get a blank slate and two real-world sample Excel files. Your job is to look at those files, understand the kind of work Finance & Accounting teams do, and design a complete demo project around them.

**Ground yourself in these docs before writing any code:**
1. `CONTEXT.md` — who the user is, who the audience is, what "good" looks like
2. `CONSTRAINTS.md` — what NOT to build (hard rules — native Excel/OneDrive features are banned)
3. `BRAND.md` — iPipeline brand styling (colors, fonts, visual rules)
4. `PLAN.md` — the project plan template you fill in at Stage 1 (it is your contract for every subsequent stage)
5. `STARTER_PROMPT.md` — the exact prompts the user will paste at each stage (read it so you know what's coming)

Then ask any clarifying questions before starting.

## How This Project Runs — Stage by Stage

This project is too large for one Codex task. It will be completed across roughly **8–15 separate Codex runs**, each driven by a prompt from the user. The flow:

1. **Stage 1 (this run):** Read the brief, inventory the sample files, and fill `PLAN.md`. **No code yet.**
2. **User reviews `PLAN.md`** and replies with the word `approved`.
3. **Stage 2+:** Each run builds a slice of `PLAN.md` (a few modules, one guide, a video script) and commits. User reviews between stages.

Rules that apply across all stages:
- Every stage, re-read `README.md`, `CONTEXT.md`, `CONSTRAINTS.md`, `BRAND.md`, and `PLAN.md` first.
- Never expand scope beyond the current stage's prompt.
- Never assume the user approved something — wait for the literal word `approved`.
- If something in `PLAN.md` needs to change, stop and propose the change first.

---

## The Two Sample Files

Located in `samples/` — **do not modify these files directly.** Treat them as representative inputs. Your code should work alongside them (in a copy, or as an add-in, or as a new workbook that references them).

| File | Represents | How it will be used in videos |
|---|---|---|
| `samples/ExcelDemoFile_adv.xlsm` | A full P&L (Profit & Loss) demo workbook — assumptions, actuals, variance reports, charts, reconciliation sheets | Feature-rich demo of file-specific automation (variance commentary, reconciliation, dashboards, PDF export, etc.) |
| `samples/Sample_Quarterly_ReportV2.xlsm` | A basic quarterly revenue report (multiple tabs: raw data, customer list, contacts, etc.) | Demo of "universal tools" that plug into *any* Excel file |

Open both, fully review every sheet, every column, every named range, every existing macro (if any). Do not assume any sheet is irrelevant. **Fully understand the inputs before designing the output.**

---

## The Two-Prong Architecture

The user wants coworkers to walk away with **two distinct deliverable types**, and you must keep them clearly separated in your project:

### Prong 1 — Universal Tool Library (plug-and-play, file-agnostic)

A library of VBA macros + Python scripts that work on **any Excel file, any data shape, any company**. No hardcoded sheet names. No hardcoded column positions. No assumptions about headers. These are the "Swiss Army knife" tools coworkers can drop into their own workbooks and immediately use.

Examples of the *kind* of tool that belongs here (not an exhaustive list — use your judgment):
- Data sanitizers (fix text-stored numbers, strip whitespace, normalize dates)
- Cross-sheet comparators
- Column splitters / combiners with configurable delimiters
- Conditional highlighters (threshold, duplicates, top/bottom N)
- Tab reorganizers (color-coding, sorting, grouping)
- Consolidation across multiple sheets
- Validation builders
- Formula builders (VLOOKUP / INDEX-MATCH helpers)
- Anything clever that Finance teams would want "on tap"

This library should be distributable as a standalone Excel Add-In (`.xlam`) or Python package. Organize it so a non-technical coworker can install it and have dozens of tools available in any file they open.

### Prong 2 — File-Specific Demo Features (tied to the P&L workbook)

Rich, polished features built specifically around `ExcelDemoFile_adv.xlsm`. These showcase what's possible when code is built for a *specific* workbook's structure — variance commentary auto-generation, reconciliation dashboards, executive briefs, PDF batch export, "what-if" scenario runners, etc.

This code **will not work on other files** without modification — and that's the point. It's the reason the CoPilot Prompt Guide (below) exists.

### The Bridge — CoPilot Prompt Guide

Because Prong-2 code is file-specific, coworkers can't just drop it into their own workbooks. **Build a CoPilot Prompt Guide** that teaches non-technical coworkers how to:
- Paste your file-specific code into Microsoft 365 CoPilot
- Describe their own file's structure in plain English
- Ask CoPilot to adapt the code to their file
- Understand the edits CoPilot suggests

This guide is the key that unlocks Prong-2 code for the other 1,999 employees whose workbooks look nothing like the sample P&L. Make it thorough, non-technical, with copy-paste prompt templates.

---

## What You Will Deliver

At minimum, the repo must contain:

```
samples/                              ← (do not modify — already provided)
    ExcelDemoFile_adv.xlsm
    Sample_Quarterly_ReportV2.xlsm

vba/                                  ← all .bas modules
    universal/                        ← Prong 1 — plug-and-play
    demo/                             ← Prong 2 — file-specific to the P&L

python/                               ← all .py scripts
    universal/                        ← file-agnostic helpers
    demo/                             ← demo-specific ETL / forecasting / reporting

sql/                                  ← .sql scripts (if applicable)

guides/
    copilot-prompt-guide.md           ← THE bridge between Prong 2 and coworkers
    universal-toolkit-user-guide.md   ← how coworkers install + use the library
    demo-walkthrough-guide.md         ← step-by-step of every demo feature
    brand-styling-reference.md        ← recap of BRAND.md for the team

videos/
    video-1-<title>.md                ← one script per video, detailed
    video-2-<title>.md
    video-3-<title>.md
    ...                               ← 3 to 5 videos total

README.md                             ← update this file once you've scoped the project
CONTEXT.md                            ← do not delete
CONSTRAINTS.md                        ← do not delete
BRAND.md                              ← do not delete
```

You may reorganize, split, or rename as you see fit — but the **two prongs must stay separated** and the CoPilot Prompt Guide must exist.

---

## Video Demo Requirement

Produce **3 to 5 video scripts** (you pick the exact count). Similar length across videos. **5–10 minutes each is a good target; go longer only if the content needs it.** Videos should feel like a coherent series — each one earns its place, no filler.

Rough shape of a good video series (you may adjust):
- **Video 1** — The hook: "What if Excel could do *this*?" Fast-paced highlight reel of the most impressive features
- **Video 2** — Full walkthrough of the file-specific demo workbook (Prong 2)
- **Video 3** — The universal toolkit in action on a plain coworker-style file (Prong 1)
- **Video 4** — Python/SQL integration: bringing external data, forecasting, ETL
- **Video 5** (optional) — CoPilot Prompt Guide walkthrough: "here's how you adapt all this to YOUR file"

Each video script must include:
- Timestamped outline
- Full narration text (speakable, conversational — not technical jargon)
- On-screen action callouts (click here, highlight this, type this)
- Closing CTA

Narration will be generated via AI voice (ElevenLabs-style). Write like a human would speak it.

---

## Audience Rules (Non-Negotiable)

- **Training guides are for non-developers.** Plain English. Every step written out, no matter how obvious. "Open Excel" level of detail when needed.
- **Code comments should be sparse.** Let good naming do the work. Only comment the *why* when a constraint is non-obvious.
- **Code should be defensive but not paranoid.** Guard against realistic Finance-file weirdness (merged cells, blank rows, text-stored numbers, formulas that return errors). Don't build validation for impossible scenarios.
- **Everything demo-facing must be iPipeline-branded.** See `BRAND.md`.

---

## Quality Bar

Before you mark anything "done," ask yourself:
1. Would the CFO be proud to see this on screen in front of the CEO?
2. Would a non-developer on the Finance team be able to use this *without asking for help*?
3. Does this feature do something Excel/OneDrive can't already do natively? (If no — see `CONSTRAINTS.md`.)
4. Is every part of the deliverable polished, or am I shipping "good enough"?

If any answer is "no," fix it before moving on.

---

## Ground Rules for Collaboration

- **Before starting major work, post a short plan and wait for confirmation.** (Small obvious edits — just do them.)
- **Ask clarifying questions** rather than inferring when something is ambiguous.
- **Never modify the files in `samples/`.** Copy them if you need to run code against them.
- **No destructive shortcuts.** If a test fails, fix the root cause — don't suppress it.
- **Explain trade-offs** when you make design decisions (e.g., "I used Python over VBA here because…").
- **Self-review your output** before declaring a feature finished.

---

## First Steps (Stage 1 — Do These Before Any Coding)

1. Read `CONTEXT.md`, `CONSTRAINTS.md`, `BRAND.md`, `PLAN.md`, `STARTER_PROMPT.md` fully.
2. Open both sample Excel files and inventory every sheet, column, named range, and existing VBA.
3. **Fill `PLAN.md`** by replacing every `<<FILL IN>>` block with your proposed answers. Preserve the section structure.
4. Commit `PLAN.md`.
5. Wait for the user to reply with the literal word `approved` before Stage 2.

Welcome to the project. Make it world-class.

---

## Current Build Status (Stage Progress)

Completed so far in-repo:
- Stage 1 planning completed in `PLAN.md`.
- Universal toolkit foundation delivered in `vba/universal/`.
- Demo-specific modules delivered in `vba/demo/`.
- Python companion utilities delivered in `python/universal/`.
- Demo workflow Python utilities delivered in `python/demo/`.
- User/training guides delivered in `guides/`.
- Git push/branch quickstart guide delivered in `guides/git-branch-push-quickstart.md`.
- Claude comparison handoff + prompt guides delivered in `guides/claude-handoff-deep-analysis.md` and `guides/claude-review-prompt.md`.
- Project execution backlog tracked in `PROJECT_TODO.md`.
- Video scripts 1–5 delivered in `videos/`.
- Repository smoke checks centralized in `tests/stage2_smoke_check.py`.
- Auto-generated code inventory available in `CODE_INVENTORY.md`.
- Universal tool catalog expanded to 160 high-quality tool candidates in `guides/universal-tool-catalog.md`.

Next focus area:
- Polishing pass for consistency, deeper runtime validation in Excel host, and final packaging readiness.

## Running Smoke Checks

Run this command from the repository root:

```bash
bash scripts/run_stage_smoke.sh
```

What it does:
1. Runs repository artifact smoke checks (`tests/stage2_smoke_check.py`).
2. Compiles Python utilities to catch syntax issues early.

## Code Inventory (All Project Code)

This repository includes an auto-generated code map at:

- `CODE_INVENTORY.md`

To refresh it after code changes:

```bash
python scripts/update_code_inventory.py
```

`tests/stage2_smoke_check.py` enforces that `CODE_INVENTORY.md` is up to date (smoke checks fail if it drifts).
