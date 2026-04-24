# PROMPT 1 — Full research review for the iPipeline Finance demo project

Paste the full block below into Claude.ai. Attach all your research files (the 10 files of code ideas for SQL / Python / VBA) to the same chat.

---

You are reviewing a collection of code-idea research files I've collected from other AI sessions. Your job is to produce a **curated list of code ideas** that genuinely fit my project's purpose — not a dump of everything.

## The project

I'm a Finance & Accounting analyst at iPipeline (a SaaS company in life insurance / financial services). I'm building a 4-video demo series for **2,000+ coworkers plus the CFO and CEO**. The series shows non-technical Finance folks what's possible when you combine Excel + VBA + Python + SQL + AI. Every output has to be polished, plain English, and CFO-grade.

**The 4 videos:**
- **Video 1** — "What's Possible" — fast highlight reel (Excel + VBA focus) [RECORDED]
- **Video 2** — "Full Demo Walkthrough" — end-to-end tour of a macro-enhanced P&L workbook with 62 automated actions [RECORDED]
- **Video 3** — "Universal Tools" — VBA toolkit that works on any Excel file (plug-and-play) [RECORDED]
- **Video 4** — "Python Automation for Finance" — command-line Python scripts for Finance tasks [PLANNING — currently being redesigned from scratch]

**Existing code I already have:**
- 23 VBA modules with ~140 universal tools (data sanitizer, compare, consolidate, highlights, pivot tools, tab organizer, column ops, sheet tools, comments, validation builder, lookup builder, command center, exec brief, finance-specific tools, audit tools, branding, etc.)
- 22+ Python scripts already built (aging report, bank reconciler, compare files, forecast rollforward, fuzzy lookup, pdf extractor, variance analysis, variance decomposition, clean data, consolidate files, multi-file consolidator, date unifier, two-file reconciler, SQL query tool, word report, batch processor, regex extractor, unpivot, pnl forecast, pnl dashboard, etc.)
- 7 new stdlib-only "zero-install" Python scripts
- 4 SQL scripts (staging, transformations, validations, enhancements)
- A Python-to-Excel `word_report.py` for branded Word report generation

## My constraints for this project (important — don't recommend things that violate these)

- **No external AI API calls** (OpenAI, Claude, etc.) — parked for later. iPipeline IT policy uncertain for now.
- **No Outlook / email automation** — too messy for this audience.
- **No Windows Task Scheduler dependencies** — not familiar enough with it for coworkers to replicate.
- **Must be iPipeline-branded** — iPipeline Blue `#0B4779`, Navy `#112E51`, Arial fonts, plain-English output.
- **Non-developer audience** — every feature must be explainable and runnable by someone with zero coding background.
- **Plug-and-play is valued** — tools that work on any coworker's file (no hardcoded sheet names) are more valuable than demo-specific ones.

## What I want you to produce

A **curated list of code ideas** from the attached files. Filter out anything that violates my constraints above. For every idea that survives, give me:

1. **Idea name** (short, human-readable)
2. **What it does** (1–2 plain-English sentences)
3. **Language** (VBA / Python / SQL / combo)
4. **Best fit** — which video + context it belongs in:
   - **Universal toolkit add** (works on any file, goes into the library)
   - **Video 4 candidate** (specifically for the upcoming Python-focused video)
   - **Future / post-demo** (parked for later — too ambitious for V4 or requires constraints I've disallowed)
   - **Already covered** (I already have something like this — note which existing tool)
5. **Effort** — S (under ~2 hours) / M (half-day) / L (>1 day)
6. **Why it's worth including** — 1 sentence on the concrete value to coworkers or the CFO
7. **Source file** — cite which of my attached research files it came from

Group your output into **4 sections**:

### Section A — Universal Toolkit Additions (new code worth adding to the library)
Tools that would live alongside my existing `modUTL_*` VBA modules or `UniversalToolkit\python\` scripts. Plug-and-play on any file, not specific to the demo workbook.

### Section B — Video 4 Candidates (Python ideas that might anchor the upcoming V4 video)
Ideas specifically suited to the "Python Automation for Finance" theme. If unclear, lean toward **including** it and letting me decide.

### Section C — Future Ideas (parked for post-demo work)
Anything ambitious, complex, or requiring things I've disallowed. Note the specific reason for parking.

### Section D — Skip (and why)
Things in the research files that shouldn't be used at all — either already covered, off-brand, or not relevant. Short explanations.

## Rules for your review

- **Be curatorial, not exhaustive.** If you find 200 ideas in 10 files, return the best ~40–60.
- **De-duplicate.** If the same idea appears across multiple files, consolidate into one row and cite all sources.
- **Flag overlaps** with my existing code. I'd rather have the research AI tell me "you already have this" than duplicate work.
- **Be strict about quality.** An idea with 3 lines of sketchy pseudocode is not the same as one with a clear, explained implementation — weigh the evidence.
- **Cite specific files.** If you say "Idea X from file 3", give the filename.
- **Markdown format output.** Use tables where they fit. Headers + bullets for longer descriptions.

## Final deliverable style

One clean markdown document I can scroll through in 5 minutes. Tables beat walls of prose. When you're done, end with a one-paragraph **Top 10 summary** — your personal top 10 picks across all 4 sections, ranked by bang-for-buck (value ÷ effort).

Ready? Start the review. Ask me one clarifying question ONLY if something in the attached files is genuinely ambiguous — otherwise go straight into producing the list.

---

**(End of prompt. Attach all 10 research files to the chat before sending.)**
