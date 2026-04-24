# PROMPT 2 — Video 4-focused research review

Paste the full block below into Claude.ai. Attach the same 10 research files. (Separate chat from Prompt 1 so the AI doesn't conflate the two missions.)

---

You are reviewing a collection of code-idea research files I've gathered from other AI sessions. Your job is focused: identify the **5–10 best ideas specifically suited for Video 4** of my iPipeline Finance demo series.

## The project

I'm a Finance & Accounting analyst at iPipeline (life insurance / financial services SaaS). I'm building a 4-video demo for **2,000+ coworkers + CFO/CEO**. Videos 1–3 are recorded and shipped. Video 4 is the one I'm planning now.

**Videos 1–3 (recorded):**
- V1: "What's Possible" — Excel + VBA highlight reel
- V2: "Full Demo Walkthrough" — 62 automated actions on a demo P&L workbook
- V3: "Universal Tools" — VBA toolkit that plugs into any Excel file

**Video 4 is different:**
- **Theme:** "Python Automation for Finance"
- **Goal:** show what Python uniquely adds on top of Excel + VBA
- **Runtime target:** 5–8 minutes
- **Audience:** equal split — coworkers (need something they'll use) + CFO/leadership (need the "wow")
- **Must ship a real tool, not just a video.** The output should be downloadable + coworker-usable the Monday after they watch.

## What I'm currently considering for Video 4 (not locked in)

Three combos I'm weighing:

1. **"Finance Copilot" menu** — one Python script with a friendly numbered menu. User types a number, Python walks them through the task with prompts. Wraps my existing 8+ Python scripts in a single approachable entry point.
2. **"Excel Button Edition" (xlwings)** — macro-enabled workbook where coworkers click Excel buttons; Python runs silently behind the scenes; results appear as new sheets. Zero exposure to Command Prompt.
3. **"Hero Demo + Cookbook"** — dramatic one-command hero demo (messy folder of files → clean consolidated PDF) + a 5-recipe "cookbook" of copy-pasteable Python scripts for coworkers to steal.

I haven't picked a combo yet. Your research review should help me decide.

## Hard constraints for Video 4 (don't recommend things that violate these)

- **No external AI API calls** (OpenAI, Claude, Gemini, etc.) — parked for later
- **No Outlook / email automation** — parked for later
- **No Windows Task Scheduler** — parked for later
- **No internet scraping of company / paid data** — stick to public APIs if at all
- **Python packages must be `pip install`-safe** — no obscure dependencies, no things likely to be IT-blocked (pandas, openpyxl, pdfplumber, python-docx, thefuzz, numpy, matplotlib, xlwings are all fine)
- **Non-developer audience** — anything too "data-sciencey" is out unless clearly explained
- **Must be iPipeline-branded** — iPipeline Blue `#0B4779`, Navy `#112E51`, Arial fonts, plain English

## Existing Python scripts I already have (don't re-suggest these)

`aging_report.py`, `bank_reconciler.py`, `compare_files.py`, `forecast_rollforward.py`, `fuzzy_lookup.py`, `pdf_extractor.py`, `variance_analysis.py`, `variance_decomposition.py`, `clean_data.py`, `consolidate_files.py`, `multi_file_consolidator.py`, `date_format_unifier.py`, `two_file_reconciler.py`, `sql_query_tool.py`, `word_report.py`, `batch_process.py`, `regex_extractor.py`, `unpivot_data.py`, `pnl_forecast.py`, `pnl_dashboard.py`, `master_data_mapper.py`, `variance_decomposition.py`, plus 7 new stdlib-only zero-install scripts (`profile_workbook.py`, `sanitize_dataset.py`, `compare_workbooks.py`, `build_exec_summary.py`, `variance_classifier.py`, `scenario_runner.py`, `sheets_to_csv.py`).

## What I want you to produce

A focused list of **5 to 10 ideas** from the research files that are the strongest Video 4 candidates. Not 50. Not 20. **5 to 10 finalists** — your best curation.

For each finalist, give me:

1. **Idea name** (short, human-readable)
2. **What it does** (2–3 plain-English sentences)
3. **Why it's perfect for Video 4** — what makes it uniquely suited vs. generic universal-toolkit material
4. **Which combo direction it best supports** — Combo 1 (Copilot menu) / Combo 2 (xlwings) / Combo 3 (Hero + Cookbook) / any / neutral
5. **CFO wow factor** (1–5, 5 = high) — would this impress leadership?
6. **Coworker usefulness** (1–5, 5 = high) — would a Finance analyst use it next week?
7. **Demo-ability on camera** (1–5, 5 = high) — does it look good on screen in 30–60 seconds?
8. **Effort to build** — S (reuse existing ~1–2 hrs) / M (new code ~4–8 hrs) / L (rewrite >1 day)
9. **Python packages needed** — must be on my approved list above
10. **Source file** — which attached research file this came from

## Output format

- Single markdown table for the 5–10 finalists with the above columns
- Underneath the table, **a 1-paragraph narrative per finalist** explaining its story on camera (how you'd actually demo it in 30–60 seconds)

## At the end — your recommendations

Close with these 3 short sections:

1. **Best 3 for Combo 1 (Copilot menu)** — which ideas would I add to a menu-driven launcher?
2. **Best 3 for Combo 2 (xlwings)** — which ideas would translate well to "click a button in Excel"?
3. **Best 3 for Combo 3 (Hero + Cookbook)** — which idea should be THE hero demo, plus 5 recipes for the cookbook?

End with your personal pick: **"If this were my video, I'd pick combo X because ___"** — one opinionated paragraph.

## Rules for your review

- **Be picky.** Out of 200+ ideas across 10 files, only 5–10 deserve to be V4 finalists.
- **De-duplicate.** Same idea across files = one entry, cite all sources.
- **Skip anything that violates the hard constraints.** Don't list "AI variance commentary via OpenAI" and note it as parked — just skip it entirely.
- **Skip anything already in my existing script list.**
- **Favor ideas that produce a shareable tool** over pure demo-only tricks.
- **If an idea is great but needs a small dependency I might not have approved** — list it and flag the package, I'll decide.

Ready? Start the review. Ask me one clarifying question ONLY if truly needed — otherwise go straight to the finalists.

---

**(End of prompt. Attach all 10 research files to the chat before sending.)**
