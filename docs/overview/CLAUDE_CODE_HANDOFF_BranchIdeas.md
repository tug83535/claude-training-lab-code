# Claude Code Handoff — Branch Ideas Review (April 2026)

> **How to use this file:**  
> Copy everything from the horizontal rule below and paste it as your first message into a new Claude Code (claude.ai/code) session.  
> Claude Code will have full context on what was reviewed, what was decided, and what to build next.

---

---

## Context for Claude Code

Hello! I need your help continuing work on the **APCLDmerge / iPipeline Finance Demo project**.  
Read everything below carefully before suggesting or building anything.

---

### Who I Am and What This Project Is

I am a Finance & Accounting analyst at iPipeline (life insurance / financial services SaaS).  
I am building a **4-video demo series** for 2,000+ coworkers plus the CFO and CEO showing what's possible when Finance combines Excel + VBA + Python + SQL + AI.

**Videos:**
- Video 1 — "What's Possible" (highlight reel) — **RECORDED ✅**
- Video 2 — "Full Demo Walkthrough" — **RECORDED ✅**
- Video 3 — "Universal Tools" — **RECORDED ✅**
- Video 4 — "Python Automation for Finance" — **PLANNING / REDESIGN 🔄**

---

### My Hard Constraints (Never Violate These)

- ❌ No external AI API calls (OpenAI, Claude API, etc.)
- ❌ No Outlook or email automation
- ❌ No Windows Task Scheduler dependencies
- ✅ Approved Python packages only: `pandas`, `openpyxl`, `pdfplumber`, `python-docx`, `thefuzz`, `numpy`, `matplotlib`, `xlwings`, and stdlib
- ✅ Non-developer audience — every output must be explainable to someone with zero coding background
- ✅ Plug-and-play preferred — tools must work on any Excel file, no hardcoded sheet names
- ✅ iPipeline branding: Blue `#0B4779`, Navy `#112E51`, Arial fonts

---

### What Already Exists (Do NOT Duplicate)

**VBA:** 23 modules, ~140 universal tools. Key modules live in `FinalExport/UniversalToolkit/vba/`.  
**Python:** 22+ scripts (aging report, bank reconciler, compare files, forecast rollforward, fuzzy lookup, pdf extractor, variance analysis, variance decomposition, clean data, consolidate files, multi-file consolidator, date unifier, two-file reconciler, SQL query tool, word report, batch processor, regex extractor, unpivot, pnl forecast, pnl dashboard, etc.)  
**Zero-Install Python (stdlib-only):** 7 scripts in `FinalExport/UniversalToolkit/python/ZeroInstall/`  
**SQL:** 4 scripts (staging, transformations, validations, enhancements)

---

### What Was Just Reviewed (Branch Review Output — April 2026)

A full review was done of all active GitHub branches (last 10 days). Here is the curated output:

---

#### SECTION A — Universal Toolkit Additions (Already Implemented or Ready to Import)

| # | Idea | Language | Status |
|---|------|----------|--------|
| A1 | Materiality Classifier — tags rows Material/Watch/Normal | VBA | ✅ Built — `modUTL_Intelligence.bas` |
| A2 | Exception Narrative Generator — plain-English row commentary | VBA | ✅ Built — `modUTL_Intelligence.bas` |
| A3 | Data Quality Scorecard — 0-100 score on blanks/errors | VBA | ✅ Built — `modUTL_Intelligence.bas` |
| A4 | Header Row Auto-Detect — finds true header row automatically | VBA | ✅ Built — `modUTL_Core.bas` |
| A5 | Quick Row Compare Count — fast hash-based mismatch pre-check | VBA | ✅ Built — `modUTL_Compare.bas` |
| A6 | Run Receipt Sheet — writes audit tab on each macro run | VBA | ✅ Built — `modUTL_Audit.bas` |
| A7 | Cover "Show Tools" Button Installer — branded launcher button | VBA | ✅ Built — `modUTL_CommandCenter.bas` |
| A8 | Static Intelligence Category in Command Center — pins tools near top | VBA | ✅ Built — `modUTL_CommandCenter.bas` |
| A9 | Zero-Install Workbook Profiler — stdlib-only workbook inventory | Python | ✅ Built — `ZeroInstall/profile_workbook.py` |
| A10 | Word Report Talking Points — `--talking-points` CLI flag | Python | ✅ Built — `UniversalToolkit/python/word_report.py` |

---

#### SECTION B — Video 4 Candidates (Build These Next)

These are the **best ideas** for the Video 4 "Python Automation for Finance" video.  
None are built yet. Prioritize these in order:

| Priority | Idea | Language | Effort | Notes |
|----------|------|----------|--------|-------|
| 🥇 1 | **Close Readiness Score View** — SQL view returning per-entity close score 0–100 | SQL | M | Highest ROI — CFO-level language |
| 🥈 2 | **Exception Triage Engine** — ranks exceptions by impact × confidence × recency | Python | M | Improves analyst workflow every month-end |
| 🥉 3 | **Control Evidence Pack Generator** — zips logs + validation results into audit bundle | Python | M | Cuts audit-prep hours directly |
| 4 | **Zero-Install Workbook Compare** — row-level diff to CSV, stdlib only | Python | S | Great live demo for locked-down laptops |
| 5 | **Zero-Install Variance Classifier** — labels rows Over/Under/On-target | Python | S | Already built — confirm if Video 4 ready |
| 6 | **Zero-Install Scenario Runner** — percentage shocks on any metric column | Python | S | Already built — confirm if Video 4 ready |
| 7 | **Sheets-to-CSV Batch Export** — every sheet → separate CSV | Python | S | Already built — confirm if Video 4 ready |
| 8 | **Finance Data Contract Checker** — YAML-based schema/quality validation | Python | M | Prevents bad data entering reports |
| 9 | **Workbook Dependency Scanner** — maps formula/range impact graph | Python | M | Safer workbook changes |
| 10 | **Executive Summary Builder** — CSV → markdown summary with stats | Python | S | Already built — confirm if Video 4 ready |

---

#### SECTION C — Future / Post-Demo (Park Until After Video 4)

Do not build these until after Video 4 is recorded. They are great ideas but would delay the demo.

- SQL: Allocation Drift Tracker, Forecast Backtest Warehouse, Subledger Completeness Matrix, Workbook-to-Source Recon Mart, Vendor Payment Velocity Baselines, JE Duplicate Ring Detection, Close Bottleneck Heatmap, SoD Audit Pack
- VBA: Formula Integrity Fingerprinting, Exception Workbench Sheet, Macro Runtime Telemetry Dashboard, Controlled Snapshot Sign-off

---

#### SECTION D — Skip Permanently

These were found in branches but violate constraints or are out of scope:

- ❌ Outlook Mail Merge / Calendar Builder (email automation rule)
- ❌ JIRA Bridge / Digest (external integration, not finance-close)
- ❌ Slack/Teams Webhook Notifiers (external platform dependency)
- ❌ AWS Cost Optimizer (out of demo scope)
- ❌ ML Churn / Ticket Triage (`scikit-learn` not in approved packages)
- ❌ PowerShell IT Admin scripts (outside audience/stack)
- ❌ Power Automate / Office Scripts flows (different surface, distracts)

---

### Top 10 Priority Ranking (Use This to Plan Work)

| Rank | Idea | Effort | Why |
|------|------|--------|-----|
| 1 | Close Readiness Score View | M | One number per entity = CFO-grade language |
| 2 | Exception Triage Engine | M | Directly improves analyst workflow monthly |
| 3 | Data Quality Scorecard (VBA) | S | Fast, visual, zero-setup demo moment |
| 4 | Control Evidence Pack Generator | M | Cuts audit hours — leadership cares |
| 5 | Materiality Classifier (VBA) | S | Turns any flat sheet into risk view in seconds |
| 6 | Word Report Talking Points | S | AI-style output, no AI calls — great CFO story |
| 7 | Zero-Install Workbook Compare | S | Convinces locked-laptop coworkers immediately |
| 8 | Header Row Auto-Detect (VBA helper) | S | Makes all tools more reliable |
| 9 | Exception Workbench Sheet | M | Creates one action hub post-demo |
| 10 | Macro Runtime Telemetry Dashboard | M | Shows toolkit is production-grade |

---

### What I Need from You Right Now

**Option 1 — Build Video 4 Python Scripts:**  
Help me build the `exception_triage.py` script (Priority #2) using only approved packages.  
Config-driven weights. CLI interface. Output to CSV or Excel. No hardcoded paths.

**Option 2 — Build the Close Readiness SQL View:**  
Help me write `close_readiness_score_vw.sql` (Priority #1) for our SQLite-based setup.  
Score per entity per day. Weighted from: failed validations + missing feeds + late postings.

**Option 3 — Review and confirm Video 4 script list:**  
Look at `FinalExport/UniversalToolkit/python/ZeroInstall/` and tell me which scripts are Video 4-ready as-is vs. need polish.

**Tell me which option you'd like me to start with, or give me a new direction.**

---

### Reference File Locations in the Repo

```
FinalExport/UniversalToolkit/vba/modUTL_Intelligence.bas   ← Materiality + Narratives + Scorecard
FinalExport/UniversalToolkit/vba/modUTL_Core.bas           ← Header detect + shared helpers
FinalExport/UniversalToolkit/vba/modUTL_Compare.bas        ← Quick row compare count
FinalExport/UniversalToolkit/vba/modUTL_Audit.bas          ← Run receipt sheet
FinalExport/UniversalToolkit/vba/modUTL_CommandCenter.bas  ← Command center + button installer
FinalExport/UniversalToolkit/python/ZeroInstall/           ← 7 stdlib-only Python scripts
FinalExport/UniversalToolkit/python/word_report.py         ← Word doc with talking points
codexreview2/CodexCodeIdeas.md                             ← Full backlog of 30 ideas
docs/overview/BranchIdeasReview_April2026.md               ← This review (full table format)
CLAUDE.md                                                  ← Full project memory and constraints
```

---

*Handoff generated: 2026-04-22. Full curated ideas list: `docs/overview/BranchIdeasReview_April2026.md`*
