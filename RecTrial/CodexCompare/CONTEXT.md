# Project Context

## The Company — iPipeline

iPipeline is a large SaaS company serving the life insurance and financial services industry. The Finance & Accounting (F&A) team supports internal operations: monthly P&L close, variance analysis, forecasting, reconciliation, headcount planning, AWS cost allocation, and financial reporting to executives and the board.

F&A lives in Excel. Data lands via SQL pulls, CSV exports from internal systems, and manual entry. Most workbooks are large, multi-sheet, and have grown organically over years. Reports are distributed as Excel files, PDFs, and PowerPoint decks.

## The User — Connor

- **Role:** Finance & Accounting, iPipeline
- **Technical background:** Not a developer. Writes VBA, basic SQL, basic Python — but treats them as tools, not as a profession.
- **Communication style:** Prefers plain English. Ask questions instead of guessing. Break big tasks into numbered steps. Confirm plans before executing.
- **Quality expectation:** World-class output every time. "Good enough" is not acceptable for anything that will be shown to coworkers or executives.
- **Working style:** Will iterate. Expects you to catch your own mistakes before delivering (self-review). After a correction, do not repeat the same mistake.

## The Audience — Two Tiers

### Tier 1 — Executives (CFO, CEO, leadership)
Will see the final product in a formal demo setting. They care about:
- **Clarity:** What does this do for the business?
- **Polish:** Does it look like a finished product or a hobby project?
- **ROI:** How many hours does this save the team? What does that translate to?
- **Scale:** Can this work across the 2,000-employee org, or is it a one-off?

### Tier 2 — Coworkers (2,000+ employees, mostly non-technical)
Will consume the videos and training guides and try to apply the universal toolkit to their own files. They care about:
- **Can I use this without calling IT?**
- **Will it break my file?**
- **Is there a button / step-by-step guide?**
- **How do I adapt the cool demo-specific code to my own workbook?** (← this is why the CoPilot Prompt Guide matters)

## Why This Project Exists

Most employees at iPipeline use Excel the way it ships — sort, filter, PivotTable, chart. They don't know what's possible when you layer VBA, Python, and SQL on top. The user wants to **open that door** for them:
- Show executives that F&A is a center of innovation
- Give coworkers a toolkit and a learning path so they can adopt these techniques themselves
- Build a reusable library that pays dividends long after the initial demo

This is **not** a one-off demo. It's the start of a Finance automation program at iPipeline.

## The Two Sample Files — What They Represent

### `samples/ExcelDemoFile_adv.xlsm` — "The Full P&L"
Represents the kind of multi-sheet, feature-heavy financial workbook that a F&A team owns and maintains monthly. Think: Assumptions tab, monthly tabs (Jan–Dec), an FY summary, Variance reports, Reconciliation checks, charts, named ranges, embedded logic. The P&L is the *beating heart* of Finance — everything ties back to it.

Your file-specific demo features should showcase what automation can do when the workbook's structure is known and stable — variance commentary, reconciliation alerts, executive briefings, month-end PDF packs, "what-if" scenario comparisons, etc.

### `samples/Sample_Quarterly_ReportV2.xlsm` — "The Coworker File"
Represents a *typical* file any coworker in the company might have — revenue broken out by region, product, rep; a customer list; a contact list. Not specialized. Not feature-heavy. Just data in tabs.

Your universal toolkit tools should demonstrate on this file: "look, I dropped these tools into a plain vanilla workbook and instantly had 50+ capabilities." This file is proof the toolkit is file-agnostic.

## Success Criteria

A successful delivery means:
1. The CFO/CEO demo lands and gets executive buy-in for more Finance automation work
2. Coworkers can install the universal toolkit and use it without help within 15 minutes of watching the video
3. The CoPilot Prompt Guide unlocks the file-specific features for anyone willing to paste a prompt
4. The code, guides, and videos are polished enough to be referenced as a model for other teams at iPipeline

## Non-Goals

- Not a SaaS product
- Not a commercial tool
- Not a one-size-fits-all platform (no need to handle every edge case in the world — handle realistic Finance-file edge cases)
- Not a training course on "how to code VBA" — it's a demo of what's possible, not a tutorial on syntax

## Working Environment

- Windows 10/11
- Excel (Microsoft 365)
- Python 3.10+ with openpyxl, pandas, pywin32 if needed for Excel automation
- SQL Server / generic SQL
- Microsoft 365 CoPilot (the coworker's AI assistant — key for Prong-2 adaptation)
- Git / GitHub for version control

## What You Will Not Have Access To

- Real production data
- Internal iPipeline systems
- Live database connections

Design code that accepts generic inputs (a path, a dataframe, a range reference) so a coworker can point it at their own data. Don't hardcode anything company-specific.
