# Video 4 — Draft Ideas (Start from Scratch)

Throwing out the current Video 4 plan (8 CMD-run Python scripts) and starting fresh. Goal: make Video 4 *genuinely useful* to the 2,000+ coworker audience + CFO/CEO, not just a tool-catalog demo reel.

**Project context:**
- Videos 1–3 covered Excel + VBA automation (universal toolkit + demo walkthrough)
- Video 4 is meant to show what Python adds on top — things Excel/VBA can't do alone
- Existing assets: 22+ Python scripts already built, 10 narration MP3s, 12 demo input files, 3 matching guides

**Started:** 2026-04-22
**Final plan:** TBD — will be picked from below, possibly combining 2-3 angles.

---

## The question Video 4 should answer

> *"Why should a Finance analyst care about Python when they already have Excel + Copilot?"*

Everything below is an angle on that question. Whatever we pick must have a clear "Excel + Copilot alone can't do this — but Python can" moment.

---

## Idea A — "One Monday Morning" (end-to-end story)

Open on a CFO's imaginary Monday 7am. A single Python script:
1. Reads last week's data from a folder
2. Reconciles bank to GL
3. Generates a variance narrative via AI (Copilot or API)
4. Builds a branded PDF report
5. Emails the CFO
Total: 40 seconds of runtime, zero human involvement.

**Story:** "This could be your Monday morning."
**Effort:** 1 new orchestrator script + integration with existing 8 demos. Moderate.
**Unique angle:** Shows orchestration + AI + scheduling — all things Excel can't do alone.
**Risk:** Email step requires Outlook automation or SMTP — possible compliance flags.

---

## Idea B — "Finance Copilot" (friendly menu-driven script)

ONE script: `finance_copilot.py`. Run it → user gets a numbered menu of 8 Finance tasks. Type a number → script walks them through with plain-English prompts. No path pasting, no flags.

**Story:** "Python doesn't have to be scary — it's just a chatbot with superpowers."
**Effort:** Wrap existing 8 scripts in a menu launcher. ~2 hours.
**Unique angle:** Positions Python as **approachable** for non-coders. High adoption potential.
**Risk:** Low. Just a launcher around existing scripts.

---

## Idea C — "AI-Powered Finance"

3–4 demos that specifically showcase Python + AI working together:
- Variance CSV → AI generates plain-English executive summary
- Expense line items → AI auto-classifies into categories
- Vendor contract PDF → AI extracts key dates, penalties, renewal terms
- 50 rows of transactions → AI flags anomalies with reasoning

**Story:** "Python + AI does what Excel + Copilot can't: automated, at scale, on your data."
**Effort:** 3–4 new scripts calling AI APIs. ~1 day. Needs API access or creative use of Copilot.
**Unique angle:** AI is top-of-mind for every CFO in 2026. Riding the wave.
**Risk:** API access + IT approval needed if using OpenAI/Claude direct. Alternative: use Copilot manually as the AI step.

---

## Idea D — "Excel + Python Power Duo" (xlwings)

Build a version of the demo workbook where coworkers never leave Excel. Excel button → Python runs silently behind the scenes → results appear as a new sheet.

**Story:** "Stay in the app you love. Python just makes it faster."
**Effort:** xlwings setup + 2–3 demo integrations. Moderate.
**Unique angle:** Lowest barrier to adoption. Coworkers don't need to touch Command Prompt.
**Risk:** xlwings installation may require IT approval. Some environments disable COM.

---

## Idea E — "The Scheduled Assistant" (set-and-forget)

Script runs on Windows Task Scheduler every Monday at 7am. Pulls data, flags anomalies, emails summary. Video shows the setup process + a day where the coworker arrives to a pre-populated inbox summary.

**Story:** "Automation while you sleep."
**Effort:** 1 script + Task Scheduler walkthrough. Low if using mock data.
**Unique angle:** "It works when I'm not at my desk" is viscerally impressive.
**Risk:** Email automation + scheduled jobs may raise IT flags.

---

## Idea F — "Starter Pack" (training over demo)

Position Video 4 as a **how-to-start** video, not a tool showcase. Walk through:
1. Installing Python on your Windows laptop (5 min)
2. Running your first script
3. 3 copy-pasteable scripts they can use today
4. Link to the toolkit folder

**Story:** "Here's how YOU get started with Python this afternoon."
**Effort:** Low — uses existing scripts, main work is writing instructions.
**Unique angle:** Converts passive viewers into active users.
**Risk:** Less "wow" factor for the exec audience.

---

## Idea G — "Data Pipeline: From Source to Dashboard"

Full data pipeline end-to-end: SQL extract → Python transform → Excel dashboard → PDF report → email. One button, one command, 60 seconds. Each stop along the pipeline briefly visible.

**Story:** "This is the whole data journey. One command. Sixty seconds."
**Effort:** Moderate — requires a real or mock SQL source.
**Unique angle:** Shows Python as **orchestrator** of the full stack iPipeline already uses.
**Risk:** SQL access may require real or simulated DB.

---

## Idea H — "Fix Your File Friday" (Python repair doctor)

Coworker drags any broken Excel file onto a script. Script fixes it: text-stored numbers, broken links, duplicate rows, phantom hyperlinks, blank rows, inconsistent date formats. Returns a clean version + a diff report of what was fixed.

**Story:** "Give it your ugliest file. Get it back perfect."
**Effort:** Wrap existing sanitize/cleanup tools into one entry point. Low.
**Unique angle:** Solves a real, painful, weekly problem for Finance coworkers.
**Risk:** Very low. Additive + immediately useful.

---

## Idea I — "One-Click Monthly Report" (hero demo)

Start with a folder of 10 messy CSV files. Run ONE command. Out pops: consolidated summary, reconciliation report, variance narrative, branded PDF deck, emailed summary.

**Story:** "Your month-end close in 30 seconds."
**Effort:** One orchestrator script that glues together existing scripts + adds PDF generation + email. Moderate.
**Unique angle:** **The** hero demo of the whole project. Collapses a 3-day task into 30 seconds.
**Risk:** Email + PDF integration has moving parts.

---

## Idea J — "Web Scraping for Finance" (Python talks to the internet)

Python fetches: exchange rates, stock prices, interest rate curves, SEC filings, industry benchmarks. Visibly pulls live data no Excel user could get.

**Story:** "Excel can't see the internet. Python can."
**Effort:** Low. Use free APIs (FRED, yfinance).
**Unique angle:** Shows the internet-talking superpower no VBA macro can match.
**Risk:** Using public data is fine; corporate API access may need IT sign-off.

---

## Idea K — "The Email Robot"

Python reads your Outlook inbox. Finds emails with "expense report" attachments. Processes the attached Excel. Auto-replies with a reconciled summary.

**Story:** "Send me your expense report — I'll reconcile it and reply in 5 seconds."
**Effort:** Moderate — Outlook MAPI or Graph API integration.
**Unique angle:** Show-stopper for anyone who hates inbox triage.
**Risk:** Outlook automation + company email policy — needs IT clearance.

---

## Idea L — "Before Python / After Python" split-screen

Literally split the screen in half. Left side: manual Excel process (tediously). Right side: Python doing the same thing in seconds. Do it for 3–4 tasks. Visceral, timed.

**Story:** "Which one would you rather do?"
**Effort:** Low — uses existing scripts + requires Excel footage of manual process.
**Unique angle:** Physically impossible to argue with the time comparison.
**Risk:** Recording the manual Excel half takes patience.

---

## Idea M — "Build a Python Tool LIVE" (with Copilot)

Record yourself writing a Python script in 60 seconds — not by coding, but by typing natural-language prompts into Copilot (Cursor, VS Code Copilot, Claude). Copilot generates the script. You run it. Done.

**Story:** "You don't need to be a coder. AI writes the code for you."
**Effort:** Low — just one live demo + a clean script at the end.
**Unique angle:** Removes the "I can't code" objection. Positions Python as **accessible**.
**Risk:** Live coding is risky on camera — one typo could tank the take. Rehearse heavily.

---

## Idea N — "Dashboard in 5 Minutes" (Streamlit)

Show Python building an interactive Finance dashboard that runs in a browser. Sliders change values. Charts update live. CFO realizes the same data they stare at every month could be interactive.

**Story:** "What if your monthly report was interactive?"
**Effort:** Moderate — new Streamlit app.
**Unique angle:** Modern, web-native feel. Very different from static Excel.
**Risk:** Streamlit install + local port access required.

---

## Idea O — "PDF Magic"

Extract structured data from messy PDFs (vendor invoices, bank statements, government filings). Table data → Excel → auto-formatted. PDFs are the #1 Finance pain point. Python crushes them.

**Story:** "You've been manually retyping PDFs for years. Stop."
**Effort:** Low — existing pdf_extractor.py handles it.
**Unique angle:** Targets a specific, well-known Finance pain.
**Risk:** Low.

---

## Idea P — "Real Case Study: Connor's Tuesday Headache"

Pick ONE real pain point you (or a coworker) actually had. Walk through the problem first. Show Python solving it. Concrete, relatable, personal.

**Story:** "Here's the specific Tuesday that made me learn Python."
**Effort:** Low — depends on picking the right story.
**Unique angle:** Human, not theoretical. Most memorable format.
**Risk:** Depends on finding a compelling real-life example.

---

## Idea Q — "Python Cookbook: 5 Recipes You Can Steal"

Short punchy format. Each recipe is 5–10 lines of code with clear comments. Viewer sees: recipe name → code snippet → output. Ends with "download the cookbook" link.

**Story:** "Here are 5 Finance recipes. Steal them all."
**Effort:** Low — curate from existing scripts.
**Unique angle:** Extremely practical, zero fluff.
**Risk:** Low.

---

## Idea R — "Python vs Excel Race" (entertaining)

Actual timed race. Same task. Excel on one side, Python on the other. Starting pistol, go. Python finishes in 3 seconds, Excel user is still clicking 30 seconds later.

**Story:** "Don't take my word for it. Watch."
**Effort:** Low-moderate.
**Unique angle:** Entertaining, shareable, low-ego way to show scale.
**Risk:** None.

---

## Combinations worth considering

- **A + C** (Monday Morning + AI-powered) — rich story with AI superpower highlighted
- **B + F** (Copilot menu + Starter Pack) — approachable + training-focused
- **H + Q** (File Doctor + Cookbook) — immediate usefulness
- **I + L** (Hero demo + before/after) — big finale with timed proof
- **M + F** (Live Copilot coding + Starter Pack) — addresses "I can't code" fear head-on

---

## My current instinct (subject to change)

Probably a **combo of I (One-Click Monthly Report) + L (Before/After)** — show the pain, then show Python killing it in one command. That's the strongest hero-demo framing and makes the previous 3 videos feel like the setup to this one big reveal.

But we're still in brainstorm mode. Nothing locked.

---

## Still to come

- User will share research files with additional code ideas
- Once those land, expand this list
- After brainstorming finishes, pick the final angle(s) and write the real Video 4 plan
- Then build any new scripts needed + record

Parking for now.
