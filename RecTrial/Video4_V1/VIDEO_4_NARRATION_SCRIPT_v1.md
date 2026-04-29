# Video 4 — Narration Script
## Python Automation for Finance — iPipeline

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Version:** v1.0 — 2026-04-28
**Target length:** 9–12 minutes
**Format:** Word-for-word narration for ElevenLabs. Read at a conversational pace (~130 words/min).
**[ON SCREEN] cues** are for recording reference only — not read aloud.

---

## HOW TO USE THIS SCRIPT

- Record each chapter as a separate ElevenLabs clip (or all at once if you prefer a single take)
- [ON SCREEN] notes tell you what to show on screen while each paragraph plays
- Pause naturally at periods — do not rush through sentences
- Aim for the same tone as Videos 1–3: plain, confident, Finance-literate

---

## CHAPTER 1 — Why Python After Excel and VBA?
**Target time:** 45 seconds | **Word count:** ~95

---

[ON SCREEN: Simple title card — "Python Automation for Finance" or a clean static slide
showing two columns: "Excel + VBA is for..." vs. "Python adds..."]

Excel and VBA are great for automating your own workbook — and you saw what that looks like
in Videos 1 through 3.

Python adds value for a different class of problems. Cross-file work — checking one data
export against another, validating a file before analysis starts, or building a repeatable
report from a raw CSV. If you've ever downloaded a billing export and spent thirty minutes
cleaning it before you could even look at it — that's exactly what Python solves.

This video shows four tools a Finance analyst can run without any programming background.

---

## CHAPTER 2 — Safety First
**Target time:** 60 seconds | **Word count:** ~130

---

[ON SCREEN: Open PYTHON_SAFETY.md in Notepad — a clean readable text file, not code.
Scroll slowly through the rules while narrating.]

Before any demo, here is the ground contract.

These scripts run entirely on your own machine. No internet connection, no external calls,
and no AI processing your data somewhere else. Your input files are never modified — every
script opens files read-only and writes all output to a separate folder.

[ON SCREEN: Show the outputs folder path — outputs/YYYYMMDD_HHMMSS_toolname/]

Each run creates a new timestamped folder. Even if you run the same tool ten times, nothing
is ever overwritten. If something goes wrong, you get a plain error message in plain English —
not a wall of code.

Keep your outputs in the outputs folder and do not share them outside the Finance and
Accounting team. Those rules are in writing right here, and this file ships with the toolkit.

---

## CHAPTER 3 — Revenue Leakage Finder
**Target time:** 2 minutes 30 seconds – 3 minutes 30 seconds | **Word count:** ~365

---

[ON SCREEN: Open FinanceTools.xlsm in Excel — the workbook is open on screen.
Show the Finance Tools button on the sheet.]

This is the Finance Automation Toolkit. It ships as a single file from SharePoint.
Everything runs from this one button.

[ON SCREEN: Click the Finance Tools button. A numbered CLI menu appears in a terminal window.]

The question the Revenue Leakage Finder answers is: which customers are being billed
without a matching contract on file — and which ones have contracts that expired months ago
but are still generating invoices?

[ON SCREEN: In the menu, type "1" and press Enter. Show the script processing briefly —
a few lines of output, then "Analysis complete." Output folder path printed at the bottom.]

[ON SCREEN: Switch to browser. The HTML summary report opens automatically.]

Here is the summary. Against 123 contracts and 336 billing records, the tool found
38 exceptions across five categories.

Twelve customers are being billed with no matching contract row. That is the most common
leakage type — and it mirrors exactly what our own AlphaTrust data shows when we run
a contract reconciliation manually.

Ten contracts are expired but still generating invoices. Nine invoices show amounts that
differ from the expected contract value by more than ten percent, with no overage explanation.

[ON SCREEN: Scroll to the ARR waterfall chart in the report.]

This is the ARR waterfall — expected revenue by customer tier compared to what is actually
billing. The gap in the mid-market band is where most of the leakage concentrates. This is
the slide you bring to a revenue review.

[ON SCREEN: Open exceptions_ranked.csv in Excel. Show the ranked rows.]

Down at the row level, each exception is ranked by priority. Row one is the highest-priority
review item. The tool tells you what the problem is, which customer, what billing period,
and what the dollar gap looks like. This is where Finance starts the investigation —
not a hundred rows to sort through manually, but a ranked list with the hardest cases at
the top.

To run this against your own data, you need two CSV files — a contracts file and a billing
export. The structure is documented in the README that ships with the toolkit.

---

## CHAPTER 4 — Data Contract Checker
**Target time:** 90 seconds | **Word count:** ~185

---

[ON SCREEN: Return to the Finance Tools menu — or show it re-opening.
Type "2" and press Enter to launch Data Contract Checker in sample mode.]

Before you run any analysis against a new data export, check that the file is structured
correctly. A renamed column silently breaks everything downstream — and you won't find out
until you're halfway through the analysis.

[ON SCREEN: Show the FAIL output in the terminal window — red text, clear error messages.
E.g., "Missing required column: amount_billed" and "invoice_date: 3 non-date values found."]

Red means something is wrong before we even start. In this example, a required column
is missing and there are non-date values in a date field.

[ON SCREEN: Open the billing CSV in Notepad. Rename one column. Save. Re-run the tool.
Show green PASS output.]

PASS means the file is safe to analyze. FAIL means fix the input first.

This takes ten seconds and saves you from discovering the problem halfway through a
thirty-minute analysis. Run it every time you get a new data export.

---

## CHAPTER 5 — Exception Triage Engine
**Target time:** 90 seconds | **Word count:** ~185

---

[ON SCREEN: Return to the Finance Tools menu. Type "3" and press Enter.
Show the script running briefly, then terminal output showing scores.]

Once you have a list of exceptions — from any analysis, not just the Revenue Leakage
Finder — Python can rank them so you know what to work through first. Not all exceptions
are equal, and you should not spend the same amount of time on every row.

[ON SCREEN: Show scored output briefly — exception classes, customer names, priority scores.]

Each exception gets a priority score based on four factors: dollar impact, how confident
the tool is in the finding, how recently it occurred, and whether the same customer
appears more than once.

[ON SCREEN: Open top_10_action_list.csv in Excel. Show the ranked rows with the
"recommended_action" column visible.]

Row one is your highest-priority review. Each row tells you what the issue is, which
customer, what period, and what action to take. You hand this list to whoever is doing
the review and they know exactly where to start. No judgment calls about what to look
at first.

---

## CHAPTER 6 — Control Evidence Pack
**Target time:** 90 seconds | **Word count:** ~185

---

[ON SCREEN: Return to the Finance Tools menu. Type "4" and press Enter.
Show the script scanning the Revenue Leakage output folder — file names and hashes appearing.]

After any significant analysis — especially one that may go to audit or to leadership —
create an evidence record. This tool scans the output folder from any previous run and
produces a tamper-evident manifest.

[ON SCREEN: Show the manifest output — file names, sizes, timestamps, SHA-256 hashes.]

It logs exactly which files were analyzed, their size, their last-modified date, and a
SHA-256 hash — a unique fingerprint for each file. If a file is changed after this
point, the hash will not match, and you will know.

[ON SCREEN: Open evidence_summary.html in browser — clean one-page summary.]

This is the evidence summary — one page that captures what was run, when, and what
the outputs were. If someone asks six months from now which files were analyzed and
whether the data was changed between the analysis and the review — this folder answers
that question precisely. Attach it to the ticket.

---

## CHAPTER 7 — Finance Automation Launcher
**Target time:** 60 seconds | **Word count:** ~130

---

[ON SCREEN: Return to FinanceTools.xlsm in Excel. Show the single Finance Tools button.
Click it to open the launcher menu one more time.]

You do not need to remember any command-line arguments. Click this button in Excel,
and you get the numbered menu you have been watching all video.

[ON SCREEN: Show the full menu — all 8 options visible.]

Pick a number, press Enter, and the tool runs in sample mode — so you can see results
before you point it at any of your own files.

[ON SCREEN: Select option 7 — outputs folder opens in Windows Explorer.]

Option 7 opens your outputs folder directly in Explorer. Option 6 shows the full
safety rules on screen. Option 8 exits.

This is the entry point. Everything you just saw in the last four chapters is one click
and one number away.

---

## CHAPTER 8 — How to Start
**Target time:** 30 seconds | **Word count:** ~70

---

[ON SCREEN: Simple text card with the four rules — no demo, just clean text on screen.]

Four rules for getting started.

Always run in sample mode first before using your own files.

Start with the supported workflows — they are the tested path.

All output goes to the outputs folder. Your input files are never touched.

And if something does not work, or a result looks wrong — contact Connor in Finance
and Accounting.

---

## WORD COUNT SUMMARY

| Chapter | Words | Est. Time |
|---|---|---|
| 1 — Why Python? | ~95 | ~44 sec |
| 2 — Safety First | ~130 | ~60 sec |
| 3 — Revenue Leakage Finder | ~365 | ~2 min 49 sec |
| 4 — Data Contract Checker | ~185 | ~1 min 26 sec |
| 5 — Exception Triage Engine | ~185 | ~1 min 26 sec |
| 6 — Control Evidence Pack | ~185 | ~1 min 26 sec |
| 7 — Finance Automation Launcher | ~130 | ~60 sec |
| 8 — How to Start | ~70 | ~32 sec |
| **Total** | **~1,345** | **~10 min 23 sec** |

Target range: 9–12 min. This draft lands at ~10:23 — solidly in range.
Adjust by trimming or expanding Chapter 3 (it has the most flex).

---

## RECORDING NOTES FOR CONNOR

**ElevenLabs:** Record narration-only (no [ON SCREEN] text). Each chapter can be one clip,
or record the whole script as one take and split in post.

**Screen recording:** Record separately from narration. Match screen actions to the
timestamps below.

**Rough sync guide:**
- 0:00–0:44 — Chapter 1 (static slide on screen)
- 0:44–1:44 — Chapter 2 (scroll PYTHON_SAFETY.md, show outputs folder)
- 1:44–4:33 — Chapter 3 (Excel button → menu → HTML report → waterfall → ranked CSV)
- 4:33–5:59 — Chapter 4 (menu → FAIL → fix in Notepad → PASS)
- 5:59–7:25 — Chapter 5 (menu → scored output → top_10_action_list.csv in Excel)
- 7:25–8:51 — Chapter 6 (menu → manifest → evidence_summary.html)
- 8:51–9:51 — Chapter 7 (Excel button → full menu → option 7 → option 8)
- 9:51–10:23 — Chapter 8 (text card with 4 rules)

**Chapter 3 demo tip:** Move through the CLI window quickly — it is not the story.
Cut to the HTML report as fast as possible. The browser report is the hero visual.

**Chapter 4 demo tip:** The FAIL → fix → PASS sequence is fast. Do not slow down.
Confidence that "this is easy" is the message — not a detailed walkthrough.

---

*End of narration script. Version v1.0 — 2026-04-28.*
*Review, adjust wording to match your natural speaking style, then record.*
