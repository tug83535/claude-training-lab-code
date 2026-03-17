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
