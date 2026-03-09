# Video 2 — "Full Demo Walkthrough" — AI Narration Script

**Runtime Target:** 15:00 to 18:00
**Audience:** Finance & Accounting team, interested power users
**Purpose:** Detailed tour of the P-and-L demo file — show how things work, step by step
**Scope:** Demo file only. No mention of universal code library (that's Video 3).
**Audio Segments:** 16 clips to generate in ElevenLabs

---

## How to Use This Script

1. Each **SEGMENT** below is one audio clip you generate in ElevenLabs
2. Copy only the text inside the **[PASTE INTO ELEVENLABS]** box
3. Generate the clip, download it, name it using the filename shown
4. The **[YOUR SCREEN ACTIONS]** section tells you what to do on screen while that audio plays
5. The **[TIMING NOTE]** tells you roughly how long each clip should be

---

## SEGMENT 2.0 — Opening

**Filename:** `V2_S0_Opening.mp3`
**Timing:** ~40 seconds
**What's on screen:** Excel file open on the Report landing page

### [PASTE INTO ELEVENLABS]:

```
Welcome to the full walkthrough of the Finance Automation file.

This is a single Excel workbook that automates the monthly P-and-L close process... from importing data, to running quality checks, to generating analysis, building dashboards, and producing final deliverables. Sixty-two actions, all accessible from one control panel.

I'm going to walk you through how it works, chapter by chapter. Everything you see is running live... no slides, no mockups. Let's get into it.
```

### [YOUR SCREEN ACTIONS]:
1. Excel is already open on the Report landing page
2. As the voice starts, slowly scroll down the Report page — let the viewer see it's a real, populated workbook
3. Keep scrolling slowly through the whole clip — no clicking yet
4. Stop scrolling right as the voice says "Let's get into it"

### [TIMING NOTE]:
- Wait 2 seconds of silence before this clip starts (add in editor)
- Measured pace — this is a longer video, no need to rush out of the gate
- "Sixty-two actions" should land clearly and confidently, same energy as Video 1

---

## SEGMENT 2.1 — Chapter 1: The Workbook & Command Center

**Filename:** `V2_S1_Workbook_CommandCenter.mp3`
**Timing:** ~2 minutes
**What's on screen:** Navigating sheet tabs, then opening the Command Center

### [PASTE INTO ELEVENLABS]:

```
Let's start with what's inside this file.

The landing page — Report — gives you a summary of the workbook and quick navigation to any section. Think of it as your table of contents.

The file has over a dozen sheets. You've got the main P-and-L Monthly Trend sheet... this is your core financial data. Revenue and expenses by month, with a full-year total and budget column.

There are Functional P-and-L Summary sheets — one for each month — that break things down by department.

A Product Line Summary sheet showing revenue by product... iGO, Affirm, InsureSight, DocFast.

An Assumptions sheet with the key financial drivers — growth rates, allocation percentages, revenue shares.

And a General Ledger sheet with the raw transaction data.

You don't need to memorize any of this... because everything runs from one place.

This is the Command Center. Every automated action in this file is listed here, organized by category. You can scroll through to browse, or use the search bar to find what you need.

Monthly Operations, Analysis and Reporting, Enterprise Features, Utilities... it's all here. Pick an action, click Run, and it handles the rest.

That's your home base. Every demo from here on out starts from this screen.
```

### [YOUR SCREEN ACTIONS]:
1. Start on the Report landing page
2. **On "P-and-L Monthly Trend sheet"** — click that tab, pause 2 seconds so the viewer sees the data layout
3. **On "Functional P-and-L Summary sheets"** — click one monthly tab (e.g., Jan), pause 1 second
4. **On "Product Line Summary"** — click that tab, pause 2 seconds
5. **On "Assumptions sheet"** — click that tab, pause 1 second
6. **On "General Ledger sheet"** — click that tab, pause 1 second
7. **On "everything runs from one place"** — press Control Shift M to open the Command Center. Wait 2 seconds after it opens.
8. **On "organized by category"** — slowly scroll through the categories (5-6 seconds of scrolling)
9. **On "search bar"** — click the search box, type "reconciliation", show the filtered results, then clear it
10. Leave the Command Center open — it bridges into Chapter 2

### [TIMING NOTE]:
- Don't linger on any sheet too long. This is orientation, not analysis. 1-2 seconds per tab is plenty.
- The Command Center opening is the payoff of this chapter. Give it room to breathe.
- When scrolling through categories, move slowly enough that a viewer could pause the video and read action names
- If ElevenLabs runs the category list too fast, add ellipses between them: "Monthly Operations... Analysis and Reporting... Enterprise Features... Utilities"

---

## SEGMENT 2.2 — Chapter 2A: GL Import

**Filename:** `V2_S2_GL_Import.mp3`
**Timing:** ~50 seconds
**What's on screen:** Running GL Import from the Command Center

### [PASTE INTO ELEVENLABS]:

```
Before you can do any analysis, you need your data.

The General Ledger Import pulls in G-L data from a C-S-V or Excel file with format validation built in. It reads the source file, validates the format, maps the columns, and loads the transactions into the workbook. If something doesn't match the expected structure... it tells you.

What used to take around forty-five minutes of manual copying, pasting, and reformatting — done in about thirty seconds.
```

### [YOUR SCREEN ACTIONS]:
1. In the Command Center, find "Import GL Data" (search or scroll)
2. **On "pulls in G-L data"** — click Run
3. If a file dialog appears, navigate to the source file (have this ready in a known, easy-to-find location)
4. Wait for the import to complete
5. **On "loads the transactions"** — navigate to the General Ledger sheet to show the loaded data
6. Slow scroll through a few rows so the viewer can see real transaction data
7. Pause on the last line — let the time savings sink in

### [TIMING NOTE]:
- Have the source file ready in an obvious location so the file dialog doesn't fumble
- If the import runs fast and there's dead air, that's fine — you'll tighten it in the editor
- This is a setup feature, not a wow feature. Keep it brisk and move on.
- "G-L" and "C-S-V" are spelled out so ElevenLabs says the letters

---

## SEGMENT 2.3 — Chapter 2B: Data Quality Scan

**Filename:** `V2_S3_Data_Quality.mp3`
**Timing:** ~50 seconds
**What's on screen:** Running Data Quality Scan from the Command Center

### [PASTE INTO ELEVENLABS]:

```
Now that the data's loaded, the first thing you want to know is... how clean is it?

The Data Quality Scan checks your entire workbook across six categories — completeness, accuracy, consistency, formatting, outliers, and cross-references.

Right at the top... a letter grade. That tells you at a glance whether your data is ready to work with.

Below that, each category gets its own score and detail. If there are issues, it tells you exactly where — which sheet, which column, what the problem is.

This scan has never been done manually. There was no practical way to do it. Now it takes about fifteen seconds.
```

### [YOUR SCREEN ACTIONS]:
1. Back to the Command Center (Control Shift M)
2. Find "Data Quality Scan" — click Run
3. Wait for completion — the Data Quality Report sheet appears
4. **On "letter grade"** — hover your cursor near the letter grade badge. Hold for 3 seconds. Don't talk over it — let the viewer read.
5. **On "each category gets its own score"** — slowly scroll down through the category breakdown
6. If specific issues are flagged, hover near one briefly
7. **On "fifteen seconds"** — stop scrolling, pause

### [TIMING NOTE]:
- The letter grade badge is a strong visual anchor. Make sure you're hovering near it when the voice says "letter grade."
- After "fifteen seconds" — leave 2 seconds of silence. Let the speed sink in.
- This feature also appeared in Video 1. That's intentional — repetition reinforces. But here you go deeper into the category breakdown.

---

## SEGMENT 2.4 — Chapter 2C: Reconciliation Checks

**Filename:** `V2_S4_Reconciliation.mp3`
**Timing:** ~50 seconds
**What's on screen:** Running Reconciliation Checks from the Command Center

### [PASTE INTO ELEVENLABS]:

```
Next step... make sure all the numbers tie out.

The reconciliation engine runs a series of validation checks across every sheet — verifying that cross-sheet totals match, that revenue and expense lines balance, and that formulas are intact.

Each check gets a clear PASS or FAIL. Green means it ties. Red means something needs attention.

Either way, you know exactly where you stand in ten seconds instead of two hours.
```

### [YOUR SCREEN ACTIONS]:
1. Back to the Command Center
2. Find "Run Reconciliation Checks" — click Run
3. Wait for completion — the Checks sheet appears with PASS/FAIL results
4. **On "PASS or FAIL"** — pause on the scorecard. The green and red colors are immediately readable. Hold 3 seconds.
5. Slowly scroll through all checks
6. If any FAIL items exist, hover your cursor near them
7. **On "ten seconds instead of two hours"** — stop scrolling, let the overlay appear

### [TIMING NOTE]:
- The PASS/FAIL color coding is visually satisfying. Let the viewer see the full list before moving on.
- Read the actual results on screen — don't script specific outcomes since they depend on the demo data state
- The "two hours to ten seconds" comparison is one of your strongest. Let it land.

---

## SEGMENT 2.5 — Chapter 3A: Variance Analysis

**Filename:** `V2_S5_Variance_Analysis.mp3`
**Timing:** ~1 minute
**What's on screen:** Running Variance Analysis from the Command Center

### [PASTE INTO ELEVENLABS]:

```
Your data's in, it's clean, and it reconciles. Now... what's actually happening in the numbers?

The Variance Analysis compares each line item month over month and flags anything that moved more than fifteen percent. Revenue, expenses, margins — it checks everything.

Items over the threshold are highlighted automatically. You can see the dollar change, the percentage change, and whether it's favorable or unfavorable.

Here's an important detail for the finance folks — for expense items, the favorable and unfavorable logic is automatically reversed. A decrease in costs is flagged as favorable, not unfavorable. The system knows the difference.

Instead of scanning hundreds of rows yourself, you get a filtered view of what actually needs your attention.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "Variance Analysis" — click Run
2. Wait for the Variance Analysis sheet to appear
3. **On "flags anything that moved"** — pause on the header row so the viewer sees the column structure
4. **On "highlighted automatically"** — scroll through slowly, pausing on flagged items
5. Hover your cursor near a highlighted item to draw attention
6. **On "filtered view"** — stop scrolling, brief pause

### [TIMING NOTE]:
- The highlighted items are the visual hook. Make sure they're clearly visible.
- The cost-line reversal detail is important for Finance people — mention it, but don't over-explain
- Keep your energy up here. You're building toward the Variance Commentary, which is the real payoff.

---

## SEGMENT 2.6 — Chapter 3B: Variance Commentary

**Filename:** `V2_S6_Variance_Commentary.mp3`
**Timing:** ~1 minute
**What's on screen:** Running Variance Commentary from the Command Center

### [PASTE INTO ELEVENLABS]:

```
This is one of the features I'm most excited about.

You've got your flagged variances. Now the system can write the commentary for you.

These are plain English narratives for the top variances. Each one identifies the line item, states the dollar and percentage change, and describes what happened — in complete sentences, ready to paste into an email, a report, or a board deck.

Let that sit for a second. Read a couple of those.

Writing these manually — pulling the numbers, doing the comparison, putting it into words — that's typically an hour of work. This takes about five seconds.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "Variance Commentary" — click Run
2. Wait for the Variance Commentary sheet to appear
3. **IMPORTANT:** Pause for 2-3 seconds of silence here. Let the viewer take in the full page before the narration continues.
4. **On "plain English narratives"** — slowly scroll through the generated commentary. There should be about 5 narratives.
5. Hover your cursor near one narrative to draw the eye
6. **On "Let that sit for a second"** — stop scrolling. Hold completely still. Give 3-4 seconds of silence.
7. **On "five seconds"** — pause again. Let the impact sit.

### [TIMING NOTE]:
- THIS IS YOUR JAW-DROP MOMENT. Treat it accordingly.
- The silence while narratives are visible is critical. Resist the urge to keep talking. Let people read.
- After "five seconds" — hold for a full beat before moving to the next segment. Don't rush.
- If the generated narratives look particularly good, consider reading one aloud as an ad-lib: "For example..." — but only if it feels natural, not scripted.

---

## SEGMENT 2.7 — Chapter 3C: YoY Variance

**Filename:** `V2_S7_YoY_Variance.mp3`
**Timing:** ~1 minute
**What's on screen:** Running YoY Variance from the Command Center

### [PASTE INTO ELEVENLABS]:

```
Variance Analysis gives you month over month. But leadership often wants year over year — how does this year compare to last year, and how are we tracking against budget?

This builds a full Year-over-Year comparison. Full-year total versus prior year, full-year total versus budget, with dollar and percentage variances for every line.

Same idea — items beyond the threshold are flagged. You get a complete picture of where you're ahead, where you're behind, and by how much.

What would normally take a couple hours of pulling data from two different periods and building the comparison... done in about ten seconds.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "YoY Variance" — click Run
2. Wait for the YoY Variance Analysis sheet to appear
3. **On "full Year-over-Year comparison"** — pause on the header row. Let the viewer see the column structure (FY Total, Prior Year, Budget, Dollar Variance, Percent Variance).
4. **On "items beyond the threshold"** — scroll through slowly, pausing on flagged items
5. **On "about ten seconds"** — stop scrolling, brief pause

### [TIMING NOTE]:
- This feature is important for the Finance audience but less dramatic than Variance Commentary. Keep it solid but don't over-dwell.
- Emphasize the "leadership wants year over year" framing — it connects the feature to a real request they get regularly
- Smooth transition into Chapter 4

---

## SEGMENT 2.8 — Chapter 4A: Dashboard Charts

**Filename:** `V2_S8_Dashboard_Charts.mp3`
**Timing:** ~50 seconds
**What's on screen:** Running Build Dashboard from the Command Center

### [PASTE INTO ELEVENLABS]:

```
You've done the analysis. Now you need to present it.

One click builds eight branded charts in a grid layout — revenue trends, expense breakdowns, margin analysis, product mix, and more. All formatted in brand colors, all properly labeled.

These are the visuals you'd normally build one at a time in a separate PowerPoint or chart tool. Here, they're generated directly from your data in about fifteen seconds.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "Build Dashboard" — click Run
2. Wait for the Charts and Visuals sheet to appear
3. **When it first loads** — pause 2-3 seconds. Let the full chart grid fill the screen. This is a visually impressive moment.
4. **On "revenue trends, expense breakdowns"** — slowly scroll through the chart grid. Pause 1-2 seconds on each chart.
5. Let the branded colors and formatting speak for themselves — no need to point at every element
6. **On "fifteen seconds"** — stop scrolling, hold

### [TIMING NOTE]:
- The chart grid is visually rich. Give the viewer a full-screen moment when it first appears.
- Don't describe every single chart. "Revenue trends, expense breakdowns, margin analysis, product mix" covers it. The visual does the work.
- "Brand colors" signals this is presentation-ready, not generic Excel charts

---

## SEGMENT 2.9 — Chapter 4B: Executive Dashboard

**Filename:** `V2_S9_Executive_Dashboard.mp3`
**Timing:** ~50 seconds
**What's on screen:** Running Executive Dashboard from the Command Center

### [PASTE INTO ELEVENLABS]:

```
For a more focused leadership view, there's the Executive Dashboard.

This puts everything on one sheet — K-P-I summary cards across the top, a waterfall chart showing how you get from budget to actual, and a product line comparison at the bottom.

This is designed to be the one sheet you pull up when the C-F-O asks... "how are we doing this month?" One click, one sheet, full picture.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "Executive Dashboard" — click Run
2. Wait for the Executive Dashboard sheet to appear
3. **On "K-P-I summary cards"** — hover near the KPI cards at the top. Hold 2 seconds.
4. **On "waterfall chart"** — scroll down to the waterfall chart. Hold 2 seconds.
5. **On "product line comparison"** — scroll down to the product comparison. Hold 2 seconds.
6. **On "how are we doing this month"** — hold the screen still. Let the full dashboard be visible.

### [TIMING NOTE]:
- The KPI cards, waterfall, and product comparison are three distinct visual elements. Give each one a moment.
- "K-P-I" and "C-F-O" are spelled out so ElevenLabs says the letters
- "How are we doing this month" — say this like you're quoting someone walking into your office. Conversational, not formal.
- Don't say "the C-F-O" in a way that feels like name-dropping. It's a relatable scenario, not a reference to your actual CFO.

---

## SEGMENT 2.10 — Chapter 4C: PDF Export

**Filename:** `V2_S10_PDF_Export.mp3`
**Timing:** ~50 seconds
**What's on screen:** Running PDF Export from the Command Center

### [PASTE INTO ELEVENLABS]:

```
When you need a final deliverable — something you can email, save to a shared drive, or print — the P-D-F Export handles it.

It takes key sheets from the workbook and compiles them into a single, clean P-D-F. Each page has proper headers and footers — the report title, the date, page numbers. Formatted for printing or sharing.

Manually formatting and exporting multiple sheets to a clean P-D-F... that's easily thirty minutes of adjusting print areas, fixing page breaks, and hoping nothing shifts. This takes about ten seconds.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "PDF Export" — click Run
2. Wait for the export to complete
3. If the PDF opens automatically, pause on the first page — let the viewer see the formatting
4. Scroll to page 2 to show headers, footers, and clean formatting
5. If it doesn't auto-open, navigate to the output file and open it briefly
6. **On "hoping nothing shifts"** — a quick smile if you feel it. This is a relatable moment.
7. **On "ten seconds"** — hold the PDF on screen for 2-3 seconds

### [TIMING NOTE]:
- "P-D-F" is spelled out so ElevenLabs says the letters, not "puhdf"
- If the PDF looks crisp and professional on screen, hold on it for a few seconds. It sells itself.
- "Hoping nothing shifts" — that small moment of relatability lands with anyone who's fought Excel print formatting

---

## SEGMENT 2.11 — Chapter 5A: Executive Mode

**Filename:** `V2_S11_Executive_Mode.mp3`
**Timing:** ~30 seconds
**What's on screen:** Toggling Executive Mode on and off

### [PASTE INTO ELEVENLABS]:

```
When leadership needs to review the file, they don't need to see every technical sheet. Executive Mode cleans it up.

One click hides all the working sheets and leaves only the presentation-ready views. Toggle it off... and everything comes back.

Simple, but it makes a big difference when you're sharing the file with someone who just wants the highlights.
```

### [YOUR SCREEN ACTIONS]:
1. First, make sure the full tab bar is visible at the bottom — lots of sheet tabs showing
2. **On "Executive Mode cleans it up"** — toggle Executive Mode ON (from the Command Center, or use the shortcut). Watch tabs disappear, leaving only key sheets.
3. Pause 2 seconds — let the viewer see the clean state
4. **On "Toggle it off"** — toggle Executive Mode OFF. Tabs reappear.
5. Brief pause

### [TIMING NOTE]:
- This is a quick hit — don't over-explain. The visual toggle is self-explanatory.
- The disappearing and reappearing tabs is a satisfying visual. Let it happen without talking over it.
- 30 seconds max. In, out, move on.

---

## SEGMENT 2.12 — Chapter 5B: Version Control

**Filename:** `V2_S12_Version_Control.mp3`
**Timing:** ~40 seconds
**What's on screen:** Saving a version snapshot

### [PASTE INTO ELEVENLABS]:

```
Version Control lets you save a snapshot of the entire workbook at any point... and compare or restore it later.

You give it a name — "Pre-Close" or "March Draft 1" — and it saves the full state. If something goes wrong, or if someone overwrites your work, you go back to Version Control, pick a snapshot, and restore it.

Every snapshot is timestamped and logged. You always know what changed and when.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find the "Version Control" area
2. Run Save Snapshot — enter a name when prompted (e.g., "March Draft 1")
3. **On "saves the full state"** — show the confirmation or the snapshot list
4. Briefly show the Restore option (don't actually restore — just show it exists)
5. Brief pause at the end

### [TIMING NOTE]:
- The use case "someone overwrites your work" is universally relatable. It earns a knowing nod.
- Don't actually demonstrate a restore — it would require setup and adds time. Just showing the option exists is enough.
- Keep it tight. 40 seconds.

---

## SEGMENT 2.13 — Chapter 5C: Scenario Management + 5D: Sensitivity Analysis

**Filename:** `V2_S13_Scenarios_Sensitivity.mp3`
**Timing:** ~1 minute 20 seconds
**What's on screen:** Scenario Management, then Sensitivity Analysis

### [PASTE INTO ELEVENLABS]:

```
Scenario Management lets you save and load different sets of assumptions. You can set up a Base Case, an Optimistic case, a Conservative case — each with different growth rates, allocation percentages, whatever drivers matter.

Switch between them with one click and the entire workbook recalculates. You can also compare scenarios side by side to see the impact of different assumptions.

Now... Sensitivity Analysis takes this a step further. Instead of switching between preset scenarios, you can run what-if analysis on any key assumption and see how it ripples through the entire P-and-L.

What happens to total revenue if growth is two percent higher? What happens to margins if your allocation changes by five points? This answers those questions instantly.

Doing this manually — changing an assumption, recalculating, recording the result, changing it back, trying the next one — that's four or more hours of tedious work. This runs all of them in about twenty seconds.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "Scenario Management"
2. Show the scenario list (if scenarios are pre-saved, show Base Case, Optimistic, etc.)
3. Load one scenario — show the Assumptions sheet values change
4. **On "Sensitivity Analysis takes this a step further"** — back to Command Center, find "Sensitivity Analysis", click Run
5. Wait for the Sensitivity Analysis sheet to appear
6. **On "What happens to total revenue"** — pause on the output, let the viewer see the structure
7. **On "answers those questions instantly"** — slowly scroll through the results
8. **On "about twenty seconds"** — stop scrolling, let the overlay appear

### [TIMING NOTE]:
- The what-if questions should sound conversational — like you're in a meeting and someone asks them out loud
- The four-hours-to-twenty-seconds comparison is dramatic. Let the overlay stay visible for 3-4 seconds.
- This is the last feature demo before "Under the Hood." End with energy.
- If the Assumptions sheet values visibly change when you load a scenario, that's your visual proof. Even a few cells changing is enough.

---

## SEGMENT 2.14 — Chapter 6A: Integration Test

**Filename:** `V2_S14_Integration_Test.mp3`
**Timing:** ~50 seconds
**What's on screen:** Running the Integration Test from the Command Center

### [PASTE INTO ELEVENLABS]:

```
With this many automated actions, you need to know the system is working correctly.

The Integration Test runs eighteen automated checks across the entire workbook — sheet existence, data integrity, formula health, macro functionality.

Eighteen out of eighteen... all passing.

This runs every time you want to verify the file is in a good state. Before a close, after making changes, anytime you want peace of mind — one click.
```

### [YOUR SCREEN ACTIONS]:
1. Command Center — find "Integration Test" — click Run
2. Wait for the test to complete — results appear
3. **On "Eighteen out of eighteen"** — pause on the 18/18 PASS result. Hold 3 seconds. This is a confidence-building moment.
4. If results are listed individually, slowly scroll through them
5. **On "one click"** — brief pause

### [TIMING NOTE]:
- "Eighteen out of eighteen... all passing" — say this with quiet confidence. Not bragging, just stating a fact.
- Don't explain what each individual test does. The categories ("sheet existence, data integrity, formula health, macro functionality") cover it. The 18/18 result is what matters.
- "Peace of mind" resonates with anyone who's worked with complex spreadsheets

---

## SEGMENT 2.15 — Chapter 6B: Audit Log

**Filename:** `V2_S15_Audit_Log.mp3`
**Timing:** ~40 seconds
**What's on screen:** Viewing the Audit Log sheet

### [PASTE INTO ELEVENLABS]:

```
Every action you run is logged automatically.

The Audit Log records a timestamp, the module that ran, and the result for every single action. If you need to know who ran what, and when... it's all here.

You'll see entries from everything we just did — every import, every scan, every export. Full traceability.
```

### [YOUR SCREEN ACTIONS]:
1. Navigate to the Audit Log sheet (unhide if hidden, or access through the Command Center)
2. **On "records a timestamp"** — show the log entries. They should include timestamps from the demo you just ran.
3. Slowly scroll through, hovering near a few entries
4. **On "everything we just did"** — pause and let the viewer see the entries match the features you demoed
5. Brief pause at the end

### [TIMING NOTE]:
- The audit log filled with entries from this very demo is a nice moment — makes it tangible and real
- Don't spend too long here. The point is: it exists, it's automatic, it's complete. 40 seconds.
- If the Audit Log is on a hidden sheet, briefly show the unhide action — that itself is a feature (hidden from casual users, available when needed)

---

## SEGMENT 2.16 — Chapter 7: Closing

**Filename:** `V2_S16_Closing.mp3`
**Timing:** ~1 minute
**What's on screen:** Report landing page, then closing title card

### [PASTE INTO ELEVENLABS]:

```
That's the full walkthrough. Let me recap what we covered.

We imported G-L data, checked data quality, ran reconciliation, analyzed variances month over month and year over year, generated written commentary, built dashboards and an executive view, exported a clean P-D-F, managed scenarios and sensitivity analysis, ran a full integration test, and reviewed the audit trail.

All from one Excel file, all through the Command Center, all in a matter of minutes.

If you want to explore the file yourself, it's available on SharePoint along with step-by-step training guides for everything you just saw. If you run into any questions... reach out. I'm happy to help.

Thanks for watching.
```

### [YOUR SCREEN ACTIONS]:
1. Navigate back to the Report landing page for the recap
2. During the recap list, keep the landing page static — don't click around. Let the words carry it.
3. **On "it's available on SharePoint"** — this is where you'll cut to the closing title card in the editor
4. Keep the title card on screen for 5+ seconds after "Thanks for watching"

### [TIMING NOTE]:
- The recap should be spoken at a slightly faster pace than the rest of the video — it's a summary, not a re-explanation
- Don't re-demo anything during the recap. Just list it.
- "G-L" and "P-D-F" spelled out for correct ElevenLabs pronunciation
- "Reach out — I'm happy to help" is the right CTA for this audience. They're the Finance team. They'll have questions. Make it easy.
- End confidently. No "so yeah, that's it" energy. Clear, professional close.
- Leave 3 seconds of silence after "Thanks for watching" before the clip ends
- Add closing music sting in the video editor (not in the audio clip)

---

## Title Card Specs

You'll add these as static images in the video editor (not audio clips):

### Opening Title Card (5 seconds, at the very start)
```
Finance Automation
Full Demo Walkthrough
```
- Background: Brand Blue (#0B4779)
- Text: White, Arial Bold
- Add company logo if permitted
- Brief music sting (2-3 seconds), then fade to silence

### Closing Title Card (5-8 seconds, at the very end)
```
Finance Automation

Demo File & Training Guides
Available on SharePoint

Questions? Contact Connor
```
- Same styling as opening card
- Brief music, fade out

---

## Chapter Card Specs

Each chapter card is a static image shown for 3 seconds between segments. Add in the video editor (not audio clips). Brief music accent or silence.

### Chapter 1 Card
```
Chapter 1
The Workbook & Command Center
Your home base for everything
```

### Chapter 2 Card
```
Chapter 2
Data Import & Quality
Getting your data in — and making sure it's clean
```

### Chapter 3 Card
```
Chapter 3
Analysis
Making sense of your numbers
```

### Chapter 4 Card
```
Chapter 4
Reporting & Visuals
Turning analysis into deliverables
```

### Chapter 5 Card
```
Chapter 5
Enterprise Features
Power tools for control and flexibility
```

### Chapter 6 Card
```
Chapter 6
Under the Hood
Built to be trusted
```

### Chapter 7 Card
```
Chapter 7
Next Steps
Where to go from here
```

**Style for all chapter cards:**
- Background: Brand Blue (#0B4779)
- Chapter number: White, Arial Bold, large
- Title: White, Arial Bold, medium
- Subtitle: Arctic White (#F9F9F9), Arial Regular, smaller
- Centered on screen

---

## Full Audio Generation Checklist

| # | Segment | Filename | Duration | Generated? |
|---|---------|----------|----------|------------|
| 0 | Opening | V2_S0_Opening.mp3 | ~40 sec | [ ] |
| 1 | Ch 1: Workbook & Command Center | V2_S1_Workbook_CommandCenter.mp3 | ~2:00 | [ ] |
| 2 | Ch 2A: GL Import | V2_S2_GL_Import.mp3 | ~50 sec | [ ] |
| 3 | Ch 2B: Data Quality Scan | V2_S3_Data_Quality.mp3 | ~50 sec | [ ] |
| 4 | Ch 2C: Reconciliation Checks | V2_S4_Reconciliation.mp3 | ~50 sec | [ ] |
| 5 | Ch 3A: Variance Analysis | V2_S5_Variance_Analysis.mp3 | ~1:00 | [ ] |
| 6 | Ch 3B: Variance Commentary | V2_S6_Variance_Commentary.mp3 | ~1:00 | [ ] |
| 7 | Ch 3C: YoY Variance | V2_S7_YoY_Variance.mp3 | ~1:00 | [ ] |
| 8 | Ch 4A: Dashboard Charts | V2_S8_Dashboard_Charts.mp3 | ~50 sec | [ ] |
| 9 | Ch 4B: Executive Dashboard | V2_S9_Executive_Dashboard.mp3 | ~50 sec | [ ] |
| 10 | Ch 4C: PDF Export | V2_S10_PDF_Export.mp3 | ~50 sec | [ ] |
| 11 | Ch 5A: Executive Mode | V2_S11_Executive_Mode.mp3 | ~30 sec | [ ] |
| 12 | Ch 5B: Version Control | V2_S12_Version_Control.mp3 | ~40 sec | [ ] |
| 13 | Ch 5C-D: Scenarios + Sensitivity | V2_S13_Scenarios_Sensitivity.mp3 | ~1:20 | [ ] |
| 14 | Ch 6A: Integration Test | V2_S14_Integration_Test.mp3 | ~50 sec | [ ] |
| 15 | Ch 6B: Audit Log | V2_S15_Audit_Log.mp3 | ~40 sec | [ ] |
| 16 | Ch 7: Closing | V2_S16_Closing.mp3 | ~1:00 | [ ] |
| | **TOTAL NARRATION** | | **~14:20** | |
| | + title cards + chapter cards + pauses | | **~16:00-17:30** | |

---

## Total Runtime Breakdown

| Section | Duration | Cumulative |
|---------|----------|------------|
| Opening Title Card | 0:05 | 0:05 |
| Opening Narration | 0:40 | 0:45 |
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
| Ch 5C-D: Scenarios + Sensitivity | 1:20 | 13:30 |
| Ch 6 Card | 0:03 | 13:33 |
| Ch 6A: Integration Test | 0:50 | 14:23 |
| Ch 6B: Audit Log | 0:40 | 15:03 |
| Ch 7 Card | 0:03 | 15:06 |
| Ch 7: Closing | 1:00 | 16:06 |
| Closing Title Card | 0:05 | 16:11 |
| **TOTAL** | | **~16:11** |

Buffer for natural pauses, macro run times, and breathing room: expect **16:30-17:30** in practice.

---

*Created: 2026-03-09 | Part of AI Narration Scripts package*
