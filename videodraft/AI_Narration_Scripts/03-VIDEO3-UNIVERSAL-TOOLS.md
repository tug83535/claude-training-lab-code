# Video 3 — "Universal Tools" — AI Narration Script

**Runtime Target:** 8:00 to 10:00
**Audience:** All employees
**Purpose:** Show that the VBA and Python tools work on ANY Excel file, not just the P&L demo. Demonstrate a handful of the ~100 tools on a regular messy spreadsheet.
**Audio Segments:** 14 clips to generate in ElevenLabs
**Demo File:** Sample_Quarterly_Report.xlsx (NOT the P&L demo file)

---

## How to Use This Script

1. Each **SEGMENT** below is one audio clip you generate in ElevenLabs
2. Copy only the text inside the **[PASTE INTO ELEVENLABS]** box
3. Generate the clip, download it, name it using the filename shown
4. The **[YOUR SCREEN ACTIONS]** section tells you what to do on screen while that audio plays
5. The **[TIMING NOTE]** tells you roughly how long each clip should be

---

## SEGMENT 3.0 — Opening

**Filename:** `V3_S0_Opening.mp3`
**Timing:** ~45 seconds
**What's on screen:** Sample_Quarterly_Report.xlsx open — a messy, realistic spreadsheet

### [PASTE INTO ELEVENLABS]:

```
In the first video, I showed you what automation can do for a P-and-L close process... inside one specific file.

But most of those tools aren't locked to that file. There are close to eighty V-B-A tools and twenty-two Python scripts that work on any Excel spreadsheet you've got.

So I'm going to open a completely different file... just a regular messy quarterly report... and show you a handful of these tools in action.

Everything I'm about to show you is available on SharePoint, with step-by-step guides and Copilot prompts to help you get started.
```

### [YOUR SCREEN ACTIONS]:
1. Have Sample_Quarterly_Report.xlsx already open — make sure the data looks messy (blank rows, merged cells, text-stored numbers, etc.)
2. Slowly scroll through the spreadsheet as the voice starts — let the viewer see it's a real, imperfect file
3. **On "regular messy quarterly report"** — pause scrolling so the viewer can see the mess
4. **On "available on SharePoint"** — stop scrolling, hold still

### [TIMING NOTE]:
- Wait 2 seconds of silence before this clip starts (add in editor)
- This opening establishes a key point: these tools are universal. Let that land.
- "eighty V-B-A tools and twenty-two Python scripts" — if ElevenLabs rushes it, add a comma: "close to eighty V-B-A tools, and twenty-two Python scripts"

---

## Chapter 1 — Data Cleanup

---

## SEGMENT 3.1A — Delete Blank Rows

**Filename:** `V3_S1A_Delete_Blank_Rows.mp3`
**Timing:** ~30 seconds
**What's on screen:** Running Delete Blank Rows on the sample file

### [PASTE INTO ELEVENLABS]:

```
First up... blank rows. Every spreadsheet has them, and they break everything... filters, formulas, sorting... all of it.

One click, and they're gone. It backs up your sheet first, so nothing's lost. But the blank rows... gone.
```

### [YOUR SCREEN ACTIONS]:
1. Scroll to show some blank rows scattered in the data — make sure they're visible
2. **On "One click"** — run the Delete Blank Rows tool
3. Wait for it to complete — let the viewer see the before and after
4. **On "gone"** — pause, let the clean result sit on screen for 2 seconds

### [TIMING NOTE]:
- The before/after contrast is the whole point. Make sure blank rows are clearly visible before you run it.
- Quick, punchy segment. Don't linger.

---

## SEGMENT 3.1B — Remove Leading/Trailing Spaces

**Filename:** `V3_S1B_Remove_Spaces.mp3`
**Timing:** ~30 seconds
**What's on screen:** Running Remove Spaces on the sample file

### [PASTE INTO ELEVENLABS]:

```
Next... invisible spaces. You can't see them, but they're there... hiding at the beginning or end of cell values.

They'll break a V-LOOKUP, mess up a filter, make two things that look identical... not match. One click strips them all out.
```

### [YOUR SCREEN ACTIONS]:
1. Click on a cell that has a leading or trailing space — show in the formula bar that there's an extra space
2. **On "One click"** — run the Remove Spaces tool
3. Click the same cell again — show in the formula bar that the space is gone
4. Hold for 2 seconds so the viewer sees the difference

### [TIMING NOTE]:
- The formula bar is the key visual here. Make sure it's visible and the viewer can see the space before and after.
- "V-LOOKUP" is spelled out for correct ElevenLabs pronunciation

---

## SEGMENT 3.1C — Convert Text to Numbers

**Filename:** `V3_S1C_Text_To_Numbers.mp3`
**Timing:** ~30 seconds
**What's on screen:** Running Convert Text to Numbers on the sample file

### [PASTE INTO ELEVENLABS]:

```
This one's a classic. Numbers stored as text. You try to SUM a column and get zero... because Excel thinks they're words, not numbers.

You can usually spot them... little green triangles in the corner of the cell. One click converts them all to real numbers.
```

### [YOUR SCREEN ACTIONS]:
1. Show a column with green triangle indicators (text-stored numbers)
2. Click on a cell with the green triangle — show the warning icon
3. Show a SUM formula returning zero or ignoring those cells
4. **On "One click"** — run the Convert Text to Numbers tool
5. Show the SUM formula now returning the correct total
6. Hold for 2 seconds

### [TIMING NOTE]:
- The green triangles are the visual proof. Make sure they're visible on screen before the fix.
- The SUM going from zero to the real number is the payoff moment.

---

## SEGMENT 3.1D — Unmerge Cells and Fill Down

**Filename:** `V3_S1D_Unmerge_Fill.mp3`
**Timing:** ~30 seconds
**What's on screen:** Running Unmerge and Fill Down on the sample file

### [PASTE INTO ELEVENLABS]:

```
Merged cells. They look nice in a printed report, but try sorting or filtering with merged cells and Excel loses its mind.

This tool unmerges everything and fills the values down into every row... so your data actually works like data.
```

### [YOUR SCREEN ACTIONS]:
1. Show a section with merged cells — click on one to highlight the merge
2. Try to sort or filter briefly to show it doesn't work properly (optional — only if it's quick)
3. **On "This tool"** — run the Unmerge and Fill Down tool
4. Show the result — each row now has its own value, no more merges
5. Hold for 2 seconds

### [TIMING NOTE]:
- "Excel loses its mind" should sound casual and real — not scripted
- The fill-down result is visually clear. Let it breathe on screen.

---

## SEGMENT 3.1E — Full Data Sanitize Preview

**Filename:** `V3_S1E_Sanitize_Preview.mp3`
**Timing:** ~30 seconds
**What's on screen:** Running the Sanitize Preview, then the Full Sanitize

### [PASTE INTO ELEVENLABS]:

```
And if you don't want to run these one at a time... there's a full sanitize option that handles everything at once.

But here's the nice part... you can preview first. It shows you exactly what it would change... without touching anything. When you're comfortable, run the full sanitize and it cleans it all up.
```

### [YOUR SCREEN ACTIONS]:
1. **On "preview first"** — run the Preview Sanitize tool
2. Wait for the preview report to appear — scroll through it slowly so the viewer can see the list of proposed changes
3. **On "run the full sanitize"** — run the Full Sanitize tool
4. Show the cleaned result — hold for 2 seconds

### [TIMING NOTE]:
- The preview report is the trust moment. The viewer needs to see it's not a black box — you can check before committing.
- Give the preview report at least 3-4 seconds on screen.

---

## Chapter 2 — Formatting and Audit

---

## SEGMENT 3.2A — Apply Company Branding

**Filename:** `V3_S2A_Branding.mp3`
**Timing:** ~40 seconds
**What's on screen:** Running the branding tool on the sample file

### [PASTE INTO ELEVENLABS]:

```
Now let's make this thing look professional.

This tool detects your headers and total rows automatically, then applies company brand colors... the blues, the alternating row shading, Arial font... the whole look. One click.

It's not going to win a design award, but it takes you from messy spreadsheet to presentation-ready in about two seconds. And it works on whatever sheet you're on.
```

### [YOUR SCREEN ACTIONS]:
1. Make sure the active sheet looks messy and unstyled — default fonts, no colors
2. **On "This tool"** — run the Apply Branding tool
3. Wait for it to complete — the transformation should be dramatic
4. **On "presentation-ready"** — slowly scroll down through the styled sheet so the viewer sees the full effect
5. **On "whatever sheet you're on"** — stop scrolling, hold

### [TIMING NOTE]:
- The before/after transformation is the visual highlight. Make sure the "before" looks genuinely messy.
- Give the styled result at least 3 seconds on screen. Let people appreciate it.
- "It's not going to win a design award" — keep this light and real. It's self-aware humor.

---

## SEGMENT 3.2B — AutoFit All Columns

**Filename:** `V3_S2B_AutoFit.mp3`
**Timing:** ~20 seconds
**What's on screen:** Running AutoFit on the sample file

### [PASTE INTO ELEVENLABS]:

```
Quick one. AutoFit every column on the sheet... so nothing's cut off, nothing's too wide. It's small, but it's satisfying.
```

### [YOUR SCREEN ACTIONS]:
1. Make sure some columns are too narrow (data cut off) and some are too wide
2. Run the AutoFit tool
3. Let the viewer see all columns snap to the right width
4. Hold for 2 seconds

### [TIMING NOTE]:
- This is a palate cleanser between the bigger demos. Keep it snappy.
- The visual snap of all columns resizing at once is the payoff.

---

## SEGMENT 3.2C — Find External Links

**Filename:** `V3_S2C_External_Links.mp3`
**Timing:** ~40 seconds
**What's on screen:** Running Find External Links on the sample file

### [PASTE INTO ELEVENLABS]:

```
This one's a lifesaver if you've ever inherited someone else's spreadsheet.

It scans your entire workbook and finds every cell that links to another file. It shows you the cell address, what file it's pointing to, and the formula behind it.

If you've ever gotten a "Update Links?" popup and had no idea where those links are... this tells you exactly where they're hiding.
```

### [YOUR SCREEN ACTIONS]:
1. **On "It scans"** — run the Find External Links tool
2. Wait for the report to appear
3. **On "cell address, what file"** — slowly scroll through the results, letting the viewer see the detail
4. **On "where they're hiding"** — stop scrolling, hold

### [TIMING NOTE]:
- "inherited someone else's spreadsheet" — this is relatable. Let it land.
- The report is the proof. Give it 3-4 seconds of screen time.

---

## SEGMENT 3.2D — Audit Hidden Sheets

**Filename:** `V3_S2D_Hidden_Sheets.mp3`
**Timing:** ~30 seconds
**What's on screen:** Running Audit Hidden Sheets on the sample file

### [PASTE INTO ELEVENLABS]:

```
Hidden sheets. Every workbook has them... and some have very hidden sheets that you can't even unhide from the right-click menu.

This tool finds all of them and gives you a report... sheet name, visibility status, whether it has data. No more surprises.
```

### [YOUR SCREEN ACTIONS]:
1. Right-click on a sheet tab to show the normal Unhide option — briefly show that some sheets are hidden
2. **On "This tool"** — run the Audit Hidden Sheets tool
3. Wait for the report — scroll through to show sheet names and visibility status
4. **On "No more surprises"** — stop scrolling, hold

### [TIMING NOTE]:
- "very hidden sheets that you can't even unhide from the right-click menu" — this is a surprise for most people. Let the voice convey that.
- Keep this segment moving. It's a quick reveal.

---

## SEGMENT 3.2E — Workbook Metadata Reporter

**Filename:** `V3_S2E_Metadata.mp3`
**Timing:** ~20 seconds
**What's on screen:** Running the Metadata Reporter on the sample file

### [PASTE INTO ELEVENLABS]:

```
And for a quick health check... the metadata reporter. It counts your sheets, formulas, named ranges, and gives you the big picture in a few seconds. Useful when you first open a file you've never seen before.
```

### [YOUR SCREEN ACTIONS]:
1. Run the Workbook Metadata Reporter tool
2. Let the report appear — show the summary stats
3. Hold for 2-3 seconds so the viewer can scan the numbers

### [TIMING NOTE]:
- Quick and clean. This is a utility, not a showpiece.
- "when you first open a file you've never seen before" — relatable. Let it connect.

---

## Chapter 3 — Python Tools

---

## SEGMENT 3.3A — Python Introduction + File Comparison

**Filename:** `V3_S3A_Python_File_Compare.mp3`
**Timing:** ~40 seconds
**What's on screen:** Running the File Comparison Python script

### [PASTE INTO ELEVENLABS]:

```
Now... a few Python tools. And I want to be clear about something — you don't need to know Python to use these. You double-click a file, it asks you a couple of questions, it runs, and it gives you an Excel file with the results. That's it.

First one... file comparison. Say you've got two versions of a report and you need to know what changed. You point it at both files, and it gives you a side-by-side diff... every cell that's different, highlighted and labeled.
```

### [YOUR SCREEN ACTIONS]:
1. Show the Python script file briefly (just the icon — don't open the code)
2. **On "You double-click"** — double-click the file comparison script
3. When prompted, select two Excel files
4. Wait for it to run — show the output Excel file opening
5. **On "side-by-side diff"** — scroll through the comparison report slowly
6. **On "highlighted and labeled"** — pause, let the result sit on screen

### [TIMING NOTE]:
- "you don't need to know Python" is the most important line in this chapter. Make sure it lands clearly.
- The double-click and the output file are the two key visuals. Show both clearly.
- If the script takes a few seconds to run, that's fine — add a brief speed-up in the editor if needed.

---

## SEGMENT 3.3B — Fuzzy Lookup

**Filename:** `V3_S3B_Fuzzy_Lookup.mp3`
**Timing:** ~40 seconds
**What's on screen:** Running the Fuzzy Lookup Python script

### [PASTE INTO ELEVENLABS]:

```
This one's really useful for reconciliation work. Fuzzy lookup.

You've got a list of company names from one system and a list from another... and they don't match exactly. One says "I-B-M Corporation" and the other says "I-B-M Corp." A V-LOOKUP won't find that.

Fuzzy lookup will. It matches names that are close but not identical... and gives you a confidence score for each match. Super helpful when you're trying to reconcile two data sources.
```

### [YOUR SCREEN ACTIONS]:
1. **On "Fuzzy lookup"** — double-click the fuzzy lookup script
2. When prompted, select the input file with mismatched names
3. Wait for it to run — show the output Excel file opening
4. **On "confidence score"** — scroll through the matches, showing the match pairs and scores
5. Hold for 2-3 seconds so the viewer can see the results

### [TIMING NOTE]:
- "I-B-M Corporation" and "I-B-M Corp" — spelled out for ElevenLabs pronunciation
- "V-LOOKUP" spelled out for correct pronunciation
- The confidence score column is the visual proof. Make sure it's visible on screen.

---

## SEGMENT 3.3C — PDF Extractor

**Filename:** `V3_S3C_PDF_Extractor.mp3`
**Timing:** ~40 seconds
**What's on screen:** Running the PDF Extractor Python script

### [PASTE INTO ELEVENLABS]:

```
Last one. You've got a P-D-F with tables in it... and you need that data in Excel. We've all been there.

This tool pulls the tables right out of the P-D-F and drops them into an Excel file. It's not perfect on every P-D-F — some are cleaner than others — but when it works, it saves you an hour of copy-pasting and reformatting.
```

### [YOUR SCREEN ACTIONS]:
1. Show a PDF file with a table in it — let the viewer see the source
2. **On "This tool"** — double-click the PDF extractor script
3. When prompted, select the PDF
4. Wait for it to run — show the output Excel file opening
5. Scroll through the extracted table — show that the data made it into Excel cleanly
6. Hold for 2-3 seconds

### [TIMING NOTE]:
- "We've all been there" — conversational and relatable. Let the voice sit in that moment.
- "It's not perfect on every P-D-F" — honesty builds trust. Don't skip this line.
- "P-D-F" spelled out every time for correct pronunciation

---

## SEGMENT 3.4 — Closing

**Filename:** `V3_S4_Closing.mp3`
**Timing:** ~30 seconds
**What's on screen:** Back on Sample_Quarterly_Report.xlsx, then closing title card

### [PASTE INTO ELEVENLABS]:

```
That's just a handful. There are about a hundred tools total between the V-B-A and the Python... and they all work on whatever spreadsheet you've got open.

Everything's on SharePoint... the tools, the guides, the Copilot prompts. Grab what you need, try it out. If you get stuck, the guides walk you through every step.

Thanks for watching.
```

### [YOUR SCREEN ACTIONS]:
1. Navigate back to the sample file — show the now-cleaned, branded spreadsheet
2. Let the viewer see the transformation from the messy file they saw at the start
3. **On "Everything's on SharePoint"** — this is where you'll cut to the closing title card in the editor
4. Keep the title card on screen for 5+ seconds after "Thanks for watching"

### [TIMING NOTE]:
- "about a hundred tools total" — confident, not boastful
- "Grab what you need, try it out" — casual invitation, not a hard sell
- Leave 3 seconds of silence after "Thanks for watching" before the clip ends
- Add closing music sting in the video editor (not in the audio clip)

---

## Title Card Specs

You'll add these as static images in the video editor (not audio clips):

### Opening Title Card (5 seconds, at the very start)
```
Finance Automation
Universal Tools
```
- Background: Brand Blue (#0B4779)
- Text: White, Arial Bold
- Add company logo if permitted
- Brief music sting (2-3 seconds), then fade to silence

### Chapter Cards (3 seconds each, between chapters)
```
Chapter 1: Data Cleanup
```
```
Chapter 2: Formatting & Audit
```
```
Chapter 3: Python Tools
```
- Background: Navy (#112E51)
- Text: White, Arial Bold
- No music — just a clean visual break

### Closing Title Card (5-8 seconds, at the very end)
```
Finance Automation

~100 Tools | VBA + Python
Guides | Copilot Prompts
Available on SharePoint

Questions? Contact Connor
```
- Same styling as opening card
- Brief music, fade out

---

## Full Audio Generation Checklist

| # | Segment | Filename | Duration | Generated? |
|---|---------|----------|----------|------------|
| 1 | Opening | V3_S0_Opening.mp3 | ~45 sec | [ ] |
| 2 | Delete Blank Rows | V3_S1A_Delete_Blank_Rows.mp3 | ~30 sec | [ ] |
| 3 | Remove Spaces | V3_S1B_Remove_Spaces.mp3 | ~30 sec | [ ] |
| 4 | Text to Numbers | V3_S1C_Text_To_Numbers.mp3 | ~30 sec | [ ] |
| 5 | Unmerge & Fill | V3_S1D_Unmerge_Fill.mp3 | ~30 sec | [ ] |
| 6 | Sanitize Preview | V3_S1E_Sanitize_Preview.mp3 | ~30 sec | [ ] |
| 7 | Branding | V3_S2A_Branding.mp3 | ~40 sec | [ ] |
| 8 | AutoFit | V3_S2B_AutoFit.mp3 | ~20 sec | [ ] |
| 9 | External Links | V3_S2C_External_Links.mp3 | ~40 sec | [ ] |
| 10 | Hidden Sheets | V3_S2D_Hidden_Sheets.mp3 | ~30 sec | [ ] |
| 11 | Metadata Reporter | V3_S2E_Metadata.mp3 | ~20 sec | [ ] |
| 12 | Python Intro + File Compare | V3_S3A_Python_File_Compare.mp3 | ~40 sec | [ ] |
| 13 | Fuzzy Lookup | V3_S3B_Fuzzy_Lookup.mp3 | ~40 sec | [ ] |
| 14 | PDF Extractor | V3_S3C_PDF_Extractor.mp3 | ~40 sec | [ ] |
| 15 | Closing | V3_S4_Closing.mp3 | ~30 sec | [ ] |
| | **TOTAL NARRATION** | | **~8:15** | |
| | + title cards + chapter cards + pauses | | **~9:00-10:00** | |

---

## Total Runtime Breakdown

| Section | Narration | Cards/Pauses | Total |
|---------|-----------|-------------|-------|
| Opening title card | — | 5 sec | 5 sec |
| Opening (3.0) | 45 sec | 2 sec lead-in | 47 sec |
| Chapter 1 card | — | 3 sec | 3 sec |
| 1A: Delete Blank Rows | 30 sec | — | 30 sec |
| 1B: Remove Spaces | 30 sec | — | 30 sec |
| 1C: Text to Numbers | 30 sec | — | 30 sec |
| 1D: Unmerge & Fill | 30 sec | — | 30 sec |
| 1E: Sanitize Preview | 30 sec | — | 30 sec |
| Chapter 2 card | — | 3 sec | 3 sec |
| 2A: Branding | 40 sec | — | 40 sec |
| 2B: AutoFit | 20 sec | — | 20 sec |
| 2C: External Links | 40 sec | — | 40 sec |
| 2D: Hidden Sheets | 30 sec | — | 30 sec |
| 2E: Metadata Reporter | 20 sec | — | 20 sec |
| Chapter 3 card | — | 3 sec | 3 sec |
| 3A: Python Intro + File Compare | 40 sec | — | 40 sec |
| 3B: Fuzzy Lookup | 40 sec | — | 40 sec |
| 3C: PDF Extractor | 40 sec | — | 40 sec |
| Closing (3.4) | 30 sec | — | 30 sec |
| Closing title card | — | 8 sec | 8 sec |
| **TOTAL** | **~8:15** | **~24 sec** | **~9:09** |

---

*Created: 2026-03-09 | Part of AI Narration Scripts package*
