# VIDEO 4 — "Python Automation for Finance"
# Full Narration Script

**Runtime Target:** 6-8 minutes
**Audience:** All iPipeline employees (2,000+)
**Tone:** Professional, conversational, confident. Same voice/style as Videos 1-3.
**Purpose:** Show that Python scripts do in seconds what takes hours — no coding needed.

---

## AUDIO CLIP NAMING CONVENTION

All clips saved to: `RecTrial/AudioClips/Video4/`

| Clip | Filename | Duration |
|------|----------|----------|
| 1 | V4_S0_Opening.mp3 | ~30 sec |
| 2 | V4_S1_CompareFiles.mp3 | ~45 sec |
| 3 | V4_S2_PDFExtractor.mp3 | ~40 sec |
| 4 | V4_S3_FuzzyLookup.mp3 | ~45 sec |
| 5 | V4_S4_BankReconciler.mp3 | ~45 sec |
| 6 | V4_S5_AgingReport.mp3 | ~40 sec |
| 7 | V4_S6_VarianceDecomp.mp3 | ~45 sec |
| 8 | V4_S7_ForecastRoll.mp3 | ~40 sec |
| 9 | V4_S8_VarianceAnalysis.mp3 | ~40 sec |
| 10 | V4_S9_Closing.mp3 | ~30 sec |

---

## ELEVENLABS SETTINGS (same as Videos 1-3)

- Model: Eleven Multilingual v2
- Stability: 50%
- Similarity: 80%
- Style: 20%
- Speaker Boost: ON
- Use the same voice as Videos 1-3 for consistency

---

## CLIP 1 — Opening

**Filename:** V4_S0_Opening.mp3
**Duration:** ~30 seconds

### [PASTE INTO ELEVENLABS]

In the last three videos, you saw what the Excel automation toolkit can do — sixty-two actions, all from one file.

But there is a second piece to this project — a library of Python scripts built specifically for Finance and Accounting.

These scripts do not require any coding knowledge. You point them at a file, run one command, and get a polished output in seconds.

In the next few minutes, I am going to walk you through eight of them.

---

## CLIP 2 — Compare Files

**Filename:** V4_S1_CompareFiles.mp3
**Duration:** ~45 seconds

### [PASTE INTO ELEVENLABS]

First — file comparison.

You have two versions of the same report. Maybe last month versus this month. Maybe your version versus someone else's. You need to know exactly what changed.

This script compares every cell across both files — and builds a color-coded diff report.

Green means a row was added. Red means it was removed. Yellow highlights every cell that changed, with the old value and the new value side by side.

One command. Every difference found. No more scrolling through two files trying to spot what moved.

---

## CLIP 3 — PDF Extractor

**Filename:** V4_S2_PDFExtractor.mp3
**Duration:** ~40 seconds

### [PASTE INTO ELEVENLABS]

Next — extracting data from PDFs.

If you have ever received a financial statement, an invoice summary, or a vendor report as a PDF, you know the problem. The data is right there on the page, but you cannot use it in Excel.

This script reads the PDF, finds every table in it, and pulls the data straight into an Excel workbook — one sheet per table. Columns, rows, numbers — all extracted automatically.

No retyping. No copy-paste errors. Just point it at the PDF and let it work.

---

## CLIP 4 — Fuzzy Lookup

**Filename:** V4_S3_FuzzyLookup.mp3
**Duration:** ~45 seconds

### [PASTE INTO ELEVENLABS]

This one is probably the most impressive.

You have two lists of vendor names — one from your system, one from a partner or a bank statement. They should match, but they do not. One says "Metropolitan Life Insurance" and the other says "MetLife." One says "JP Morgan Chase" and the other says "JPMorgan Chase and Co."

A normal VLOOKUP fails on these because the names are not identical. This script uses fuzzy matching to find the closest match — even when the spelling is different.

It shows you every match with a confidence score. Exact matches in green. Fuzzy matches in yellow with the score. No match in red.

One command replaces hours of manual matching.

---

## CLIP 5 — Bank Reconciler

**Filename:** V4_S4_BankReconciler.mp3
**Duration:** ~45 seconds

### [PASTE INTO ELEVENLABS]

Bank reconciliation — the task nobody enjoys.

You have your general ledger on one side and your bank statement on the other. The descriptions never match exactly. Your ledger says "Office Supplies - Staples" and the bank says "STAPLES STORE 4521."

This script matches them using a combination of amount, date, and fuzzy description matching. It assigns a confidence score to every match and flags anything it cannot reconcile.

The output is a clean report — matched items in green, fuzzy matches in yellow with the confidence score, and unmatched items in red for you to investigate.

What used to take a full afternoon now takes about ten seconds.

---

## CLIP 6 — Aging Report

**Filename:** V4_S5_AgingReport.mp3
**Duration:** ~40 seconds

### [PASTE INTO ELEVENLABS]

Aging reports — accounts receivable, accounts payable, or any date-based tracking.

Give this script a file with dates and amounts, and it automatically buckets everything into Current, zero to thirty days, thirty-one to sixty, sixty-one to ninety, and ninety-plus.

The output is a color-coded Excel workbook with a detail sheet, a summary by bucket, and a pivot by vendor or customer. Green for current, yellow for aging, red for anything past ninety days.

One command, and you have a board-ready aging report.

---

## CLIP 7 — Variance Decomposition

**Filename:** V4_S6_VarianceDecomp.mp3
**Duration:** ~45 seconds

### [PASTE INTO ELEVENLABS]

This is the one for the FP&A team.

When revenue is up or down versus budget, leadership does not just want to know the total variance. They want to know why. Was it a price change? A volume change? A product mix shift?

This script takes your actual and budget data — units and prices by product — and decomposes the variance into three components: price effect, volume effect, and mix effect.

The output is color-coded. Favorable variances in green, unfavorable in red. Each component broken out separately so you can tell the full story in your next board presentation.

---

## CLIP 8 — Forecast Rollforward

**Filename:** V4_S7_ForecastRoll.mp3
**Duration:** ~40 seconds

### [PASTE INTO ELEVENLABS]

Rolling forecasts — updated every month, always looking twelve months ahead.

Give this script your historical actuals and it builds a twelve-month rolling forecast automatically. You can choose the method — moving average, growth rate, or flat projection.

The output includes a combined actual-plus-forecast view and a line chart showing where you have been and where the model projects you are going. Actuals in blue, forecast in green.

No more manually extending formulas every month. One command and the forecast rolls forward.

---

## CLIP 9 — Variance Analysis

**Filename:** V4_S8_VarianceAnalysis.mp3
**Duration:** ~40 seconds

### [PASTE INTO ELEVENLABS]

Last one — variance analysis across multiple files.

Point this script at a folder of budget files — one per department, one per entity, however you organize them — and it consolidates everything into a single actual versus budget report.

Dollar variance. Percent variance. Favorable or unfavorable flag on every line. Plus a bar chart showing the top variances by department.

It handles ten files or fifty files — does not matter. One command, one consolidated report, ready for leadership review.

---

## CLIP 10 — Closing

**Filename:** V4_S9_Closing.mp3
**Duration:** ~30 seconds

### [PASTE INTO ELEVENLABS]

That is eight Python scripts — file comparison, PDF extraction, fuzzy matching, bank reconciliation, aging reports, variance decomposition, forecasting, and variance analysis.

Every one of them runs from a single command. No coding required. Just point it at your files and go.

All scripts, documentation, and sample files are available on SharePoint in the Finance Automation folder. If you want to try any of these on your own data, the guides walk you through every step.

Thanks for watching.

---

## SCRIPT REVIEW CHECKLIST

Before generating audio, verify:

- [ ] Every clip reads naturally when spoken aloud (no awkward phrasing)
- [ ] No technical jargon that a non-technical Finance person would not understand
- [ ] Consistent tone across all 10 clips (professional, confident, conversational)
- [ ] Each clip starts with a clear topic sentence so the viewer knows what is coming
- [ ] No clip exceeds ~45 seconds of narration (keeps pacing tight)
- [ ] The opening sets expectations (8 scripts, no coding needed)
- [ ] The closing recaps all 8 and points to SharePoint
- [ ] No mention of specific file paths, Python versions, or technical setup
- [ ] Product names (iGO, Affirm, InsureSight, DocFast) NOT mentioned — these scripts work on any data, not iPipeline-specific

---

## NOTES FOR CONNOR

1. **Generate all 10 clips in one ElevenLabs session** so the voice is consistent
2. **Listen to each clip with headphones** before saving — check for robotic pauses or mispronunciations
3. **Watch for:** "FP&A" should be pronounced "F-P-and-A" not "fpa." "VLOOKUP" should be "V-lookup." "PDF" should be "P-D-F."
4. **Save each clip** with the exact filename from the table above
5. **Create the folder** `RecTrial/AudioClips/Video4/` and put all 10 clips there

---

*Script created: 2026-04-01*
*Part of the iPipeline Finance Automation Video Demo Project*
*Video 4 of 4: Python Automation for Finance*
