"""
Excel Automation Videos — ElevenLabs Audio Generator
Video 4: Python Automation for Finance
10 clips

FIXES APPLIED FROM SCRIPT REVIEW:
- FP&A → F-P-and-A in clip text
- VLOOKUP → V-LOOKUP in clip text
- PDF/PDFs → P-D-F / P-D-Fs in clip text
- "sixty-two actions" → "sixty-five actions" in opening
- Settings matched to Videos 1-3 (eleven_v3, stability 0.35, similarity 0.75, style 0.30)
- V4_S9 closing recap split into two sentences for better pacing

SETUP:
1. pip install elevenlabs
2. Add your API key on line 24 below
3. Run script
4. All MP3s saved to ./ElevenLabs_Clips/Video4/

NOTE: Script skips already-generated files.
      Delete ElevenLabs_Clips/Video4/ folder to regenerate from scratch.
"""

from elevenlabs.client import ElevenLabs
from elevenlabs import VoiceSettings
import os
import time

# ─────────────────────────────────────────────────────────────
# CONFIG — same settings as Videos 1-3 for voice consistency
# ─────────────────────────────────────────────────────────────

API_KEY  = "YOUR_ELEVENLABS_API_KEY"   # elevenlabs.io → Profile → API Keys
VOICE_ID = "GzE4TcXfh9rYCU9gVgPp"     # Same built-in voice as Videos 1-3

SETTINGS = VoiceSettings(
    stability=0.35,        # Matches Videos 1-3
    similarity_boost=0.75, # Matches Videos 1-3
    style=0.30,            # Matches Videos 1-3
    use_speaker_boost=True
)

MODEL_ID = "eleven_v3"     # Matches Videos 1-3 — DO NOT change to multilingual v2

# ─────────────────────────────────────────────────────────────
# ALL 10 CLIPS — VIDEO 4
# ─────────────────────────────────────────────────────────────

CLIPS = [

    {
        "folder": "Video4",
        "filename": "V4_S0_Opening.mp3",
        "text": """In the last three videos, you saw what the Excel automation toolkit can do — sixty-five actions, all from one file.

But there is a second piece to this project — a library of Python scripts built specifically for Finance and Accounting.

These scripts do not require any coding knowledge. You point them at a file, run one command, and get a polished output in seconds.

In the next few minutes, I am going to walk you through eight of them."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S1_CompareFiles.mp3",
        "text": """First — file comparison.

You have two versions of the same report. Maybe last month versus this month. Maybe your version versus someone else's. You need to know exactly what changed.

This script compares every cell across both files — and builds a color-coded diff report.

Green means a row was added. Red means it was removed. Yellow highlights every cell that changed, with the old value and the new value side by side.

One command. Every difference found. No more scrolling through two files trying to spot what moved."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S2_PDFExtractor.mp3",
        "text": """Next — extracting data from P-D-Fs.

If you have ever received a financial statement, an invoice summary, or a vendor report as a P-D-F, you know the problem. The data is right there on the page, but you cannot use it in Excel.

This script reads the P-D-F, finds every table in it, and pulls the data straight into an Excel workbook — one sheet per table. Columns, rows, numbers — all extracted automatically.

No retyping. No copy-paste errors. Just point it at the P-D-F and let it work."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S3_FuzzyLookup.mp3",
        "text": """This one is probably the most impressive.

You have two lists of vendor names — one from your system, one from a partner or a bank statement. They should match, but they do not. One says "Metropolitan Life Insurance" and the other says "MetLife." One says "JP Morgan Chase" and the other says "JPMorgan Chase and Co."

A normal V-LOOKUP fails on these because the names are not identical. This script uses fuzzy matching to find the closest match — even when the spelling is different.

It shows you every match with a confidence score. Exact matches in green. Fuzzy matches in yellow with the score. No match in red.

One command... and every match is found in seconds."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S4_BankReconciler.mp3",
        "text": """Bank reconciliation — the task nobody enjoys.

You have your general ledger on one side and your bank statement on the other. The descriptions never match exactly. Your ledger says "Office Supplies - Staples" and the bank says "STAPLES STORE 4521."

This script matches them using a combination of amount, date, and fuzzy description matching. It assigns a confidence score to every match and flags anything it cannot reconcile.

The output is a clean report — matched items in green, fuzzy matches in yellow with the confidence score, and unmatched items in red for you to investigate.

What used to take a full afternoon now takes about ten seconds."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S5_AgingReport.mp3",
        "text": """Aging reports — accounts receivable, accounts payable, or any date-based tracking.

Give this script a file with dates and amounts, and it automatically buckets everything into Current, zero to thirty days, thirty-one to sixty, sixty-one to ninety, and ninety-plus.

The output is a color-coded Excel workbook with a detail sheet, a summary by bucket, and a pivot by vendor or customer. Green for current, yellow for aging, red for anything past ninety days.

One command, and you have a board-ready aging report."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S6_VarianceDecomp.mp3",
        "text": """This is the one for the F-P-and-A team.

When revenue is up or down versus budget, leadership does not just want to know the total variance. They want to know why. Was it a price change? A volume change? A product mix shift?

This script takes your actual and budget data — units and prices by product — and decomposes the variance into three components: price effect, volume effect, and mix effect.

The output is color-coded. Favorable variances in green, unfavorable in red. Each component broken out separately so you can tell the full story in your next board presentation."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S7_ForecastRoll.mp3",
        "text": """Rolling forecasts — updated every month, always looking twelve months ahead.

Give this script your historical actuals and it builds a twelve-month rolling forecast automatically. You can choose the method — moving average, growth rate, or flat projection.

The output includes a combined actual-plus-forecast view and a line chart showing where you have been and where the model projects you are going. Actuals in blue, forecast in green.

No more manually extending formulas every month. One command and the forecast rolls forward."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S8_VarianceAnalysis.mp3",
        "text": """Last one — variance analysis across multiple files.

Point this script at a folder of budget files — one per department, one per entity, however you organize them — and it consolidates everything into a single actual versus budget report.

Dollar variance. Percent variance. Favorable or unfavorable flag on every line. Plus a bar chart showing the top variances by department.

It handles ten files or fifty files — does not matter. One command, one consolidated report, ready for leadership review."""
    },
    {
        "folder": "Video4",
        "filename": "V4_S9_Closing.mp3",
        "text": """That is eight Python scripts — file comparison, P-D-F extraction, fuzzy matching, and bank reconciliation.

Plus aging reports, variance decomposition, forecasting, and variance analysis.

Every one of them runs from a single command. No coding required. Just point it at your files and go.

All scripts, documentation, and sample files are available on SharePoint in the Finance Automation folder. If you want to try any of these on your own data, the guides walk you through every step.

Thanks for watching."""
    },
]


# ─────────────────────────────────────────────────────────────
# GENERATE
# ─────────────────────────────────────────────────────────────

def main():
    client = ElevenLabs(api_key=API_KEY)

    base_dir = "ElevenLabs_Clips"
    os.makedirs(os.path.join(base_dir, "Video4"), exist_ok=True)

    total = len(CLIPS)
    failed = []

    print(f"\n{'='*55}")
    print(f"  Excel Automation — Video 4 Clip Generator")
    print(f"  {total} clips | Voice ID: {VOICE_ID[:8]}...")
    print(f"{'='*55}\n")

    for i, clip in enumerate(CLIPS, 1):
        path = os.path.join(base_dir, clip["folder"], clip["filename"])

        if os.path.exists(path):
            print(f"[{i:02d}/{total}] SKIPPED (exists)  {clip['filename']}")
            continue

        try:
            print(f"[{i:02d}/{total}] Generating...    {clip['filename']}", end="", flush=True)

            audio = client.text_to_speech.convert(
                voice_id=VOICE_ID,
                text=clip["text"],
                model_id=MODEL_ID,
                voice_settings=SETTINGS,
            )

            with open(path, "wb") as f:
                for chunk in audio:
                    f.write(chunk)

            size_kb = os.path.getsize(path) // 1024
            print(f"  ✓  ({size_kb} KB)")

            if i < total:
                time.sleep(1.5)

        except Exception as e:
            print(f"  ✗  FAILED: {e}")
            failed.append(clip["filename"])

    print(f"\n{'='*55}")
    done = total - len(failed)
    print(f"  Done: {done}/{total} clips generated")
    print(f"  Saved to: ./{base_dir}/Video4/")
    if failed:
        print(f"\n  Failed clips (re-run to retry):")
        for f in failed:
            print(f"    - {f}")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
