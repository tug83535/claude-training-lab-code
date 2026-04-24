# VIDEO 4 SCRIPT REVIEW — Feedback for Claude Code
# Date: April 2026
# Reviewer: Connor Atlee (via Claude video production assistant)

---

## SUMMARY

The Video 4 narration script is solid overall — clean pacing, correct tone, good structure.
The issues below are mostly pronunciation fixes and one number inconsistency.
No major rewrites needed.

---

## FIXES ALREADY APPLIED (in generate_clips_video4.py)

These were caught in review and already fixed in the Colab generation script.
Claude Code should apply the same fixes to the master script file if it maintains one.

### 1. Number inconsistency — Opening clip
- ORIGINAL: "sixty-two actions"
- FIXED: "sixty-five actions"
- REASON: Videos 1-3 were updated to 65 actions. Video 4 must match.
- LOCATION: V4_S0_Opening.mp3

### 2. FP&A pronunciation
- ORIGINAL: "FP&A"
- FIXED: "F-P-and-A"
- REASON: ElevenLabs reads "FP&A" as "fpa" — not intelligible
- LOCATION: V4_S6_VarianceDecomp.mp3

### 3. VLOOKUP pronunciation
- ORIGINAL: "VLOOKUP"
- FIXED: "V-LOOKUP"
- REASON: ElevenLabs may mispronounce — hyphen format forces correct letter-by-letter read
- LOCATION: V4_S3_FuzzyLookup.mp3

### 4. PDF pronunciation
- ORIGINAL: "PDF" / "PDFs"
- FIXED: "P-D-F" / "P-D-Fs"
- REASON: ElevenLabs reads "PDF" as a word ("pdf") — hyphen format forces correct letters
- LOCATION: V4_S2_PDFExtractor.mp3, V4_S9_Closing.mp3

### 5. Closing recap pacing
- ORIGINAL: Single long sentence listing all 8 scripts
- FIXED: Split into two sentences — first 4 scripts, then second 4
- REASON: ElevenLabs rushes through long lists — splitting creates natural pause
- LOCATION: V4_S9_Closing.mp3

### 6. ElevenLabs settings mismatch
- ORIGINAL SCRIPT SPECIFIED: Eleven Multilingual v2, Stability 50%, Similarity 80%, Style 20%
- CORRECT SETTINGS (matching Videos 1-3):
  - Model: eleven_v3
  - Stability: 0.35
  - Similarity: 0.75
  - Style: 0.30
  - Speaker Boost: ON
- REASON: Using different settings would make Video 4 sound noticeably different from Videos 1-3
- ACTION: If Claude Code maintains an ElevenLabs config file, update it for Video 4

### 7. Minor closer improvement — V4_S3
- ORIGINAL: "One command replaces hours of manual matching."
- FIXED: "One command... and every match is found in seconds."
- REASON: Slightly abrupt as a closer — new version flows better with trailing ellipsis

---

## NO ACTION NEEDED — These are fine as written

- "JP Morgan Chase" / "JPMorgan Chase and Co" — ElevenLabs handles naturally
- "MetLife" / "Metropolitan Life Insurance" — fine
- "Excel" — fine
- "iGO", "Affirm", "InsureSight", "DocFast" — correctly NOT mentioned in Video 4 (these scripts are not iPipeline-specific)
- All clip durations are within target (30-45 sec)
- Tone is consistent with Videos 1-3
- No technical jargon that non-technical finance staff wouldn't understand
- No file paths, Python versions, or setup instructions mentioned

---

## OPEN QUESTION FOR CLAUDE CODE

The script mentions 8 Python scripts being demoed. Please confirm:

1. Are all 8 of these scripts fully built and available on SharePoint?
   - File Comparison
   - PDF Extractor
   - Fuzzy Lookup
   - Bank Reconciler
   - Aging Report
   - Variance Decomposition
   - Forecast Rollforward
   - Variance Analysis (multi-file)

2. Does the Bank Reconciler (V4_S4) exist as a standalone script or is it part of the existing fuzzy lookup module?

3. Does the Forecast Rollforward (V4_S7) exist or is it planned? The script references it as complete.

4. The closing says "All scripts, documentation, and sample files are available on SharePoint in the Finance Automation folder." — please confirm this is accurate before video goes live.

---

## DELIVERABLES FROM THIS REVIEW

- generate_clips_video4.py — ready to run in Google Colab with all fixes applied
- This feedback document — for Claude Code to update master script files

---

*Review completed: April 2026*
*Video 4 is part of the iPipeline Finance Automation Video Demo Project (4 videos total)*
