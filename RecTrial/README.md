# RecTrial/ — Working Folder Snapshot

**Snapshot date:** 2026-04-23
**Branch:** April23CLD
**What this is:** A mirror of Connor's `C:\Users\connor.atlee\RecTrial\` local working folder at this point in time. The folder is NOT normally committed — this branch is a one-time snapshot so the full project state can be reviewed on GitHub in one place.

**Master narrative doc:** See `PROJECT_OVERVIEW.md` in this folder for the full project story, the 4-video plan, current V4 direction, and open decisions.

## Folder map

| Folder / file | What's inside | Notes |
|---|---|---|
| **`AGENTS.md`** | Agent configuration for Connor's local setup | Environment-specific |
| **`PROJECT_OVERVIEW.md`** | ⭐ **Master narrative doc** — the 15-section project overview | Read this first |
| **`build_sample_v2.py`** | Python script that builds `Sample_Quarterly_ReportV2.xlsm` | Used during Video 3 prep |
| **`build_video4_demo_files.py`** | Python script that generates Video 4 demo input files | Source of V4 demo data |
| **`AudioClips/`** | Per-video narration MP3s (Video1, Video2, Video3, Video4) | ⚠️ **MP3s excluded from git**. Folders kept via `.gitkeep`. Live locally at `C:\Users\connor.atlee\RecTrial\AudioClips\` |
| **`Brainstorm/`** | All Video 4 + project planning docs | Key files: `VIDEO_4_CURRENT_PROPOSAL.md`, `VIDEO_4_DRAFT_IDEAS.md`, `FUTURE_AUTOMATION_IDEAS.md`, `NewCodeResearch/` |
| **`Brainstorm/NewCodeResearch/ResearchComplied/`** | 6 compiled research docs (V1–V6) synthesizing ~156 automation ideas | Used to shape the V4 plan |
| **`Brainstorm/NewCodeResearch/ResearchFiles/`** | 14 raw research files from parallel AI sessions | Sources for the compiled synthesis |
| **`CodexCompare/`** | Parallel Codex build comparison + cherry-pick tracker | `COMPARISON_REPORT.md`, `CHERRY_PICK_TRACKER.md` |
| **`DemoFile/`** | `ExcelDemoFile_adv.xlsm` — main P&L demo workbook for Videos 1 & 2 | |
| **`DemoPython/`** | Python scripts used in the demo workbook | |
| **`DemoVBA/`** | 32 demo-file VBA modules + synced copy of `modDirector.bas` | Imported into `ExcelDemoFile_adv.xlsm` |
| **`Feedback/`** | 4 rounds of Gemini AI bug reports for Video 3 + the v3.3 review prompt | Shows the iteration history for V3 |
| **`Guide/`** | Connor's recording / planning guides (Gemini review prompts, interactive guides, etc.) | Not shipped to coworkers |
| **`Guides/`** | Folder of user-facing training guides (Word versions) | Coworker-facing |
| **`Random Claude Built Files-Unneeded/`** | Scratch folder of AI-built files Connor hasn't needed | Kept for archival reasons |
| **`Recordings/`** | Final recorded MP4s per video | ⚠️ **MP4s excluded from git.** Folders kept via `.gitkeep`. Live locally at `C:\Users\connor.atlee\RecTrial\Recordings\` |
| **`SampleFile/`** | `Sample_Quarterly_ReportV2.xlsm` — the V3 universal toolkit demo | Plus backup copies from prior versions |
| **`UniversalToolkit/vba/`** | 23 universal toolkit VBA modules (~140 tools) | The plug-and-play Excel toolkit |
| **`UniversalToolkit/python/`** | 28 Python scripts (includes `ZeroInstall/` subfolder of 7 stdlib-only scripts) | The plug-and-play Python toolkit |
| **`UniversalToolkit/sql/`** | 4 SQL scripts (staging, transformations, validations, pnl_enhancements) | Reference SQL templates |
| **`VBABackup_PrePathA/`** | Backup of ~10 VBA files before the Path A silent-wrapper refactor | Rollback safety net |
| **`VBABackup_PreV2.2Fix/`** | Backup of VBA + sample file before the Video 3 v2.2 fix cycle | Rollback safety net |
| **`VBAToImport/`** | Active VBA files Connor imports into the sample workbook during iteration | `modDirector.bas` is authoritative here |
| **`Video4DemoFiles/`** | 12 input files for the original Video 4 demos | Currently being reassessed as V4 plan pivots |
| **`VideoScripts/`** | Video scripts (master plans + per-video scripts) | |
| **`VideoTitleCards/`** | Originals (V1–V3 + disclaimer, all "OF 3" format) | Superseded by v2 below |
| **`VideoTitleCards_v2/`** | Regenerated title cards — all 4 videos + disclaimer, "OF 4" format, iPipeline-branded | Generator script: `VideoTitleCards/generate_title_cards.py` |

## What's deliberately excluded from git

### Audio (`AudioClips/*.mp3`)
20 MB total. Too large to version in git, not code, regenerable from ElevenLabs if needed.

### Recorded videos (`Recordings/*.mp4`)
~707 MB total. Binary video files belong in cloud storage, not git.

Both live on Connor's laptop. This branch keeps empty folder structure via `.gitkeep` files so the layout is still visible.

## How to use this branch

### If you're giving a second-opinion review

Start with `PROJECT_OVERVIEW.md` → Section 14 "How to review this project" lists the angles worth pushing on.

Then read whichever specific doc matches what you want to push on:

- Video 4 direction: `Brainstorm/VIDEO_4_CURRENT_PROPOSAL.md`
- Code ideas evaluation: `Brainstorm/NewCodeResearch/ResearchComplied/CodeReview_V5.md` (most comprehensive of the 6)
- Cherry-pick campaign: `CodexCompare/CHERRY_PICK_TRACKER.md`
- Post-demo plans: `Brainstorm/FUTURE_AUTOMATION_IDEAS.md`

### If you're Connor

This is just a point-in-time backup of the working folder. Don't work directly in this `RecTrial/` folder inside the repo — keep editing in your local `C:\Users\connor.atlee\RecTrial\` and re-snapshot when needed.

### If you're a fresh Claude Code session

Do not treat this `RecTrial/` folder as the source of truth for live code — the authoritative VBA + Python lives at `C:\Users\connor.atlee\RecTrial\` (Connor's local path). This snapshot is for GitHub review only.

The `FinalExport/` folder at the repo root is the source of truth for what ships to coworkers.

## Main branch vs this branch

- **`main` / `April19update`:** live working branches with the actual code + docs
- **`April23CLD` (this branch):** read-only snapshot of the working folder as of 2026-04-23 for external review

If you want to merge anything FROM this branch INTO main, it should only be individual planning docs (like `PROJECT_OVERVIEW.md`), not the whole `RecTrial/` folder.
