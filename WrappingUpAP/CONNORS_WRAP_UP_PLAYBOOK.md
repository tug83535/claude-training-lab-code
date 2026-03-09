# Connor's Wrap-Up Playbook — Start to Finish

**Purpose:** This is YOUR personal step-by-step guide. It walks you through everything from downloading your files off GitHub all the way through the final SharePoint upload. Every step is written out — nothing skipped.

**Last Updated:** 2026-03-09

---

## How This Guide Works

- Steps are in order. Do them top to bottom.
- Each step has a checkbox. Check it off when done.
- If something goes wrong, stop and ask Claude for help — don't push through.
- Estimated total time: 3-5 hours across multiple sittings (not counting video recording practice).

---

## PHASE 1: Get Your Files Ready on Your Computer

**Goal:** Download everything from GitHub so you have all the latest files on your local machine.

### Step 1 — Download the Branch from GitHub

- [ ] 1. Open your web browser and go to your GitHub repo: `https://github.com/tug83535/claude-training-lab-code`
- [ ] 2. Near the top left, you'll see a dropdown that says the branch name. Click it
- [ ] 3. Switch to the branch: `claude/resume-ipipeline-demo-qKRHn`
- [ ] 4. Confirm the branch name shows at the top
- [ ] 5. Click the green **Code** button
- [ ] 6. Click **Download ZIP**
- [ ] 7. Save the ZIP file to your Desktop (or Downloads folder)
- [ ] 8. Right-click the ZIP file → **Extract All** → pick a folder you'll remember (like `Desktop\APCLDmerge`)
- [ ] 9. Open the extracted folder and confirm you see all the subfolders: `vba/`, `python/`, `FinalRoughGuides/`, `UniversalToolsForAllFiles/`, etc.

**You now have every file from the branch on your computer.**

---

## PHASE 2: Fix the Minor Guide Updates

**Goal:** Fix the small text issues found during the guide review before converting to PDF.

These are all quick find-and-replace fixes. Open each file in a text editor (Notepad, VS Code, or even Word).

### Step 2 — Fix Guide 01 (Command Center Guide)

- [ ] 1. Open `FinalRoughGuides/01-How-to-Use-the-Command-Center.md`
- [ ] 2. Use **Ctrl+H** (Find and Replace)
- [ ] 3. Find: `7-sheet PDF` → Replace with: `multi-sheet PDF`
- [ ] 4. Click **Replace All** — should find 4 instances
- [ ] 5. Save the file

### Step 3 — Fix Guide 03 (Leadership Overview)

- [ ] 1. Open `FinalRoughGuides/03-What-This-File-Does-Leadership-Overview.md`
- [ ] 2. Find: `~99` → Replace with: `~100` → Replace All (should find ~4 instances)
- [ ] 3. Find: `7-sheet report` → Replace with: `multi-sheet report` → Replace All (should find ~2 instances)
- [ ] 4. Find the line that says `21` bugs found — update the number to `30+`
- [ ] 5. Save the file

### Step 4 — Fix Guide 05 (Video Script)

- [ ] 1. Open `FinalRoughGuides/05-Video-Demo-Script-and-Storyboard.md`
- [ ] 2. Find: `~99 tools` → Replace with: `~100 tools` → Replace All (should find 1 instance)
- [ ] 3. Save the file

### Step 5 — Fix Guide 06 (Universal Toolkit Guide)

- [ ] 1. Open `FinalRoughGuides/06-Universal-Toolkit-Guide.md`
- [ ] 2. Find the **"Which Modules to Import"** table (around line 117)
- [ ] 3. Add 4 new rows to the table for these modules:
   - `modUTL_DataCleaningPlus` — Extra cleaning tools (whitespace, non-printable chars, text case)
   - `modUTL_AuditPlus` — Extra audit tools (data boundary, header validation, formula errors)
   - `modUTL_DuplicateDetection` — Exact duplicate finder
   - `modUTL_NumberFormat` — Enhanced text-to-number converter, workbook metadata reporter
- [ ] 4. Save the file

### Step 6 — Fix VIDEO_DEMO_PLAN.md

- [ ] 1. Open `videodraft/VIDEO_DEMO_PLAN.md`
- [ ] 2. Find: `99 tools` → Replace with: `~100 tools` (should find 1 instance, around line 221)
- [ ] 3. Save the file

---

## PHASE 3: Review and Approve the Training Guides

**Goal:** Read each guide yourself. Make sure they're accurate and ready for coworkers.

### Step 7 — Read Each Guide

Go through these one at a time. For each one, ask yourself: "Would a non-technical coworker be able to follow this without calling me for help?"

- [ ] 1. Read `01-How-to-Use-the-Command-Center.md` — all 62 actions documented, workflows, troubleshooting
- [ ] 2. Read `02-Getting-Started-First-Time-Setup.md` — download, open, enable macros, first 5 actions
- [ ] 3. Read `03-What-This-File-Does-Leadership-Overview.md` — CFO/CEO briefing, cost savings, rollout
- [ ] 4. Read `04-Quick-Reference-Card.md` — 1-page cheat sheet, shortcuts, quick reference
- [ ] 5. Read `06-Universal-Toolkit-Guide.md` — all ~100 universal tools, setup, playbooks

**If anything needs changes:** Note what needs fixing. You can either fix it yourself or ask Claude to fix it in the next session.

**If a guide looks good:** Move on. You'll convert to PDF in Phase 5.

---

## PHASE 4: Lock Down the Excel Demo File

**Goal:** Make sure the Excel file works perfectly on a clean machine, with no leftover debug data.

### Step 8 — Clean Up the Excel File

- [ ] 1. Open your demo Excel file (the `.xlsm` with all 34 VBA modules already imported)
- [ ] 2. Press **Alt+F11** to open the VBA Editor
- [ ] 3. Click **Debug** → **Compile VBAProject** — make sure it compiles with zero errors
- [ ] 4. Close the VBA Editor (Alt+Q)

### Step 9 — Check for Leftover Test Data

- [ ] 1. Look at every sheet tab at the bottom of the workbook
- [ ] 2. Delete any test sheets that shouldn't be there (sheets created during testing that aren't part of the demo)
- [ ] 3. Make sure these sheets exist and look correct:
   - Report--> (the landing page)
   - P&L Summary (or your main P&L sheet)
   - All monthly tabs (Jan, Feb, Mar at minimum)
   - Assumptions
   - GL detail
   - Checks (Reconciliation)
   - Home (if created by modSheetIndex)
   - Charts & Visuals
- [ ] 4. Right-click any sheet tab → **Unhide** — check if there are hidden sheets that shouldn't be there (VBA_AuditLog is fine to keep hidden)
- [ ] 5. Make sure no personal file paths are visible anywhere in the sheets

### Step 10 — Check the VBA Code for Debug Leftovers

- [ ] 1. Press **Alt+F11** to open VBA Editor again
- [ ] 2. Press **Ctrl+F** (Find)
- [ ] 3. Search for `Debug.Print` — if you find any, they're harmless but you can delete them for cleanliness
- [ ] 4. Search for `MsgBox "test` or `MsgBox "debug` — delete any test message boxes
- [ ] 5. Search for `Stop` — if any `Stop` statements are in the code, delete them (they pause execution like a breakpoint)
- [ ] 6. Close VBA Editor

### Step 11 — Test on a Fresh Excel Session

- [ ] 1. Save the file (**Ctrl+S**)
- [ ] 2. Close Excel completely (not just the file — close the whole application)
- [ ] 3. Reopen the Excel file from scratch
- [ ] 4. If you get a security warning about macros, click **Enable Content**
- [ ] 5. Open the Command Center (click the button or run LaunchCommandCenter)
- [ ] 6. Test 3-4 key actions to confirm everything works:
   - **Action 1** — Run Reconciliation (should show PASS/FAIL on Checks sheet)
   - **Action 7** — Run Data Quality Scan (should produce a report)
   - **Action 10** — PDF Export (should generate a PDF)
   - **Action 44** — Run Integration Tests (should show 18/18 PASS)
- [ ] 7. If everything works, save the file one final time

### Step 12 — Save the Final Copy

- [ ] 1. **File → Save As**
- [ ] 2. Save to your `CompletePackageStorage/production/` folder
- [ ] 3. Name it something clean: `iPipeline_PnL_Demo_FINAL.xlsm`
- [ ] 4. Make a second copy to `CompletePackageStorage/backups/` with a date: `iPipeline_PnL_Demo_BACKUP_2026-03-XX.xlsm`

---

## PHASE 5: Convert Guides to PDF

**Goal:** Coworkers get PDFs, not markdown files. Convert the approved guides.

### Step 13 — Convert Each Guide to PDF

There are a few ways to do this. Pick whichever is easiest for you:

**Option A — Using VS Code (if you have it):**
1. Open the `.md` file in VS Code
2. Install the "Markdown PDF" extension (if not already installed): Click Extensions icon (left sidebar) → search "Markdown PDF" → Install
3. Open the `.md` file → press **Ctrl+Shift+P** → type "Markdown PDF: Export (pdf)" → press Enter
4. PDF saves to the same folder

**Option B — Using a Website (no install needed):**
1. Go to `https://www.markdowntopdf.com` (or similar free converter)
2. Paste the markdown content
3. Click Convert → Download PDF

**Option C — Using Word:**
1. Open the `.md` file in Notepad, select all, copy
2. Paste into a new Word document
3. Clean up the formatting (headings, bold, tables)
4. File → Save As → PDF

**Convert these files:**
- [ ] 1. `01-How-to-Use-the-Command-Center.md` → `01 - How to Use the Command Center.pdf`
- [ ] 2. `02-Getting-Started-First-Time-Setup.md` → `02 - Getting Started First Time Setup.pdf`
- [ ] 3. `03-What-This-File-Does-Leadership-Overview.md` → `03 - What This File Does (Leadership Overview).pdf`
- [ ] 4. `04-Quick-Reference-Card.md` → `04 - Quick Reference Card.pdf`
- [ ] 5. `06-Universal-Toolkit-Guide.md` → `06 - Universal Toolkit Guide.pdf`
- [ ] 6. `CoPilotPromptGuide/AP_Copilot_PromptGuideHelpV2.md` → `Copilot Prompts Guide.pdf` (or use the .docx version and save as PDF from Word)

### Step 14 — Save PDFs to Production Folder

- [ ] 1. Copy all finished PDFs to `CompletePackageStorage/production/`
- [ ] 2. Also keep copies in the `training/` folder in your repo

---

## PHASE 6: Prepare for the Recording

**Goal:** Get everything set up so when you hit record, it goes smoothly.

### Step 15 — Answer the Open Questions

Before recording, decide these things (from `videodraft/VIDEO_DEMO_PLAN.md`):

- [ ] 1. **Recording software:** What will you use? Options:
   - **OBS Studio** — Free, powerful, most popular. Download from obsproject.com
   - **PowerPoint** — Built-in screen recording (Insert → Screen Recording). Simple but limited
   - **Camtasia** — Paid, easiest editing. Check if iPipeline has a license
   - **Xbox Game Bar** — Built into Windows 10/11. Press Win+G. Very basic but works
- [ ] 2. **Webcam or no webcam?** Recommendation: No webcam — screen only. Simpler to record and edit
- [ ] 3. **Background music?** Recommendation: Soft music on intro/outro title cards only. Silent during live demo. Free music at YouTube Audio Library
- [ ] 4. **How many videos?** The plan calls for 3 videos:
   - Video 1: "What's Possible" — 4-5 min overview for everyone
   - Video 2: "Full Demo Walkthrough" — 15-18 min for Finance team
   - Video 3: "Universal Tools" — 8-10 min for anyone who wants tools for their own files
   - You can start with just Video 1 if you want to keep it simple
- [ ] 5. **Where will videos be hosted?** SharePoint is the plan. Confirm you have upload access to create a new folder

### Step 16 — Set Up Your Computer for Recording

- [ ] 1. Close ALL applications except Excel and your recording software
- [ ] 2. Turn off notifications:
   - **Teams:** Click your profile picture → Set status to "Do Not Disturb"
   - **Outlook:** File → Options → Mail → uncheck "Display a Desktop Alert"
   - **Windows:** Settings → System → Notifications → turn off
   - **Phone:** Put on silent or airplane mode
- [ ] 3. Clean your desktop — hide or move all desktop icons (right-click desktop → View → uncheck "Show desktop icons")
- [ ] 4. Set your taskbar to auto-hide (right-click taskbar → Taskbar settings → toggle "Automatically hide the taskbar")
- [ ] 5. Set Excel zoom to **100%** (or 110% if your screen is large)
- [ ] 6. Set display scaling to **100%** (Settings → Display → Scale → 100%)
- [ ] 7. Set screen resolution to **1920 x 1080** (Settings → Display → Resolution)

### Step 17 — Prepare the Demo File for Recording

- [ ] 1. Open the final `.xlsm` file from `CompletePackageStorage/production/`
- [ ] 2. Make sure you're on the **Report-->** landing page (or whatever your starting sheet is)
- [ ] 3. Close the Command Center if it's open (you'll open it on camera)
- [ ] 4. Delete any leftover output sheets from previous runs (Variance Analysis results, dashboard outputs, etc.) — start clean
- [ ] 5. Make sure the file is saved in this clean starting state

### Step 18 — Audio Test

- [ ] 1. Open your recording software
- [ ] 2. Record a 30-second test clip — just talk normally about anything
- [ ] 3. Play it back with headphones
- [ ] 4. Listen for: background noise, volume too low, echo, keyboard/mouse clicks
- [ ] 5. If audio quality is bad:
   - Use a USB headset or plug-in microphone instead of laptop mic
   - Close windows/doors to reduce background noise
   - Put a towel or soft surface under your mouse to reduce click sounds
- [ ] 6. Record another 30-second test and confirm it sounds good

---

## PHASE 7: Record the Demo Video(s)

**Goal:** Record the actual video following the scripts.

### Step 19 — Do a Dry Run First (No Recording)

- [ ] 1. Pull up the video script on your second monitor (or print it)
   - For Video 1: `videodraft/COMPILED_VIDEO_PACKAGE.md` → Section 2
   - For Video 2: `videodraft/COMPILED_VIDEO_PACKAGE.md` → Section 3
   - Or use `FinalRoughGuides/05-Video-Demo-Script-and-Storyboard.md`
- [ ] 2. Walk through the entire demo WITHOUT recording — just practice
- [ ] 3. Time yourself — are you close to the target time?
- [ ] 4. Note any spots where you stumble or the flow feels awkward
- [ ] 5. Practice those spots 2-3 more times
- [ ] 6. Do one more full dry run — should feel smooth

### Step 20 — Record

- [ ] 1. Start your recording software
- [ ] 2. Confirm it's capturing the right screen at 1920×1080
- [ ] 3. Start recording
- [ ] 4. Follow the script — but don't read it word-for-word robotically. Know the bullet points, talk naturally
- [ ] 5. **Key tips during recording:**
   - After each action runs, **pause for 2-3 seconds** — let the result sink in
   - Call out time savings: "That used to take 2 hours. We just did it in 10 seconds."
   - If you make a mistake, just pause, take a breath, and redo that section. You can edit it out later
   - Don't rush. Slower is better for video
- [ ] 6. When done, stop recording
- [ ] 7. Watch the recording back — check for:
   - Audio clear and loud enough?
   - Screen visible and readable?
   - Any notifications pop up that need to be cropped?
   - Any long dead air that should be trimmed?
- [ ] 8. If it's good, save it. If not, redo the sections that need fixing

### Step 21 — Save the Video

- [ ] 1. Export/save as MP4 (H.264 format)
- [ ] 2. Name it clearly:
   - `1 - Whats Possible (Overview).mp4`
   - `2 - Full Demo Walkthrough.mp4`
   - `3 - Universal Tools.mp4`
- [ ] 3. Save to `CompletePackageStorage/production/`
- [ ] 4. Also save a backup copy to `CompletePackageStorage/backups/`

---

## PHASE 8: Build the START HERE Document

**Goal:** Create a one-page PDF that sits at the top of the SharePoint folder and tells anyone landing there exactly what they're looking at.

### Step 22 — Create the START HERE PDF

- [ ] 1. Open Word (or ask Claude to generate this for you)
- [ ] 2. Create a one-page document that includes:
   - **Title:** "iPipeline Finance Automation — Start Here"
   - **What is this?** One paragraph explaining the project
   - **What's in this folder?** Bullet list of each subfolder and what's inside
   - **Which video should I watch?**
     - Leadership/Everyone → Video 1 (5 min)
     - Finance team → Video 2 (18 min)
     - Want tools for your own files → Video 3 (10 min)
   - **How do I get started?** Download the `.xlsm`, enable macros, open Command Center
   - **Questions?** Contact Connor
- [ ] 3. Apply iPipeline branding (blue headers, Arial font — reference `docs/ipipeline-brand-styling.md`)
- [ ] 4. Save as `START HERE.pdf`
- [ ] 5. Copy to `CompletePackageStorage/production/`

---

## PHASE 9: Upload to SharePoint

**Goal:** Get everything live and accessible for coworkers.

### Step 23 — Create the SharePoint Folder Structure

- [ ] 1. Go to your SharePoint site
- [ ] 2. Navigate to the document library where you want to host this
- [ ] 3. Create a new folder called: `iPipeline Finance Automation`
- [ ] 4. Inside that folder, create these subfolders:
   - `Videos`
   - `Demo File`
   - `Training Guides`
   - `Universal Code Library` (optional — can wait until after demo)

### Step 24 — Upload the Files

- [ ] 1. Upload `START HERE.pdf` to the **root** of `iPipeline Finance Automation/` (not inside a subfolder)
- [ ] 2. Upload video files to `Videos/`:
   - `1 - Whats Possible (Overview).mp4`
   - `2 - Full Demo Walkthrough.mp4`
   - `3 - Universal Tools.mp4`
- [ ] 3. Upload the demo file to `Demo File/`:
   - `iPipeline_PnL_Demo_FINAL.xlsm`
- [ ] 4. Upload the training PDFs to `Training Guides/`:
   - `01 - How to Use the Command Center.pdf`
   - `02 - Getting Started First Time Setup.pdf`
   - `03 - What This File Does (Leadership Overview).pdf`
   - `04 - Quick Reference Card.pdf`
   - `Copilot Prompts Guide.pdf`
- [ ] 5. (Optional — after demo) Upload universal tools to `Universal Code Library/`:
   - Create subfolders: `VBA/`, `Python/`, `SQL/`
   - Upload all `.bas` files from `UniversalToolsForAllFiles/vba/` to `VBA/`
   - Upload all `.py` files from `UniversalToolsForAllFiles/python/` to `Python/`
   - Upload all `.sql` files from `sql/` to `SQL/`
   - Upload `06 - Universal Toolkit Guide.pdf` to `Training Guides/`

### Step 25 — Set Permissions

- [ ] 1. Right-click the `iPipeline Finance Automation` folder → **Manage Access**
- [ ] 2. Add the appropriate group (your team, your department, or "Everyone" depending on CFO guidance)
- [ ] 3. Set permission level to **Read** (viewers) or **Edit** (if you want people to download)
- [ ] 4. Confirm the permissions look correct

### Step 26 — Make It Easy to Find

- [ ] 1. Pin the folder to your SharePoint site's Quick Links (if you have access)
- [ ] 2. Copy the SharePoint link to the folder
- [ ] 3. Save that link — you'll include it in any Teams messages, emails, or announcements

---

## PHASE 10: Final Announcement

**Goal:** Tell people it's there.

### Step 27 — Send the Announcement

- [ ] 1. Draft a Teams message or email to your audience. Something like:

> **Subject: iPipeline Finance Automation — Now Live on SharePoint**
>
> Hi everyone,
>
> I'm excited to share the iPipeline Finance Automation toolkit — a set of 62 automated actions built in Excel that handle month-end close tasks like reconciliation, variance analysis, data quality checks, and report generation.
>
> Everything is on SharePoint: [paste link]
>
> - **Short on time?** Watch Video 1 (5 min overview)
> - **On the Finance team?** Watch Video 2 (full walkthrough) and download the demo file
> - **Want tools for your own files?** Watch Video 3 and check out the Universal Code Library
>
> There's a START HERE document in the folder that explains everything.
>
> Questions? Reach out anytime.
>
> Connor

- [ ] 2. Send to your team / department / company (based on CFO's guidance on audience)
- [ ] 3. Share the link with the CFO/CEO directly (separate message — personal touch)

---

## YOU'RE DONE

When you've checked off everything above:

- [ ] The Excel demo file is tested, clean, and saved to production
- [ ] All training guides are converted to PDF
- [ ] The video(s) are recorded and saved
- [ ] Everything is uploaded to SharePoint with the right permissions
- [ ] The announcement has been sent
- [ ] The CFO/CEO has the link

**Congratulations — the project is delivered.**

---

## Quick Reference — Where Everything Lives

| What | Where |
|------|-------|
| VBA source code | `vba/` folder in your repo |
| Python scripts | `python/` folder in your repo |
| Training guides (markdown) | `FinalRoughGuides/` folder in your repo |
| Training guides (PDF) | `training/` + `CompletePackageStorage/production/` |
| Video scripts | `videodraft/COMPILED_VIDEO_PACKAGE.md` |
| Video plan & tips | `videodraft/VIDEO_DEMO_PLAN.md` |
| Universal tools (VBA) | `UniversalToolsForAllFiles/vba/` |
| Universal tools (Python) | `UniversalToolsForAllFiles/python/` |
| CoPilot prompt guide | `CoPilotPromptGuide/` |
| iPipeline brand colors | `docs/ipipeline-brand-styling.md` |
| Final production files | `CompletePackageStorage/production/` |
| Backups | `CompletePackageStorage/backups/` |
| This playbook | `WrappingUpAP/CONNORS_WRAP_UP_PLAYBOOK.md` |

---

*Created: 2026-03-09 | Branch: claude/resume-ipipeline-demo-qKRHn*
