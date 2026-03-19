# Video Production Guide — Start to Finish

**What This Is:** Your complete, step-by-step playbook for producing all 3 demo videos. Follow this document from top to bottom — it covers everything from Day 1 prep work all the way through uploading to SharePoint.

**How Long This Will Take:** Plan for about 5 working days spread across 1-2 weeks. You don't need to do it all in one sitting. Each day is a self-contained block you can start and stop.

**What You're Making:**
- **Video 1:** "What's Possible" — 4-5 minute overview for all 2,000+ employees
- **Video 2:** "Full Demo Walkthrough" — 15-18 minute deep dive for Finance & Accounting
- **Video 3:** "Universal Tools" — 8-10 minute demo for anyone who uses Excel

---

## Table of Contents

1. [Day 1 — Get Everything Ready (No Recording Yet)](#day-1--get-everything-ready)
2. [Day 2 — Generate All AI Audio Clips](#day-2--generate-all-ai-audio-clips)
3. [Day 3 — Record All Screen Videos](#day-3--record-all-screen-videos)
4. [Day 4 — Edit and Assemble the Videos](#day-4--edit-and-assemble-the-videos)
5. [Day 5 — Review, Export, and Upload to SharePoint](#day-5--review-export-and-upload)
6. [Quick Reference — File Naming Cheat Sheet](#quick-reference--file-naming-cheat-sheet)
7. [Quick Reference — Software Links](#quick-reference--software-links)
8. [Troubleshooting — What to Do If Something Goes Wrong](#troubleshooting)

---
---

# Day 1 — Get Everything Ready

**Goal:** By the end of today, your computer is set up, your software is installed, your Excel demo file is tested, and you have title cards ready. No recording today — just prep.

**Time estimate:** 2-3 hours

---

## Step 1 — Install the Software You Need

You need 3 pieces of software. All have free versions.

### 1A. Screen Recorder — OBS Studio (Free)

- [ ] 1. Open your browser and go to **obsproject.com**
- [ ] 2. Click **Windows** to download the installer
- [ ] 3. Run the installer — click Next through all screens, accept defaults
- [ ] 4. When it asks about auto-configuration, click **Cancel** — we'll set it up manually
- [ ] 5. OBS opens. Leave it open for now — we'll configure it in Step 4

### 1B. Video Editor — CapCut Desktop (Free)

CapCut is the simplest editor and is free. If you already have Camtasia or DaVinci Resolve, you can use those instead, but this guide writes instructions for CapCut.

- [ ] 1. Open your browser and go to **capcut.com**
- [ ] 2. Click **Download** for Windows
- [ ] 3. Run the installer — accept defaults
- [ ] 4. Open CapCut. It will ask you to sign in — you can use Google or create an account
- [ ] 5. Close CapCut for now — we'll use it on Day 4

### 1C. AI Voice Generator — ElevenLabs

- [ ] 1. Open your browser and go to **elevenlabs.io**
- [ ] 2. Click **Sign Up** (top right corner)
- [ ] 3. Create an account with your email (or sign in with Google)
- [ ] 4. You start on the free tier — about 10 minutes of audio per month
- [ ] 5. **You will need the Starter plan ($5/month)** — your 3 videos need about 25-30 minutes of total audio
- [ ] 6. To upgrade: Click your profile icon (bottom left) → **Subscription** → choose **Starter**
- [ ] 7. You can cancel after one month — you only need it for this project

**What you should see:** You're logged into ElevenLabs and can see the **Text to Speech** option in the left sidebar.

---

## Step 2 — Set Up Your AI Voice

You have two options. Pick one.

### Option A — Clone Your Own Voice (Recommended)

This makes the narration sound like you. It takes about 5 minutes.

- [ ] 1. You need a 1-minute recording of yourself talking naturally
   - Open the **Voice Recorder** app on your phone or Windows (search "Voice Recorder" in the Start menu)
   - Talk naturally for about 60 seconds — describe what you did last weekend, explain what you had for lunch, describe your desk. **Don't read a script** — just talk like you're explaining something to a coworker
   - Make sure you're in a quiet room with no background noise
   - Save the recording as an audio file (MP3, WAV, or M4A all work)
- [ ] 2. In ElevenLabs, click **Voices** in the left sidebar
- [ ] 3. Click **Add Voice** (top right)
- [ ] 4. Click **Instant Voice Cloning**
- [ ] 5. Click **Upload** and select your audio recording
- [ ] 6. Name the voice: **Connor - Demo**
- [ ] 7. Check the box that says you have rights to use this voice
- [ ] 8. Click **Add Voice**
- [ ] 9. Wait about 30 seconds — your voice will appear in your voice library

**What you should see:** "Connor - Demo" appears in your Voices list.

### Option B — Use a Built-In Voice

If you don't want to clone your voice:

- [ ] 1. In ElevenLabs, click **Text to Speech** in the left sidebar
- [ ] 2. Click the voice dropdown at the top
- [ ] 3. Browse the Voice Library — listen to previews
- [ ] 4. Pick a voice that sounds: male, American English, professional but natural, not robotic
- [ ] 5. Good starting options: **Adam**, **Josh**, or **Marcus**
- [ ] 6. Click the voice to select it

---

## Step 3 — Test Your Voice with a Sample Clip

Before generating all 37 audio clips, test one to make sure it sounds right.

- [ ] 1. In ElevenLabs, click **Text to Speech** in the left sidebar
- [ ] 2. Select your voice from the dropdown (either your cloned voice or the built-in one you picked)
- [ ] 3. Copy and paste this test text into the box:

```
This is a single Excel file. Nothing to install, nothing to configure...
you just open it and go. Inside are sixty-two automated actions that handle
reporting, analysis, data quality checks, charts, exports, and more...
each one triggered with a single click.
```

- [ ] 4. Set these settings (look for the gear icon or sliders near the Generate button):
   - **Model:** Eleven Multilingual v2
   - **Stability:** 50%
   - **Similarity:** 80%
   - **Style:** 20%
   - **Speaker Boost:** ON
- [ ] 5. Click **Generate**
- [ ] 6. Listen to the clip with headphones
- [ ] 7. Ask yourself:
   - Does it sound natural? Not robotic?
   - Is the pacing comfortable — not too fast, not too slow?
   - Does "sixty-two" come through clearly?
   - Would you be proud to play this for the CFO?

### If it doesn't sound right:

| Problem | What to Change |
|---------|----------------|
| Too robotic / monotone | Lower **Stability** to 40-45%, raise **Style** to 25% |
| Too dramatic / overdone | Raise **Stability** to 60%, lower **Style** to 10% |
| Doesn't sound like you (cloned voice) | Re-record your voice sample in a quieter room. Talk more naturally. Re-upload. |
| Pronunciation of a word is wrong | Spell it phonetically in the text. Example: "iPipeline" → "eye-Pipeline" |

- [ ] 8. Once it sounds good, **write down your exact settings** (Stability, Similarity, Style). You'll use the same settings for ALL clips so they sound consistent.

**What you should see:** A test clip that sounds professional, natural, and clear.

---

## Step 4 — Configure OBS (Screen Recorder)

- [ ] 1. Open OBS Studio
- [ ] 2. You should see a black preview screen with "Sources" at the bottom

### Add your screen as a source:
- [ ] 3. In the **Sources** panel at the bottom, click the **+** button
- [ ] 4. Click **Display Capture**
- [ ] 5. Click **OK** on the popup (you can leave the name as "Display Capture")
- [ ] 6. Click **OK** again — you should now see your screen in the preview

### Set recording quality:
- [ ] 7. Click **Settings** (bottom right)
- [ ] 8. Click **Output** on the left sidebar
- [ ] 9. Change these settings:
   - **Recording Path:** Click **Browse** and choose a folder you'll remember (create one called `VideoRecordings` on your Desktop)
   - **Recording Format:** Select **mp4**
   - Leave everything else as default
- [ ] 10. Click **Video** on the left sidebar
- [ ] 11. Set:
   - **Base (Canvas) Resolution:** 1920x1080
   - **Output (Scaled) Resolution:** 1920x1080
   - **FPS:** 30
- [ ] 12. Click **Audio** on the left sidebar
- [ ] 13. **THIS IS CRITICAL — Set both of these to Disabled:**
   - **Desktop Audio:** Disabled
   - **Mic/Auxiliary Audio:** Disabled
   - (You do NOT want OBS recording any sound. The AI audio gets added later in the editor.)
- [ ] 14. Click **Apply**, then **OK**

### Test a quick recording:
- [ ] 15. Click **Start Recording** (bottom right)
- [ ] 16. Move your mouse around, open a random window, close it
- [ ] 17. After 10 seconds, click **Stop Recording**
- [ ] 18. Go to your `VideoRecordings` folder — open the file
- [ ] 19. Verify: the video is clear, 1080p, no audio, smooth mouse movement

**What you should see:** A clean, silent screen recording in your folder.

---

## Step 5 — Prepare the Excel Demo File

The demo file needs to be in a perfect "clean state" before you record anything.

- [ ] 1. Open your demo `.xlsm` file (the one with all 39 VBA modules imported)
- [ ] 2. Make sure macros are enabled (if you see a yellow bar at the top, click **Enable Content**)
- [ ] 3. Go to each sheet tab at the bottom and **delete any leftover output sheets** from previous runs:
   - Delete: Variance Analysis (if it exists)
   - Delete: Variance Commentary (if it exists)
   - Delete: Data Quality Report (if it exists)
   - Delete: Executive Dashboard (if it exists)
   - Delete: YoY Variance Analysis (if it exists)
   - Delete: Sensitivity Analysis (if it exists)
   - Delete: any other output sheets that shouldn't be there
   - Right-click the tab → Delete → confirm
- [ ] 4. Navigate to the **Report-->** landing page (this is your starting point for recording)
- [ ] 5. Close the Command Center if it's open
- [ ] 6. Set zoom to **100%** (View tab → Zoom → 100%)
- [ ] 7. Make sure the ribbon is visible at the top (if it's collapsed, double-click any ribbon tab to expand it)
- [ ] 8. Save the file (**Ctrl+S**)

### Test that key macros actually work:
- [ ] 9. Press **Ctrl+Shift+M** to open the Command Center
   - **What you should see:** The Command Center form pops up with all actions listed
- [ ] 10. Run **Action 1 — Data Quality Scan** from the Command Center
   - **What you should see:** A "Data Quality Report" sheet gets created with results and a letter grade
- [ ] 11. Delete the Data Quality Report sheet (right-click tab → Delete)
- [ ] 12. Close the Command Center
- [ ] 13. Save the file again

**If any macro errors:** Stop here. Fix the issue before proceeding. You cannot record a demo with broken macros.

---

## Step 6 — Prepare the Sample File for Video 3

Video 3 uses a separate file — `Sample_Quarterly_Report.xlsx` — to show that the universal tools work on any file, not just the demo.

- [ ] 1. Check if the sample file already exists in the repo: look in `FinalExport/` or `Archive/videodraft/`
- [ ] 2. If it exists, open it and verify it has messy data (blank rows, text-stored numbers, formatting issues) — this is intentional
- [ ] 3. If it doesn't exist, you'll need to create a simple messy spreadsheet:
   - Open a new Excel file
   - Add 3-4 columns: Date, Description, Amount, Category
   - Add 30-50 rows of sample data
   - Intentionally add some problems: a few blank rows, some numbers stored as text (put an apostrophe before a number), inconsistent date formats, a duplicate row
   - Save as `Sample_Quarterly_Report.xlsx`
- [ ] 4. Import the universal toolkit VBA modules into this file so you can run tools on it during recording
- [ ] 5. Save the file

---

## Step 7 — Create Title Cards

You need simple title cards for the start and end of each video, plus chapter cards for Videos 2 and 3.

### How to Make Title Cards in PowerPoint:

- [ ] 1. Open PowerPoint
- [ ] 2. Set slide size to **Widescreen (16:9)** — this matches 1920x1080 video
   - Design tab → Slide Size → Widescreen (16:9)
- [ ] 3. Create these slides (one per title card):

**Slide 1 — Video 1 Opening Title:**
- Background color: iPipeline Blue (#0B4779)
- Text (centered, white, Arial Bold, large):
  ```
  Finance Automation
  What's Possible
  ```

**Slide 2 — Video 1 Closing Title:**
- Background color: iPipeline Blue (#0B4779)
- Text (centered, white, Arial Bold):
  ```
  Want to learn more?
  See the full walkthrough in the "Full Demo" video
  All files and guides available on SharePoint
  ```

**Slide 3 — Video 2 Opening Title:**
- Background color: iPipeline Blue (#0B4779)
- Text: `Finance Automation — Full Demo Walkthrough`

**Slides 4-10 — Video 2 Chapter Cards** (one for each chapter):
- Background: Navy (#112E51)
- Text (white, Arial Bold):
  ```
  Chapter 1: The Workbook & Command Center
  Chapter 2: Data Import & Quality
  Chapter 3: Analysis
  Chapter 4: Reporting & Visuals
  Chapter 5: Enterprise Features
  Chapter 6: Under the Hood
  Chapter 7: Next Steps
  ```

**Slide 11 — Video 2 Closing Title:**
- Same as Video 1 closing but says "See the Universal Tools video" instead

**Slide 12 — Video 3 Opening Title:**
- Background: iPipeline Blue (#0B4779)
- Text: `Universal Tools — For Any Excel File`

**Slide 13 — Video 3 Chapter Cards** (as needed — similar to Video 2)

**Slide 14 — Video 3 Closing Title:**
- Text: `All tools and guides available on SharePoint`

### Export each slide as an image:

- [ ] 4. For each slide: File → Save As → choose **PNG** format
- [ ] 5. When asked, click **Just This Slide** (not all slides)
- [ ] 6. Save each image in your `VideoRecordings` folder with a clear name:
   - `V1_Title_Open.png`
   - `V1_Title_Close.png`
   - `V2_Title_Open.png`
   - `V2_Chapter1.png`
   - `V2_Chapter2.png`
   - etc.

**What you should see:** A folder of PNG images, one for each title/chapter card.

---

## Step 8 — Set Up Your Folder Structure

Create this folder structure on your Desktop (or wherever you're working):

```
VideoProject/
├── Audio/
│   ├── Video1/          ← ElevenLabs clips for Video 1 go here
│   ├── Video2/          ← ElevenLabs clips for Video 2 go here
│   └── Video3/          ← ElevenLabs clips for Video 3 go here
├── ScreenRecordings/
│   ├── Video1/          ← OBS screen recordings for Video 1
│   ├── Video2/          ← OBS screen recordings for Video 2
│   └── Video3/          ← OBS screen recordings for Video 3
├── TitleCards/           ← All PNG title/chapter card images
└── FinalExport/          ← Where your finished MP4 videos go
```

- [ ] 1. Create the `VideoProject` folder on your Desktop
- [ ] 2. Create all the subfolders listed above
- [ ] 3. Move your title card PNGs into `TitleCards/`
- [ ] 4. Update OBS recording path to point to the appropriate `ScreenRecordings/` subfolder (you'll change this before each video)

**What you should see:** An organized folder structure ready for all your files.

---

## Step 9 — Print or Open Your Scripts

You'll need the narration scripts visible while you record.

- [ ] 1. The scripts are in the repo at:
   - `Archive/videodraft/AI_Narration_Scripts/01-VIDEO1-WHATS-POSSIBLE.md`
   - `Archive/videodraft/AI_Narration_Scripts/02-VIDEO2-FULL-DEMO-WALKTHROUGH.md`
   - `Archive/videodraft/AI_Narration_Scripts/03-VIDEO3-UNIVERSAL-TOOLS.md`
- [ ] 2. **Option A (Best):** If you have a second monitor, open the script on the second monitor while Excel is on the primary monitor
- [ ] 3. **Option B:** Print the scripts — each segment on its own page so you can flip through them
- [ ] 4. **Option C:** Open the script on your phone or tablet propped up next to your monitor

**What you should see:** Easy access to the screen action instructions for each segment.

---

### Day 1 Checklist — Confirm Before Moving On

Before moving to Day 2, confirm all of these:

- [ ] OBS is installed and tested (silent 1080p recording works)
- [ ] CapCut is installed (don't need to configure yet)
- [ ] ElevenLabs account is set up with Starter plan ($5/month)
- [ ] Your AI voice is selected and tested (settings written down)
- [ ] Excel demo file opens, macros work, clean state saved
- [ ] Sample file for Video 3 is ready
- [ ] Title card images are created and saved as PNGs
- [ ] Folder structure is set up
- [ ] Scripts are accessible (second monitor, printed, or on phone)

**If all boxes are checked — you're done for today. Nice work.**

---
---

# Day 2 — Generate All AI Audio Clips

**Goal:** By the end of today, you have all 37 audio clips generated, downloaded, organized in folders, and reviewed for quality.

**Time estimate:** 2-3 hours

---

## How This Works

Each video script is broken into numbered **segments**. Each segment is one audio clip. You'll copy the text from each segment, paste it into ElevenLabs, generate the clip, listen to it, and download it.

**Clip counts:**
- Video 1: **7 segments** (7 audio clips)
- Video 2: **16 segments** (16 audio clips)
- Video 3: **14 segments** (14 audio clips)
- **Total: 37 audio clips**

---

## Step 1 — Open Your Script and ElevenLabs Side by Side

- [ ] 1. Open the Video 1 script: `Archive/videodraft/AI_Narration_Scripts/01-VIDEO1-WHATS-POSSIBLE.md`
- [ ] 2. Open ElevenLabs in your browser: **elevenlabs.io** → click **Text to Speech**
- [ ] 3. Select your voice from the dropdown
- [ ] 4. Set your settings to the exact values you wrote down on Day 1:
   - Stability: ____%
   - Similarity: ____%
   - Style: ____%
   - Speaker Boost: ON

**IMPORTANT:** Do NOT change these settings between clips. Keep them identical for all 37 clips so the voice sounds consistent across all 3 videos.

---

## Step 2 — Generate Video 1 Clips (7 clips)

For each segment in the Video 1 script, do this:

- [ ] 1. Find the **[PASTE INTO ELEVENLABS]** box in the script
- [ ] 2. Copy the text inside the box (just the narration text — not the section headers or screen actions)
- [ ] 3. Paste it into the ElevenLabs text box
- [ ] 4. Click **Generate**
- [ ] 5. Wait for it to finish (usually 5-15 seconds)
- [ ] 6. **Listen to the clip with headphones** — the whole thing, start to finish
- [ ] 7. Ask yourself:
   - Does it sound natural?
   - Are there any weird pauses or mispronunciations?
   - Does it match the energy of your previous clips?
- [ ] 8. If it sounds good → click the **download arrow** → save as MP3
- [ ] 9. Save it in `VideoProject/Audio/Video1/` with the filename from the script:
   - `V1_S1_Opening_Hook.mp3`
   - `V1_S2_Command_Center.mp3`
   - etc.
- [ ] 10. If it sounds off → **don't change settings**. Instead, try these:
   - Re-generate with the exact same text (ElevenLabs produces slightly different results each time)
   - If a specific word sounds wrong, spell it phonetically in the text
   - If there's a weird pause, remove or add punctuation around that spot

### Video 1 Segments Checklist:

- [ ] V1_S1 — Opening Hook (~25 sec)
- [ ] V1_S2 — Command Center Introduction (~30 sec)
- [ ] V1_S3 — Data Quality Scan (~40 sec)
- [ ] V1_S4 — Variance Commentary (~40 sec)
- [ ] V1_S5 — Dashboard / Executive Dashboard (~40 sec)
- [ ] V1_S6 — Bridge to Universal Tools (~30 sec)
- [ ] V1_S7 — Closing + Call to Action (~30 sec)

**After all 7 clips are downloaded:** Play them back-to-back in order (just open them in Windows Media Player or VLC one by one). Make sure the voice, pacing, and energy feel consistent. If one clip sounds noticeably different, re-generate it.

---

## Step 3 — Generate Video 2 Clips (16 clips)

Same process. Use the Video 2 script: `02-VIDEO2-FULL-DEMO-WALKTHROUGH.md`

Save clips in `VideoProject/Audio/Video2/`

### Video 2 Segments Checklist:

- [ ] V2_S0 — Opening (~40 sec)
- [ ] V2_S1 — Chapter 1: Workbook & Command Center
- [ ] V2_S2 — Chapter 1 continued (Command Center search/launch)
- [ ] V2_S3 — Chapter 2: Data Import
- [ ] V2_S4 — Chapter 2: Data Quality Scan + Letter Grade
- [ ] V2_S5 — Chapter 2: Reconciliation Checks
- [ ] V2_S6 — Chapter 3: Variance Analysis
- [ ] V2_S7 — Chapter 3: Variance Commentary
- [ ] V2_S8 — Chapter 3: YoY Variance
- [ ] V2_S9 — Chapter 4: Dashboard Charts
- [ ] V2_S10 — Chapter 4: Executive Dashboard
- [ ] V2_S11 — Chapter 4: PDF Export
- [ ] V2_S12 — Chapter 5: Executive Mode + Version Control
- [ ] V2_S13 — Chapter 5: Scenario + Sensitivity
- [ ] V2_S14 — Chapter 6: Integration Test + Audit Log
- [ ] V2_S15 — Chapter 7: Closing

**After all 16 clips:** Play them back-to-back. Check for consistency.

---

## Step 4 — Generate Video 3 Clips (14 clips)

Same process. Use the Video 3 script: `03-VIDEO3-UNIVERSAL-TOOLS.md`

Save clips in `VideoProject/Audio/Video3/`

### Video 3 Segments Checklist:

- [ ] V3_S0 — Opening (~45 sec)
- [ ] V3_S1 through V3_S13 — (follow the script segment numbers)

**After all 14 clips:** Play them back-to-back. Check for consistency.

---

## Step 5 — Final Audio Review

Before moving on, do one final quality pass:

- [ ] 1. Play ALL Video 1 clips in order (should total ~4-5 minutes)
- [ ] 2. Play ALL Video 2 clips in order (should total ~15-18 minutes)
- [ ] 3. Play ALL Video 3 clips in order (should total ~8-10 minutes)
- [ ] 4. Listen for:
   - Any clip that sounds noticeably different in tone or energy
   - Any mispronounced words (especially "iPipeline", "P-and-L", product names)
   - Any unnatural pauses
   - Any clip that's too fast or too slow
- [ ] 5. Re-generate and replace any clips that don't pass

**What you should see:** 37 MP3 files organized in 3 folders, all sounding consistent and professional.

---

### Day 2 Checklist — Confirm Before Moving On

- [ ] All 7 Video 1 clips generated and saved in `Audio/Video1/`
- [ ] All 16 Video 2 clips generated and saved in `Audio/Video2/`
- [ ] All 14 Video 3 clips generated and saved in `Audio/Video3/`
- [ ] Played all clips back-to-back — consistent voice and energy
- [ ] No mispronounced words or weird audio artifacts
- [ ] Total audio is roughly 27-33 minutes across all 3 videos

**If all boxes are checked — audio is done. That's a huge piece finished.**

---
---

# Day 3 — Record All Screen Videos

**Goal:** By the end of today, you have silent screen recordings for every segment of all 3 videos.

**Time estimate:** 3-4 hours (including practice runs)

---

## Step 1 — Lock Down Your Computer

Do this EVERY time before you start recording. Every single time. No exceptions.

- [ ] 1. **Close everything** except Excel and OBS
   - Close all browser tabs
   - Close Outlook completely (not minimized — closed)
   - Close Teams completely (or set to "Do Not Disturb": click your profile picture → Do Not Disturb)
   - Close any other apps
- [ ] 2. **Turn off ALL notifications:**
   - Windows: Click the notification icon in the bottom right → click **Focus Assist** → select **Alarms Only**
   - Or: Settings → System → Notifications → toggle OFF
- [ ] 3. **Clean your desktop:**
   - Right-click your desktop → View → uncheck **Show desktop icons**
   - (You can turn this back on after recording)
- [ ] 4. **Auto-hide the taskbar:**
   - Right-click the taskbar → **Taskbar settings** → turn on **Automatically hide the taskbar**
- [ ] 5. **Set display to 1920x1080 at 100% scaling:**
   - Settings → Display → **Scale:** 100% (not 125% or 150%)
   - Settings → Display → **Display resolution:** 1920 x 1080
- [ ] 6. **Plug in your laptop** (prevent battery throttling / sleep)
- [ ] 7. **Put your phone on silent or airplane mode**

**Why this matters:** One Teams notification popping up during recording ruins the take and you have to re-do it.

---

## Step 2 — Set Up Excel for Recording

- [ ] 1. Open your demo `.xlsm` file
- [ ] 2. **Maximize** Excel to fill the entire screen (click the maximize button or double-click the title bar)
- [ ] 3. Set **zoom to 100%** (View tab → Zoom → 100%)
- [ ] 4. Make sure the **ribbon is visible** at the top
- [ ] 5. Navigate to the **Report-->** landing page
- [ ] 6. **Close** the Command Center if it's open
- [ ] 7. Delete any leftover output sheets from previous test runs
- [ ] 8. **Save** the file (Ctrl+S)

**This is your "clean state."** After each segment that creates output (like a Data Quality Report), you'll need to delete that output and return to this clean state before the next segment.

---

## Step 3 — How the Recording Process Works

Here's the flow for recording each segment:

```
1. Put on headphones
2. Open the audio clip for this segment (don't play yet)
3. Open OBS — click "Start Recording"
4. Wait 2 seconds (silent lead-in)
5. Play the audio clip in your headphones
6. Watch the screen and click along with what the voice says
7. When the audio ends, wait 2 seconds
8. Click "Stop Recording" in OBS
9. Name/move the recording to the right folder
10. Reset Excel if needed (delete output sheets, go back to starting point)
11. Move to the next segment
```

**Key principle:** The audio controls the pacing. You follow the audio — not the other way around. When the voice says "I'm opening the Command Center," that's when you press Ctrl+Shift+M. When the voice pauses, you pause.

---

## Step 4 — Practice Run (Do This Before Recording for Real)

Pick Segment 1.1 (Opening Hook) from Video 1.

- [ ] 1. Read the **[YOUR SCREEN ACTIONS]** section in the script
- [ ] 2. Put on headphones
- [ ] 3. Play `V1_S1_Opening_Hook.mp3`
- [ ] 4. Practice doing the screen actions while listening — but **don't record** this time
   - The script says: "slowly scroll down the Report page"
   - Practice scrolling at the right speed — smooth, deliberate, matching the audio pace
- [ ] 5. Do it twice. Get comfortable with the timing
- [ ] 6. Now do it for real with OBS recording

**Tips for smooth mouse movement:**
- Move your mouse **slowly and deliberately** — fast jerky movements look bad on video
- When clicking a button, move to it, pause for half a second, then click — this gives the viewer time to see what you're about to click
- Scroll smoothly — use the scroll wheel gently, not in big jumps
- If a macro takes a few seconds to run, just wait. Don't click anything. Let the viewer see it working.

---

## Step 5 — Record Video 1 (7 segments)

Update OBS recording path to `VideoProject/ScreenRecordings/Video1/`

For each segment, follow the process from Step 3. Use the script's **[YOUR SCREEN ACTIONS]** section to know exactly what to do on screen.

### Video 1 Recording Checklist:

| Segment | What Happens On Screen | Reset After? |
|---------|----------------------|--------------|
| V1_S1 | Scroll down Report page slowly | No |
| V1_S2 | Open Command Center (Ctrl+Shift+M), browse categories, use search | Close Command Center |
| V1_S3 | Run Data Quality Scan from Command Center — show the report + letter grade | Delete Data Quality Report sheet |
| V1_S4 | Run Variance Commentary — show the auto-generated narratives | Delete output sheet |
| V1_S5 | Run Executive Dashboard — show charts and visuals | Delete output sheet |
| V1_S6 | Brief mention of universal tools — maybe show the SharePoint folder or guide | No |
| V1_S7 | Closing — hold on Report page | No |

- [ ] V1_S1 recorded
- [ ] V1_S2 recorded
- [ ] V1_S3 recorded
- [ ] V1_S4 recorded
- [ ] V1_S5 recorded
- [ ] V1_S6 recorded
- [ ] V1_S7 recorded

**After each recording:** Play it back quickly to make sure it looks good. If you messed up, just re-do that one segment. That's the beauty of recording in segments.

---

## Step 6 — Record Video 2 (16 segments)

Update OBS recording path to `VideoProject/ScreenRecordings/Video2/`

This is the longest video. Take a break between chapters if you need to.

### Key Reset Points for Video 2:

- After Data Quality Scan → delete the output sheet
- After Reconciliation → delete output sheet (if created)
- After each Analysis action → delete the output sheet
- After Executive Dashboard → delete the output sheet
- After PDF Export → close the PDF viewer
- After each Enterprise Feature demo → undo/reset as needed

### Video 2 Recording Checklist:

- [ ] V2_S0 — Opening (scroll Report page)
- [ ] V2_S1 — Workbook overview (flip through sheet tabs)
- [ ] V2_S2 — Command Center (open, search, browse)
- [ ] V2_S3 — Data Import demo
- [ ] V2_S4 — Data Quality Scan + Letter Grade
- [ ] V2_S5 — Reconciliation Checks (show PASS/FAIL)
- [ ] V2_S6 — Variance Analysis
- [ ] V2_S7 — Variance Commentary (auto narratives)
- [ ] V2_S8 — YoY Variance
- [ ] V2_S9 — Dashboard Charts
- [ ] V2_S10 — Executive Dashboard
- [ ] V2_S11 — PDF Export
- [ ] V2_S12 — Executive Mode + Version Control
- [ ] V2_S13 — Scenario + Sensitivity
- [ ] V2_S14 — Integration Test (18/18 PASS) + Audit Log
- [ ] V2_S15 — Closing

---

## Step 7 — Record Video 3 (14 segments)

Update OBS recording path to `VideoProject/ScreenRecordings/Video3/`

**IMPORTANT:** Video 3 uses the **Sample_Quarterly_Report.xlsx** file — NOT the demo file. Switch files before recording.

- [ ] 1. Close the demo .xlsm file
- [ ] 2. Open Sample_Quarterly_Report.xlsx
- [ ] 3. Make sure the universal toolkit VBA modules are imported into this file
- [ ] 4. Maximize Excel, zoom 100%
- [ ] 5. Record all 14 segments following the Video 3 script

### Video 3 Recording Checklist:

- [ ] V3_S0 — Opening (show the messy spreadsheet)
- [ ] V3_S1 through V3_S13 — Follow the script

---

### Day 3 Checklist — Confirm Before Moving On

- [ ] All 7 Video 1 screen recordings saved in `ScreenRecordings/Video1/`
- [ ] All 16 Video 2 screen recordings saved in `ScreenRecordings/Video2/`
- [ ] All 14 Video 3 screen recordings saved in `ScreenRecordings/Video3/`
- [ ] Quick playback of each recording — screen is clear, no notifications, actions visible
- [ ] No accidental audio recorded (all recordings should be silent)

**If all boxes are checked — the hard part is done. You have all the raw materials.**

---
---

# Day 4 — Edit and Assemble the Videos

**Goal:** By the end of today, you have 3 complete videos assembled in CapCut with audio synced to video, title cards inserted, and any timing adjustments made.

**Time estimate:** 3-4 hours

---

## Step 1 — Open CapCut and Create Your First Project

Start with Video 1 (it's the shortest and simplest — good for learning the editor).

- [ ] 1. Open CapCut
- [ ] 2. Click **New Project**
- [ ] 3. Set project resolution to **1920 x 1080** and **30 fps** (it may default to this)

---

## Step 2 — Import All Your Files for Video 1

- [ ] 1. Click **Import** (or drag files into the media panel)
- [ ] 2. Import everything from these folders:
   - All 7 audio clips from `Audio/Video1/`
   - All 7 screen recordings from `ScreenRecordings/Video1/`
   - The title card images: `V1_Title_Open.png` and `V1_Title_Close.png`

**What you should see:** All your files listed in the media panel on the left side.

---

## Step 3 — Build the Timeline for Video 1

The timeline is where you arrange everything in order. It has tracks — video goes on the top track, audio goes on the bottom track.

### 3A — Add the Opening Title Card:

- [ ] 1. Drag `V1_Title_Open.png` to the beginning of the **video track** on the timeline
- [ ] 2. Click on it and set its **duration to 5 seconds** (drag the right edge or type in the duration)
- [ ] 3. This will be the first thing the viewer sees — 5 seconds of the branded title card

### 3B — Add Segment 1 (and repeat for all segments):

- [ ] 4. Drag `V1_S1` screen recording onto the video track, right after the title card
- [ ] 5. Drag `V1_S1_Opening_Hook.mp3` onto the **audio track** directly below the video
- [ ] 6. **Align them:** The audio should start about 1-2 seconds after the video starts (this gives a moment of silent visual before the voice kicks in)
- [ ] 7. **Watch this segment:** Press Play on the timeline and watch it. Does the voice say "I'm scrolling" right when you're scrolling? Does it feel natural?
- [ ] 8. **Adjust timing if needed:**
   - If the voice is ahead of your actions → drag the audio clip to the right (later in time)
   - If your actions are ahead of the voice → drag the audio clip to the left (earlier in time)
   - Small adjustments of 0.5-2 seconds are normal and expected
- [ ] 9. Repeat for segments S2 through S7

### 3C — Add the Closing Title Card:

- [ ] 10. After the last segment, drag `V1_Title_Close.png` onto the video track
- [ ] 11. Set its duration to **5 seconds**

### 3D — Add Transitions (Optional but Professional):

- [ ] 12. Between the title card and the first segment, add a **crossfade transition** (in CapCut: click **Transitions** → drag a crossfade between the two clips)
- [ ] 13. Set transition duration to **0.5 seconds**
- [ ] 14. Add the same transition between the last segment and the closing title card
- [ ] 15. You do NOT need transitions between every segment — just between title cards and content

---

## Step 4 — Add Time Savings Text Overlays (Video 2 Only)

The master plan calls for 3-4 "time savings" text overlays in Video 2, shown after key macro demos. Example:

```
Manual: 2 hours → Automated: 10 seconds
```

In CapCut:
- [ ] 1. Click **Text** in the top toolbar
- [ ] 2. Click **Add Text**
- [ ] 3. Type the time savings message
- [ ] 4. Style it: white text, dark semi-transparent background, bottom-right corner of the screen
- [ ] 5. Set it to appear for **3 seconds** after the macro finishes running
- [ ] 6. Add these after: Data Quality Scan, Variance Commentary, Executive Dashboard, PDF Export

---

## Step 5 — Build Videos 2 and 3

Repeat the same process for Video 2 and Video 3:

### Video 2:
- [ ] 1. Create a new project in CapCut
- [ ] 2. Import all 16 audio clips, 16 screen recordings, and title/chapter card images
- [ ] 3. Build the timeline in this order:
   - Opening title card (5 sec)
   - Segment 0 (opening)
   - Chapter 1 card (3 sec) → Segments 1-2
   - Chapter 2 card (3 sec) → Segments 3-5
   - Chapter 3 card (3 sec) → Segments 6-8
   - Chapter 4 card (3 sec) → Segments 9-11
   - Chapter 5 card (3 sec) → Segments 12-13
   - Chapter 6 card (3 sec) → Segment 14
   - Segment 15 (closing)
   - Closing title card (5 sec)
- [ ] 4. Sync all audio clips to their matching screen recordings
- [ ] 5. Add time savings overlays (3-4 total)
- [ ] 6. Add crossfade transitions between title/chapter cards and content

### Video 3:
- [ ] 1. Create a new project in CapCut
- [ ] 2. Import all 14 audio clips, 14 screen recordings, and title/chapter card images
- [ ] 3. Build the timeline (same pattern: title card → segments with chapter cards → closing card)
- [ ] 4. Sync all audio clips
- [ ] 5. Add transitions

---

## Step 6 — Watch Each Video All the Way Through

This is your first real preview. Watch each video from start to finish as if you're a coworker seeing it for the first time.

### Watch for:

- [ ] Audio and video are in sync (voice matches the screen actions)
- [ ] No awkward gaps longer than 3 seconds of silence
- [ ] No jump cuts where something appears/disappears suddenly without explanation
- [ ] Title cards are readable and on screen long enough
- [ ] The overall flow makes sense — one section leads naturally into the next
- [ ] No leftover output sheets visible that shouldn't be there
- [ ] Total runtime is in the target range:
   - Video 1: 4-5 minutes
   - Video 2: 15-18 minutes
   - Video 3: 8-10 minutes

### Common fixes at this stage:

| Issue | How to Fix |
|-------|-----------|
| Gap between segments feels too long | Trim the end of the previous clip or the start of the next clip |
| Voice says something but the action already happened | Nudge the audio clip slightly later on the timeline |
| A chapter card flashes by too fast | Extend its duration to 3-4 seconds |
| Total video is too long | Check for segments where you waited too long before clicking — trim those pauses |
| A specific segment looks bad | Don't re-edit — re-record just that one screen segment (Day 3 Step 3 process) and swap it in |

---

### Day 4 Checklist — Confirm Before Moving On

- [ ] Video 1 fully assembled — audio synced, title cards, transitions
- [ ] Video 2 fully assembled — audio synced, chapter cards, time savings overlays, transitions
- [ ] Video 3 fully assembled — audio synced, chapter cards, transitions
- [ ] Watched all 3 videos start to finish — they flow well and look professional
- [ ] Runtimes are within target ranges
- [ ] No sync issues, no notification pop-ups, no artifacts

**If all boxes are checked — your videos are built. One more day to polish and deliver.**

---
---

# Day 5 — Review, Export, and Upload

**Goal:** Final quality check, export the finished MP4 files, and upload everything to SharePoint.

**Time estimate:** 2-3 hours

---

## Step 1 — Final Quality Review

Watch each video one more time with fresh eyes. Pretend you're the CFO seeing this for the first time.

### Video 1 Review:
- [ ] Would the CFO get the point in 4-5 minutes?
- [ ] Does the opening hook grab attention in the first 10 seconds?
- [ ] Are the 4 feature demos clear and impressive?
- [ ] Does it end with a clear "where to find more" message?

### Video 2 Review:
- [ ] Does each chapter flow naturally into the next?
- [ ] Are the time savings overlays visible and impactful?
- [ ] Can a Finance team member follow along and understand what each feature does?
- [ ] Does the Integration Test (18/18 PASS) moment land as a confidence builder?

### Video 3 Review:
- [ ] Is it clear that these tools work on ANY file (not just the demo)?
- [ ] Are the before/after moments clear for each tool demo?
- [ ] Does it end with clear instructions on where to get the tools?

### Ask yourself for ALL 3 videos:
- [ ] Would I be proud to show this to 2,000+ people?
- [ ] Is the audio clear and professional?
- [ ] Are the screen actions visible and easy to follow?
- [ ] Is this truly world-class?

---

## Step 2 — Export Final MP4 Files

### In CapCut:

For each video:

- [ ] 1. Click **Export** (top right corner)
- [ ] 2. Set these settings:
   - **Resolution:** 1920 x 1080
   - **Frame Rate:** 30 fps
   - **Format:** MP4
   - **Quality:** High (or "Recommended")
- [ ] 3. Set the export location to `VideoProject/FinalExport/`
- [ ] 4. Name the files exactly:
   - `1 - What's Possible (Overview).mp4`
   - `2 - Full Demo Walkthrough.mp4`
   - `3 - Universal Tools.mp4`
- [ ] 5. Click **Export** and wait (may take several minutes per video)
- [ ] 6. After export, **play each file** from the folder to verify it exported correctly

**What you should see:** 3 MP4 files in your FinalExport folder, all playing perfectly.

---

## Step 3 — Build the SharePoint Folder

This is the folder your coworkers will access. Everything they need goes here.

### Create the folder structure:

- [ ] 1. Go to your SharePoint site (or Teams → Files)
- [ ] 2. Create a top-level folder: **Finance Automation**
- [ ] 3. Inside that, create these subfolders:
   ```
   Finance Automation/
   ├── Videos/
   ├── Demo File/
   ├── Universal Code Library/
   │   ├── VBA/
   │   ├── Python/
   │   └── SQL/
   └── Training Guides/
   ```

### Upload files:

- [ ] 4. Upload to **Videos/**:
   - `1 - What's Possible (Overview).mp4`
   - `2 - Full Demo Walkthrough.mp4`
   - `3 - Universal Tools.mp4`

- [ ] 5. Upload to **Demo File/**:
   - Your final `.xlsm` demo file
   - `Sample_Quarterly_Report.xlsx` (the file used in Video 3)

- [ ] 6. Upload to **Universal Code Library/VBA/**:
   - All universal toolkit `.bas` files from `FinalExport/UniversalToolkit/vba/`

- [ ] 7. Upload to **Universal Code Library/Python/**:
   - All Python scripts from `FinalExport/UniversalToolkit/python/`

- [ ] 8. Upload to **Universal Code Library/SQL/**:
   - All SQL scripts from `FinalExport/DemoPython/sql/`

- [ ] 9. Upload to **Training Guides/**:
   - All PDF guides from `FinalExport/Guides/`

- [ ] 10. Upload to the **root** of Finance Automation/:
   - The START HERE PDF (if created)
   - The CoPilot Prompt Guide PDF

---

## Step 4 — Set Permissions

- [ ] 1. Right-click the **Finance Automation** folder → **Manage Access**
- [ ] 2. Set permissions so your target audience can view/download but not edit
- [ ] 3. Consider: one permission group for the whole folder (simpler than per-subfolder)

---

## Step 5 — Test Everything as a Viewer

- [ ] 1. Open the SharePoint folder in an **InPrivate/Incognito browser window** (to see it as a coworker would)
- [ ] 2. Or ask a trusted coworker to try opening the folder
- [ ] 3. Verify:
   - [ ] All 3 videos play directly in SharePoint (or download and play)
   - [ ] The demo `.xlsm` file downloads correctly
   - [ ] The PDF guides open correctly
   - [ ] The folder structure makes sense to someone seeing it for the first time

---

## Step 6 — Send the Announcement

You're done! Time to tell people about it.

- [ ] 1. Draft a short Teams message or email to your audience:
   - Subject: "Finance Automation Toolkit — Now Available"
   - Include: link to the SharePoint folder
   - Mention: "Start with the 5-minute overview video"
   - Keep it short — let the video do the talking

---

### Day 5 Checklist — Confirm You're Done

- [ ] 3 MP4 videos exported and verified
- [ ] SharePoint folder created with correct structure
- [ ] All files uploaded (videos, demo file, code library, guides, prompts)
- [ ] Permissions set
- [ ] Tested as a viewer — everything opens and plays
- [ ] Announcement drafted

**If all boxes are checked — you did it. Congratulations.**

---
---

# Quick Reference — File Naming Cheat Sheet

## Audio Clips (from ElevenLabs)
```
Audio/Video1/V1_S1_Opening_Hook.mp3
Audio/Video1/V1_S2_Command_Center.mp3
Audio/Video1/V1_S3_Data_Quality.mp3
...etc (follow naming in each script)

Audio/Video2/V2_S0_Opening.mp3
Audio/Video2/V2_S1_Chapter1_Workbook.mp3
...etc

Audio/Video3/V3_S0_Opening.mp3
...etc
```

## Screen Recordings (from OBS)
```
ScreenRecordings/Video1/V1_S1_Opening_Hook.mp4
ScreenRecordings/Video1/V1_S2_Command_Center.mp4
...etc (match the audio clip names)
```

## Title Cards (from PowerPoint)
```
TitleCards/V1_Title_Open.png
TitleCards/V1_Title_Close.png
TitleCards/V2_Title_Open.png
TitleCards/V2_Chapter1.png
TitleCards/V2_Chapter2.png
...etc
```

## Final Exports
```
FinalExport/1 - What's Possible (Overview).mp4
FinalExport/2 - Full Demo Walkthrough.mp4
FinalExport/3 - Universal Tools.mp4
```

---

# Quick Reference — Software Links

| Software | URL | Cost |
|----------|-----|------|
| OBS Studio (screen recorder) | obsproject.com | Free |
| CapCut Desktop (video editor) | capcut.com | Free |
| ElevenLabs (AI voice) | elevenlabs.io | $5/month (Starter) |
| DaVinci Resolve (advanced editor) | blackmagicdesign.com | Free |
| VLC Media Player | videolan.org | Free |

---

# Troubleshooting

## "A macro crashed during screen recording"
- Stop recording. Fix the issue in Excel (or delete the output sheet and try again). Re-record just that segment. You don't have to redo the whole video.

## "The audio and video don't line up"
- This is normal. That's what the video editor is for. Drag clips on the timeline to adjust. Small adjustments of 0.5-2 seconds are typical and expected.

## "OBS recorded audio I didn't want"
- Go to OBS Settings → Audio → set Desktop Audio and Mic/Auxiliary Audio both to "Disabled". Re-record the affected segment.

## "One audio clip sounds different from the others"
- Re-generate it in ElevenLabs with the exact same settings. ElevenLabs produces slightly different results each time — just keep generating until it matches. Do NOT change the Stability/Similarity/Style settings.

## "ElevenLabs says I'm out of characters"
- Upgrade to Starter plan ($5/month). You can cancel after the project.

## "CapCut won't import my OBS recordings"
- Make sure OBS is recording in MP4 format (Settings → Output → Recording Format → mp4). If you recorded in MKV, convert to MP4 first (OBS has a built-in converter: File → Remux Recordings).

## "The video editor is confusing"
- CapCut is the simplest option. If you're stuck, search YouTube for "CapCut beginner tutorial" — there are great 10-minute walkthroughs.

## "I want to re-record a segment but the Excel state has changed"
- Reload the clean-state version of the demo file (the one you saved on Day 1 Step 5). Run any macros needed to get back to the right starting point for that segment, then re-record.

## "The title cards look blurry"
- Make sure your PowerPoint slide size is 16:9 (Widescreen) and you export as PNG (not JPG). PNG preserves sharp text.

## "I missed a step and need to go back"
- That's fine. Each Day is independent once you have the outputs from the previous day. Go back to that Day's checklist and pick up where you left off.

---

**You've got this. One day at a time, one step at a time.**

*Created: 2026-03-18 | Part of the iPipeline P&L Demo Project*
