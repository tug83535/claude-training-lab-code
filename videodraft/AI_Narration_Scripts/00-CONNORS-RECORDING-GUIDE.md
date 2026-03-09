# Connor's Recording Guide — AI Narration + Screen Recording

**What This Is:** Your personal step-by-step guide for recording all 3 demo videos using AI-generated voice narration (ElevenLabs) synced with screen recording.

**How It Works:** You generate audio clips first, then record your screen while listening to the audio in headphones and clicking along. No talking required from you during recording.

---

## Table of Contents

1. [The Big Picture — How This All Works](#1-the-big-picture)
2. [What You Need Before Starting](#2-what-you-need)
3. [Step-by-Step: Setting Up ElevenLabs](#3-elevenlabs-setup)
4. [Step-by-Step: Generating Audio Clips](#4-generating-audio-clips)
5. [Step-by-Step: Recording Your Screen](#5-recording-your-screen)
6. [Step-by-Step: Syncing Audio + Video](#6-syncing-audio-and-video)
7. [ElevenLabs Tips for Natural Sound](#7-elevenlabs-tips)
8. [Troubleshooting](#8-troubleshooting)

---

## 1. The Big Picture

Here is the full workflow, start to finish:

```
STEP 1: Clone your voice in ElevenLabs (one-time, 1 minute of recording)
    ↓
STEP 2: Generate audio clips from the scripts (one clip per segment)
    ↓
STEP 3: Listen to each clip — re-generate any that sound off
    ↓
STEP 4: Set up your screen for recording (Excel open, clean desktop)
    ↓
STEP 5: Play each audio clip in headphones while screen recording
    ↓
STEP 6: Click through the demo in sync with what you hear
    ↓
STEP 7: Drop audio + screen recording into a video editor
    ↓
STEP 8: Adjust timing, add title cards, export as MP4
```

**Why this order matters:** You generate audio FIRST because the audio controls the pacing. You then match your clicking to the audio — not the other way around. This is much easier than trying to talk and click at the same time.

---

## 2. What You Need

### Software (All Free Options Available)

| What | Free Option | Paid Option | Notes |
|------|-------------|-------------|-------|
| **AI Voice** | ElevenLabs free tier (~10 min/month) | ElevenLabs Starter $5/mo (~30 min) | You need ~25-30 min total audio across all 3 videos. The $5 tier covers it. |
| **Screen Recorder** | OBS Studio (obsproject.com) | Camtasia ($300) | OBS is free and records at 1080p. Camtasia has better editing. |
| **Video Editor** | DaVinci Resolve (free) or CapCut (free) | Camtasia (if you bought it) | You need a timeline editor to sync audio + video. |
| **Audio Player** | Windows Media Player or VLC | — | To play audio clips in headphones during recording |

### Hardware

| What | Why |
|------|-----|
| **Headphones or earbuds** | You listen to the AI narration while recording. Must not be audible to the screen recorder. Use wired headphones (not speakers). |
| **Second monitor (optional but helpful)** | Script on one screen, Excel on the other |

### Files

| What | Where |
|------|-------|
| Video 1 script | `videodraft/AI_Narration_Scripts/01-VIDEO1-WHATS-POSSIBLE.md` |
| Video 2 script | `videodraft/AI_Narration_Scripts/02-VIDEO2-FULL-DEMO-WALKTHROUGH.md` |
| Video 3 script | `videodraft/AI_Narration_Scripts/03-VIDEO3-UNIVERSAL-TOOLS.md` |
| Demo Excel file | Your local `.xlsm` with all VBA modules imported |

---

## 3. ElevenLabs Setup

### Step 1 — Create an Account

- [ ] 1. Go to **elevenlabs.io**
- [ ] 2. Click **Sign Up** (top right)
- [ ] 3. Create an account with your email (or sign in with Google)
- [ ] 4. You start on the free tier — 10 minutes of audio per month

### Step 2 — Clone Your Voice (Optional but Recommended)

If you want the narration to sound like you:

- [ ] 1. Click **Voices** in the left sidebar
- [ ] 2. Click **Add Voice** → **Instant Voice Cloning**
- [ ] 3. You need to upload a clean audio recording of yourself talking naturally for about 1 minute
   - Record yourself on your phone or computer reading a paragraph from one of the scripts
   - Speak at your normal pace and tone — don't try to sound "professional"
   - Make sure there's no background noise
- [ ] 4. Upload the recording
- [ ] 5. Name the voice (e.g., "Connor - Demo")
- [ ] 6. Check the box confirming you have rights to the voice
- [ ] 7. Click **Add Voice**
- [ ] 8. Your cloned voice now appears in your voice library

**If you skip cloning:** Use one of ElevenLabs' built-in voices instead. Go to the **Voice Library** and preview voices until you find one you like. Look for:
- Male, American English
- Professional but not overly formal
- Clear and natural-sounding (not robotic, not dramatic)
- Good options to try: "Adam", "Josh", "Marcus"

### Step 3 — Choose Your Plan

- The **free tier** gives you ~10,000 characters/month (~10 minutes of audio)
- Video 1 alone is ~4-5 minutes of audio, so you'll need the **Starter plan ($5/month)** which gives ~30 minutes
- You can cancel after one month — just need it for this project
- [ ] Upgrade if needed: Click your profile → Subscription → Starter

---

## 4. Generating Audio Clips

### How the Scripts Are Organized

Each video script is broken into numbered **segments**. Each segment is one chunk of narration that goes with one set of screen actions.

Example:
```
SEGMENT 1.1 — Opening Hook
[AUDIO TEXT]:
"This is a single Excel file..."

[YOUR SCREEN ACTIONS]:
- File is open on Report page
- Slowly scroll down
```

You generate one audio clip per segment. This gives you:
- Individual files you can re-generate if one sounds wrong
- Easy alignment with screen actions
- Flexibility to adjust timing in the editor

### Step-by-Step: Generating One Clip

- [ ] 1. Go to **elevenlabs.io** → click **Text to Speech** in the left sidebar
- [ ] 2. Select your voice (your cloned voice or a built-in one) from the dropdown at the top
- [ ] 3. Copy the **[AUDIO TEXT]** section from one segment in the script
- [ ] 4. Paste it into the text box
- [ ] 5. Set these settings:
   - **Model:** Eleven Multilingual v2 (or Eleven Turbo v2 for faster generation)
   - **Stability:** 50-60% (lower = more expressive, higher = more consistent)
   - **Similarity:** 75-85% (how closely it matches the voice sample)
   - **Style:** 15-25% (adds natural variation — don't go too high or it gets dramatic)
- [ ] 6. Click **Generate**
- [ ] 7. Listen to the clip with headphones
- [ ] 8. If it sounds good → click the **download arrow** → save as MP3
- [ ] 9. Name the file to match the segment: `V1_S1_Opening_Hook.mp3`
- [ ] 10. If it sounds off → adjust settings slightly and re-generate (see Tips section below)

### Naming Convention for Files

Use this pattern so files sort correctly:

```
V1_S1_Opening_Hook.mp3
V1_S2_Command_Center.mp3
V1_S3_Data_Quality.mp3
...
V2_S1_Opening.mp3
V2_S2_Chapter1_Workbook.mp3
...
```

### Do All Clips for One Video Before Moving On

Generate ALL clips for Video 1 first. Listen to them in order. Make sure the voice, pacing, and energy feel consistent across all segments. Then move to Video 2, then Video 3.

---

## 5. Recording Your Screen

### Before You Record — Computer Setup

Do this EVERY time before recording, even if you think it's already done:

- [ ] 1. Close ALL applications except Excel and your screen recorder (OBS or other)
- [ ] 2. Close all browser tabs
- [ ] 3. Turn off notifications:
   - **Teams:** Profile picture → "Do Not Disturb"
   - **Outlook:** Close it completely
   - **Windows:** Settings → System → Notifications → turn off
   - **Phone:** Silent or airplane mode
- [ ] 4. Clean desktop: Right-click desktop → View → uncheck "Show desktop icons"
- [ ] 5. Auto-hide taskbar: Right-click taskbar → Taskbar settings → "Automatically hide"
- [ ] 6. Set display scaling to **100%**: Settings → Display → Scale → 100%
- [ ] 7. Set resolution to **1920 x 1080**: Settings → Display → Resolution
- [ ] 8. Plug in your laptop (no battery throttling)

### Before You Record — Excel Setup

- [ ] 1. Open the demo `.xlsm` file
- [ ] 2. Maximize Excel to fill the entire screen
- [ ] 3. Set zoom to **100%** (View → Zoom → 100%)
- [ ] 4. Make sure the ribbon is visible
- [ ] 5. Navigate to the **Report-->** landing page
- [ ] 6. Close the Command Center if it's open
- [ ] 7. Delete any leftover output sheets from previous runs:
   - Variance Analysis, Variance Commentary, Data Quality Report, Executive Dashboard, YoY Variance Analysis, Sensitivity Analysis
- [ ] 8. Save the file in this clean state

### Before You Record — Screen Recorder Setup (OBS)

If using OBS Studio:

- [ ] 1. Open OBS
- [ ] 2. Under **Sources**, click **+** → **Display Capture** → OK
- [ ] 3. In **Settings** → **Output**:
   - Recording Path: Choose a folder you'll remember
   - Recording Format: **mp4**
   - Encoder: Use default (x264)
- [ ] 4. In **Settings** → **Video**:
   - Base Resolution: 1920x1080
   - Output Resolution: 1920x1080
   - FPS: 30
- [ ] 5. In **Settings** → **Audio**:
   - Set **Desktop Audio** to **Disabled** (you don't want OBS recording the AI audio from your headphones)
   - Set **Mic/Auxiliary Audio** to **Disabled** (no mic needed — AI does the talking)
- [ ] 6. Click **OK** to save settings

**IMPORTANT — Audio Settings:** You do NOT want OBS to record any audio. The final video will have the AI audio clips added in the video editor. OBS should record ONLY the screen — silent video.

### The Recording Process

For each segment:

1. **Put on your headphones**
2. **Start OBS recording** (click "Start Recording" or press the hotkey)
3. **Play the audio clip** for the current segment on your computer (it plays through headphones only)
4. **Watch the screen and click along** as the narration progresses
   - When the voice says "I press Control Shift M" → you press Ctrl+Shift+M
   - When the voice says "Let me run a Data Quality scan" → you click Run on the action
   - When the voice pauses → you pause your clicking (let the viewer see the result)
5. **After the segment's audio ends, wait 2-3 seconds**, then stop recording
6. **Move to the next segment**

### Tips for Smooth Recording

- **Practice each segment once** before recording. Play the audio, practice the clicks without recording. Then do it for real.
- **It's OK if timing isn't perfect.** Small gaps (1-2 seconds) between voice and action are normal and can be adjusted in the editor.
- **If you mess up, just re-record that segment.** That's the whole point of doing segments — you don't have to redo the entire video.
- **Between segments, reset if needed.** If a macro created an output sheet and the next segment needs a clean state, delete the output sheet and save before recording the next segment.
- **Keep your mouse movements slow and deliberate.** Fast, jerky mouse movements look bad on screen recording.

---

## 6. Syncing Audio and Video

### Using DaVinci Resolve (Free)

- [ ] 1. Download and install DaVinci Resolve from blackmagicdesign.com (free version)
- [ ] 2. Create a new project
- [ ] 3. Import all your files:
   - All screen recording clips (from OBS)
   - All audio clips (from ElevenLabs)
   - Title card images (if you made them)
- [ ] 4. Drag everything to the timeline in order:
   - Title card image (5 seconds)
   - Segment 1 video + Segment 1 audio (aligned)
   - Segment 2 video + Segment 2 audio
   - etc.
- [ ] 5. For each segment:
   - Put the video clip on the video track
   - Put the matching audio clip on the audio track directly below it
   - Adjust the start point of either clip so the action matches the narration
   - If the voice says "I'm running the scan" and you clicked 1 second later, nudge the video clip 1 second earlier
- [ ] 6. Add transitions between segments if desired (simple crossfade, 0.5 seconds)
- [ ] 7. Add title cards between chapters (for Videos 2 and 3)
- [ ] 8. Preview the full video — check that audio and actions line up
- [ ] 9. Export: **Deliver** tab → MP4, H.264, 1920x1080, 30fps

### Using CapCut (Free, Simpler)

- [ ] 1. Download CapCut from capcut.com
- [ ] 2. Create a new project at 1920x1080
- [ ] 3. Import all screen recordings and audio clips
- [ ] 4. Drag clips to the timeline in order
- [ ] 5. Align audio with video for each segment (drag to adjust timing)
- [ ] 6. Add text overlays for time savings (e.g., "Manual: 2 hours → Automated: 10 seconds")
- [ ] 7. Export as MP4, 1080p

### If Timing Is Off

Don't worry about perfection. Here's how to fix common issues:

| Problem | Fix |
|---------|-----|
| Voice says something before you click | Split the video clip, add a small gap (freeze frame) to delay |
| You click before the voice says to | Nudge the audio clip earlier, or trim the start of the video clip |
| There's dead air while a macro runs | Add a freeze frame or slow-motion effect, or leave it — natural pauses are fine |
| One segment is way off | Re-record just that screen segment. The audio is already done. |

---

## 7. ElevenLabs Tips for Natural Sound

### Making It Sound Like a Real Person (Not a Robot)

1. **Use contractions.** The scripts already use "I'm", "you've", "that's", "it's", "don't", "won't", "here's" — these sound natural. If you see a spot that doesn't have a contraction, add one.

2. **Add breathing pauses with ellipses.** The scripts use `...` to create natural pauses. ElevenLabs reads ellipses as short pauses. Two or three dots = short pause. A period followed by a new line = longer pause.

3. **Use em dashes for mid-sentence pauses.** "This is a single Excel file — nothing to install" sounds more natural than a comma there. The scripts already do this.

4. **Don't use ALL CAPS for emphasis.** ElevenLabs may shout it. Instead, the scripts use natural phrasing where emphasis falls in the right place.

5. **Numbers and abbreviations:**
   - "62" will be read as "sixty-two" — that's fine
   - "P&L" may be read as "P and L" — test it. If it sounds weird, change to "P-and-L" or "P. and L."
   - "Ctrl+Shift+M" — change to "Control Shift M" (already done in scripts)
   - "1920x1080" — change to "nineteen twenty by ten eighty" (already done in scripts)
   - "CFO" will be read as "C-F-O" — that's fine
   - "VBA" will be read as "V-B-A" — that's fine
   - "PDF" will be read as "P-D-F" — that's fine
   - "CSV" may sound like "C-S-V" or "cee-ess-vee" — test it

6. **Pronunciation adjustments:**
   - "iGO" → test it. If it says "eye-go" that's correct. If it says "ih-go", change to "i-GO"
   - "Affirm" → should be fine
   - "InsureSight" → test it. May need "Insure-Sight" with a hyphen
   - "DocFast" → test it. May need "Doc-Fast"
   - ".xlsm" → change to "dot X-L-S-M" (already done in scripts)
   - "xlsx" → change to "dot X-L-S-X"

7. **Pacing controls:**
   - Short sentence = faster delivery
   - Longer sentence with commas = natural breathing points
   - `...` = 0.5-1 second pause
   - New paragraph = 1-2 second pause
   - The scripts are already written with these patterns

### ElevenLabs Settings Cheat Sheet

| Setting | What It Does | Recommended Value |
|---------|-------------|-------------------|
| **Stability** | Higher = more monotone, Lower = more expressive | 50-55% for narration |
| **Similarity** | How closely it matches the voice | 75-80% (higher can sound stiff) |
| **Style** | Adds natural expression | 15-20% (too high = overdramatic) |
| **Speaker Boost** | Clarity enhancement | ON |

### If a Clip Sounds Off

| Problem | Fix |
|---------|-----|
| Too monotone / robotic | Lower Stability to 40-45%, raise Style to 25% |
| Too dramatic / overdone | Raise Stability to 60%, lower Style to 10% |
| Weird pronunciation of a word | Spell it phonetically: "Reconciliation" → "Recon-sill-ee-ay-shun" (only if needed) |
| Unnatural pause in wrong spot | Rewrite the sentence to move the pause. Split into two shorter sentences. |
| Voice sounds different between clips | Keep ALL settings identical across clips. Don't change Stability/Similarity between segments. |

---

## 8. Troubleshooting

### Common Problems and Solutions

**"ElevenLabs says I'm out of characters"**
- Free tier = ~10,000 characters. The full Video 2 script alone is ~6,000 characters.
- Upgrade to Starter ($5/month) for ~100,000 characters. Cancel after the project.

**"The audio timing doesn't match my screen recording"**
- This is normal and expected. That's what the video editor is for.
- Nudge clips on the timeline until they line up. 1-2 seconds of adjustment is typical.

**"A macro errored during screen recording"**
- Stop recording. Fix the issue. Delete any output sheets. Save. Re-record that segment.
- This is why we record segment-by-segment, not all at once.

**"OBS recorded audio I didn't want"**
- Go to OBS Settings → Audio → set Desktop Audio and Mic to "Disabled"
- You want silent screen recording. Audio comes from the ElevenLabs clips.

**"The video editor is confusing"**
- CapCut is simpler than DaVinci Resolve. Start there.
- YouTube has great 10-minute tutorials for both. Search "CapCut beginner tutorial 2026"

**"My cloned voice sounds weird"**
- The source recording matters a lot. Record 1 minute of natural talking in a quiet room.
- Avoid reading — just talk naturally about something (describe your weekend, explain what you had for lunch)
- Re-upload the better recording

**"I don't want to clone my voice"**
- That's fine. Use a built-in ElevenLabs voice. Preview several and pick one that sounds natural and professional.
- Good starting points: "Adam", "Josh", "Marcus" (all male, American English)

---

## Quick Reference — Full Workflow Checklist

### Phase 1: Audio Generation
- [ ] Set up ElevenLabs account
- [ ] Clone voice (or pick a built-in voice)
- [ ] Generate all Video 1 clips → listen → re-generate any that sound off
- [ ] Generate all Video 2 clips → listen → re-generate any that sound off
- [ ] Generate all Video 3 clips → listen → re-generate any that sound off
- [ ] Organize all clips in folders: `Video1/`, `Video2/`, `Video3/`

### Phase 2: Screen Recording
- [ ] Set up computer (notifications off, desktop clean, resolution 1920x1080)
- [ ] Set up OBS (silent recording, 1080p, 30fps, mp4)
- [ ] Set up Excel (clean state, Report page, Command Center closed)
- [ ] Record all Video 1 segments (play audio in headphones, click along)
- [ ] Record all Video 2 segments
- [ ] Record all Video 3 segments

### Phase 3: Editing
- [ ] Open DaVinci Resolve or CapCut
- [ ] Import all audio clips and screen recordings
- [ ] Build Video 1: arrange clips on timeline, sync audio to video, add title cards
- [ ] Build Video 2: same process, add chapter cards between sections
- [ ] Build Video 3: same process, add chapter cards
- [ ] Add time savings overlays where indicated in scripts
- [ ] Preview each video start to finish
- [ ] Export all 3 as MP4, 1920x1080, 30fps

### Phase 4: Delivery
- [ ] Save videos to your output folder
- [ ] Upload to SharePoint
- [ ] Test that videos play correctly from SharePoint

---

*Created: 2026-03-09 | Part of AI Narration Scripts package*
