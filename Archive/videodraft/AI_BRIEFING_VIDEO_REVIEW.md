# AI Briefing Document — Video Demo Review Request

**Purpose:** This document provides everything another AI system needs to give informed second opinions on planning and producing a demo video for a corporate Excel automation toolkit.

**What I need from you:** Ideas, feedback, and recommendations on how to produce the best possible demo video. Challenge my plan, suggest things I haven't thought of, and tell me what would make this video genuinely impressive at a Fortune 100 company.

---

## Project Summary

### What Was Built
A single Excel file (`.xlsm`) containing 62 automated VBA macros, organized into a "Command Center" — a searchable, categorized control panel. The file automates the entire monthly P&L (Profit & Loss) close process for the company's Finance & Accounting team.

### Key Stats
- **62 automated actions** accessible from one Command Center
- **34 VBA modules** (~12,700 lines of VBA code)
- **14 Python scripts** (data quality, forecasting, Monte Carlo simulation)
- **~99 universal tools** (additional VBA/Python tools that work on any Excel file)
- **Zero cost** — built entirely in-house using free tools
- **Zero installation** — just open the Excel file and go
- **6 training guides** written for non-technical Finance staff

### Who Built It
One Finance & Accounting team member (Connor) built this with AI assistance (Claude). Connor is not a developer — he works on guides, training docs, VBA, SQL, and Python demos.

### The Audience
- **Primary:** 2,000+ employees
- **Key stakeholders:** CFO and CEO
- **Target users:** Finance & Accounting staff (non-technical)

---

## What the Toolkit Does — Feature Inventory

### Monthly Operations
| Action | What It Does | Time Savings |
|--------|-------------|--------------|
| Generate Monthly Tabs | Creates individual monthly P&L sheets for each month | Manual: 30 min → Auto: 10 sec |
| Run Reconciliation Checks | Validates all cross-sheet totals match, shows PASS/FAIL scorecard | Manual: 2 hours → Auto: 10 sec |
| Import GL Data | Imports GL data from CSV/Excel with format validation | Manual: 45 min → Auto: 30 sec |

### Analysis & Reporting
| Action | What It Does | Time Savings |
|--------|-------------|--------------|
| Data Quality Scan | 6-category scan with A-F letter grade | Manual: N/A (never done) → Auto: 15 sec |
| Variance Analysis | Month-over-month variance, flags items over 15% | Manual: 3 hours → Auto: 15 sec |
| Variance Commentary | Auto-generates English narratives for top 5 variances | Manual: 1 hour → Auto: 5 sec |
| YoY Variance | Year-over-Year comparison with FY Total vs Prior Year/Budget | Manual: 2 hours → Auto: 10 sec |
| Sensitivity Analysis | What-if analysis on key assumptions | Manual: 4+ hours → Auto: 20 sec |
| Build Dashboard | Creates branded charts (revenue trends, margins, product mix) | Manual: 2 hours → Auto: 15 sec |
| Executive Dashboard | KPI cards + waterfall + product comparison on one sheet | Manual: 3 hours → Auto: 20 sec |
| PDF Export | 7-sheet professional PDF with headers/footers | Manual: 30 min → Auto: 10 sec |

### Enterprise Features
| Action | What It Does |
|--------|-------------|
| Version Control | Save/compare/restore workbook snapshots |
| Scenario Management | Save/load/compare named assumption sets (Base Case, Optimistic, etc.) |
| Executive Mode | Toggle to hide all technical sheets — clean view for leadership |
| Integration Test | 18 automated tests covering sheet existence, data integrity, formula health |
| Audit Log | Every action logged to a hidden sheet with timestamp, module, and result |
| Keyboard Shortcuts | Ctrl+Shift+M (Command Center), Ctrl+Shift+H (Home), Ctrl+Shift+R (Exec Mode) |

### Bonus / Supporting
| Feature | Description |
|---------|------------|
| Data Quality Letter Grade | A-F grade displayed in 28pt colored badge at top of quality report |
| Forecast Accuracy (MAPE) | Leave-one-out backtest with accuracy metrics |
| Cost-Line Reversal | Automatically reverses favorable/unfavorable for expense items |
| Disclaimer Sheet | Professional sheet stating all financial data is fictional |
| 12 Sheet Tools | AutoFit, Sort, Protect, Find/Replace across sheets — work on any file |

---

## The Company

- Insurance technology (insurtech) company
- Products: iGO, Affirm, InsureSight, DocFast (life insurance quoting, e-applications, underwriting, document management)
- Large enough to have 2,000+ employees
- Professional, corporate culture
- Brand colors: Brand Blue (#0B4779), Navy (#112E51), Lime Green (#BFF18C), Aqua (#2BCCD3)
- Fonts: Arial family only

---

## The Demo Workbook

### Sheet Structure (13+ sheets)
1. **Report-->** — Main landing page with summary and navigation
2. **P&L - Monthly Trend** — Revenue and expenses by month (Jan-Dec), FY Total, Budget columns
3. **Functional P&L Summary - Jan/Feb/Mar** — Monthly snapshots by function/department
4. **Product Line Summary** — Revenue by product (iGO, Affirm, InsureSight, DocFast)
5. **Assumptions** — Key financial drivers (revenue shares, allocation percentages, growth rates)
6. **General Ledger** — Raw GL transaction data (fictional)
7. **Checks** — Reconciliation results (PASS/FAIL for each check)
8. **Charts & Visuals** — 8 interactive charts in a grid layout
9. Additional sheets created by macros: Variance Analysis, Sensitivity Analysis, Executive Dashboard, Variance Commentary, Data Quality Report, YoY Variance Analysis, Exec Summary - Print, Disclaimer

### What's Fictional
ALL financial data is completely made up. Product names are real company products but all numbers, vendors, GL accounts, and transactions are fabricated for demo purposes.

### What's Real
The macros, automation logic, reporting templates, and formatting are all production-ready and can be applied to actual financial data.

---

## Current Video Plan

### Format
Screen recording with voice-over narration (Connor's voice). No slides — all live Excel.

### Length Options Under Consideration
- **Option A:** Three 6-7 minute videos (Introduction, Power Features, Advanced)
- **Option B:** One 18-22 minute video
- **Option C (leaning toward):** One continuous 18-20 min video with chapter markers

### Planned Demo Flow (18-20 minutes)
1. Hook / opening stat (30 sec)
2. Workbook tour — scroll through sheets (60 sec)
3. Command Center — open, browse, search (90 sec)
4. Data Quality Scan + Letter Grade (90 sec)
5. Reconciliation Checks — PASS/FAIL (60 sec)
6. Variance Analysis — flagged items (90 sec)
7. Variance Commentary — auto-generated narratives (90 sec)
8. Dashboard Charts — branded visuals (60 sec)
9. PDF Export — deliverable output (45 sec)
10. Executive Mode toggle (30 sec)
11. Version Control — save snapshot (60 sec)
12. Sensitivity Analysis — what-if (60 sec)
13. Integration Test — 18/18 PASS (45 sec)
14. Closing — recap, where to get the file, call to action (60 sec)

### Tone
Professional but approachable. Like a trusted colleague showing something genuinely useful. Confident, practical, focused on "here's what this does for YOU."

### Key Messages
1. Single Excel file — nothing to install
2. 62 automated actions — one click each
3. What used to take hours now takes seconds
4. Zero cost, zero IT involvement
5. All data is demo — the tools are real and production-ready

---

## What I Want Your Opinion On

### Video Strategy
1. Is 18-20 minutes too long? Should we aim shorter? What's the ideal length for this type of corporate demo?
2. Three videos vs one with chapters — which is better for our audience (2,000+ non-technical finance people)?
3. Should we also produce a 3-5 minute "highlight reel" cut for leadership? Or is the full video fine for everyone?
4. What's the best platform for hosting internally? SharePoint video? Microsoft Stream? Something else?

### Content & Flow
5. Is the demo flow above in the right order? What would you rearrange?
6. What are we missing? What features should we highlight that we're not?
7. Should we show the Disclaimer sheet in the video? It's professional but takes time.
8. Is 14 out of 62 actions the right number to demo? Too many? Too few?
9. Should we mention the Universal Toolkit (99 tools for any file) or keep focus on just the demo file?

### Presentation Style
10. Webcam overlay (face in corner) — yes or no for this type of video?
11. Should we add title cards / chapter cards between sections?
12. Timer on screen while actions run — good idea or gimmicky?
13. Background music — yes or no? If yes, only intro/outro or throughout?
14. Should Connor read from a script or use bullet points and talk naturally?

### "Before vs After" Concept
15. Should the video include a "before" segment showing the old manual process? Or jump straight into the new automated way?
16. If we do "before vs after", how? Side-by-side screenshots? A quick story? Numbers on screen?

### Production Quality
17. What recording software would you recommend? (Options: OBS Studio (free), Camtasia (paid), PowerPoint screen recording, Loom)
18. Any tips for making a screen recording look professional without a video production team?
19. What resolution and format? We're planning 1920x1080, 30fps, MP4.
20. Audio — USB headset mic or should we rent/buy a proper desk mic?

### Distribution & Follow-Up
21. What should the email/Teams announcement say when the video is released?
22. Should we do a live walkthrough session in addition to the video?
23. Should we create a feedback form so viewers can ask questions or request features?
24. Any other distribution or engagement ideas?

### Things I Might Be Missing
25. What mistakes do people commonly make with corporate demo videos?
26. What would make this video go from "good" to "the best internal demo video at this company"?
27. Is there anything about this project that should be presented differently than how we're planning?
28. If you were the CFO watching this video, what would impress you most? What would bore you?

---

## Existing Draft Materials

We already have a detailed 550-line video script and storyboard at `FinalRoughGuides/05-Video-Demo-Script-and-Storyboard.md`. It includes:
- Pre-recording checklist (workbook, computer, audio, script prep)
- Equipment and software recommendations
- Screen layout and Excel settings
- 3-part shot lists with timing breakdowns
- Full word-for-word speaker script for all 3 parts
- B-roll and transition ideas
- Common FAQ to address in the video
- Post-recording and distribution checklists

We also have a brainstorming document at `videodraft/VIDEO_DEMO_PLAN.md` with:
- Format pros/cons analysis (3 videos vs 1 vs hybrid)
- Tiered feature ranking (must show, should show, nice to show)
- Ideas to make the video stand out (before/after, timer, hook strategies)
- Three script style options (casual, formal, problem-solution)
- Practical recording tips (what kills demos, what makes them great)
- Flow recommendations and the "10-second rule"
- Multiple version options (short cut for leadership, full demo, hands-on tutorial)

---

## Technical Details (For Context Only)

### VBA Module List (34 modules)
modConfig, modFormBuilder, modMasterMenu, modNavigation, modDashboard, modDashboardAdvanced, modDataQuality, modReconciliation, modVarianceAnalysis, modPDFExport, modPerformance, modMonthlyTabGenerator, modSearch, modUtilities, modLogger, modSensitivity, modAWSRecompute, modImport, modForecast, modScenario, modAllocation, modConsolidation, modVersionControl, modAdmin, modIntegrationTest, modDemoTools, modDataGuards, modDrillDown, modAuditTools, modETLBridge, modTrendReports, modDataSanitizer, modSheetIndex

### Python Script List (14 scripts)
pnl_config.py, pnl_etl.py, pnl_report.py, pnl_forecast.py, pnl_allocation.py, pnl_what_if.py, pnl_data_quality.py, pnl_monte_carlo.py, pnl_cli.py, pnl_dashboard.py, pnl_reconciliation.py, pnl_variance.py, pnl_consolidation.py, pnl_budget_loader.py

### Test Status
- T1 (Compilation): 8/8 PASS
- T2 (Foundation): 4/7 PASS (3 not yet run)
- T4 (Python): pytest 99 passed, 15 skipped, 0 failures
- T5 (Advanced VBA): 2/6 PASS (4 not yet run)
- Remaining: T3, T6, T7, T8 not yet run

---

## Summary

We have a world-class Excel automation toolkit ready for demo. The code is solid (34 VBA modules, 14 Python scripts, 99 universal tools). We have 6 training guides written. We have a detailed script and storyboard. Now we need to produce the actual video.

**The question is: How do we make this video as impactful and professional as possible, given that we are one person with no video production experience or budget?**

Give us your best thinking. Challenge our plan. Tell us what we're missing. Help us make this the best internal demo video the company has ever seen.

---

*Document created: 2026-03-05 | Branch: claude/resume-ipipeline-demo-qKRHn*
