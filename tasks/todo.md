# Project Todo

## Current Session — 2026-02-27
### Completed This Session
- [x] Full comprehensive audit of every file, folder, VBA module, SQL script, Python script, and NewTesting ideas
- [x] Delivered full audit report to user (read-only — no changes made)
- [x] Corrected stale items in prior todo.md (gitignore and folders already existed)

### Still To Do — Cleanup (carry from 2026-02-26)
- [ ] Update CLAUDE.md with correct folder descriptions (Repo Structure section is slightly outdated)
- [ ] Rewrite README.md with professional project overview
- [ ] Commit and push cleanup changes to branch `claude/review-project-status-ntucB`

### Newly Identified — Bugs & Gaps (found this session)
- [ ] Fix SQL table name mismatch: pnl_enhancements.sql references `fact_gl_transactions` but staging.sql creates `fact_gl` — will cause errors if run
- [ ] Reconcile revenue share percentages: Python pnl_config.py has iGO=55%/Affirm=28%/InsureSight=12%/DocFast=5% but SQL transformations.sql has iGO=50%/Affirm=25%/InsureSight=15%/DocFast=10% — decide correct values and update both files
- [ ] Update frmCommandCenter_code.txt to show all 62 actions (currently only shows 50 — modUtilities v2.1 added actions 51-62)
- [ ] Commit missing VBA .bas files for modules referenced in ExecuteAction() but not in repo: modLogger, modImport, modForecast, modScenario, modAllocation, modConsolidation, modVersionControl, modAdmin, modIntegrationTest, modAWSRecompute, modSensitivity (and others — see CLAUDE.md session summary)
- [ ] Update docs/overview/CODE_COMPARISON_REPORT.md — it is outdated, was written before modUtilities_v2.1 was committed; 12 of the 25 "Not Yet Built" items are now actually built
- [ ] Build: Timestamp Audit Trail on Cell Changes (Worksheet_Change event) — highest priority remaining VBA gap
- [ ] Build: Backup Workbook with Timestamp macro
- [ ] Build: Export All Charts to PowerPoint macro
- [ ] Fully read and audit remaining 9 Python scripts (pnl_allocation_simulator, pnl_forecast, pnl_month_end, pnl_ap_matcher, pnl_dashboard, pnl_email_report, pnl_cli, pnl_tests, pnl_snapshot)
- [ ] Audit DemofileChartBuild/ folder contents
- [ ] Audit qa/ folder contents
- [ ] Audit docs/day-to-day/ and docs/setup/ contents

## Backlog
- [ ] Upload existing files from local APCLDmerge_ALL folder into GitHub
- [ ] Connect repo to Claude Code on the web
- [ ] Set up claude-master-config repo
- [ ] Populate /training/ folder with coworker training materials

## Completed (All Sessions)
- [x] Set up GitHub repo and folder structure
- [x] Created CLAUDE.md
- [x] Created tasks folder with todo.md and lessons.md
- [x] Created .gitignore at root (completed in prior session, commit c31d0bb)
- [x] Created CompletePackageStorage/production/ and CompletePackageStorage/backups/ subfolders (completed in prior session, commit c31d0bb)
- [x] Repo structure audit (2026-02-26)
- [x] Full comprehensive audit of all code, docs, and NewTesting files (2026-02-27)
