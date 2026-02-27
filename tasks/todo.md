# Project Todo

## Current Session — 2026-02-27
### Completed This Session (2026-02-27 continuation)
- [x] Full comprehensive audit of every file, folder, VBA module, SQL script, Python script, and NewTesting ideas
- [x] Delivered full audit report to user (read-only — no changes made)
- [x] Corrected stale items in prior todo.md (gitignore and folders already existed)
- [x] Removed "Backup Workbook" from all plans — user confirmed NOT wanted
- [x] Explained GitHub Pull Request / merge process to user (new to GitHub)
- [x] FIXED SQL bug: renamed fact_gl_transactions → fact_gl in pnl_enhancements.sql (all occurrences)
- [x] BUILT modLogger_v2.1.bas from scratch — logs to VBA_AuditLog hidden sheet; has LogAction, ClearLog, ExportLog, GetLogSheet
- [x] UPDATED frmCommandCenter_code.txt — now shows all 62 actions; added "Sheet Tools" category with actions 51-62 from modUtilities

### Still To Do — Next Priority
- [ ] DECISION NEEDED: Reconcile revenue share percentages — Python pnl_config.py has iGO=55%/Affirm=28%/InsureSight=12%/DocFast=5% but SQL transformations.sql has iGO=50%/Affirm=25%/InsureSight=15%/DocFast=10% — ask user which set is correct, then update both files to match
- [ ] Update docs/overview/CODE_COMPARISON_REPORT.md — outdated; 12 "Not Yet Built" items are now actually built (modUtilities); also modLogger is now built
- [ ] Rewrite README.md with professional project overview
- [ ] Update CLAUDE.md with correct folder descriptions (Repo Structure section is slightly outdated)
- [ ] Build: Timestamp Audit Trail on Cell Changes (Worksheet_Change event) — highest priority remaining VBA feature
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
