# Project Todo

## Session — 2026-02-28

### Completed This Session (2026-02-28)
- [x] FIXED SQL revenue share mismatch: updated transformations.sql iGO=55%/Affirm=28%/InsureSight=12%/DocFast=5% to match Python pnl_config.py (synchronized with comment timestamp)
- [x] BUILT pnl_monte_carlo.py — full Monte Carlo P&L risk simulation engine: 10,000+ iterations, Dirichlet-distributed revenue shares, Normal-distributed expense amounts, P5/P25/P50/P75/P95 output, Value at Risk, 4-panel chart, Excel export, optional shock event modeling
- [x] WIRED monte-carlo into pnl_cli.py — full subcommand with --iterations, --seed, --concentration, --shock-prob, --shock-size, --export, --chart-path, --no-chart flags
- [x] REWROTE README.md — professional project overview: all 14 VBA modules, 4 SQL scripts, 12 Python scripts including Monte Carlo, CLI usage examples, repo structure, v2.1.0 status table
- [x] UPDATED docs/overview/CODE_COMPARISON_REPORT.md — all Section 2 items corrected; 10 new "ALREADY BUILT" entries for modUtilities macros; Section 3 gap list rewritten (quick-wins removed, user decisions logged); Section 4 priority order updated to reflect current state; Section 5 scorecard updated (53% built, up from 33%)
- [x] DROPPED by user: Timestamp Audit Trail VBA (user declined 2026-02-28); SQL layer audit trail remains
- [x] DROPPED by user: Export All Charts to PowerPoint (user dropped permanently)
- [x] DROPPED by user: Backup Workbook with Timestamp (user declined, logged in lessons.md)

### Still To Do — Next Priority
- [ ] Update CLAUDE.md with correct folder descriptions (Repo Structure section has Python listed as "12 scripts" now — was already updated in README.md)
- [ ] Fully read and audit remaining Python scripts not yet reviewed (pnl_allocation_simulator, pnl_forecast, pnl_month_end, pnl_ap_matcher, pnl_dashboard, pnl_email_report, pnl_tests, pnl_snapshot)
- [ ] Audit DemofileChartBuild/ folder contents
- [ ] Audit qa/ folder contents
- [ ] Audit docs/day-to-day/ and docs/setup/ contents

## Backlog
- [ ] Upload existing files from local APCLDmerge_ALL folder into GitHub
- [ ] Connect repo to Claude Code on the web
- [ ] Set up claude-master-config repo
- [ ] Populate /training/ folder with coworker training materials
- [ ] Build: Dynamic Progress Bar KPI Shape (next highest-value VBA feature)
- [ ] Build: Financial Statement Generator from Trial Balance (requires account mapping design first)
- [ ] Build: VBA Outlook Email Integration (completes PDF → Email in one workflow)

## Completed (All Sessions)
- [x] Set up GitHub repo and folder structure
- [x] Created CLAUDE.md
- [x] Created tasks folder with todo.md and lessons.md
- [x] Created .gitignore at root (commit c31d0bb)
- [x] Created CompletePackageStorage/production/ and CompletePackageStorage/backups/ (commit c31d0bb)
- [x] Repo structure audit (2026-02-26)
- [x] Full comprehensive audit of all code, docs, and NewTesting files (2026-02-27)
- [x] Fixed SQL bug: fact_gl_transactions → fact_gl in pnl_enhancements.sql
- [x] Built modLogger_v2.1.bas — VBA runtime audit log to hidden VBA_AuditLog sheet
- [x] Updated frmCommandCenter_code.txt — 62 actions, Sheet Tools category added
- [x] Fixed revenue share mismatch: SQL synced to Python values (2026-02-28)
- [x] Built pnl_monte_carlo.py — Monte Carlo P&L risk simulation (2026-02-28)
- [x] Wired monte-carlo into pnl_cli.py (2026-02-28)
- [x] Rewrote README.md professionally (2026-02-28)
- [x] Updated CODE_COMPARISON_REPORT.md — scorecard now 53% built (2026-02-28)
