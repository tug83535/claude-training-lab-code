# Lessons Learned - APCLDmerge Project

## Guides & Documentation
- Always number every single step no matter how small or obvious
- Never summarize — full detail on every step
- Add "what you should see" confirmations after key steps
- World class quality only — CFO/CEO and 2,000+ employees will see this

## Excel & VBA
- Always review EVERY sheet before starting any work
- Confirm all sheet names found before proceeding
- Never assume which sheet is most important
- Always confirm plan before touching the Excel file

## General Workflow
- Break ALL large tasks into a numbered action plan first
- Present plan and wait for approval before executing
- Execute one step at a time
- Stop immediately if something goes wrong — re-plan and check in
- Update todo.md when asked or at end of session
- Never infer — always ask if unclear

## Communication
- Plain English only — user is not a developer
- Always confirm what you will do in bullet points before doing it
- Be proactive with new ideas and recommendations

## Session Handoff & Account Switching
- When switching Claude accounts due to usage limits, the new account starts with zero context
- Always update CLAUDE.md with a "Current Session Summary" section at end of session — this is the single most important handoff document
- The session summary in CLAUDE.md should include: branch name, what was done, key findings, exact bugs found, and prioritized next steps
- Phrase next steps clearly enough that a brand new Claude account can pick up without any explanation from the user

## Keeping Records Accurate
- tasks/todo.md can become stale — always check what's actually in the repo (git log, file listing) before reporting what's done vs. not done
- Never trust todo.md item status without verifying against the actual repo — items can be completed in one session and never marked done
- docs/overview/CODE_COMPARISON_REPORT.md (and similar analysis files) can go out of date when new commits are made — always check git log for the most recent commits before relying on any comparison report
- Before saying something is "missing" or "not built," check the most recently committed files first

## Reading Binary Files
- .xlsx files are binary and cannot be read with any text tool — use ARCHITECTURE_DIAGRAM.md and modConfig_v2.1.bas constants as the source of truth for Excel sheet structure and naming
- Always note in session summary that the Excel file was not directly readable and which reference files were used instead

## Cross-File Consistency Checks
- Always compare config values across layers: VBA constants, Python config, and SQL tables can all define the same values (fiscal year, revenue shares, sheet names) — mismatches cause silent errors
- Revenue shares defined in Python pnl_config.py must match the allocation_shares table in SQL transformations.sql — found mismatch this session (iGO=55% Python vs 50% SQL)
- SQL scripts that reference tables must use the exact table name from the script that creates them — found `fact_gl_transactions` vs `fact_gl` mismatch in pnl_enhancements.sql

## User Feature Preferences
- User confirmed they do NOT want a "Backup Workbook with Timestamp" macro — do not propose or rebuild this
- User confirmed they do NOT want a VBA Timestamp Audit Trail (Worksheet_Change event) — SQL layer audit trail in pnl_enhancements.sql is sufficient; do not propose rebuilding in VBA
- User confirmed they do NOT want "Export All Charts to PowerPoint" — dropped permanently; do not re-propose
- Always log user feature rejections here so future Claude accounts don't re-propose them
- When proposing new features, review this section first to avoid suggesting features the user has already declined

## Context Recovery After Session Limits
- When resuming from a compressed context, always read the current file state FIRST before making any edits — do not trust the summary's description of what was already edited
- Prior session summaries may describe edits as "complete" when the session was cut off mid-task — verify with a fresh file read
- Read both the source file AND the task summary; if they conflict, the file is the truth
- For large multi-section documents, read the full file before editing any section — avoids duplicate edits or missing context

## Verifying Claims Before Reporting Gaps
- Always check git log and actual committed files before saying something is "missing" or "not built"
- In this project: modLogger was reported missing but already existed; pnl_enhancements.sql bug was reported but already fixed; frmCommandCenter_code.txt was reported outdated but already updated — all stale CLAUDE.md entries
- The session summary in CLAUDE.md can itself become stale — treat it as a starting point, not gospel

## Monte Carlo / Statistical Modeling
- For allocation share randomization, use Dirichlet distribution (not uniform random) — Dirichlet guarantees shares always sum to exactly 1.0 and respects the shape/concentration of the original distribution
- Dirichlet alpha vector = base_shares * concentration_parameter (higher concentration = tighter clustering around the mean)
- For expense randomization, Normal distribution is appropriate (truncated at zero for amounts that can't go negative)
- Always expose key simulation parameters (iterations, seed, concentration, shock_prob) as CLI arguments for reproducibility and testing
