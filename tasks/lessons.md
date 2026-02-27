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
