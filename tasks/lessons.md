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
- ALWAYS validate the Excel file opens without errors after any script modifies it
  - Re-load with openpyxl and re-save to a temp file to confirm no corruption
  - Check all sheets still present, correct row/column counts, no broken merges
  - Verify data validations, charts, formulas all survived the save
  - Unmerge cells BEFORE clearing content (MergedCell objects are read-only)
  - Test that the file can be opened in Excel without repair prompts

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
