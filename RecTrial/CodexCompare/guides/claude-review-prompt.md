# Claude Code Comparison Prompt (Copy/Paste)

Use this exact prompt in Claude Code:

---
You are reviewing a Finance automation repository and comparing it against another version built in Claude Code.

## Objective
Perform a deep comparative analysis between:
1) this repository state, and
2) my separate Claude-built project.

I need a practical reviewer report I can act on.

## Required outputs
1. **Executive summary** (plain English).
2. **Feature parity matrix** (Universal VBA, Demo VBA, Python utilities, SQL templates, docs, tests, CI).
3. **Code quality review** (error handling, modularity, assumptions, maintainability).
4. **Validation review** (smoke checks, unit tests, gaps, false confidence risks).
5. **Operations review** (onboarding flow, branch/push PR usability, reproducibility).
6. **Top 10 differences** ranked by impact.
7. **Recommendation**: which implementation is stronger now, and what to merge from each.
8. **90-minute improvement plan** (quick wins).
9. **1-week improvement plan** (structural upgrades).

## Constraints
- Be concrete; cite exact files and functions.
- Highlight risks that matter for non-developer users.
- Separate “must fix now” from “nice to have”.
- Keep recommendations compatible with GitHub + VS Code/Codespaces workflows.

## Project context
Use `guides/claude-handoff-deep-analysis.md` as your primary context map before judging.

## Deliverable style
- Use markdown.
- Use clear section headers.
- Use bullet points and tables.
- End with a final go/no-go recommendation.
---
