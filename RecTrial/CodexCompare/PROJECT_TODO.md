# Project To-Do (Execution Backlog)

## Priority 0 — Immediate Workflow Stability
- [ ] Confirm remote branch strategy (`work` vs `codex/*`) and standardize naming.
- [ ] Ensure successful push from local/Codespaces terminal to GitHub remote.
- [ ] Validate PR template consistency for future updates.

## Priority 1 — Verification Depth
- [ ] Add additional unit tests for edge cases in demo Python utilities.
- [ ] Add negative-path tests for malformed CSV/workbook inputs.
- [ ] Add a short “manual Excel runtime verification checklist” for host-dependent VBA behavior.

## Priority 2 — Documentation / Review Readiness
- [x] Add AI handoff package for deep comparative review (`guides/claude-handoff-deep-analysis.md`).
- [x] Add ready-to-use Claude review prompt (`guides/claude-review-prompt.md`).
- [ ] Expand architecture guide with sequence diagram (text/mermaid).
- [ ] Add “decision log” section documenting tradeoffs and deferred items.

## Priority 3 — Packaging / Distribution
- [ ] Define final packaging path (xlam-first vs workbook-first).
- [ ] Add release checklist items for signed macro distribution.
- [ ] Add semantic versioning policy and tag format.

## Priority 4 — Comparative Analysis Follow-up
- [ ] Run Claude-vs-Codex comparative review and capture findings.
- [ ] Build merged action plan from both implementations.
- [ ] Prioritize top 5 high-impact improvements and implement.

## Completion Criteria
- [ ] Push path is stable and documented for non-developers.
- [ ] Smoke + unit checks pass reliably in local and CI runs.
- [ ] Comparative review completed with clear go-forward roadmap.
