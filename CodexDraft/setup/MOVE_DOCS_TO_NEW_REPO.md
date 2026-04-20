# Move These Markdown Docs to a New Repo

## Short Answer
You have 3 good options:
1. **Copy files directly** (simplest)
2. **Cherry-pick specific commits** (clean Git history)
3. **Push current branch to new remote** (move everything on this branch)

Use Option 1 if you are new to Git and want the safest path.

---

## Files to Move
From this repo, these are the docs created in this workstream:

- `Archive/docs/overview/DEMO_STORY_RECOMMENDATION.md`
- `Archive/docs/overview/FOUR_VIDEO_DEMO_SERIES_PLAN.md`
- `Archive/docs/overview/FOUR_VIDEO_SHOT_BY_SHOT_RUN_SHEET.md`
- `Archive/docs/overview/TOTAL_CODEBASE_100_PLUS_EXAMPLES_PLAN.md`
- `Archive/docs/setup/BUILD_VIDEO_DEMO_WORKBOOK.md`
- `Archive/docs/setup/CODEX_GIT_BRANCH_FAQ.md`
- `Archive/docs/setup/MOVE_DOCS_TO_NEW_REPO.md`

---

## Option 1 — Direct Copy (Recommended for beginners)

### A) Copy from old repo into new repo working folder

```bash
# run from OLD repo root
cp Archive/docs/overview/DEMO_STORY_RECOMMENDATION.md /path/to/new-repo/Archive/docs/overview/
cp Archive/docs/overview/FOUR_VIDEO_DEMO_SERIES_PLAN.md /path/to/new-repo/Archive/docs/overview/
cp Archive/docs/overview/FOUR_VIDEO_SHOT_BY_SHOT_RUN_SHEET.md /path/to/new-repo/Archive/docs/overview/
cp Archive/docs/overview/TOTAL_CODEBASE_100_PLUS_EXAMPLES_PLAN.md /path/to/new-repo/Archive/docs/overview/
cp Archive/docs/setup/BUILD_VIDEO_DEMO_WORKBOOK.md /path/to/new-repo/Archive/docs/setup/
cp Archive/docs/setup/CODEX_GIT_BRANCH_FAQ.md /path/to/new-repo/Archive/docs/setup/
```

### B) Commit in new repo

```bash
cd /path/to/new-repo
git add Archive/docs/overview/*.md Archive/docs/setup/*.md
git commit -m "Add demo planning/setup markdown docs"
git push
```

---

## Option 2 — Cherry-Pick Specific Commits (Keeps authorship/history)

If the new repo has this old repo added as a remote/fetch source:

```bash
cd /path/to/new-repo
git remote add oldrepo /path/to/old/repo   # or URL
git fetch oldrepo
```

Then cherry-pick only the relevant commits:

```bash
git cherry-pick 2f8422d 586cffe 1aa8422 6685a1a 4c1206d d185d5c f4cec12
```

If conflict happens:

```bash
# resolve files
git add .
git cherry-pick --continue
```

Then push:

```bash
git push
```

---

## Option 3 — Push Current Branch to New Remote (Moves all branch commits)

```bash
# from current repo
git remote add neworigin <NEW_REPO_URL>
git push -u neworigin work
```

If new repo expects `main`:

```bash
git push neworigin work:main
```

Use this only if you want the entire branch history, not just selected files.

---

## Quick Safety Checklist
- Confirm destination repo path before copying
- Run `git status` before and after copy
- Open each moved file in new repo to verify formatting
- Commit with clear message
- Push and verify on Git host UI

---

## Recommendation
For your situation (new clone + new repo), use **Option 1 (Direct Copy)** first.
It is simplest, low-risk, and easy to verify.
