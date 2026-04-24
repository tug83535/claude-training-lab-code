# Git Branch + Push Quickstart (Beginner-Safe)

This guide shows exactly how to move local commits in this repo to GitHub and open a PR.

## 1) Confirm where you are

Run in terminal:

```bash
pwd
git status -sb
git branch --show-current
```

Expected:
- you are inside this repo folder,
- you see your current branch name,
- no unexpected unstaged changes before pushing.

## 2) Check whether remote is configured

```bash
git remote -v
```

If you see no output, add your repo remote once:

```bash
git remote add origin https://github.com/<your-user-or-org>/<your-repo>.git
```

If origin exists but URL is wrong:

```bash
git remote set-url origin https://github.com/<your-user-or-org>/<your-repo>.git
```

## 3) Push your branch

If local and remote branch names should match:

```bash
git push -u origin <your-local-branch>
```

Example:

```bash
git push -u origin work
```

If you want a different branch name on GitHub:

```bash
git push -u origin <local-branch>:<remote-branch>
```

Example:

```bash
git push -u origin work:codex/task-title
```

## 4) Open the pull request

After push:
1. Open your GitHub repository.
2. Click **Compare & pull request**.
3. Review file list and test evidence.
4. Submit PR.

## 5) Common errors and fixes

### Error: `No configured push destination`
Fix: add `origin` remote (Step 2), then push again.

### Error: `remote origin already exists`
Fix: use `git remote set-url origin ...`.

### Error: permission/auth failed
Fix:
- sign in to GitHub in VS Code/Codespaces,
- or run `gh auth login`,
- then retry push.

### Pushed wrong branch
Fix: push branch mapping explicitly:

```bash
git push -u origin <local>:<correct-remote>
```

## 6) Recommended mini-checklist before clicking “Create PR”

- `bash scripts/run_stage_smoke.sh` passes.
- `git status -sb` is clean.
- Branch is pushed (`git branch -vv` shows upstream).
- PR description includes summary + testing commands.
