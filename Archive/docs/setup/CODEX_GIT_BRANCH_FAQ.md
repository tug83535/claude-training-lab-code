# Codex + Git Branch FAQ (Beginner-Friendly)

## Q1) If I push this, is it a new branch?
**It depends on which branch you are currently on.**

- If you push while on branch `work`, your commits go to `work`.
- That is **not** a new branch unless you create one first.

So pushing does **not automatically** create a new branch by itself.

---

## Q2) How do I check what branch I’m on?
Run:

```bash
git status --short --branch
```

You will see something like:

```text
## work
```

That means your active branch is `work`.

---

## Q3) How do I create and push a new branch safely?
Run these commands in order:

```bash
git checkout -b your-new-branch-name
git push -u origin your-new-branch-name
```

After that, future `git push` commands will push to that new branch.

---

## Q4) What does Codex usually do?
Codex makes commits on the **current checked-out branch** in your repo environment.
It does not automatically pick a random new branch unless instructed.

---

## Q5) Safe workflow before pushing
1. Confirm branch: `git status --short --branch`
2. Review commits: `git log --oneline -n 5`
3. Push intentionally to that branch.

This avoids surprises and keeps your history clean.
