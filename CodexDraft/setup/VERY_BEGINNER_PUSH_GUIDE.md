# Very Beginner Guide: Push Your Branch to GitHub

If previous steps felt confusing, use this exact flow.

## What "push" means
"Push" means: send your local commits from this Codespace/VS Code terminal up to GitHub.

---

## Step 0 — Open terminal in your repo
In VS Code/Codespaces, open **Terminal**.
Make sure you are inside your repo folder.

---

## Step 1 — Check your branch name
Run:

```bash
git status --short --branch
```

You might see:

```text
## work
```

That means your branch is named `work`.

---

## Step 2 — Push this branch
Run:

```bash
git push -u origin work
```

What this does:
- `origin` = your GitHub repo remote
- `work` = your branch name
- `-u` = remembers this link for next time

If your branch is not `work`, replace `work` with your branch name from Step 1.

---

## Step 3 — Next time, push is easier
After first push, usually just run:

```bash
git push
```

---

## Step 4 — Verify on GitHub
Go to your repo page on GitHub:
1. Click branch dropdown
2. Select `work` (or your branch)
3. Confirm latest commit appears
4. Confirm `CodexDraft/` folder is there

---

## If Step 2 fails

### Error: "origin does not appear to be a git repository"
You need to add a remote first:

```bash
git remote -v
git remote add origin <YOUR_GITHUB_REPO_URL>
git push -u origin work
```

### Error: "permission denied"
Your GitHub auth/token is not connected in this environment.
Reconnect GitHub auth in Codespaces/VS Code, then retry Step 2.

### Error: "src refspec work does not match any"
Branch name is different. Re-run Step 1 and use the actual name.

---

## 30-second quick version
If you just want minimal commands:

```bash
git status --short --branch
git push -u origin work
```

(Replace `work` if your branch is named differently.)
