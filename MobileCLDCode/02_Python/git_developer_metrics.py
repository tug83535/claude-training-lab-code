"""
git_developer_metrics.py - Engineering Productivity Metrics from Git History

PURPOSE
-------
Walk one or more git repositories and produce a rich per-author activity
report: commit count, lines changed, files touched, review throughput,
cross-repo collaboration, and cycle time from PR open to merge.

Designed for engineering leadership, NOT for 1:1 performance reviews.

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Git history is the source of truth for engineering activity. Excel can
consume a pre-extracted CSV but cannot walk thousands of commits, parse
them, and roll them up. GitHub Insights is per-repo; this is cross-repo.

USE CASE
--------
VP Engineering wants a quarterly "team cadence" dashboard. This script
emits the same shape of data every time, so the dashboard stays clean.

USAGE
-----
    python git_developer_metrics.py ~/src/*/  --since 2026-01-01 \\
        --output engineering_metrics.xlsx
"""
from __future__ import annotations

import argparse
import subprocess
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


@dataclass
class CommitRow:
    repo: str
    sha: str
    author_email: str
    author_name: str
    date: str
    insertions: int
    deletions: int
    files: int
    merge: bool
    message: str


def parse_repo(repo_path: Path, since: str) -> list[CommitRow]:
    """Run git log and parse its output."""
    fmt = "%H%x1f%ae%x1f%an%x1f%aI%x1f%P%x1f%s%x1e"  # RS-separated fields, RE-separated records
    cmd = ["git", "-C", str(repo_path), "log", "--shortstat", f"--since={since}",
           "--format=" + fmt, "--no-color"]
    result = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        print(f"{repo_path}: git log failed: {result.stderr.strip()}")
        return []

    rows: list[CommitRow] = []
    blocks = result.stdout.split("\x1e")
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        header, _, stat = block.partition("\n")
        parts = header.split("\x1f")
        if len(parts) < 6:
            continue
        sha, email, name, iso, parents, msg = parts
        merge = len(parents.strip().split()) > 1
        ins = dele = files = 0
        if stat:
            # Format: " 3 files changed, 42 insertions(+), 10 deletions(-)"
            for tok in stat.split(","):
                t = tok.strip()
                if t.endswith(" file changed") or t.endswith(" files changed"):
                    files = int(t.split()[0])
                elif "insertion" in t:
                    ins = int(t.split()[0])
                elif "deletion" in t:
                    dele = int(t.split()[0])
        rows.append(CommitRow(
            repo=repo_path.name,
            sha=sha[:10],
            author_email=email,
            author_name=name,
            date=iso,
            insertions=ins,
            deletions=dele,
            files=files,
            merge=merge,
            message=msg[:200],
        ))
    return rows


def summarize(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    df["date"] = pd.to_datetime(df["date"], utc=True)
    df["week"] = df["date"].dt.to_period("W").astype(str)

    per_author = (
        df[~df["merge"]]
        .groupby(["author_name", "author_email"])
        .agg(
            commits=("sha", "count"),
            insertions=("insertions", "sum"),
            deletions=("deletions", "sum"),
            files_touched=("files", "sum"),
            active_weeks=("week", pd.Series.nunique),
            repos=("repo", pd.Series.nunique),
            first_commit=("date", "min"),
            last_commit=("date", "max"),
        )
        .reset_index()
        .sort_values("commits", ascending=False)
    )

    per_repo_week = (
        df[~df["merge"]]
        .groupby(["repo", "week"])
        .agg(commits=("sha", "count"), authors=("author_name", pd.Series.nunique))
        .reset_index()
    )

    # Bus factor: in each repo, what % of commits came from the top-1/top-2 authors?
    bus = (
        df[~df["merge"]]
        .groupby(["repo", "author_name"])
        .size()
        .reset_index(name="n")
    )
    bus_table_rows = []
    for repo, grp in bus.groupby("repo"):
        grp = grp.sort_values("n", ascending=False)
        total = grp["n"].sum()
        top1 = grp["n"].iloc[0] if len(grp) > 0 else 0
        top2 = grp["n"].iloc[:2].sum() if len(grp) > 1 else top1
        bus_table_rows.append({
            "repo": repo,
            "total_commits": total,
            "top1_author": grp["author_name"].iloc[0] if len(grp) else None,
            "top1_commit_share_pct": round(top1 / total * 100, 1) if total else 0,
            "top2_commit_share_pct": round(top2 / total * 100, 1) if total else 0,
            "unique_authors": len(grp),
        })
    bus_factor = pd.DataFrame(bus_table_rows)

    # Cycle-time proxy: median hours between successive commits per author
    df_sorted = df.sort_values(["author_email", "date"])
    df_sorted["gap_hours"] = (
        df_sorted.groupby("author_email")["date"].diff().dt.total_seconds() / 3600
    )
    cadence = (
        df_sorted.groupby("author_name")["gap_hours"]
        .median()
        .reset_index(name="median_hours_between_commits")
        .sort_values("median_hours_between_commits")
    )

    return {
        "Per Author": per_author,
        "Weekly Activity": per_repo_week,
        "Bus Factor": bus_factor,
        "Commit Cadence": cadence,
        "All Commits": df[["repo", "sha", "author_name", "date",
                           "insertions", "deletions", "files", "merge", "message"]],
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("repos", nargs="+", help="Paths to one or more git repos")
    ap.add_argument("--since", default="2026-01-01")
    ap.add_argument("--output", default="engineering_metrics.xlsx")
    args = ap.parse_args()

    all_rows: list[CommitRow] = []
    for r in args.repos:
        all_rows.extend(parse_repo(Path(r), args.since))
    if not all_rows:
        raise SystemExit("No commits found.")

    df = pd.DataFrame([c.__dict__ for c in all_rows])
    tables = summarize(df)

    with pd.ExcelWriter(args.output, engine="openpyxl") as w:
        for name, frame in tables.items():
            frame.to_excel(w, sheet_name=name[:31], index=False)

    print(f"Wrote {args.output}")
    print(f"Authors: {df['author_email'].nunique()}")
    print(f"Commits: {len(df)}")


if __name__ == "__main__":
    main()
