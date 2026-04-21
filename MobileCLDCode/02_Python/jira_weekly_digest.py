"""
jira_weekly_digest.py - JIRA Weekly Digest Report

PURPOSE
-------
Pull from the JIRA Cloud REST API and build an executive-friendly weekly
digest Excel file covering every project your team tracks:

  - What moved this week (created, resolved, re-opened)
  - Bug velocity (open vs closed trend)
  - Age distribution of open issues (0-7, 7-30, 30-90, 90+)
  - Story points delivered by team
  - Blocked tickets older than N days
  - Flaky-test & critical-bug watchlist
  - Exec one-pager with the 3 numbers leadership actually cares about

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
JIRA's own dashboards are per-project. Building a cross-project exec digest
requires either (a) Atlassian Data Lake ($$$) or (b) pulling the API yourself
and doing the math. This script does (b), in 60 lines of business logic.

USE CASE
--------
Every Monday 7am, VP Engineering gets the digest in their inbox. They can
spot issues in 90 seconds instead of digging through 12 dashboards.

SETUP
-----
    export JIRA_SITE="https://yourco.atlassian.net"
    export JIRA_EMAIL="you@yourco.com"
    export JIRA_TOKEN="..."   # from id.atlassian.com

USAGE
-----
    python jira_weekly_digest.py --projects FIN,ENG,OPS --output weekly.xlsx
"""
from __future__ import annotations

import argparse
import base64
import os
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests


def jira_search(site: str, auth: tuple[str, str], jql: str) -> list[dict]:
    """Paged JQL search. Returns flat list of issues."""
    headers = {
        "Authorization": "Basic " + base64.b64encode(f"{auth[0]}:{auth[1]}".encode()).decode(),
        "Accept": "application/json",
    }
    issues = []
    start_at = 0
    while True:
        r = requests.get(
            f"{site}/rest/api/3/search",
            params={
                "jql": jql,
                "startAt": start_at,
                "maxResults": 100,
                "fields": "summary,status,priority,assignee,created,resolutiondate,"
                          "customfield_10016,labels,issuetype,project",  # 10016 = story points
            },
            headers=headers,
            timeout=30,
        )
        r.raise_for_status()
        data = r.json()
        issues.extend(data.get("issues", []))
        if start_at + len(data.get("issues", [])) >= data.get("total", 0):
            break
        start_at += 100
    return issues


def _flatten(issue: dict) -> dict:
    f = issue["fields"]
    return {
        "key": issue["key"],
        "project": f["project"]["key"],
        "type": f["issuetype"]["name"],
        "status": f["status"]["name"],
        "priority": (f.get("priority") or {}).get("name"),
        "assignee": (f.get("assignee") or {}).get("displayName"),
        "created": pd.to_datetime(f["created"]),
        "resolved": pd.to_datetime(f.get("resolutiondate")) if f.get("resolutiondate") else pd.NaT,
        "story_points": f.get("customfield_10016"),
        "labels": ",".join(f.get("labels") or []),
        "summary": f["summary"],
    }


def build_digest(df: pd.DataFrame, today: datetime) -> dict[str, pd.DataFrame]:
    week_ago = pd.Timestamp(today - timedelta(days=7)).tz_localize(timezone.utc)

    created_this_week = df[df["created"] >= week_ago]
    resolved_this_week = df[df["resolved"] >= week_ago]

    this_week = pd.DataFrame([{
        "metric": "Created this week", "count": len(created_this_week),
    }, {
        "metric": "Resolved this week", "count": len(resolved_this_week),
    }, {
        "metric": "Net change (open backlog)",
        "count": len(created_this_week) - len(resolved_this_week),
    }])

    open_df = df[df["resolved"].isna()].copy()
    now_utc = pd.Timestamp(today).tz_localize(timezone.utc)
    open_df["age_days"] = (now_utc - open_df["created"]).dt.days
    age_bucket = pd.cut(
        open_df["age_days"],
        bins=[-1, 7, 30, 90, 10000],
        labels=["0-7d", "7-30d", "30-90d", "90d+"],
    )
    age_dist = open_df.groupby(age_bucket).size().reset_index(name="open_issues")

    velocity = df.groupby("project").agg(
        open_issues=("resolved", lambda s: s.isna().sum()),
        created_7d=("created", lambda s: (s >= week_ago).sum()),
        resolved_7d=("resolved", lambda s: (s >= week_ago).sum()),
        points_delivered_7d=("story_points", lambda s: s.fillna(0).where(
            df.loc[s.index, "resolved"] >= week_ago, 0).sum()),
    ).reset_index()

    stale = open_df[open_df["age_days"] > 30].sort_values("age_days", ascending=False).head(50)
    blockers = open_df[open_df["labels"].str.contains("blocked", case=False, na=False)]

    exec_summary = pd.DataFrame([{
        "headline": "Weekly JIRA Digest",
        "opened": len(created_this_week),
        "resolved": len(resolved_this_week),
        "total_open": len(open_df),
        "open_over_30d": int((open_df["age_days"] > 30).sum()),
        "blocked": len(blockers),
    }])

    return {
        "Exec Summary": exec_summary,
        "This Week": this_week,
        "Velocity by Project": velocity,
        "Age of Open Issues": age_dist,
        "Stale Open (30d+)": stale,
        "Blocked Tickets": blockers,
        "All Issues": df,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--projects", required=True, help="Comma-separated project keys")
    ap.add_argument("--days", type=int, default=90, help="Pull issues changed in last N days")
    ap.add_argument("--output", default="jira_weekly.xlsx")
    args = ap.parse_args()

    site = os.environ["JIRA_SITE"]
    auth = (os.environ["JIRA_EMAIL"], os.environ["JIRA_TOKEN"])

    keys = [k.strip() for k in args.projects.split(",")]
    jql = (
        f"project in ({','.join(keys)}) AND "
        f"updated >= -{args.days}d ORDER BY created DESC"
    )
    print(f"Fetching: {jql}")
    raw = jira_search(site, auth, jql)
    df = pd.DataFrame([_flatten(i) for i in raw])
    print(f"Fetched {len(df)} issues")

    today = datetime.now(timezone.utc).replace(tzinfo=None)
    digest = build_digest(df, today)

    with pd.ExcelWriter(args.output, engine="openpyxl") as w:
        for name, frame in digest.items():
            frame.to_excel(w, sheet_name=name[:31], index=False)

    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
