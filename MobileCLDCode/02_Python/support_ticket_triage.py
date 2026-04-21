"""
support_ticket_triage.py - Support Ticket Auto-Triage & Theme Summarizer

PURPOSE
-------
Ingest a week/month of customer support tickets (Zendesk, Salesforce, Freshdesk,
ServiceNow export) and produce:

  - Auto-classification into business categories (Billing, Bug, How-To, Outage,
    Feature Request, Abuse) using keyword + TF-IDF nearest-centroid
  - Sentiment score per ticket (negative / neutral / positive)
  - Top 10 recurring themes (noun phrase clustering)
  - Escalation scoring: which tickets deserve a human read today
  - Excel deliverable the CS ops lead can drop into Monday standup

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Excel + OneDrive have zero NLP. Power BI's "text analytics" add-in is a
paid Azure Cognitive Services hit per row - expensive at 50K tickets.
This runs locally with scikit-learn for free and in seconds.

USE CASE
--------
Every Monday, CS ops gets a digest: "The top 3 new themes this week were
'SSO timeout', 'renewal quote PDF broken', and 'invoice email not received'.
Here are 12 tickets Sales should see today."

INPUT: tickets.csv with columns
    ticket_id, created_at, customer_id, subject, body, priority, status

USAGE
-----
    python support_ticket_triage.py tickets.csv --output digest.xlsx
"""
from __future__ import annotations

import argparse
import re
from collections import Counter

import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans


CATEGORY_KEYWORDS = {
    "Billing": ["invoice", "charge", "refund", "payment", "credit", "tax", "past due", "ach"],
    "Bug": ["error", "broken", "crash", "500", "exception", "doesn't work", "fails"],
    "How-To": ["how do i", "how to", "where is", "tutorial", "can you help"],
    "Outage": ["down", "can't log in", "timeout", "unavailable", "503", "slow"],
    "FeatureRequest": ["feature request", "wish", "would love", "could you add"],
    "Abuse": ["spam", "hack", "phish", "unauthorized", "suspicious activity"],
}


POSITIVE_WORDS = {"thanks", "thank you", "great", "awesome", "love", "perfect", "appreciate"}
NEGATIVE_WORDS = {"angry", "frustrated", "terrible", "worst", "unacceptable", "cancel", "sue", "refund"}


def classify_category(text: str) -> str:
    """Crude-but-effective keyword match with tiebreaker."""
    t = text.lower()
    scores = {cat: sum(1 for kw in kws if kw in t) for cat, kws in CATEGORY_KEYWORDS.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "Other"


def sentiment(text: str) -> str:
    t = set(re.findall(r"\b\w+\b", text.lower()))
    pos = len(t & POSITIVE_WORDS)
    neg = len(t & NEGATIVE_WORDS)
    if neg > pos + 1:
        return "negative"
    if pos > neg + 1:
        return "positive"
    return "neutral"


def escalation_score(row: pd.Series) -> float:
    s = 0.0
    if row["sentiment"] == "negative":
        s += 3
    if row.get("priority", "").lower() in ("urgent", "high", "critical"):
        s += 3
    if row["category"] in ("Outage", "Abuse"):
        s += 4
    if "refund" in str(row.get("body", "")).lower():
        s += 2
    if "cancel" in str(row.get("body", "")).lower():
        s += 3
    return s


def extract_themes(df: pd.DataFrame, n_themes: int = 10) -> pd.DataFrame:
    """TF-IDF + KMeans = low-cost topic model."""
    docs = (df["subject"].fillna("") + " " + df["body"].fillna("")).tolist()
    if len(docs) < n_themes:
        return pd.DataFrame()

    vec = TfidfVectorizer(
        max_features=2000,
        ngram_range=(1, 2),
        stop_words="english",
        min_df=3,
    )
    X = vec.fit_transform(docs)
    km = KMeans(n_clusters=n_themes, n_init=10, random_state=42).fit(X)

    terms = np.array(vec.get_feature_names_out())
    rows = []
    for c in range(n_themes):
        top_idx = km.cluster_centers_[c].argsort()[::-1][:6]
        label = ", ".join(terms[top_idx])
        size = int((km.labels_ == c).sum())
        rows.append({"theme": f"Theme {c+1}", "top_terms": label, "tickets": size})
    return pd.DataFrame(rows).sort_values("tickets", ascending=False)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("tickets_csv")
    ap.add_argument("--output", default="support_digest.xlsx")
    ap.add_argument("--top-escalations", type=int, default=25)
    args = ap.parse_args()

    df = pd.read_csv(args.tickets_csv, parse_dates=["created_at"])
    df["combined"] = (df["subject"].fillna("") + " " + df["body"].fillna("")).str.strip()
    df["category"] = df["combined"].apply(classify_category)
    df["sentiment"] = df["combined"].apply(sentiment)
    df["escalation_score"] = df.apply(escalation_score, axis=1)

    themes = extract_themes(df)

    cat_summary = (
        df.groupby("category")
        .agg(tickets=("ticket_id", "count"),
             negative=("sentiment", lambda s: (s == "negative").sum()))
        .reset_index()
        .sort_values("tickets", ascending=False)
    )

    top_escalations = df.sort_values("escalation_score", ascending=False).head(args.top_escalations)
    top_escalations = top_escalations[
        ["ticket_id", "created_at", "customer_id", "subject",
         "category", "sentiment", "priority", "escalation_score"]
    ]

    with pd.ExcelWriter(args.output, engine="openpyxl") as w:
        cat_summary.to_excel(w, sheet_name="Category Summary", index=False)
        themes.to_excel(w, sheet_name="Themes", index=False)
        top_escalations.to_excel(w, sheet_name="Top Escalations", index=False)
        df[["ticket_id", "created_at", "customer_id", "subject", "category",
            "sentiment", "priority", "escalation_score"]].to_excel(
            w, sheet_name="All Tickets", index=False
        )

    print(f"Wrote {args.output}")
    print(f"Total tickets: {len(df)}  |  Negative sentiment: {(df['sentiment']=='negative').sum()}")
    print(cat_summary.to_string(index=False))


if __name__ == "__main__":
    main()
