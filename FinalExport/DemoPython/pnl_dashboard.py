#!/usr/bin/env python3
"""
pnl_dashboard.py — Interactive Streamlit Dashboard
====================================================

PURPOSE: Web-based interactive dashboard for the P&L model.
         No Python knowledge required — just run and open browser.

USAGE:
    streamlit run pnl_dashboard.py
    streamlit run pnl_dashboard.py -- --file my_model.xlsx

REQUIREMENTS:
    pip install streamlit plotly pandas openpyxl
"""

import os
import sys
import argparse
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import numpy as np

try:
    from pnl_config import *
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pnl_config import *

try:
    import streamlit as st
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    HAS_STREAMLIT = True
except ImportError:
    HAS_STREAMLIT = False


def load_data(file_path: str) -> pd.DataFrame:
    """Load and clean the GL data."""
    base = PnLBase(verbose=False)
    return base._load_gl(file_path)


def build_dashboard():
    """Main dashboard builder."""
    # ── Page Config ──
    st.set_page_config(
        page_title=f"{APP_NAME} Dashboard",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.title(f"📊 {APP_NAME} — P&L Dashboard")
    st.caption(f"{FY_LABEL} | {APP_VERSION}")

    # ── Load Data ──
    file_path = SOURCE_FILE
    # Check CLI args (streamlit passes args after --)
    for i, arg in enumerate(sys.argv):
        if arg == "--file" and i + 1 < len(sys.argv):
            file_path = sys.argv[i + 1]

    if not os.path.exists(file_path):
        st.error(f"Source file not found: {file_path}")
        st.info("Place the Excel file in the same directory or pass --file argument")
        return

    with st.spinner("Loading P&L data..."):
        gl = load_data(file_path)

    # ── Sidebar Filters ──
    st.sidebar.header("Filters")

    months_available = sorted(gl["Month"].dropna().unique().astype(int))
    month_names = {m: MONTH_ABBREVS[m-1] for m in months_available if 1 <= m <= 12}

    selected_months = st.sidebar.multiselect(
        "Months",
        options=months_available,
        default=months_available,
        format_func=lambda m: month_names.get(m, f"M{m}")
    )

    selected_products = st.sidebar.multiselect(
        "Products",
        options=PRODUCTS,
        default=PRODUCTS
    )

    selected_depts = st.sidebar.multiselect(
        "Departments",
        options=DEPARTMENTS,
        default=DEPARTMENTS
    )

    # Apply filters
    mask = (
        gl["Month"].isin(selected_months) &
        gl["Product"].isin(selected_products + [""]) &
        gl["Department"].isin(selected_depts + [""])
    )
    gf = gl[mask].copy()

    # ── KPI Row ──
    st.markdown("---")
    k1, k2, k3, k4, k5 = st.columns(5)

    total_spend = gf["Amount"].sum()
    total_abs = gf["Abs_Amount"].sum()
    txn_count = len(gf)
    unique_vendors = gf["Vendor"].nunique()
    avg_txn = gf["Abs_Amount"].mean() if txn_count > 0 else 0

    k1.metric("Net Spend", format_currency(total_spend))
    k2.metric("Gross Spend", format_currency(total_abs))
    k3.metric("Transactions", format_number(txn_count))
    k4.metric("Vendors", format_number(unique_vendors))
    k5.metric("Avg Transaction", format_currency(avg_txn))

    st.markdown("---")

    # ── Row 1: Trend + Product Mix ──
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("Monthly Spend Trend")
        monthly = gf.groupby("Month").agg(
            Spend=("Amount", "sum"),
            Gross=("Abs_Amount", "sum"),
            Count=("Amount", "count")
        ).reset_index()
        monthly["Month_Name"] = monthly["Month"].apply(
            lambda m: MONTH_ABBREVS[int(m)-1] if 1 <= m <= 12 else f"M{int(m)}"
        )

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Bar(
            x=monthly["Month_Name"], y=monthly["Gross"],
            name="Gross Spend", marker_color=COLORS["blue"],
            opacity=0.6
        ))
        fig_trend.add_trace(go.Scatter(
            x=monthly["Month_Name"], y=monthly["Spend"],
            name="Net Spend", line=dict(color=COLORS["navy"], width=3),
            mode="lines+markers"
        ))
        fig_trend.update_layout(
            height=380, margin=dict(t=30, b=40),
            legend=dict(orientation="h", y=1.12),
            yaxis_title="Amount ($)"
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    with col2:
        st.subheader("Product Mix")
        prod_spend = gf[gf["Product"].isin(PRODUCTS)].groupby("Product")["Abs_Amount"].sum().reset_index()
        if len(prod_spend) > 0:
            colors = [PRODUCT_COLORS.get(p, COLORS["grey"]) for p in prod_spend["Product"]]
            fig_pie = go.Figure(go.Pie(
                labels=prod_spend["Product"],
                values=prod_spend["Abs_Amount"],
                marker_colors=colors,
                hole=0.4,
                textinfo="label+percent",
                textposition="outside"
            ))
            fig_pie.update_layout(height=380, margin=dict(t=30, b=30), showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No product data for selected filters")

    # ── Row 2: Department + Expense Category ──
    col3, col4 = st.columns(2)

    with col3:
        st.subheader("Department Spend")
        dept_spend = gf[gf["Department"].isin(DEPARTMENTS)].groupby("Department")["Abs_Amount"].sum()
        dept_spend = dept_spend.sort_values(ascending=True).reset_index()
        if len(dept_spend) > 0:
            fig_dept = go.Figure(go.Bar(
                y=dept_spend["Department"],
                x=dept_spend["Abs_Amount"],
                orientation="h",
                marker_color=COLORS["blue"],
                text=dept_spend["Abs_Amount"].apply(lambda x: format_currency(x)),
                textposition="outside"
            ))
            fig_dept.update_layout(height=380, margin=dict(t=30, l=140, r=80), xaxis_title="Spend ($)")
            st.plotly_chart(fig_dept, use_container_width=True)

    with col4:
        st.subheader("Expense Categories")
        cat_spend = gf[gf["Expense Category"] != ""].groupby("Expense Category")["Abs_Amount"].sum()
        cat_spend = cat_spend.sort_values(ascending=False).head(8).reset_index()
        if len(cat_spend) > 0:
            fig_cat = go.Figure(go.Bar(
                x=cat_spend["Expense Category"],
                y=cat_spend["Abs_Amount"],
                marker_color=COLORS["amber"],
                text=cat_spend["Abs_Amount"].apply(lambda x: format_currency(x)),
                textposition="outside"
            ))
            fig_cat.update_layout(height=380, margin=dict(t=30, b=100), yaxis_title="Spend ($)",
                                  xaxis_tickangle=-35)
            st.plotly_chart(fig_cat, use_container_width=True)

    # ── Row 3: Heatmap ──
    st.subheader("Department × Product Heatmap")
    heatmap_data = gf[
        gf["Department"].isin(DEPARTMENTS) & gf["Product"].isin(PRODUCTS)
    ].groupby(["Department", "Product"])["Abs_Amount"].sum().reset_index()

    if len(heatmap_data) > 0:
        pivot = heatmap_data.pivot_table(
            index="Department", columns="Product", values="Abs_Amount", fill_value=0
        )
        # Reorder products
        for p in PRODUCTS:
            if p not in pivot.columns:
                pivot[p] = 0
        pivot = pivot[PRODUCTS]

        fig_heat = go.Figure(go.Heatmap(
            z=pivot.values,
            x=pivot.columns.tolist(),
            y=pivot.index.tolist(),
            colorscale="Blues",
            text=[[format_currency(v) for v in row] for row in pivot.values],
            texttemplate="%{text}",
            textfont={"size": 11},
            hoverongaps=False
        ))
        fig_heat.update_layout(height=350, margin=dict(t=30, b=40))
        st.plotly_chart(fig_heat, use_container_width=True)

    # ── Row 4: Top Vendors + Data Table ──
    col5, col6 = st.columns(2)

    with col5:
        st.subheader("Top 10 Vendors")
        vendor_spend = gf[gf["Vendor"] != ""].groupby("Vendor").agg(
            Spend=("Abs_Amount", "sum"),
            Txns=("Amount", "count")
        ).sort_values("Spend", ascending=False).head(10).reset_index()

        if len(vendor_spend) > 0:
            fig_vendor = go.Figure(go.Bar(
                y=vendor_spend["Vendor"],
                x=vendor_spend["Spend"],
                orientation="h",
                marker_color=COLORS["green"],
                text=vendor_spend["Spend"].apply(lambda x: format_currency(x)),
                textposition="outside"
            ))
            fig_vendor.update_layout(height=380, margin=dict(t=30, l=160, r=80),
                                     xaxis_title="Spend ($)", yaxis_autorange="reversed")
            st.plotly_chart(fig_vendor, use_container_width=True)

    with col6:
        st.subheader("Transaction Detail")
        st.dataframe(
            gf[["Date", "Department", "Product", "Vendor", "Amount"]].head(100)
            .style.format({"Amount": "${:,.2f}"}),
            height=380,
            use_container_width=True
        )

    # ── Footer ──
    st.markdown("---")
    st.caption(
        f"{APP_NAME} v{APP_VERSION} | "
        f"Data: {file_path} | "
        f"{len(gl):,} total records | "
        f"Showing {len(gf):,} filtered records"
    )


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    if not HAS_STREAMLIT:
        print("Streamlit not installed. Run: pip install streamlit plotly")
        print("Then: streamlit run pnl_dashboard.py")
        sys.exit(1)
    build_dashboard()
