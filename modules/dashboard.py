"""
GST Input Reconciliation System – Enterprise Edition
Dashboard & Analytics Module
Prepared & Developed by Karthik LVN

Provides:
  - KPI cards (Home Dashboard)
  - Interactive Plotly charts: vendor-wise, monthly, GST pie, heatmap
  - Top 20 vendors table
  - Recent uploads and processing history widgets
"""

import datetime
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from modules.utils import format_currency, safe_float, setup_logging
from modules.reconciliation import (
    get_kpi_summary,
    get_vendor_summary,
    get_monthly_summary,
    STATUS_COLORS,
)

logger = setup_logging()

# ---------------------------------------------------------------------------
# Chart Theme
# ---------------------------------------------------------------------------

CHART_THEME = "plotly_dark"
CHART_BG = "rgba(0,0,0,0)"
CHART_PAPER_BG = "rgba(0,0,0,0)"
PRIMARY_COLOR = "#00D4FF"
SECONDARY_COLOR = "#A78BFA"
ACCENT_COLOR = "#34D399"

_CHART_LAYOUT = dict(
    plot_bgcolor=CHART_BG,
    paper_bgcolor=CHART_PAPER_BG,
    font=dict(color="#EAEAEA", family="sans-serif"),
    margin=dict(l=10, r=10, t=40, b=10),
)


def _apply_dark_layout(fig):
    """Apply standard dark theme layout to a Plotly figure."""
    fig.update_layout(**_CHART_LAYOUT)
    fig.update_xaxes(gridcolor="#1A1A2E", zerolinecolor="#1A1A2E")
    fig.update_yaxes(gridcolor="#1A1A2E", zerolinecolor="#1A1A2E")
    return fig


# ---------------------------------------------------------------------------
# KPI Cards HTML
# ---------------------------------------------------------------------------

def _kpi_card_html(
    icon: str,
    label: str,
    value: str,
    subtitle: str = "",
    color: str = "#00D4FF",
    bg_color: str = "rgba(26,26,46,0.9)",
) -> str:
    """Render a styled KPI card as HTML."""
    subtitle_html = (
        "<div style=\"font-size:0.72rem; color:#64748B; margin-top:2px;\">" + subtitle + "</div>"
        if subtitle else ""
    )
    return (
        "<div style=\"background:" + bg_color + "; border:1px solid " + color + "44;"
        " border-left:4px solid " + color + "; border-radius:12px; padding:16px 14px;"
        " margin:4px 0; box-shadow:0 4px 16px " + color + "22;\">"
        "<div style=\"display:flex; align-items:center; gap:10px;\">"
        "<div style=\"font-size:1.8rem;\">" + icon + "</div>"
        "<div>"
        "<div style=\"font-size:1.4rem; font-weight:800; color:" + color + "; line-height:1.2;\">"
        + value +
        "</div>"
        "<div style=\"font-size:0.78rem; color:#94A3B8; margin-top:2px;\">" + label + "</div>"
        + subtitle_html +
        "</div>"
        "</div>"
        "</div>"
    )


# ---------------------------------------------------------------------------
# Dashboard Home
# ---------------------------------------------------------------------------

def render_home_dashboard() -> None:
    """Render the main dashboard home page with KPI cards and recent activity."""

    company = st.session_state.get("app_settings", {}).get("company_name", "My Company Pvt Ltd")
    fy = st.session_state.get("app_settings", {}).get("financial_year", "2025-26")
    username = st.session_state.get("full_name", st.session_state.get("username", "User"))
    now = datetime.datetime.now().strftime("%d %b %Y, %I:%M %p")

    # ── Welcome Header ──────────────────────────────────────────────────────
    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg, rgba(0,212,255,0.12) 0%,
             rgba(167,139,250,0.08) 100%); border-radius:16px; padding:20px 24px;
             border:1px solid rgba(0,212,255,0.2); margin-bottom:20px;">
            <div style="font-size:1.5rem; font-weight:800; color:#00D4FF;">
                📊 GST Input Reconciliation System
            </div>
            <div style="font-size:0.9rem; color:#94A3B8; margin-top:4px;">
                {company} &nbsp;·&nbsp; FY {fy} &nbsp;·&nbsp;
                Welcome, <strong style="color:#A78BFA;">{username}</strong>
                &nbsp;·&nbsp; <span style="color:#64748B;">{now}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── KPI Cards ───────────────────────────────────────────────────────────
    master_df: Optional[pd.DataFrame] = st.session_state.get("master_df")
    recon_results = st.session_state.get("recon_results")

    if master_df is not None and not master_df.empty:
        kpis = get_kpi_summary(master_df)
    else:
        kpis = {k: 0 for k in [
            "total_purchase_value", "total_gst", "matched_count", "pending_count",
            "missing_books_count", "missing_gstr2b_count", "gst_difference_total",
            "duplicate_count", "manual_review_count", "match_rate_percent",
            "fuzzy_match_count", "total_invoices",
        ]}

    # Row 1
    r1c1, r1c2, r1c3, r1c4, r1c5 = st.columns(5)
    r1c1.markdown(
        _kpi_card_html("💰", "Total Purchase Value", format_currency(kpis["total_purchase_value"]), "", "#00D4FF"),
        unsafe_allow_html=True,
    )
    r1c2.markdown(
        _kpi_card_html("🧾", "Total GST", format_currency(kpis["total_gst"]), "", "#A78BFA"),
        unsafe_allow_html=True,
    )
    r1c3.markdown(
        _kpi_card_html("✅", "Matched Invoices", f"{kpis['matched_count']:,}", f"{kpis['match_rate_percent']:.1f}% match rate", "#34D399"),
        unsafe_allow_html=True,
    )
    r1c4.markdown(
        _kpi_card_html("⏳", "Pending Invoices", f"{kpis['pending_count']:,}", "", "#FB923C"),
        unsafe_allow_html=True,
    )
    r1c5.markdown(
        _kpi_card_html("📚", "Missing in Books", f"{kpis['missing_books_count']:,}", "", "#F87171"),
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # Row 2
    r2c1, r2c2, r2c3, r2c4, r2c5 = st.columns(5)
    r2c1.markdown(
        _kpi_card_html("🏛️", "Missing in GSTR-2B", f"{kpis['missing_gstr2b_count']:,}", "", "#F97316"),
        unsafe_allow_html=True,
    )
    r2c2.markdown(
        _kpi_card_html("💸", "GST Difference", format_currency(kpis["gst_difference_total"]), "", "#FBBF24"),
        unsafe_allow_html=True,
    )
    r2c3.markdown(
        _kpi_card_html("🔁", "Duplicate Invoices", f"{kpis['duplicate_count']:,}", "", "#A78BFA"),
        unsafe_allow_html=True,
    )
    r2c4.markdown(
        _kpi_card_html("🔍", "Manual Review", f"{kpis['manual_review_count']:,}", "", "#60A5FA"),
        unsafe_allow_html=True,
    )
    r2c5.markdown(
        _kpi_card_html("📋", "Total Invoices", f"{kpis['total_invoices']:,}", "", "#67E8F9"),
        unsafe_allow_html=True,
    )

    st.divider()

    # ── Recent Uploads ──────────────────────────────────────────────────────
    col_uploads, col_quick = st.columns([2, 1])

    with col_uploads:
        st.subheader("📁 Recent Uploads")
        history = st.session_state.get("upload_history", [])
        if history:
            h_df = pd.DataFrame(history[::-1][:5])
            h_df.columns = ["Timestamp", "File", "Source", "Rows"]
            st.dataframe(h_df, use_container_width=True, hide_index=True)
        else:
            st.info("No uploads yet in this session.")

        if st.button("📤 Go to Upload", key="dash_goto_upload"):
            st.session_state["current_page"] = "Upload Data"
            st.rerun()

    with col_quick:
        st.subheader("⚡ Quick Actions")
        if st.button("Upload Files", use_container_width=True, key="qa_upload"):
            st.session_state["current_page"] = "Upload Data"
            st.rerun()
        if st.button("Column Mapping", use_container_width=True, key="qa_mapping"):
            st.session_state["current_page"] = "Column Mapping"
            st.rerun()
        if st.button("Reconcile Now", use_container_width=True, key="qa_recon"):
            st.session_state["current_page"] = "Reconcile"
            st.rerun()
        if st.button("View Reports", use_container_width=True, key="qa_reports"):
            st.session_state["current_page"] = "Reports"
            st.rerun()

    # ── Charts (if data available) ──────────────────────────────────────────
    if master_df is not None and not master_df.empty:
        st.divider()
        render_analytics_charts(master_df)


# ---------------------------------------------------------------------------
# Analytics Charts
# ---------------------------------------------------------------------------

def render_analytics_charts(master_df: pd.DataFrame) -> None:
    """Render all interactive analytics charts for the dashboard."""

    st.markdown(
        "<h3 style='color:#00D4FF;'>📈 Dashboard Analytics</h3>",
        unsafe_allow_html=True,
    )

    # ── Status Pie Chart ───────────────────────────────────────────────────
    if "status" in master_df.columns:
        col_pie, col_bar = st.columns(2)

        with col_pie:
            status_counts = master_df["status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            colors = [STATUS_COLORS.get(s, "#64748B") for s in status_counts["Status"]]

            fig_pie = go.Figure(data=[
                go.Pie(
                    labels=status_counts["Status"],
                    values=status_counts["Count"],
                    marker=dict(colors=colors),
                    hole=0.45,
                    textinfo="label+percent",
                    textfont=dict(color="#EAEAEA", size=11),
                )
            ])
            fig_pie.update_layout(
                title="Status Distribution",
                **_CHART_LAYOUT,
                height=380,
                showlegend=False,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_bar:
            # Vendor-wise top 10
            vendor_sum = get_vendor_summary(master_df)
            if not vendor_sum.empty:
                top10 = vendor_sum.head(10)
                vendor_name_col = "vendor_name" if "vendor_name" in top10.columns else top10.columns[0]

                fig_vendor = px.bar(
                    top10,
                    x="total_invoices",
                    y=vendor_name_col,
                    orientation="h",
                    title="Top 10 Vendors by Invoices",
                    color="matched" if "matched" in top10.columns else "total_invoices",
                    color_continuous_scale=[[0, "#1A1A2E"], [1, "#00D4FF"]],
                    template=CHART_THEME,
                )
                _apply_dark_layout(fig_vendor)
                fig_vendor.update_layout(height=380, yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_vendor, use_container_width=True)

    # ── Monthly Trend ──────────────────────────────────────────────────────
    monthly = get_monthly_summary(master_df)
    if not monthly.empty and "month_label" in monthly.columns:
        fig_monthly = go.Figure()

        if "total_invoices" in monthly.columns:
            fig_monthly.add_trace(go.Bar(
                x=monthly["month_label"],
                y=monthly["total_invoices"],
                name="Total Invoices",
                marker_color="#00D4FF44",
            ))
        if "matched" in monthly.columns:
            fig_monthly.add_trace(go.Scatter(
                x=monthly["month_label"],
                y=monthly["matched"],
                name="Matched",
                mode="lines+markers",
                line=dict(color="#34D399", width=3),
                marker=dict(size=8),
            ))
        if "missing_books" in monthly.columns:
            fig_monthly.add_trace(go.Scatter(
                x=monthly["month_label"],
                y=monthly["missing_books"],
                name="Missing in Books",
                mode="lines+markers",
                line=dict(color="#F87171", width=2, dash="dot"),
                marker=dict(size=6),
            ))

        fig_monthly.update_layout(
            title="Monthly Invoice Trend",
            barmode="overlay",
            **_CHART_LAYOUT,
            height=350,
            legend=dict(
                bgcolor="rgba(26,26,46,0.8)",
                bordercolor="#00D4FF44",
                borderwidth=1,
            ),
        )
        st.plotly_chart(fig_monthly, use_container_width=True)

    # ── GST Summary Chart ──────────────────────────────────────────────────
    gst_cols_present = [
        c for c in ["cgst_pr", "sgst_pr", "igst_pr", "cess_pr"] if c in master_df.columns
    ]
    if gst_cols_present:
        gst_totals = {col.replace("_pr", "").upper(): master_df[col].sum() for col in gst_cols_present}
        gst_df = pd.DataFrame(
            {"Tax Type": list(gst_totals.keys()), "Amount": list(gst_totals.values())}
        )
        colors_gst = ["#00D4FF", "#A78BFA", "#34D399", "#FBBF24"]

        fig_gst = px.bar(
            gst_df,
            x="Tax Type",
            y="Amount",
            title="GST Breakdown",
            color="Tax Type",
            color_discrete_sequence=colors_gst,
            template=CHART_THEME,
        )
        _apply_dark_layout(fig_gst)
        fig_gst.update_layout(height=300, showlegend=False)
        st.plotly_chart(fig_gst, use_container_width=True)

    # ── Top 20 Vendors Table ───────────────────────────────────────────────
    st.subheader("🏆 Top 20 Vendors by Invoice Volume")
    vendor_sum = get_vendor_summary(master_df)
    if not vendor_sum.empty:
        st.dataframe(vendor_sum.head(20), use_container_width=True, height=350)
    else:
        st.info("Run reconciliation to see vendor summary.")


# ---------------------------------------------------------------------------
# Standalone Analytics Page
# ---------------------------------------------------------------------------

def render_analytics_page() -> None:
    """Render the full Dashboard Analytics page."""

    st.markdown(
        "<h2 style='color:#00D4FF;'>📈 Dashboard Analytics</h2>",
        unsafe_allow_html=True,
    )

    master_df: Optional[pd.DataFrame] = st.session_state.get("master_df")

    if master_df is None or master_df.empty:
        st.info(
            "📊 Analytics will appear here after reconciliation is complete.\n\n"
            "Please upload files, map columns, and run reconciliation first."
        )
        if st.button("⚙️ Go to Reconciliation"):
            st.session_state["current_page"] = "Reconciliation"
            st.rerun()
        return

    render_analytics_charts(master_df)
