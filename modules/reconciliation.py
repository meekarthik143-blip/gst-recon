"""
GST Input Reconciliation System – Enterprise Edition
Reconciliation Results & Classification Module
Prepared & Developed by Karthik LVN

Provides:
  - Classification of matched records (Perfect Match / GST Diff / Date Diff / etc.)
  - Master reconciliation DataFrame builder
  - Vendor summary and monthly summary aggregations
  - KPI summary for dashboard
  - Full reconciliation results Streamlit page with filters, search, and export
"""

from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st

from modules.matching import (
    STATUS_PERFECT_MATCH,
    STATUS_MISSING_BOOKS,
    STATUS_MISSING_GSTR2B,
    STATUS_GST_DIFF,
    STATUS_TAXABLE_DIFF,
    STATUS_INV_VALUE_DIFF,
    STATUS_DATE_DIFF,
    STATUS_DUPLICATE,
    STATUS_GSTIN_MISMATCH,
    STATUS_VENDOR_DIFF,
    STATUS_MANUAL_REVIEW,
    STATUS_FUZZY_MATCH,
    CONFIDENCE_HIGH,
    CONFIDENCE_MEDIUM,
    CONFIDENCE_LOW,
)
from modules.utils import (
    format_currency,
    safe_float,
    round_gst,
    get_month_name,
    setup_logging,
)
from modules.audit import log_event

logger = setup_logging()

# ---------------------------------------------------------------------------
# Status → Suggested Action Map
# ---------------------------------------------------------------------------

STATUS_ACTIONS: dict[str, str] = {
    STATUS_PERFECT_MATCH: "No Action Required",
    STATUS_MISSING_BOOKS: "Book Purchase Entry",
    STATUS_MISSING_GSTR2B: "Vendor Follow-up",
    STATUS_GST_DIFF: "Request Amendment / Verify GST",
    STATUS_TAXABLE_DIFF: "Verify Invoice Taxable Amount",
    STATUS_INV_VALUE_DIFF: "Check Invoice Total Value",
    STATUS_DATE_DIFF: "Verify Invoice Date with Vendor",
    STATUS_DUPLICATE: "Remove Duplicate Entry",
    STATUS_GSTIN_MISMATCH: "Verify GSTIN with Vendor",
    STATUS_VENDOR_DIFF: "Confirm Vendor Identity",
    STATUS_MANUAL_REVIEW: "Manual Verification Required",
    STATUS_FUZZY_MATCH: "Confirm Fuzzy Match Manually",
}

# Color coding for statuses (used in styling)
STATUS_COLORS: dict[str, str] = {
    STATUS_PERFECT_MATCH: "#34D399",       # green
    STATUS_MISSING_BOOKS: "#FB923C",        # orange
    STATUS_MISSING_GSTR2B: "#F87171",       # red
    STATUS_GST_DIFF: "#FBBF24",            # yellow
    STATUS_TAXABLE_DIFF: "#FDE68A",        # light yellow
    STATUS_INV_VALUE_DIFF: "#FCD34D",      # amber
    STATUS_DATE_DIFF: "#C4B5FD",           # purple-ish
    STATUS_DUPLICATE: "#A78BFA",           # purple
    STATUS_GSTIN_MISMATCH: "#F472B6",      # pink
    STATUS_VENDOR_DIFF: "#67E8F9",         # cyan
    STATUS_MANUAL_REVIEW: "#60A5FA",       # blue
    STATUS_FUZZY_MATCH: "#4ADE80",         # lime
}

CONFIDENCE_COLORS = {
    CONFIDENCE_HIGH: "#34D399",
    CONFIDENCE_MEDIUM: "#FBBF24",
    CONFIDENCE_LOW: "#F87171",
}


# ---------------------------------------------------------------------------
# Classification of Matched Records
# ---------------------------------------------------------------------------

def classify_matched_records(
    matched_df: pd.DataFrame,
    gst_tolerance: float = 1.0,
) -> pd.DataFrame:
    """
    Classify each matched record into a specific status based on field comparison.

    Comparison priority:
      1. GST difference > tolerance → GST Difference
      2. Taxable value difference   → Taxable Difference
      3. Invoice date mismatch      → Date Difference
      4. Invoice value difference   → Invoice Value Difference
      5. All match                  → Perfect Match

    Args:
        matched_df:    DataFrame of matched records (has _pr and _gstr2b suffixed columns).
        gst_tolerance: Maximum allowed GST difference in ₹ before flagging.

    Returns:
        DataFrame with added columns: status, action, gst_difference,
        taxable_difference, invoice_diff, date_mismatch
    """
    if matched_df.empty:
        return matched_df

    df = matched_df.copy()

    def _get(col_pr: str, col_gstr: str, row: pd.Series) -> tuple:
        val_pr = safe_float(row.get(col_pr, 0))
        val_gstr = safe_float(row.get(col_gstr, 0))
        return val_pr, val_gstr, abs(val_pr - val_gstr)

    statuses = []
    actions = []
    gst_diffs = []
    tax_diffs = []
    inv_diffs = []
    date_mismatches = []

    for _, row in df.iterrows():
        # GST difference
        gst_pr, gst_gstr, gst_diff = _get("total_gst_pr", "total_gst_gstr2b", row)
        # Taxable difference
        tax_pr, tax_gstr, tax_diff = _get("taxable_value_pr", "taxable_value_gstr2b", row)
        # Invoice value difference
        inv_pr, inv_gstr, inv_diff = _get("invoice_value_pr", "invoice_value_gstr2b", row)
        # Date comparison
        date_pr = str(row.get("invoice_date_pr", "")).strip()
        date_gstr = str(row.get("invoice_date_gstr2b", "")).strip()
        date_mismatch = date_pr != date_gstr and date_pr and date_gstr

        gst_diffs.append(round(gst_diff, 2))
        tax_diffs.append(round(tax_diff, 2))
        inv_diffs.append(round(inv_diff, 2))
        date_mismatches.append(date_mismatch)

        # Status determination (priority order)
        if gst_diff > gst_tolerance:
            statuses.append(STATUS_GST_DIFF)
            actions.append(STATUS_ACTIONS[STATUS_GST_DIFF])
        elif tax_diff > gst_tolerance:
            statuses.append(STATUS_TAXABLE_DIFF)
            actions.append(STATUS_ACTIONS[STATUS_TAXABLE_DIFF])
        elif date_mismatch:
            statuses.append(STATUS_DATE_DIFF)
            actions.append(STATUS_ACTIONS[STATUS_DATE_DIFF])
        elif inv_diff > gst_tolerance:
            statuses.append(STATUS_INV_VALUE_DIFF)
            actions.append(STATUS_ACTIONS[STATUS_INV_VALUE_DIFF])
        else:
            statuses.append(STATUS_PERFECT_MATCH)
            actions.append(STATUS_ACTIONS[STATUS_PERFECT_MATCH])

    df["status"] = statuses
    df["action"] = actions
    df["gst_difference"] = gst_diffs
    df["taxable_difference"] = tax_diffs
    df["invoice_diff"] = inv_diffs
    df["date_mismatch"] = date_mismatches

    return df


# ---------------------------------------------------------------------------
# Master DataFrame Builder
# ---------------------------------------------------------------------------

def build_master_reconciliation(recon_results: dict, gst_tolerance: float = 1.0) -> pd.DataFrame:
    """
    Build a single master DataFrame combining all reconciliation results.

    Args:
        recon_results: Dict returned by run_reconciliation().
        gst_tolerance: GST tolerance for classification.

    Returns:
        Master DataFrame with standardized columns including status, action, confidence_score.
    """
    parts = []

    # ── Matched records ────────────────────────────────────────────────────
    matched = recon_results.get("matched", pd.DataFrame())
    if not matched.empty:
        classified = classify_matched_records(matched, gst_tolerance)
        # Flatten to single-row representation (use PR values as primary)
        flat_rows = []
        for _, row in classified.iterrows():
            flat = {
                "vendor_name": row.get("vendor_name_pr", row.get("vendor_name", "")),
                "gstin": row.get("gstin_pr", row.get("gstin", "")),
                "invoice_number": row.get("invoice_number_pr", row.get("invoice_number", "")),
                "invoice_date_pr": row.get("invoice_date_pr", ""),
                "invoice_date_gstr2b": row.get("invoice_date_gstr2b", ""),
                "taxable_value_pr": safe_float(row.get("taxable_value_pr", 0)),
                "taxable_value_gstr2b": safe_float(row.get("taxable_value_gstr2b", 0)),
                "total_gst_pr": safe_float(row.get("total_gst_pr", 0)),
                "total_gst_gstr2b": safe_float(row.get("total_gst_gstr2b", 0)),
                "invoice_value_pr": safe_float(row.get("invoice_value_pr", 0)),
                "invoice_value_gstr2b": safe_float(row.get("invoice_value_gstr2b", 0)),
                "gst_difference": row.get("gst_difference", 0.0),
                "taxable_difference": row.get("taxable_difference", 0.0),
                "invoice_diff": row.get("invoice_diff", 0.0),
                "status": row.get("status", STATUS_PERFECT_MATCH),
                "action": row.get("action", ""),
                "match_tier": row.get("match_tier", 1),
                "match_key": row.get("match_key", ""),
                "confidence": CONFIDENCE_HIGH,
                "source": "BOTH",
                "remarks": "",
            }
            flat_rows.append(flat)
        if flat_rows:
            parts.append(pd.DataFrame(flat_rows))

    # ── Missing in Books (GSTR-2B only) ───────────────────────────────────
    miss_books = recon_results.get("missing_in_books", pd.DataFrame())
    if not miss_books.empty:
        mb = miss_books.copy()
        mb["status"] = STATUS_MISSING_BOOKS
        mb["action"] = STATUS_ACTIONS[STATUS_MISSING_BOOKS]
        mb["confidence"] = CONFIDENCE_HIGH
        mb["source"] = "GSTR2B"
        mb["match_tier"] = 0
        mb["match_key"] = "Unmatched"
        mb["gst_difference"] = 0.0
        mb["taxable_difference"] = 0.0
        mb["invoice_diff"] = 0.0
        mb["remarks"] = "Present in GSTR-2B but not in Purchase Register"
        # Rename for uniformity
        for col in ["taxable_value", "total_gst", "invoice_value"]:
            if col in mb.columns:
                mb[f"{col}_pr"] = 0.0
                mb[f"{col}_gstr2b"] = mb[col]
        parts.append(mb)

    # ── Missing in GSTR-2B (PR only) ──────────────────────────────────────
    miss_gstr = recon_results.get("missing_in_gstr2b", pd.DataFrame())
    if not miss_gstr.empty:
        mg = miss_gstr.copy()
        mg["status"] = STATUS_MISSING_GSTR2B
        mg["action"] = STATUS_ACTIONS[STATUS_MISSING_GSTR2B]
        mg["confidence"] = CONFIDENCE_HIGH
        mg["source"] = "PR"
        mg["match_tier"] = 0
        mg["match_key"] = "Unmatched"
        mg["gst_difference"] = 0.0
        mg["taxable_difference"] = 0.0
        mg["invoice_diff"] = 0.0
        mg["remarks"] = "Present in Purchase Register but not in GSTR-2B"
        for col in ["taxable_value", "total_gst", "invoice_value"]:
            if col in mg.columns:
                mg[f"{col}_pr"] = mg[col]
                mg[f"{col}_gstr2b"] = 0.0
        parts.append(mg)

    # ── Fuzzy candidates ────────────────────────────────────────────────────
    fuzzy = recon_results.get("fuzzy_candidates", pd.DataFrame())
    if not fuzzy.empty:
        fz = fuzzy.copy()
        fz["status"] = STATUS_FUZZY_MATCH
        fz["action"] = STATUS_ACTIONS[STATUS_FUZZY_MATCH]
        fz["source"] = "BOTH"
        fz["match_key"] = "Fuzzy Match"
        fz["remarks"] = fz.get("fuzzy_score", pd.Series("")).apply(
            lambda s: f"Fuzzy score: {float(s)*100:.1f}%" if s else ""
        )
        parts.append(fz)

    if not parts:
        return pd.DataFrame()

    master = pd.concat(parts, ignore_index=True)
    logger.info(f"Master reconciliation built: {len(master)} total rows")
    return master


# ---------------------------------------------------------------------------
# Summary Functions
# ---------------------------------------------------------------------------

def get_vendor_summary(master_df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate reconciliation results by vendor.

    Args:
        master_df: Master reconciliation DataFrame.

    Returns:
        Vendor-level summary DataFrame.
    """
    if master_df.empty:
        return pd.DataFrame()

    vendor_col = "vendor_name" if "vendor_name" in master_df.columns else None
    gstin_col = "gstin" if "gstin" in master_df.columns else None

    if not vendor_col:
        return pd.DataFrame()

    group_cols = [c for c in [vendor_col, gstin_col] if c]

    agg = master_df.groupby(group_cols).agg(
        total_invoices=("status", "count"),
        matched=(
            "status",
            lambda x: (x == STATUS_PERFECT_MATCH).sum(),
        ),
        missing_books=(
            "status",
            lambda x: (x == STATUS_MISSING_BOOKS).sum(),
        ),
        missing_gstr2b=(
            "status",
            lambda x: (x == STATUS_MISSING_GSTR2B).sum(),
        ),
        gst_diff_amount=("gst_difference", "sum"),
        taxable_diff_amount=("taxable_difference", "sum"),
    )

    # Add purchase value and GST columns where available
    if "taxable_value_pr" in master_df.columns:
        tv = master_df.groupby(group_cols)["taxable_value_pr"].sum()
        agg["total_purchase_value"] = tv
    if "total_gst_pr" in master_df.columns:
        tg = master_df.groupby(group_cols)["total_gst_pr"].sum()
        agg["total_gst"] = tg

    agg = agg.reset_index()
    agg["gst_diff_amount"] = agg["gst_diff_amount"].round(2)
    agg["taxable_diff_amount"] = agg["taxable_diff_amount"].round(2)

    return agg.sort_values("total_invoices", ascending=False)


def get_monthly_summary(master_df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate reconciliation results by month.

    Args:
        master_df: Master reconciliation DataFrame.

    Returns:
        Month-level summary DataFrame.
    """
    if master_df.empty:
        return pd.DataFrame()

    date_col = "invoice_date_pr" if "invoice_date_pr" in master_df.columns else "invoice_date"

    if date_col not in master_df.columns:
        return pd.DataFrame()

    df = master_df.copy()
    df["_parsed_date"] = pd.to_datetime(df[date_col], format="%d-%m-%Y", errors="coerce")
    df = df.dropna(subset=["_parsed_date"])
    df["month_num"] = df["_parsed_date"].dt.month
    df["year"] = df["_parsed_date"].dt.year
    df["month_label"] = df.apply(
        lambda r: f"{get_month_name(int(r['month_num']))}-{int(r['year'])}", axis=1
    )

    agg = df.groupby(["year", "month_num", "month_label"]).agg(
        total_invoices=("status", "count"),
        matched=("status", lambda x: (x == STATUS_PERFECT_MATCH).sum()),
        missing_books=("status", lambda x: (x == STATUS_MISSING_BOOKS).sum()),
        missing_gstr2b=("status", lambda x: (x == STATUS_MISSING_GSTR2B).sum()),
    )

    if "taxable_value_pr" in df.columns:
        tv = df.groupby(["year", "month_num", "month_label"])["taxable_value_pr"].sum()
        agg["total_purchase"] = tv
    if "total_gst_pr" in df.columns:
        tg = df.groupby(["year", "month_num", "month_label"])["total_gst_pr"].sum()
        agg["total_gst"] = tg

    agg = agg.reset_index().sort_values(["year", "month_num"])
    return agg.drop(columns=["year", "month_num"])


def get_kpi_summary(master_df: pd.DataFrame) -> dict:
    """
    Calculate KPI metrics from the master reconciliation DataFrame.

    Args:
        master_df: Master reconciliation DataFrame.

    Returns:
        Dict with KPI values.
    """
    if master_df.empty:
        return {k: 0 for k in [
            "total_purchase_value", "total_gst", "matched_count", "pending_count",
            "missing_books_count", "missing_gstr2b_count", "gst_difference_total",
            "duplicate_count", "manual_review_count", "match_rate_percent",
            "fuzzy_match_count", "total_invoices",
        ]}

    status_counts = master_df["status"].value_counts().to_dict() if "status" in master_df.columns else {}

    total = len(master_df)
    matched = status_counts.get(STATUS_PERFECT_MATCH, 0)

    return {
        "total_invoices": total,
        "total_purchase_value": safe_float(
            master_df.get("taxable_value_pr", pd.Series(0)).sum()
        ),
        "total_gst": safe_float(
            master_df.get("total_gst_pr", pd.Series(0)).sum()
        ),
        "matched_count": matched,
        "pending_count": total - matched,
        "missing_books_count": status_counts.get(STATUS_MISSING_BOOKS, 0),
        "missing_gstr2b_count": status_counts.get(STATUS_MISSING_GSTR2B, 0),
        "gst_difference_total": safe_float(
            master_df.get("gst_difference", pd.Series(0)).sum()
        ),
        "duplicate_count": status_counts.get(STATUS_DUPLICATE, 0),
        "manual_review_count": status_counts.get(STATUS_MANUAL_REVIEW, 0),
        "fuzzy_match_count": status_counts.get(STATUS_FUZZY_MATCH, 0),
        "match_rate_percent": round(matched / max(total, 1) * 100, 2),
    }


# ---------------------------------------------------------------------------
# Streamlit Results Page
# ---------------------------------------------------------------------------

def _kpi_card(label: str, value: str, color: str = "#00D4FF", icon: str = "") -> str:
    """Generate HTML for a KPI metric card."""
    return f"""
    <div style="background:rgba(26,26,46,0.8); border:1px solid {color}44;
         border-radius:12px; padding:16px 12px; text-align:center;
         box-shadow: 0 4px 12px {color}22;">
        <div style="font-size:1.4rem;">{icon}</div>
        <div style="font-size:1.5rem; font-weight:800; color:{color};">{value}</div>
        <div style="font-size:0.75rem; color:#94A3B8; margin-top:4px;">{label}</div>
    </div>
    """


def render_reconciliation_results_page() -> None:
    """Render the full reconciliation results page with filters, tabs, and export."""

    st.markdown(
        "<h2 style='color:#00D4FF;'>📊 Reconciliation Results</h2>",
        unsafe_allow_html=True,
    )

    recon_results = st.session_state.get("recon_results")
    if not recon_results:
        st.warning("⚠️ No reconciliation results available. Please run Reconciliation first.")
        if st.button("Go to Reconciliation", key="goto_recon_from_results"):
            st.session_state["current_page"] = "Reconciliation"
            st.rerun()
        return

    # Build / retrieve master DF
    if "master_df" not in st.session_state or st.session_state.get("master_df") is None:
        gst_tol = st.session_state.get("app_settings", {}).get("gst_tolerance", 1.0)
        with st.spinner("Building master reconciliation…"):
            master_df = build_master_reconciliation(recon_results, gst_tolerance=gst_tol)
        st.session_state["master_df"] = master_df
    else:
        master_df = st.session_state["master_df"]

    if master_df is None or master_df.empty:
        st.error("Could not build master reconciliation DataFrame.")
        return

    kpis = get_kpi_summary(master_df)

    # ── KPI Row ────────────────────────────────────────────────────────────
    kpi_cols = st.columns(5)
    kpi_data = [
        ("💰 Total Purchase", format_currency(kpis["total_purchase_value"]), "#00D4FF"),
        ("🧾 Total GST", format_currency(kpis["total_gst"]), "#A78BFA"),
        ("✅ Matched", f"{kpis['matched_count']:,}", "#34D399"),
        ("⚠️ Pending", f"{kpis['pending_count']:,}", "#FB923C"),
        ("📈 Match Rate", f"{kpis['match_rate_percent']:.1f}%", "#60A5FA"),
    ]
    for col, (label, val, color) in zip(kpi_cols, kpi_data):
        col.markdown(_kpi_card(label, val, color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    kpi_cols2 = st.columns(5)
    kpi_data2 = [
        ("📚 Missing in Books", f"{kpis['missing_books_count']:,}", "#F87171"),
        ("🏛️ Missing in GSTR-2B", f"{kpis['missing_gstr2b_count']:,}", "#FB923C"),
        ("💸 GST Difference", format_currency(kpis["gst_difference_total"]), "#FBBF24"),
        ("🔁 Duplicates", f"{kpis['duplicate_count']:,}", "#A78BFA"),
        ("🔍 Manual Review", f"{kpis['manual_review_count']:,}", "#60A5FA"),
    ]
    for col, (label, val, color) in zip(kpi_cols2, kpi_data2):
        col.markdown(_kpi_card(label, val, color), unsafe_allow_html=True)

    st.divider()

    # ── Global Filters ─────────────────────────────────────────────────────
    with st.expander("🔧 Filters & Search", expanded=False):
        gf1, gf2, gf3, gf4 = st.columns(4)

        all_statuses = ["All"] + sorted(master_df["status"].dropna().unique().tolist()) if "status" in master_df.columns else ["All"]
        filter_status = gf1.selectbox("Status", all_statuses, key="res_status_filter")

        all_vendors = ["All"] + sorted(
            master_df["vendor_name"].dropna().unique().tolist()
        )[:200] if "vendor_name" in master_df.columns else ["All"]
        filter_vendor = gf2.selectbox("Vendor", all_vendors, key="res_vendor_filter")

        filter_conf = gf3.selectbox(
            "Confidence", ["All", CONFIDENCE_HIGH, CONFIDENCE_MEDIUM, CONFIDENCE_LOW],
            key="res_conf_filter"
        )

        search_query = gf4.text_input(
            "🔍 Search (Invoice / Vendor / GSTIN)", key="res_search"
        )

    # Apply filters
    filtered_df = master_df.copy()

    if filter_status != "All" and "status" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["status"] == filter_status]
    if filter_vendor != "All" and "vendor_name" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["vendor_name"] == filter_vendor]
    if filter_conf != "All" and "confidence" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["confidence"] == filter_conf]
    if search_query:
        sq = search_query.lower()
        mask = (
            filtered_df.apply(lambda r: sq in str(r).lower(), axis=1)
        )
        filtered_df = filtered_df[mask]

    st.markdown(f"**Showing {len(filtered_df):,} of {len(master_df):,} records**")

    # ── Results Tabs ───────────────────────────────────────────────────────
    status_filter_options = master_df["status"].value_counts().to_dict() if "status" in master_df.columns else {}

    tab_all, tab_match, tab_books, tab_gstr, tab_gst_diff, tab_dup, tab_manual, tab_vendor = st.tabs([
        f"📋 All ({len(filtered_df):,})",
        f"✅ Perfect Match ({status_filter_options.get(STATUS_PERFECT_MATCH, 0)})",
        f"📚 Missing Books ({status_filter_options.get(STATUS_MISSING_BOOKS, 0)})",
        f"🏛️ Missing GSTR-2B ({status_filter_options.get(STATUS_MISSING_GSTR2B, 0)})",
        f"💸 GST Diff ({status_filter_options.get(STATUS_GST_DIFF, 0)})",
        f"🔁 Duplicates",
        f"🔍 Manual Review ({status_filter_options.get(STATUS_FUZZY_MATCH, 0) + status_filter_options.get(STATUS_MANUAL_REVIEW, 0)})",
        "🏢 Vendor Summary",
    ])

    def _styled_table(df: pd.DataFrame, key: str):
        """Display a styled DataFrame with download option."""
        if df.empty:
            st.info("No records for this category.")
            return
        display_cols = [c for c in df.columns if not c.startswith("_")]
        st.dataframe(df[display_cols].head(1000), use_container_width=True, height=400)
        csv = df[display_cols].to_csv(index=False).encode("utf-8")
        st.download_button(
            f"📥 Download CSV",
            data=csv,
            file_name=f"{key}.csv",
            mime="text/csv",
            key=f"dl_results_{key}",
        )

    def _get_status_df(status_val: str) -> pd.DataFrame:
        if "status" not in master_df.columns:
            return pd.DataFrame()
        return master_df[master_df["status"] == status_val].copy()

    with tab_all:
        _styled_table(filtered_df, "all_records")

    with tab_match:
        _styled_table(_get_status_df(STATUS_PERFECT_MATCH), "perfect_match")

    with tab_books:
        mb_df = _get_status_df(STATUS_MISSING_BOOKS)
        if not mb_df.empty:
            st.info(
                f"💡 **{len(mb_df)} invoices** found in GSTR-2B but not in Purchase Register. "
                "Action: Book purchase entries."
            )
        _styled_table(mb_df, "missing_books")

    with tab_gstr:
        mg_df = _get_status_df(STATUS_MISSING_GSTR2B)
        if not mg_df.empty:
            st.info(
                f"💡 **{len(mg_df)} invoices** found in Purchase Register but not in GSTR-2B. "
                "Action: Follow up with vendors."
            )
        _styled_table(mg_df, "missing_gstr2b")

    with tab_gst_diff:
        gd_df = _get_status_df(STATUS_GST_DIFF)
        if "gst_difference" in gd_df.columns and not gd_df.empty:
            total_diff = gd_df["gst_difference"].sum()
            st.metric("Total GST Difference", format_currency(total_diff))
        _styled_table(gd_df, "gst_difference")

    with tab_dup:
        dup_pr = recon_results.get("pr_duplicates", pd.DataFrame())
        dup_gstr = recon_results.get("gstr2b_duplicates", pd.DataFrame())
        st.subheader(f"PR Duplicates ({len(dup_pr)})")
        _styled_table(dup_pr, "pr_duplicates")
        st.subheader(f"GSTR-2B Duplicates ({len(dup_gstr)})")
        _styled_table(dup_gstr, "gstr2b_duplicates")

    with tab_manual:
        manual_df = pd.concat([
            _get_status_df(STATUS_FUZZY_MATCH),
            _get_status_df(STATUS_MANUAL_REVIEW),
        ], ignore_index=True)
        if not manual_df.empty:
            st.info(
                f"💡 **{len(manual_df)} records** require manual review. "
                "Verify each record and confirm or reject the match."
            )
        _styled_table(manual_df, "manual_review")

    with tab_vendor:
        vendor_summary = get_vendor_summary(master_df)
        if not vendor_summary.empty:
            st.dataframe(vendor_summary.head(200), use_container_width=True, height=400)
            csv = vendor_summary.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download Vendor Summary (CSV)",
                data=csv,
                file_name="vendor_summary.csv",
                mime="text/csv",
                key="dl_vendor_summary",
            )

    # ── Export Panel ───────────────────────────────────────────────────────
    st.divider()
    st.subheader("📤 Export Reports")
    exp1, exp2, exp3 = st.columns(3)

    if exp1.button("📊 Export Excel (Full)", use_container_width=True, key="exp_excel_full"):
        with st.spinner("Generating Excel report…"):
            try:
                from modules.reports import export_excel
                excel_bytes = export_excel(master_df, recon_results)
                st.download_button(
                    "⬇️ Download Excel Report",
                    data=excel_bytes,
                    file_name=f"GST_Reconciliation_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel_full",
                )
                log_event("EXPORT", "Excel report generated")
            except Exception as e:
                st.error(f"Excel export failed: {e}")

    if exp2.button("📄 Export PDF Report", use_container_width=True, key="exp_pdf"):
        with st.spinner("Generating PDF report…"):
            try:
                from modules.reports import export_pdf
                pdf_bytes = export_pdf(master_df, recon_results)
                st.download_button(
                    "⬇️ Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"GST_Reconciliation_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    key="dl_pdf",
                )
                log_event("EXPORT", "PDF report generated")
            except Exception as e:
                st.error(f"PDF export failed: {e}")

    if exp3.button("📋 Export CSV", use_container_width=True, key="exp_csv"):
        csv = master_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download CSV",
            data=csv,
            file_name=f"GST_Reconciliation_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            key="dl_csv_master",
        )
        log_event("EXPORT", "CSV export generated")
