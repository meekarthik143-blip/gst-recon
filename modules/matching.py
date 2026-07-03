"""
GST Input Reconciliation System – Enterprise Edition
Reconciliation Matching Engine
Prepared & Developed by Karthik LVN

Provides:
  - 5-tier matching pipeline (exact → fuzzy)
  - Duplicate detection within each source
  - Confidence score assignment
  - Reconciliation runner with progress callback
  - Streamlit reconciliation control page
"""

import re
import time
from typing import Callable, Optional

import numpy as np
import pandas as pd
import streamlit as st

from modules.utils import setup_logging, safe_float, round_gst

logger = setup_logging()

# ---------------------------------------------------------------------------
# Status Constants
# ---------------------------------------------------------------------------

STATUS_PERFECT_MATCH = "Perfect Match"
STATUS_MISSING_BOOKS = "Missing in Books"
STATUS_MISSING_GSTR2B = "Missing in GSTR-2B"
STATUS_GST_DIFF = "GST Difference"
STATUS_TAXABLE_DIFF = "Taxable Difference"
STATUS_INV_VALUE_DIFF = "Invoice Value Difference"
STATUS_DATE_DIFF = "Date Difference"
STATUS_DUPLICATE = "Duplicate"
STATUS_GSTIN_MISMATCH = "GSTIN Mismatch"
STATUS_VENDOR_DIFF = "Vendor Name Difference"
STATUS_MANUAL_REVIEW = "Manual Review"
STATUS_FUZZY_MATCH = "Fuzzy Match"
STATUS_BLOCKED_ITC = "Blocked ITC"
STATUS_REVERSE_CHARGE = "Reverse Charge"
STATUS_IMPORT = "Import"
STATUS_SEZ = "SEZ"

CONFIDENCE_HIGH = "High"
CONFIDENCE_MEDIUM = "Medium"
CONFIDENCE_LOW = "Low"

# Columns that must exist in the standardized DataFrames
REQUIRED_COLS = [
    "vendor_name", "gstin", "invoice_number", "invoice_date",
    "taxable_value", "cgst", "sgst", "igst", "cess", "total_gst",
    "invoice_value",
]


# ---------------------------------------------------------------------------
# Normalization Helpers
# ---------------------------------------------------------------------------

def normalize_for_match(s: str) -> str:
    """
    Normalize a string for matching: lowercase, remove all non-alphanumeric chars.

    Args:
        s: Input string.

    Returns:
        Normalized string.
    """
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _safe_col(df: pd.DataFrame, col: str, default="") -> pd.Series:
    """Return column as string series or default if column missing."""
    if col in df.columns:
        return df[col].fillna(default).astype(str)
    return pd.Series([default] * len(df), index=df.index)


def _add_norm_keys(df: pd.DataFrame) -> pd.DataFrame:
    """Add normalized key columns for fast matching."""
    df = df.copy().reset_index(drop=True)
    df["_norm_gstin"] = _safe_col(df, "gstin").apply(normalize_for_match)
    df["_norm_inv"] = _safe_col(df, "invoice_number").apply(normalize_for_match)
    df["_norm_date"] = _safe_col(df, "invoice_date").apply(normalize_for_match)

    # Use explicit column check instead of df.get() to avoid pandas 3.x reindex errors
    if "total_gst" in df.columns:
        df["_total_gst_r"] = pd.to_numeric(df["total_gst"], errors="coerce").fillna(0.0).round(1)
    else:
        df["_total_gst_r"] = 0.0

    if "taxable_value" in df.columns:
        df["_taxable_r"] = pd.to_numeric(df["taxable_value"], errors="coerce").fillna(0.0).round(1)
    else:
        df["_taxable_r"] = 0.0

    return df


# ---------------------------------------------------------------------------
# Duplicate Detection
# ---------------------------------------------------------------------------

def detect_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mark duplicate invoices within the same source DataFrame.

    Duplicates are identified by GSTIN + invoice_number combination.

    Args:
        df: Source DataFrame (PR or GSTR-2B) with standard columns.

    Returns:
        DataFrame with added 'is_duplicate' boolean column.
    """
    df = df.copy()
    if "gstin" in df.columns and "invoice_number" in df.columns:
        key = df["gstin"].astype(str) + "|" + df["invoice_number"].astype(str)
        df["is_duplicate"] = key.duplicated(keep=False)
    else:
        df["is_duplicate"] = False
    return df


# ---------------------------------------------------------------------------
# Matching Tiers
# ---------------------------------------------------------------------------

def _split_on_key(
    pr_df: pd.DataFrame,
    gstr2b_df: pd.DataFrame,
    key_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Inner-join on a common key column, then split into matched / unmatched.

    Returns:
        (matched_df, pr_unmatched, gstr2b_unmatched)
    """
    pr_df = pr_df.reset_index(drop=True)
    gstr2b_df = gstr2b_df.reset_index(drop=True)

    pr_df["_pr_idx"] = pr_df.index
    gstr2b_df["_gstr_idx"] = gstr2b_df.index

    merged = pd.merge(
        pr_df,
        gstr2b_df,
        on=key_col,
        suffixes=("_pr", "_gstr2b"),
        how="inner",
    )

    matched_pr_idx = set(merged["_pr_idx"].tolist())
    matched_gstr_idx = set(merged["_gstr_idx"].tolist())

    pr_unmatched = pr_df[~pr_df["_pr_idx"].isin(matched_pr_idx)].drop(columns=["_pr_idx"])
    gstr2b_unmatched = gstr2b_df[~gstr2b_df["_gstr_idx"].isin(matched_gstr_idx)].drop(
        columns=["_gstr_idx"]
    )
    matched = merged.drop(columns=["_pr_idx", "_gstr_idx"])

    return matched, pr_unmatched.reset_index(drop=True), gstr2b_unmatched.reset_index(drop=True)


def exact_match_tier1(
    pr_df: pd.DataFrame, gstr2b_df: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Tier 1: Match on normalized GSTIN + normalized Invoice Number.

    Args:
        pr_df:     Unmatched Purchase Register rows.
        gstr2b_df: Unmatched GSTR-2B rows.

    Returns:
        (matched_df, pr_unmatched, gstr2b_unmatched)
    """
    pr_df = _add_norm_keys(pr_df)
    gstr2b_df = _add_norm_keys(gstr2b_df)

    pr_df["_t1_key"] = pr_df["_norm_gstin"] + "||" + pr_df["_norm_inv"]
    gstr2b_df["_t1_key"] = gstr2b_df["_norm_gstin"] + "||" + gstr2b_df["_norm_inv"]

    matched, pr_um, gstr_um = _split_on_key(pr_df, gstr2b_df, "_t1_key")

    if not matched.empty:
        matched["match_tier"] = 1
        matched["match_key"] = "GSTIN + Invoice Number"

    # Clean up temp columns
    for col in ["_norm_gstin", "_norm_inv", "_norm_date", "_total_gst_r", "_taxable_r", "_t1_key"]:
        for df_target in [matched, pr_um, gstr_um]:
            if col in df_target.columns:
                df_target.drop(columns=[col], inplace=True, errors="ignore")

    return matched, pr_um, gstr_um


def exact_match_tier2(
    pr_df: pd.DataFrame, gstr2b_df: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Tier 2: GSTIN + Invoice Number + Date."""
    pr_df = _add_norm_keys(pr_df)
    gstr2b_df = _add_norm_keys(gstr2b_df)

    pr_df["_t2_key"] = pr_df["_norm_gstin"] + "||" + pr_df["_norm_inv"] + "||" + pr_df["_norm_date"]
    gstr2b_df["_t2_key"] = (
        gstr2b_df["_norm_gstin"] + "||" + gstr2b_df["_norm_inv"] + "||" + gstr2b_df["_norm_date"]
    )

    matched, pr_um, gstr_um = _split_on_key(pr_df, gstr2b_df, "_t2_key")
    if not matched.empty:
        matched["match_tier"] = 2
        matched["match_key"] = "GSTIN + Invoice Number + Date"

    for col in ["_norm_gstin", "_norm_inv", "_norm_date", "_total_gst_r", "_taxable_r", "_t2_key"]:
        for df_target in [matched, pr_um, gstr_um]:
            df_target.drop(columns=[col], inplace=True, errors="ignore")

    return matched, pr_um, gstr_um


def exact_match_tier3(
    pr_df: pd.DataFrame, gstr2b_df: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Tier 3: GSTIN + Invoice Number + Total GST (rounded to 1 decimal)."""
    pr_df = _add_norm_keys(pr_df)
    gstr2b_df = _add_norm_keys(gstr2b_df)

    pr_df["_t3_key"] = (
        pr_df["_norm_gstin"] + "||" + pr_df["_norm_inv"] + "||" + pr_df["_total_gst_r"].astype(str)
    )
    gstr2b_df["_t3_key"] = (
        gstr2b_df["_norm_gstin"]
        + "||"
        + gstr2b_df["_norm_inv"]
        + "||"
        + gstr2b_df["_total_gst_r"].astype(str)
    )

    matched, pr_um, gstr_um = _split_on_key(pr_df, gstr2b_df, "_t3_key")
    if not matched.empty:
        matched["match_tier"] = 3
        matched["match_key"] = "GSTIN + Invoice Number + GST Amount"

    for col in ["_norm_gstin", "_norm_inv", "_norm_date", "_total_gst_r", "_taxable_r", "_t3_key"]:
        for df_target in [matched, pr_um, gstr_um]:
            df_target.drop(columns=[col], inplace=True, errors="ignore")

    return matched, pr_um, gstr_um


def exact_match_tier4(
    pr_df: pd.DataFrame, gstr2b_df: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Tier 4: GSTIN + Invoice Number + Taxable Value (rounded to 1 decimal)."""
    pr_df = _add_norm_keys(pr_df)
    gstr2b_df = _add_norm_keys(gstr2b_df)

    pr_df["_t4_key"] = (
        pr_df["_norm_gstin"] + "||" + pr_df["_norm_inv"] + "||" + pr_df["_taxable_r"].astype(str)
    )
    gstr2b_df["_t4_key"] = (
        gstr2b_df["_norm_gstin"]
        + "||"
        + gstr2b_df["_norm_inv"]
        + "||"
        + gstr2b_df["_taxable_r"].astype(str)
    )

    matched, pr_um, gstr_um = _split_on_key(pr_df, gstr2b_df, "_t4_key")
    if not matched.empty:
        matched["match_tier"] = 4
        matched["match_key"] = "GSTIN + Invoice Number + Taxable Value"

    for col in ["_norm_gstin", "_norm_inv", "_norm_date", "_total_gst_r", "_taxable_r", "_t4_key"]:
        for df_target in [matched, pr_um, gstr_um]:
            df_target.drop(columns=[col], inplace=True, errors="ignore")

    return matched, pr_um, gstr_um


# ---------------------------------------------------------------------------
# Fuzzy Matching
# ---------------------------------------------------------------------------

def assign_confidence(score: float) -> str:
    """
    Assign a confidence label based on a normalized similarity score (0-1).

    Args:
        score: Match score in [0.0, 1.0].

    Returns:
        'High', 'Medium', or 'Low'.
    """
    if score >= 0.95:
        return CONFIDENCE_HIGH
    elif score >= 0.80:
        return CONFIDENCE_MEDIUM
    else:
        return CONFIDENCE_LOW


def fuzzy_match(
    pr_df: pd.DataFrame,
    gstr2b_df: pd.DataFrame,
    threshold: float = 0.85,
) -> pd.DataFrame:
    """
    Perform fuzzy matching between unmatched PR and GSTR-2B rows.

    Uses RapidFuzz's process.extractOne for batch-optimized matching.
    Combined match key: vendor_name + " " + invoice_number

    Args:
        pr_df:     Unmatched Purchase Register rows.
        gstr2b_df: Unmatched GSTR-2B rows.
        threshold: Minimum similarity score (0.0–1.0).

    Returns:
        DataFrame of fuzzy match candidates with 'fuzzy_score' and 'confidence' columns.
    """
    if pr_df.empty or gstr2b_df.empty:
        return pd.DataFrame()

    try:
        from rapidfuzz import process as rfprocess, fuzz as rffuzz
    except ImportError:
        logger.warning("RapidFuzz not installed — skipping fuzzy matching.")
        return pd.DataFrame()

    score_cutoff = threshold * 100  # rapidfuzz uses 0–100

    # Build combined search keys for GSTR-2B
    gstr2b_df = gstr2b_df.reset_index(drop=True)
    gstr2b_keys = (
        _safe_col(gstr2b_df, "vendor_name").apply(normalize_for_match)
        + " "
        + _safe_col(gstr2b_df, "invoice_number").apply(normalize_for_match)
    ).tolist()

    pr_df = pr_df.reset_index(drop=True)
    results = []

    for pr_idx, pr_row in pr_df.iterrows():
        pr_key = normalize_for_match(str(pr_row.get("vendor_name", ""))) + " " + normalize_for_match(
            str(pr_row.get("invoice_number", ""))
        )

        if not pr_key.strip():
            continue

        match = rfprocess.extractOne(
            pr_key,
            gstr2b_keys,
            scorer=rffuzz.WRatio,
            score_cutoff=score_cutoff,
        )

        if match is not None:
            matched_key, score, gstr_idx = match
            fuzzy_score = score / 100.0

            row = {}
            # PR columns with _pr suffix
            for col in pr_df.columns:
                row[f"{col}_pr"] = pr_row.get(col, "")
            # GSTR2B columns with _gstr2b suffix
            gstr_row = gstr2b_df.iloc[gstr_idx]
            for col in gstr2b_df.columns:
                row[f"{col}_gstr2b"] = gstr_row.get(col, "")

            row["fuzzy_score"] = round(fuzzy_score, 4)
            row["confidence"] = assign_confidence(fuzzy_score)
            row["match_tier"] = 5
            row["match_key"] = f"Fuzzy (score={score:.1f}%)"
            results.append(row)

    if results:
        return pd.DataFrame(results)
    return pd.DataFrame()


# ---------------------------------------------------------------------------
# Main Reconciliation Runner
# ---------------------------------------------------------------------------

def run_reconciliation(
    pr_df: pd.DataFrame,
    gstr2b_df: pd.DataFrame,
    match_threshold: float = 0.85,
    run_tier1: bool = True,
    run_tier2: bool = True,
    run_tier3: bool = True,
    run_tier4: bool = True,
    run_fuzzy: bool = True,
    progress_callback: Optional[Callable[[int, str], None]] = None,
) -> dict:
    """
    Orchestrate the full reconciliation pipeline.

    Processing order:
      1. Detect duplicates in PR and GSTR-2B
      2. Tier 1 matching (GSTIN + Invoice No)
      3. Tier 2 matching on remaining unmatched
      4. Tier 3 matching
      5. Tier 4 matching
      6. Fuzzy matching on remaining unmatched
      7. Classify remaining unmatched as Missing in Books / Missing in GSTR-2B

    Args:
        pr_df:             Cleaned, standardized Purchase Register DataFrame.
        gstr2b_df:         Cleaned, standardized GSTR-2B DataFrame.
        match_threshold:   Fuzzy match threshold (0.0–1.0).
        run_tier1..4:      Enable/disable individual exact-match tiers.
        run_fuzzy:         Enable/disable fuzzy matching.
        progress_callback: Optional callable(percent: int, message: str).

    Returns:
        Dict with keys:
            matched, missing_in_books, missing_in_gstr2b,
            fuzzy_candidates, pr_duplicates, gstr2b_duplicates, stats
    """

    def _progress(pct: int, msg: str):
        if progress_callback:
            progress_callback(pct, msg)
        logger.info(f"[{pct}%] {msg}")

    _progress(2, "Starting reconciliation…")

    # ── Detect duplicates ──────────────────────────────────────────────────
    pr_with_dup = detect_duplicates(pr_df)
    gstr_with_dup = detect_duplicates(gstr2b_df)

    pr_duplicates = pr_with_dup[pr_with_dup["is_duplicate"]].copy()
    gstr2b_duplicates = gstr_with_dup[gstr_with_dup["is_duplicate"]].copy()

    # Work with full DataFrames (including duplicates) for matching
    pr_remaining = pr_df.copy()
    gstr2b_remaining = gstr2b_df.copy()

    all_matched: list[pd.DataFrame] = []

    _progress(10, "Running Tier 1: GSTIN + Invoice Number…")

    # ── Tier 1 ─────────────────────────────────────────────────────────────
    if run_tier1 and not pr_remaining.empty and not gstr2b_remaining.empty:
        t1_matched, pr_remaining, gstr2b_remaining = exact_match_tier1(
            pr_remaining, gstr2b_remaining
        )
        if not t1_matched.empty:
            all_matched.append(t1_matched)
        _progress(25, f"Tier 1 complete — {len(t1_matched)} matched")

    # ── Tier 2 ─────────────────────────────────────────────────────────────
    if run_tier2 and not pr_remaining.empty and not gstr2b_remaining.empty:
        _progress(30, "Running Tier 2: GSTIN + Invoice Number + Date…")
        t2_matched, pr_remaining, gstr2b_remaining = exact_match_tier2(
            pr_remaining, gstr2b_remaining
        )
        if not t2_matched.empty:
            all_matched.append(t2_matched)
        _progress(45, f"Tier 2 complete — {len(t2_matched)} matched")

    # ── Tier 3 ─────────────────────────────────────────────────────────────
    if run_tier3 and not pr_remaining.empty and not gstr2b_remaining.empty:
        _progress(48, "Running Tier 3: GSTIN + Invoice Number + GST Amount…")
        t3_matched, pr_remaining, gstr2b_remaining = exact_match_tier3(
            pr_remaining, gstr2b_remaining
        )
        if not t3_matched.empty:
            all_matched.append(t3_matched)
        _progress(60, f"Tier 3 complete — {len(t3_matched)} matched")

    # ── Tier 4 ─────────────────────────────────────────────────────────────
    if run_tier4 and not pr_remaining.empty and not gstr2b_remaining.empty:
        _progress(62, "Running Tier 4: GSTIN + Invoice Number + Taxable Value…")
        t4_matched, pr_remaining, gstr2b_remaining = exact_match_tier4(
            pr_remaining, gstr2b_remaining
        )
        if not t4_matched.empty:
            all_matched.append(t4_matched)
        _progress(73, f"Tier 4 complete — {len(t4_matched)} matched")

    # ── Fuzzy Matching ─────────────────────────────────────────────────────
    fuzzy_candidates = pd.DataFrame()
    if run_fuzzy and not pr_remaining.empty and not gstr2b_remaining.empty:
        _progress(75, f"Running Tier 5 Fuzzy Match (threshold={match_threshold * 100:.0f}%)…")
        fuzzy_candidates = fuzzy_match(pr_remaining, gstr2b_remaining, threshold=match_threshold)
        _progress(90, f"Fuzzy matching complete — {len(fuzzy_candidates)} candidates")

    # ── Combine all exact matches ──────────────────────────────────────────
    matched_df = pd.concat(all_matched, ignore_index=True) if all_matched else pd.DataFrame()

    # ── Missing categories ─────────────────────────────────────────────────
    missing_in_books = gstr2b_remaining.copy()   # In GSTR-2B but not in PR
    missing_in_gstr2b = pr_remaining.copy()       # In PR but not in GSTR-2B

    _progress(95, "Computing statistics…")

    total_pr = len(pr_df)
    total_gstr = len(gstr2b_df)
    total_matched = len(matched_df)

    stats = {
        "total_pr": total_pr,
        "total_gstr2b": total_gstr,
        "total_matched": total_matched,
        "missing_in_books_count": len(missing_in_books),
        "missing_in_gstr2b_count": len(missing_in_gstr2b),
        "fuzzy_candidates_count": len(fuzzy_candidates),
        "pr_duplicates_count": len(pr_duplicates),
        "gstr2b_duplicates_count": len(gstr2b_duplicates),
        "match_rate": round(total_matched / max(total_pr, 1) * 100, 2),
    }

    _progress(100, "Reconciliation complete!")

    return {
        "matched": matched_df,
        "missing_in_books": missing_in_books,
        "missing_in_gstr2b": missing_in_gstr2b,
        "fuzzy_candidates": fuzzy_candidates,
        "pr_duplicates": pr_duplicates,
        "gstr2b_duplicates": gstr2b_duplicates,
        "stats": stats,
    }


# ---------------------------------------------------------------------------
# Streamlit Reconciliation Control Page
# ---------------------------------------------------------------------------

def render_reconciliation_page() -> None:
    """Render the reconciliation configuration and run page."""

    st.markdown(
        "<h2 style='color:#00D4FF;'>⚙️ Reconciliation Engine</h2>",
        unsafe_allow_html=True,
    )

    pr_df = st.session_state.get("pr_mapped")
    gstr2b_df = st.session_state.get("gstr2b_mapped")

    if pr_df is None or gstr2b_df is None:
        st.warning("⚠️ Please complete **Column Mapping** for both files before running reconciliation.")
        if st.button("Go to Column Mapping"):
            st.session_state["current_page"] = "Column Mapping"
            st.rerun()
        return

    # ── Configuration panel ────────────────────────────────────────────────
    with st.expander("🔧 Reconciliation Configuration", expanded=True):
        cfg1, cfg2 = st.columns(2)

        match_pct = cfg1.slider(
            "Fuzzy Match Threshold (%)",
            min_value=80,
            max_value=100,
            value=st.session_state.get("app_settings", {}).get("default_match_percentage", 85),
            step=5,
            key="recon_match_pct",
        )
        match_threshold = match_pct / 100.0

        cfg2.markdown("**Matching Tiers to Run**")
        run_t1 = cfg2.checkbox("Tier 1: GSTIN + Invoice No", value=True, key="tier1")
        run_t2 = cfg2.checkbox("Tier 2: + Date", value=True, key="tier2")
        run_t3 = cfg2.checkbox("Tier 3: + GST Amount", value=True, key="tier3")
        run_t4 = cfg2.checkbox("Tier 4: + Taxable Value", value=True, key="tier4")
        run_fuzzy = cfg2.checkbox("Tier 5: Fuzzy Matching", value=True, key="tier5")

    # ── Data summary ───────────────────────────────────────────────────────
    dc1, dc2 = st.columns(2)
    dc1.metric("📋 Purchase Register Rows", f"{len(pr_df):,}")
    dc2.metric("🏛️ GSTR-2B Rows", f"{len(gstr2b_df):,}")

    st.divider()

    # ── Run button ─────────────────────────────────────────────────────────
    if st.button(
        "🚀 Run Reconciliation",
        type="primary",
        use_container_width=True,
        key="run_recon_btn",
    ):
        progress_bar = st.progress(0, text="Initializing…")
        status_text = st.empty()

        def callback(pct: int, msg: str):
            progress_bar.progress(pct, text=msg)
            status_text.markdown(
                f"<span style='color:#94A3B8;'>{msg}</span>", unsafe_allow_html=True
            )

        start_time = time.time()

        with st.spinner("Running reconciliation…"):
            results = run_reconciliation(
                pr_df=pr_df,
                gstr2b_df=gstr2b_df,
                match_threshold=match_threshold,
                run_tier1=run_t1,
                run_tier2=run_t2,
                run_tier3=run_t3,
                run_tier4=run_t4,
                run_fuzzy=run_fuzzy,
                progress_callback=callback,
            )

        elapsed = time.time() - start_time
        progress_bar.progress(100, text="Complete!")

        st.session_state["recon_results"] = results
        st.session_state["recon_elapsed"] = elapsed

        # Build master_df so dashboard KPIs update
        frames = []
        if not results.get("matched", pd.DataFrame()).empty:
            frames.append(results["matched"].assign(status="Perfect Match"))
        if not results.get("missing_in_books", pd.DataFrame()).empty:
            frames.append(results["missing_in_books"].assign(status="Missing in Books"))
        if not results.get("missing_in_gstr2b", pd.DataFrame()).empty:
            frames.append(results["missing_in_gstr2b"].assign(status="Missing in GSTR-2B"))
        if not results.get("fuzzy_candidates", pd.DataFrame()).empty:
            frames.append(results["fuzzy_candidates"].assign(status="Fuzzy Match"))
        if frames:
            st.session_state["master_df"] = pd.concat(frames, ignore_index=True)
        else:
            st.session_state["master_df"] = pd.DataFrame()

        from modules.audit import log_event
        log_event(
            "PROCESS",
            f"Reconciliation completed: {results['stats']['total_matched']} matched, "
            f"{results['stats']['missing_in_books_count']} missing in books, "
            f"{results['stats']['missing_in_gstr2b_count']} missing in GSTR-2B. "
            f"Time: {elapsed:.1f}s",
        )

        st.rerun()

    # ── Results (if available) ─────────────────────────────────────────────
    results = st.session_state.get("recon_results")
    if results:
        stats = results["stats"]
        elapsed = st.session_state.get("recon_elapsed", 0)

        st.success(
            f"✅ Reconciliation completed in **{elapsed:.1f}s** — "
            f"Match rate: **{stats['match_rate']:.1f}%**"
        )

        # KPI cards
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("✅ Matched", f"{stats['total_matched']:,}")
        k2.metric("📚 Missing in Books", f"{stats['missing_in_books_count']:,}")
        k3.metric("🏛️ Missing in GSTR-2B", f"{stats['missing_in_gstr2b_count']:,}")
        k4.metric("🔍 Fuzzy Candidates", f"{stats['fuzzy_candidates_count']:,}")
        k5.metric("🔁 PR Duplicates", f"{stats['pr_duplicates_count']:,}")
        k6.metric("🔁 GSTR-2B Duplicates", f"{stats['gstr2b_duplicates_count']:,}")

        st.divider()

        # Results tabs
        tab_matched, tab_books, tab_gstr, tab_fuzzy, tab_dup = st.tabs([
            f"✅ Matched ({stats['total_matched']})",
            f"📚 Missing in Books ({stats['missing_in_books_count']})",
            f"🏛️ Missing in GSTR-2B ({stats['missing_in_gstr2b_count']})",
            f"🔍 Fuzzy Candidates ({stats['fuzzy_candidates_count']})",
            f"🔁 Duplicates (PR:{stats['pr_duplicates_count']} / GSTR:{stats['gstr2b_duplicates_count']})",
        ])

        def _show_table_with_download(df: pd.DataFrame, label: str, key: str):
            if df.empty:
                st.info(f"No {label} records.")
                return
            st.dataframe(df.head(500), use_container_width=True)
            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button(
                f"📥 Download {label} (CSV)",
                data=csv,
                file_name=f"{label.lower().replace(' ', '_')}.csv",
                mime="text/csv",
                key=f"dl_{key}",
            )

        with tab_matched:
            _show_table_with_download(results["matched"], "Matched", "matched")
        with tab_books:
            _show_table_with_download(results["missing_in_books"], "Missing in Books", "miss_books")
        with tab_gstr:
            _show_table_with_download(results["missing_in_gstr2b"], "Missing in GSTR-2B", "miss_gstr")
        with tab_fuzzy:
            _show_table_with_download(results["fuzzy_candidates"], "Fuzzy Candidates", "fuzzy")
        with tab_dup:
            st.subheader("PR Duplicates")
            _show_table_with_download(results["pr_duplicates"], "PR Duplicates", "pr_dup")
            st.subheader("GSTR-2B Duplicates")
            _show_table_with_download(results["gstr2b_duplicates"], "GSTR2B Duplicates", "gst_dup")

        st.divider()
        if st.button(
            "View Full Report & Download Excel",
            type="primary",
            use_container_width=True,
            key="goto_recon_results",
        ):
            st.session_state["current_page"] = "Reports"
            st.rerun()
