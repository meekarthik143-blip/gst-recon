"""
GST Input Reconciliation System – Enterprise Edition
Data Cleaning Module
Prepared & Developed by Karthik LVN

Provides:
  - Automated data cleaning pipeline for Purchase Register and GSTR-2B DataFrames
  - Cleaning steps: blank rows, duplicates, GSTIN normalization, invoice cleaning,
    vendor name cleaning, date standardization, GST rounding, hidden char removal
  - Summary reporting of cleaning results
"""

import re
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st

from modules.utils import (
    validate_gstin,
    parse_date,
    standardize_date,
    clean_invoice_number,
    clean_vendor_name,
    normalize_gstin,
    round_gst,
    safe_float,
    setup_logging,
)
from modules.audit import log_event

logger = setup_logging()

# Standard GST value columns that should be treated as numeric
GST_NUMERIC_COLS = [
    "taxable_value",
    "cgst",
    "sgst",
    "igst",
    "cess",
    "total_gst",
    "invoice_value",
]


# ---------------------------------------------------------------------------
# Individual Cleaning Steps
# ---------------------------------------------------------------------------

def remove_blank_rows(df: pd.DataFrame, key_cols: Optional[list[str]] = None) -> tuple[pd.DataFrame, int]:
    """
    Drop rows where all specified key columns are NaN/empty.

    Args:
        df:       Input DataFrame.
        key_cols: Columns to check; defaults to all columns.

    Returns:
        (cleaned_df, number_of_rows_removed)
    """
    before = len(df)
    if key_cols:
        # Drop rows where ALL key_cols are NaN
        existing_key_cols = [c for c in key_cols if c in df.columns]
        if existing_key_cols:
            df = df.dropna(subset=existing_key_cols, how="all")
    else:
        df = df.dropna(how="all")

    # Also drop rows that are all-empty strings — use apply with explicit bool check
    def _all_empty(row):
        for v in row:
            s = str(v).strip().lower()
            if s not in ("", "nan", "none", "nat"):
                return False
        return True

    str_mask = df.apply(_all_empty, axis=1)
    df = df[~str_mask]
    df = df.reset_index(drop=True)

    removed = before - len(df)
    logger.info(f"Blank rows removed: {removed}")
    return df, removed


def remove_duplicate_rows(
    df: pd.DataFrame, subset_cols: Optional[list[str]] = None
) -> tuple[pd.DataFrame, int]:
    """
    Remove exact duplicate rows.

    Args:
        df:          Input DataFrame.
        subset_cols: Columns to use for duplicate detection (None = all columns).

    Returns:
        (deduplicated_df, number_of_duplicates_removed)
    """
    before = len(df)
    existing_cols = [c for c in (subset_cols or []) if c in df.columns] or None
    df = df.drop_duplicates(subset=existing_cols, keep="first")
    df = df.reset_index(drop=True)
    removed = before - len(df)
    logger.info(f"Duplicate rows removed: {removed}")
    return df, removed


def clean_gstin_column(
    df: pd.DataFrame, col: str
) -> tuple[pd.DataFrame, int]:
    """
    Normalize GSTIN column: uppercase, strip whitespace, remove spaces/dashes.
    Track how many values are invalid after normalization.

    Args:
        df:  Input DataFrame.
        col: Name of the GSTIN column.

    Returns:
        (df_with_cleaned_gstin, invalid_count)
    """
    if col not in df.columns:
        return df, 0

    df = df.copy()
    df[col] = df[col].apply(lambda x: normalize_gstin(x) if pd.notna(x) else "")

    # Count invalid GSTINs (non-empty but failing format check)
    invalid_count = df[col].apply(
        lambda x: bool(x) and not validate_gstin(x)
    ).sum()

    logger.info(f"GSTIN column '{col}' cleaned. Invalid count: {invalid_count}")
    return df, int(invalid_count)


def clean_invoice_column(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """
    Normalize invoice number column: remove special chars, uppercase.

    Args:
        df:  Input DataFrame.
        col: Name of the invoice number column.

    Returns:
        DataFrame with cleaned invoice column.
    """
    if col not in df.columns:
        return df
    df = df.copy()
    df[col] = df[col].apply(lambda x: clean_invoice_number(x) if pd.notna(x) else "")
    return df


def clean_vendor_column(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """
    Normalize vendor name column: strip, title-case, collapse spaces.

    Args:
        df:  Input DataFrame.
        col: Name of the vendor name column.

    Returns:
        DataFrame with cleaned vendor column.
    """
    if col not in df.columns:
        return df
    df = df.copy()
    df[col] = df[col].apply(lambda x: clean_vendor_name(x) if pd.notna(x) else "")
    return df


def standardize_date_column(
    df: pd.DataFrame, col: str
) -> tuple[pd.DataFrame, int]:
    """
    Parse and standardize date column to DD-MM-YYYY string format.

    Args:
        df:  Input DataFrame.
        col: Name of the date column.

    Returns:
        (df_with_standardized_dates, number_of_failed_parses)
    """
    if col not in df.columns:
        return df, 0

    df = df.copy()
    failed = 0

    def _standardize(val):
        nonlocal failed
        if pd.isna(val) or str(val).strip() in ("", "nan", "None", "NaT"):
            return ""
        result = standardize_date(val)
        if result == str(val):  # standardize_date returns original on failure
            parsed = parse_date(val)
            if parsed is None:
                failed += 1
        return result

    df[col] = df[col].apply(_standardize)
    logger.info(f"Date column '{col}' standardized. Failures: {failed}")
    return df, failed


def round_gst_columns(
    df: pd.DataFrame, gst_cols: Optional[list[str]] = None
) -> pd.DataFrame:
    """
    Convert GST/value columns to numeric and round to 2 decimal places.

    Args:
        df:       Input DataFrame.
        gst_cols: Specific columns to round; defaults to all standard GST cols present.

    Returns:
        DataFrame with rounded numeric columns.
    """
    df = df.copy()
    cols_to_round = gst_cols or GST_NUMERIC_COLS

    for col in cols_to_round:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0).round(2)

    return df


def remove_hidden_characters(df: pd.DataFrame) -> pd.DataFrame:
    """
    Strip non-printable / hidden Unicode characters from all string columns.

    Args:
        df: Input DataFrame.

    Returns:
        Cleaned DataFrame.
    """
    df = df.copy()
    hidden_re = re.compile(r"[\x00-\x1F\x7F-\x9F\u200b-\u200f\ufeff\u00a0]")

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).apply(
            lambda x: hidden_re.sub("", x).strip() if x not in ("nan", "None") else x
        )

    return df


def _compute_total_gst(df: pd.DataFrame) -> pd.DataFrame:
    """
    If total_gst column is 0/empty, compute it from CGST + SGST + IGST + CESS.
    """
    df = df.copy()
    if "total_gst" in df.columns:
        computed = (
            df.get("cgst", 0).fillna(0)
            + df.get("sgst", 0).fillna(0)
            + df.get("igst", 0).fillna(0)
            + df.get("cess", 0).fillna(0)
        )
        # Fill zeros in total_gst with computed value
        mask = df["total_gst"].fillna(0) == 0
        df.loc[mask, "total_gst"] = computed[mask]
    return df


# ---------------------------------------------------------------------------
# Master Cleaning Function
# ---------------------------------------------------------------------------

def clean_dataframe(
    df: pd.DataFrame,
    column_mapping: dict,
) -> tuple[pd.DataFrame, dict]:
    """
    Master cleaning function: apply all cleaning steps in sequence.

    Args:
        df:             Raw DataFrame (after column mapping → standard column names).
        column_mapping: Dict mapping standard_col → original_col (used for reference only).

    Returns:
        (cleaned_df, cleaning_summary_dict)

    Cleaning summary keys:
        original_rows, final_rows, blank_removed, duplicates_removed,
        invalid_gstins, date_errors, steps_applied
    """
    summary = {
        "original_rows": len(df),
        "final_rows": 0,
        "blank_removed": 0,
        "duplicates_removed": 0,
        "invalid_gstins": 0,
        "date_errors": 0,
        "steps_applied": [],
    }

    try:
        # ── Step 1: Remove hidden characters ─────────────────────────────
        df = remove_hidden_characters(df)
        summary["steps_applied"].append("Hidden character removal")

        # ── Step 2: Remove blank rows ─────────────────────────────────────
        key_cols = ["gstin", "invoice_number", "invoice_date"]
        df, blank_removed = remove_blank_rows(
            df, key_cols=[c for c in key_cols if c in df.columns]
        )
        summary["blank_removed"] = blank_removed
        summary["steps_applied"].append(f"Blank rows removed: {blank_removed}")

        # ── Step 3: Remove duplicates ─────────────────────────────────────
        dup_cols = ["gstin", "invoice_number", "taxable_value"]
        df, dup_removed = remove_duplicate_rows(
            df, subset_cols=[c for c in dup_cols if c in df.columns]
        )
        summary["duplicates_removed"] = dup_removed
        summary["steps_applied"].append(f"Duplicate rows removed: {dup_removed}")

        # ── Step 4: Clean GSTIN ───────────────────────────────────────────
        if "gstin" in df.columns:
            df, invalid_gstins = clean_gstin_column(df, "gstin")
            summary["invalid_gstins"] = invalid_gstins
            summary["steps_applied"].append(f"GSTIN cleaned (invalid: {invalid_gstins})")

        # ── Step 5: Clean Invoice Number ──────────────────────────────────
        if "invoice_number" in df.columns:
            df = clean_invoice_column(df, "invoice_number")
            summary["steps_applied"].append("Invoice numbers normalized")

        # ── Step 6: Clean Vendor Name ─────────────────────────────────────
        if "vendor_name" in df.columns:
            df = clean_vendor_column(df, "vendor_name")
            summary["steps_applied"].append("Vendor names cleaned")

        # ── Step 7: Standardize Dates ─────────────────────────────────────
        if "invoice_date" in df.columns:
            df, date_errors = standardize_date_column(df, "invoice_date")
            summary["date_errors"] = date_errors
            summary["steps_applied"].append(f"Dates standardized (errors: {date_errors})")

        # ── Step 8: Round GST values ──────────────────────────────────────
        df = round_gst_columns(df)
        summary["steps_applied"].append("GST values rounded to 2 decimals")

        # ── Step 9: Compute total_gst if missing ──────────────────────────
        df = _compute_total_gst(df)
        summary["steps_applied"].append("Total GST computed where missing")

        df = df.reset_index(drop=True)
        summary["final_rows"] = len(df)

        log_event(
            "PROCESS",
            f"Data cleaning completed: {summary['original_rows']} → {summary['final_rows']} rows",
        )

    except Exception as e:
        logger.error(f"Cleaning pipeline failed: {e}")
        summary["steps_applied"].append(f"ERROR: {e}")

    return df, summary


# ---------------------------------------------------------------------------
# Streamlit Cleaning Summary Component
# ---------------------------------------------------------------------------

def render_cleaning_summary(summary: dict, source_label: str = "") -> None:
    """
    Display a styled cleaning summary in Streamlit.

    Args:
        summary:      Dict returned by clean_dataframe.
        source_label: Optional label like 'Purchase Register' for the header.
    """
    label = f" — {source_label}" if source_label else ""
    st.markdown(
        f"<h4 style='color:#00D4FF;'>🧹 Cleaning Summary{label}</h4>",
        unsafe_allow_html=True,
    )

    # Guard: summary must be a dict
    if not isinstance(summary, dict):
        st.info("No cleaning summary available yet.")
        return

    orig = summary.get("original_rows", 0)
    final = summary.get("final_rows", 0)
    blank = summary.get("blank_removed", 0)
    dupes = summary.get("duplicates_removed", 0)
    bad_gst = summary.get("invalid_gstins", 0)
    date_err = summary.get("date_errors", 0)

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("📥 Original Rows", f"{orig:,}")
    col2.metric("📤 Final Rows", f"{final:,}", delta=f"-{orig - final:,}" if orig != final else "0")
    col3.metric(
        "🗑️ Blank Removed",
        blank,
        delta=str(-blank) if blank else None,
        delta_color="inverse" if blank else "off",
    )
    col4.metric(
        "🔁 Duplicates",
        dupes,
        delta=str(-dupes) if dupes else None,
        delta_color="inverse" if dupes else "off",
    )
    col5.metric(
        "⚠️ Invalid GSTINs",
        bad_gst,
        delta=str(bad_gst) if bad_gst else None,
        delta_color="inverse" if bad_gst else "off",
    )
    col6.metric(
        "📅 Date Errors",
        date_err,
        delta=str(date_err) if date_err else None,
        delta_color="inverse" if date_err else "off",
    )

    # Steps log
    steps = summary.get("steps_applied", [])
    if steps:
        with st.expander("🔍 Cleaning Steps Applied"):
            for i, step in enumerate(steps, 1):
                color = "#EF4444" if "ERROR" in step else "#34D399"
                st.markdown(
                    f"<span style='color:{color};'>{'✅' if 'ERROR' not in step else '❌'} "
                    f"{i}. {step}</span>",
                    unsafe_allow_html=True,
                )

    # Overall health indicator
    issues = blank + dupes + bad_gst + date_err
    if issues == 0:
        st.success("🌟 Data quality is excellent — no issues found!")
    elif issues < 10:
        st.warning(f"⚠️ Minor issues found: {issues} total. Review above for details.")
    else:
        st.error(f"❌ Significant data quality issues: {issues} total. Please review before reconciliation.")
