"""
GST Input Reconciliation System – Enterprise Edition
Column Mapping Module
Prepared & Developed by Karthik LVN

Provides:
  - Automatic column detection using keyword matching + rapidfuzz
  - Manual column mapping override via Streamlit UI
  - Mapping profile save/load
  - Application of mapping to raw DataFrames
"""

import json
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from modules.utils import get_project_root, save_json_config, load_json_config, setup_logging
from modules.audit import log_event
from modules.cleaning import clean_dataframe, render_cleaning_summary

logger = setup_logging()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MAPPING_PROFILE_FILE: Path = get_project_root() / "data" / "mapping_profile.json"

# Standard column → list of recognized keyword aliases (all lowercase)
STANDARD_COLUMNS: dict[str, list[str]] = {
    "vendor_name": [
        "vendor", "supplier", "party", "vendor name", "supplier name",
        "party name", "buyer", "company name", "name",
    ],
    "gstin": [
        "gstin", "gst number", "gst no", "gstin number", "supplier gstin",
        "vendor gstin", "gstn", "tax id",
    ],
    "invoice_number": [
        "invoice", "inv no", "invoice no", "invoice number", "bill no",
        "bill number", "voucher no", "document no", "doc no", "ref no",
    ],
    "invoice_date": [
        "date", "invoice date", "bill date", "voucher date", "inv date",
        "document date", "doc date", "transaction date",
    ],
    "taxable_value": [
        "taxable", "taxable value", "taxable amount", "base amount",
        "assessable value", "taxable val", "basic amount",
    ],
    "cgst": [
        "cgst", "cgst amount", "central gst", "central tax", "cgst amt",
    ],
    "sgst": [
        "sgst", "sgst amount", "state gst", "state tax", "ugst", "utgst",
        "sgst amt",
    ],
    "igst": [
        "igst", "igst amount", "integrated gst", "integrated tax", "igst amt",
    ],
    "cess": [
        "cess", "cess amount", "compensation cess", "cess amt",
    ],
    "total_gst": [
        "total gst", "total tax", "gst amount", "tax amount",
        "total igst cgst sgst", "gst", "total taxes",
    ],
    "invoice_value": [
        "invoice value", "total value", "total amount", "grand total",
        "bill value", "invoice total", "net amount", "total invoice",
    ],
}

# Human-friendly labels for display
COLUMN_LABELS: dict[str, str] = {
    "vendor_name": "Vendor / Supplier Name",
    "gstin": "GSTIN",
    "invoice_number": "Invoice Number",
    "invoice_date": "Invoice Date",
    "taxable_value": "Taxable Value (₹)",
    "cgst": "CGST (₹)",
    "sgst": "SGST / UGST (₹)",
    "igst": "IGST (₹)",
    "cess": "CESS (₹)",
    "total_gst": "Total GST (₹)",
    "invoice_value": "Invoice Value / Total (₹)",
}

NOT_AVAILABLE = "-- Not Available --"


# ---------------------------------------------------------------------------
# Auto-Detection
# ---------------------------------------------------------------------------

def auto_detect_columns(df_columns: list[str]) -> dict[str, Optional[str]]:
    """
    Automatically detect which DataFrame column maps to each standard column.

    Uses:
    1. Exact lowercase match against known aliases (excluding already-mapped cols).
    2. Partial contains match.
    3. RapidFuzz similarity (threshold 70) as a fallback.

    Args:
        df_columns: List of actual column names in the DataFrame.

    Returns:
        Dict {standard_col: detected_df_col_or_None}
    """
    try:
        from rapidfuzz import process as rfprocess, fuzz as rffuzz
        _has_rapidfuzz = True
    except ImportError:
        _has_rapidfuzz = False

    normalized_cols = {col.lower().strip(): col for col in df_columns}
    result: dict[str, Optional[str]] = {}
    # Track which source columns have already been claimed
    used_source_cols: set[str] = set()

    # Process specific (short-name) columns before catch-all ones
    # This ensures cgst/sgst/igst are matched before total_gst steals them
    PRIORITY_ORDER = [
        "vendor_name", "gstin", "invoice_number", "invoice_date",
        "taxable_value", "cgst", "sgst", "igst", "cess",
        "total_gst", "invoice_value",
    ]

    for std_col in PRIORITY_ORDER:
        aliases = STANDARD_COLUMNS.get(std_col, [])
        detected: Optional[str] = None

        # Pass 0: check if the column name itself exactly matches the standard col name
        if std_col in normalized_cols and normalized_cols[std_col] not in used_source_cols:
            detected = normalized_cols[std_col]

        # Pass 1: exact alias match
        if detected is None:
            for alias in aliases:
                if alias in normalized_cols and normalized_cols[alias] not in used_source_cols:
                    detected = normalized_cols[alias]
                    break

        # Pass 2: partial/contains match (use the MOST SPECIFIC alias, avoid short ones for total_gst)
        if detected is None:
            for alias in sorted(aliases, key=len, reverse=True):  # longest alias first
                if len(alias) < 4:  # skip very short aliases like "gst" to avoid false matches
                    continue
                for norm_col, orig_col in normalized_cols.items():
                    if orig_col in used_source_cols:
                        continue
                    if alias == norm_col or (alias in norm_col and len(alias) >= len(norm_col) - 3):
                        detected = orig_col
                        break
                if detected:
                    break

        # Pass 3: rapidfuzz similarity
        if detected is None and _has_rapidfuzz:
            best_score = 0
            best_col = None
            available_cols = {k: v for k, v in normalized_cols.items() if v not in used_source_cols}
            for alias in aliases:
                if not available_cols:
                    break
                match = rfprocess.extractOne(
                    alias,
                    list(available_cols.keys()),
                    scorer=rffuzz.token_sort_ratio,
                    score_cutoff=75,  # raised from 70 to reduce false positives
                )
                if match and match[1] > best_score:
                    best_score = match[1]
                    best_col = available_cols[match[0]]
            detected = best_col

        result[std_col] = detected
        if detected:
            used_source_cols.add(detected)

    return result


# ---------------------------------------------------------------------------
# Profile Save / Load
# ---------------------------------------------------------------------------

def save_mapping_profile(mapping: dict[str, Optional[str]], source: str) -> None:
    """
    Save a column mapping profile for future use.

    Args:
        mapping: {standard_col: df_col} mapping dict.
        source:  'PR' or 'GSTR2B'.
    """
    profiles = load_json_config(MAPPING_PROFILE_FILE) or {}
    profiles[source] = mapping
    save_json_config(MAPPING_PROFILE_FILE, profiles)
    logger.info(f"Mapping profile saved for source: {source}")


def load_mapping_profile(source: str) -> Optional[dict[str, Optional[str]]]:
    """
    Load a saved column mapping profile.

    Args:
        source: 'PR' or 'GSTR2B'.

    Returns:
        Saved mapping dict, or None if not found.
    """
    profiles = load_json_config(MAPPING_PROFILE_FILE) or {}
    return profiles.get(source)


# ---------------------------------------------------------------------------
# Apply Mapping
# ---------------------------------------------------------------------------

def apply_mapping(
    df: pd.DataFrame,
    mapping: dict[str, Optional[str]],
    source_tag: str = "",
) -> pd.DataFrame:
    """
    Rename DataFrame columns to standard names and add missing standard columns.

    Args:
        df:         Raw DataFrame.
        mapping:    {standard_col: original_df_col or None} dict.
        source_tag: 'PR' or 'GSTR2B' — added as 'source' column for traceability.

    Returns:
        Standardized DataFrame with exactly the standard column set.
    """
    df = df.copy()
    rename_map: dict[str, str] = {}

    for std_col, orig_col in mapping.items():
        if orig_col and orig_col in df.columns and orig_col != std_col:
            rename_map[orig_col] = std_col

    df = df.rename(columns=rename_map)

    # Ensure all standard columns exist; add missing ones as empty/zero
    for std_col in STANDARD_COLUMNS:
        if std_col not in df.columns:
            if std_col in ("taxable_value", "cgst", "sgst", "igst", "cess", "total_gst", "invoice_value"):
                df[std_col] = 0.0
            else:
                df[std_col] = ""

    # Add source tag
    df["source"] = source_tag

    # Keep only standard columns + source
    keep_cols = list(STANDARD_COLUMNS.keys()) + ["source"]
    df = df[[c for c in keep_cols if c in df.columns]]

    return df


# ---------------------------------------------------------------------------
# Streamlit Mapping Page
# ---------------------------------------------------------------------------

def _render_mapping_tab(
    source_label: str,
    df_key: str,
    mapped_key: str,
    source_tag: str,
) -> None:
    """
    Render the mapping UI for one source (PR or GSTR-2B).

    Args:
        source_label: Human-readable label for display.
        df_key:       Session state key for the raw DataFrame.
        mapped_key:   Session state key to store the mapped+cleaned DataFrame.
        source_tag:   Short tag ('PR' or 'GSTR2B') used internally.
    """
    df: Optional[pd.DataFrame] = st.session_state.get(df_key)

    if df is None:
        st.warning(
            f"⚠️ {source_label} data not uploaded yet. "
            "Please go to **Upload** first."
        )
        return

    df_cols = list(df.columns)
    options = [NOT_AVAILABLE] + df_cols

    # ── Auto-detect ────────────────────────────────────────────────────────
    auto_key = f"auto_mapping_{source_tag}"
    if auto_key not in st.session_state or not isinstance(st.session_state.get(auto_key), dict):
        st.session_state[auto_key] = auto_detect_columns(df_cols) or {}

    auto_map: dict = st.session_state[auto_key]
    if not isinstance(auto_map, dict):
        auto_map = {}
        st.session_state[auto_key] = auto_map

    # ── Saved profile / Re-detect buttons ─────────────────────────────────
    saved_profile = load_mapping_profile(source_tag)
    col_load, col_detect = st.columns([1, 1])
    if saved_profile and col_load.button(
        "Load Saved Profile", key=f"load_profile_{source_tag}"
    ):
        st.session_state[auto_key] = {
            k: v for k, v in saved_profile.items() if v in df_cols or v is None
        }
        st.success("Saved profile loaded!")
        st.rerun()

    if col_detect.button("Re-run Auto Detection", key=f"redetect_{source_tag}"):
        st.session_state[auto_key] = auto_detect_columns(df_cols) or {}
        st.rerun()

    st.markdown(
        f"<p style='color:#94A3B8; font-size:0.85rem;'>Map each standard field "
        f"to the corresponding column in your <strong>{source_label}</strong> file.</p>",
        unsafe_allow_html=True,
    )

    # ── Mapping selectboxes ────────────────────────────────────────────────
    user_mapping: dict[str, Optional[str]] = {}

    for std_col, label in COLUMN_LABELS.items():
        detected = auto_map.get(std_col) if auto_map else None
        if detected and detected in df_cols:
            default_idx = options.index(detected)
            indicator = "OK"
            help_text = f"Auto-detected: '{detected}'"
        else:
            default_idx = 0
            indicator = "?"
            help_text = "Not auto-detected — please select manually if available."

        col_label, col_select = st.columns([2, 3])
        col_label.markdown(
            f"<div style='padding-top:8px;'><b>{indicator}</b> {label}</div>",
            unsafe_allow_html=True,
        )
        selected = col_select.selectbox(
            label,
            options,
            index=default_idx,
            key=f"map_{source_tag}_{std_col}",
            help=help_text,
            label_visibility="collapsed",
        )
        user_mapping[std_col] = selected if selected != NOT_AVAILABLE else None

    st.divider()

    # ── Apply & Save buttons ────────────────────────────────────────────────
    btn_col1, btn_col2 = st.columns(2)

    if btn_col1.button(
        f"✅ Apply Mapping & Clean Data",
        key=f"apply_mapping_{source_tag}",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner(f"Applying mapping and cleaning {source_label}…"):
            try:
                mapped_df = apply_mapping(df, user_mapping, source_tag=source_tag)
                cleaned_df, summary = clean_dataframe(mapped_df, user_mapping)
                st.session_state[mapped_key] = cleaned_df
                st.session_state[f"cleaning_summary_{source_tag}"] = summary
                log_event(
                    "PROCESS",
                    f"Column mapping applied for {source_label}: {len(cleaned_df)} rows",
                )
                st.success(f"✅ {source_label} mapped and cleaned — {len(cleaned_df):,} rows ready.")
            except Exception as e:
                st.error(f"❌ Mapping failed: {e}")
                logger.error(f"Mapping apply failed for {source_tag}: {e}")

    if btn_col2.button(
        "💾 Save as Profile",
        key=f"save_profile_{source_tag}",
        use_container_width=True,
    ):
        save_mapping_profile(user_mapping, source_tag)
        st.success("Profile saved for future use!")

    # ── Show cleaning summary if already applied ────────────────────────────
    summary_key = f"cleaning_summary_{source_tag}"
    if summary_key in st.session_state:
        render_cleaning_summary(st.session_state[summary_key], source_label)

    # ── Preview of mapped data ─────────────────────────────────────────────
    if mapped_key in st.session_state and st.session_state[mapped_key] is not None:
        with st.expander(f"🔍 Preview Mapped & Cleaned Data (first 5 rows)"):
            st.dataframe(
                st.session_state[mapped_key].head(5), use_container_width=True
            )


def render_mapping_page() -> None:
    """Render the full Column Mapping page with tabs for PR and GSTR-2B."""

    st.markdown(
        "<h2 style='color:#00D4FF;'>Column Mapping</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='color:#94A3B8;'>Map your file's columns to the standard fields "
        "required for reconciliation.</p>",
        unsafe_allow_html=True,
    )

    # ── One-click Apply Both ───────────────────────────────────────────────
    pr_df   = st.session_state.get("pr_df")
    gstr_df = st.session_state.get("gstr2b_df")
    pr_mapped   = st.session_state.get("pr_mapped")
    gstr_mapped = st.session_state.get("gstr2b_mapped")

    # Status badges
    p_stat = ("<span style='background:#34D39922; color:#34D399; border:1px solid #34D39944; "
              "border-radius:20px; padding:2px 10px; font-size:0.72rem;'>MAPPED</span>"
              if pr_mapped is not None else
              "<span style='background:#F8717122; color:#F87171; border:1px solid #F8717144; "
              "border-radius:20px; padding:2px 10px; font-size:0.72rem;'>NOT MAPPED</span>")
    g_stat = ("<span style='background:#34D39922; color:#34D399; border:1px solid #34D39944; "
              "border-radius:20px; padding:2px 10px; font-size:0.72rem;'>MAPPED</span>"
              if gstr_mapped is not None else
              "<span style='background:#F8717122; color:#F87171; border:1px solid #F8717144; "
              "border-radius:20px; padding:2px 10px; font-size:0.72rem;'>NOT MAPPED</span>")
    st.markdown(
        f"""
        <div style="display:flex; gap:12px; margin-bottom:12px;">
            <div style="flex:1; background:rgba(26,26,46,0.6); border:1px solid rgba(0,212,255,0.2);
                 border-radius:8px; padding:8px 14px; display:flex; justify-content:space-between; align-items:center;">
                <span style="color:#00D4FF; font-size:0.88rem; font-weight:600;">Purchase Register</span>
                {p_stat}
            </div>
            <div style="flex:1; background:rgba(26,26,46,0.6); border:1px solid rgba(167,139,250,0.2);
                 border-radius:8px; padding:8px 14px; display:flex; justify-content:space-between; align-items:center;">
                <span style="color:#A78BFA; font-size:0.88rem; font-weight:600;">GSTR-2B</span>
                {g_stat}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if pr_df is not None and gstr_df is not None:
        if st.button(
            "Apply Auto-Mapping for BOTH Files (One Click)",
            type="primary",
            use_container_width=True,
            key="apply_both_btn",
        ):
            errors = []
            for src_label, df_key, mapped_key, src_tag in [
                ("Purchase Register", "pr_df", "pr_mapped", "PR"),
                ("GSTR-2B", "gstr2b_df", "gstr2b_mapped", "GSTR2B"),
            ]:
                raw_df = st.session_state.get(df_key)
                if raw_df is None:
                    errors.append(f"{src_label} not uploaded")
                    continue
                try:
                    auto_m = auto_detect_columns(list(raw_df.columns)) or {}
                    mapped_df = apply_mapping(raw_df, auto_m, source_tag=src_tag)
                    cleaned_df, summary = clean_dataframe(mapped_df, auto_m)
                    st.session_state[mapped_key] = cleaned_df
                    st.session_state[f"cleaning_summary_{src_tag}"] = summary
                    log_event("PROCESS", f"Auto-mapping applied for {src_label}: {len(cleaned_df)} rows")
                except Exception as e:
                    errors.append(f"{src_label}: {e}")
            if errors:
                st.error("Errors: " + " | ".join(errors))
            else:
                st.success("Both files mapped and cleaned successfully!")
                st.rerun()
        st.markdown(
            "<div style='text-align:center; color:#64748B; font-size:0.78rem; margin:4px 0 14px 0;'>"
            "— or map columns manually in the tabs below —</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<div style='background:rgba(248,113,113,0.08); border:1px solid #F8717133; "
            "border-radius:8px; padding:10px 16px; margin-bottom:14px;'>"
            "<span style='color:#F87171; font-weight:600;'>Files not loaded.</span> "
            "<span style='color:#94A3B8;'>Please go to </span>"
            "<strong style='color:#00D4FF;'>Upload Data</strong>"
            "<span style='color:#94A3B8;'> first and upload both files, then come back here.</span>"
            "</div>",
            unsafe_allow_html=True,
        )

    tab_pr, tab_gstr = st.tabs(
        ["Purchase Register Mapping", "GSTR-2B Mapping"]
    )

    with tab_pr:
        _render_mapping_tab(
            source_label="Purchase Register",
            df_key="pr_df",
            mapped_key="pr_mapped",
            source_tag="PR",
        )

    with tab_gstr:
        _render_mapping_tab(
            source_label="GSTR-2B",
            df_key="gstr2b_df",
            mapped_key="gstr2b_mapped",
            source_tag="GSTR2B",
        )

    st.divider()

    # ── Proceed button ─────────────────────────────────────────────────────
    pr_mapped = st.session_state.get("pr_mapped") is not None
    gstr_mapped = st.session_state.get("gstr2b_mapped") is not None

    if pr_mapped and gstr_mapped:
        st.success("Both files mapped and cleaned. Ready for reconciliation!")
        if st.button(
            "Proceed to Reconcile",
            type="primary",
            use_container_width=True,
            key="proceed_recon_btn",
        ):
            st.session_state["current_page"] = "Reconcile"
            st.rerun()
    else:
        missing = []
        if not pr_mapped:
            missing.append("Purchase Register")
        if not gstr_mapped:
            missing.append("GSTR-2B")
        st.warning(
            f"Please apply mapping for: {', '.join(missing)} before proceeding."
        )
