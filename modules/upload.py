"""
GST Input Reconciliation System – Enterprise Edition
File Upload Module
Prepared & Developed by Karthik LVN

Provides:
  - Excel (.xlsx, .xls) and CSV file upload
  - Drag & drop UI with progress feedback
  - File validation (type, size)
  - Upload history management
  - Multi-sheet Excel handling
"""

import io
import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from modules.utils import get_temp_path, setup_logging, safe_float
from modules.audit import log_event

logger = setup_logging()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
MAX_FILE_SIZE_BYTES = 500 * 1024 * 1024  # 500 MB


# ---------------------------------------------------------------------------
# File Validation
# ---------------------------------------------------------------------------

def validate_file(uploaded_file) -> tuple[bool, str]:
    """
    Validate an uploaded file for type and size.

    Args:
        uploaded_file: Streamlit UploadedFile object.

    Returns:
        (True, "") if valid, else (False, error_message).
    """
    if uploaded_file is None:
        return False, "No file uploaded."

    name = uploaded_file.name
    ext = Path(name).suffix.lower()

    if ext not in ALLOWED_EXTENSIONS:
        return False, (
            f"Unsupported file type '{ext}'. "
            f"Allowed: {', '.join(sorted(ALLOWED_EXTENSIONS))}"
        )

    size = uploaded_file.size
    if size == 0:
        return False, "The uploaded file is empty."
    if size > MAX_FILE_SIZE_BYTES:
        size_mb = size / (1024 * 1024)
        return False, f"File size {size_mb:.1f} MB exceeds the 500 MB limit."

    return True, ""


# ---------------------------------------------------------------------------
# File Reading
# ---------------------------------------------------------------------------

def read_uploaded_file(
    uploaded_file,
    sheet_name: Optional[str] = None,
) -> tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Read an uploaded Excel or CSV file into a pandas DataFrame.

    For multi-sheet Excel files, the user is prompted to choose a sheet
    via a selectbox (handled in render_upload_page).

    Args:
        uploaded_file: Streamlit UploadedFile object.
        sheet_name:    Specific Excel sheet to read (None = first sheet).

    Returns:
        (DataFrame, None) on success, (None, error_message) on failure.
    """
    try:
        name = uploaded_file.name
        ext = Path(name).suffix.lower()
        content = uploaded_file.read()
        uploaded_file.seek(0)  # reset for re-reads

        if ext in (".xlsx", ".xls"):
            xl = pd.ExcelFile(io.BytesIO(content), engine="openpyxl" if ext == ".xlsx" else None)
            sheets = xl.sheet_names

            if sheet_name is None:
                sheet_name = sheets[0]

            df = pd.read_excel(
                io.BytesIO(content),
                sheet_name=sheet_name,
                engine="openpyxl" if ext == ".xlsx" else None,
                dtype=str,          # read all as string; type conversion done in cleaning
            )
            return df, None

        elif ext == ".csv":
            # Try UTF-8 first, fall back to latin-1
            try:
                df = pd.read_csv(io.BytesIO(content), dtype=str, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(content), dtype=str, encoding="latin-1")
            return df, None

    except Exception as e:
        logger.error(f"Failed to read file '{uploaded_file.name}': {e}")
        return None, f"Error reading file: {e}"

    return None, "Unknown file type."


def get_excel_sheets(uploaded_file) -> list[str]:
    """
    Return the list of sheet names for an Excel file.

    Args:
        uploaded_file: Streamlit UploadedFile object.

    Returns:
        List of sheet name strings, or ['Sheet1'] for CSV.
    """
    ext = Path(uploaded_file.name).suffix.lower()
    if ext not in (".xlsx", ".xls"):
        return []

    try:
        content = uploaded_file.read()
        uploaded_file.seek(0)
        xl = pd.ExcelFile(io.BytesIO(content), engine="openpyxl" if ext == ".xlsx" else None)
        return xl.sheet_names
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Temp File Management
# ---------------------------------------------------------------------------

def save_temp_file(uploaded_file) -> Path:
    """
    Save an uploaded file to the temp directory with a timestamp prefix.

    Args:
        uploaded_file: Streamlit UploadedFile object.

    Returns:
        Path to the saved temp file.
    """
    temp_path = get_temp_path(uploaded_file.name)
    temp_path.parent.mkdir(parents=True, exist_ok=True)

    content = uploaded_file.read()
    uploaded_file.seek(0)

    with open(temp_path, "wb") as f:
        f.write(content)

    logger.info(f"Temp file saved: {temp_path}")
    return temp_path


# ---------------------------------------------------------------------------
# Upload History
# ---------------------------------------------------------------------------

def get_upload_history() -> list[dict]:
    """
    Return the upload history from session state.

    Returns:
        List of upload record dicts.
    """
    return st.session_state.get("upload_history", [])


def add_to_upload_history(filename: str, rows: int, source: str) -> None:
    """
    Append an upload record to session state history.

    Args:
        filename: Name of the uploaded file.
        rows:     Number of data rows in the file.
        source:   'Purchase Register' or 'GSTR-2B'.
    """
    if "upload_history" not in st.session_state:
        st.session_state["upload_history"] = []

    st.session_state["upload_history"].append(
        {
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "filename": filename,
            "source": source,
            "rows": rows,
        }
    )


# ---------------------------------------------------------------------------
# Helper: Format file size
# ---------------------------------------------------------------------------

def _format_size(size_bytes: int) -> str:
    """Return a human-readable file size string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 ** 2:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 ** 2):.1f} MB"


# ---------------------------------------------------------------------------
# Streamlit Upload Page
# ---------------------------------------------------------------------------

def _render_single_uploader(
    source_label: str,
    session_key: str,
    icon: str,
    color: str,
) -> None:
    """
    Render a single file uploader widget for one source (PR or GSTR-2B).

    Args:
        source_label: Human-readable source name.
        session_key:  Key to store the DataFrame in session state.
        icon:         Emoji icon for the header.
        color:        Hex color for the header border.
    """
    st.markdown(
        f"""
        <div style="border-left: 4px solid {color}; padding: 8px 16px;
             background: rgba(26,26,46,0.6); border-radius: 8px; margin-bottom:12px;">
            <span style="font-size:1.1rem; font-weight:700; color:{color};">
                {icon} {source_label}
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Upload widget ──────────────────────────────────────────────────────
    uploaded = st.file_uploader(
        f"Upload {source_label} (Excel / CSV)",
        type=["xlsx", "xls", "csv"],
        key=f"uploader_{session_key}",
        help=f"Drag & drop or click to browse. Max size: 500 MB",
        label_visibility="collapsed",
    )

    st.caption(
        "📂 Drag & drop your file here or click to browse  ·  "
        "Supported: .xlsx, .xls, .csv  ·  Max: 500 MB"
    )

    if uploaded is not None:
        # Validate
        valid, err_msg = validate_file(uploaded)
        if not valid:
            st.error(f"❌ {err_msg}")
            return

        # File info
        fi1, fi2, fi3 = st.columns(3)
        fi1.metric("📄 File", uploaded.name[:28] + ("…" if len(uploaded.name) > 28 else ""))
        fi2.metric("💾 Size", _format_size(uploaded.size))

        # Handle multi-sheet Excel
        ext = Path(uploaded.name).suffix.lower()
        sheet_name = None

        if ext in (".xlsx", ".xls"):
            sheets = get_excel_sheets(uploaded)
            if len(sheets) > 1:
                sheet_name = st.selectbox(
                    "📋 Select Sheet",
                    sheets,
                    key=f"sheet_{session_key}",
                )
            elif sheets:
                sheet_name = sheets[0]

        # Progress bar simulation
        progress_placeholder = st.empty()
        progress_bar = progress_placeholder.progress(0, text="Reading file…")

        df, read_err = read_uploaded_file(uploaded, sheet_name=sheet_name)

        if read_err:
            progress_placeholder.empty()
            st.error(f"❌ {read_err}")
            return

        if df is None or df.empty:
            progress_placeholder.empty()
            st.error("❌ The file appears to be empty or unreadable.")
            return

        # Remove fully blank rows immediately
        df.dropna(how="all", inplace=True)
        df.reset_index(drop=True, inplace=True)

        progress_bar.progress(100, text="✅ File loaded successfully!")
        import time
        time.sleep(0.4)
        progress_placeholder.empty()

        fi3.metric("📊 Rows", f"{len(df):,}")

        # Store in session state
        st.session_state[session_key] = df

        # Log upload event
        log_event(
            "UPLOAD",
            f"Uploaded {source_label}: '{uploaded.name}' ({len(df)} rows)",
            file_name=uploaded.name,
        )
        add_to_upload_history(uploaded.name, len(df), source_label)

        st.success(
            f"✅ **{source_label}** loaded — **{len(df):,} rows** × **{len(df.columns)} columns**"
        )

        # Data preview
        with st.expander(f"🔍 Preview: {uploaded.name} (first 10 rows)"):
            st.dataframe(df.head(10), use_container_width=True)

    elif session_key in st.session_state and st.session_state[session_key] is not None:
        existing_df = st.session_state[session_key]
        st.info(
            f"✅ Previously loaded — **{len(existing_df):,} rows** × "
            f"**{len(existing_df.columns)} columns**"
        )
        with st.expander("🔍 Preview loaded data"):
            st.dataframe(existing_df.head(10), use_container_width=True)

        if st.button(f"🗑️ Clear {source_label}", key=f"clear_{session_key}"):
            st.session_state[session_key] = None
            st.rerun()
    else:
        st.info(f"⬆️ Upload your **{source_label}** file to get started.")


def render_upload_page() -> None:
    """Render the full File Upload page with two uploaders and history."""

    st.markdown(
        "<h2 style='color:#00D4FF;'>📤 File Upload</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='color:#94A3B8;'>Upload your Purchase Register and GSTR-2B files "
        "to begin reconciliation.</p>",
        unsafe_allow_html=True,
    )

    # ── Two-column layout ──────────────────────────────────────────────────
    col_pr, col_gstr = st.columns(2, gap="large")

    with col_pr:
        _render_single_uploader(
            source_label="Purchase Register",
            session_key="pr_df",
            icon="📋",
            color="#00D4FF",
        )

    with col_gstr:
        _render_single_uploader(
            source_label="GSTR-2B",
            session_key="gstr2b_df",
            icon="🏛️",
            color="#A78BFA",
        )

    st.divider()

    # ── Proceed button ─────────────────────────────────────────────────────
    pr_ready = st.session_state.get("pr_df") is not None
    gstr_ready = st.session_state.get("gstr2b_df") is not None

    if pr_ready and gstr_ready:
        st.success(
            "✅ Both files uploaded. Click below to proceed to Column Mapping."
        )
        if st.button(
            "➡️ Proceed to Column Mapping",
            type="primary",
            use_container_width=True,
            key="proceed_mapping_btn",
        ):
            st.session_state["current_page"] = "Column Mapping"
            st.rerun()
    else:
        missing = []
        if not pr_ready:
            missing.append("Purchase Register")
        if not gstr_ready:
            missing.append("GSTR-2B")
        st.warning(f"⚠️ Waiting for: **{', '.join(missing)}**")

    # ── Upload History ─────────────────────────────────────────────────────
    history = get_upload_history()
    if history:
        st.divider()
        st.subheader("📜 Upload History (this session)")
        history_df = pd.DataFrame(history[::-1])  # most recent first
        history_df.columns = ["Timestamp", "File Name", "Source", "Rows"]
        st.dataframe(history_df, use_container_width=True, hide_index=True)
