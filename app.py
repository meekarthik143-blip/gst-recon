"""
GST Input Reconciliation System – Enterprise Edition
Main Application Entry Point
Prepared & Developed by Karthik LVN

Features:
  - Splash screen (3-second animated loader on first run)
  - Session state initialization
  - Sidebar navigation
  - Page routing
  - Global CSS injection for enterprise dark theme
  - Persistent branding and footer
"""

import time
import uuid
import datetime

import streamlit as st

# ── Page Config (must be first Streamlit call) ─────────────────────────────
st.set_page_config(
    page_title="GST Input Reconciliation System",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": "GST Input Reconciliation System v1.0 Enterprise Edition\nPrepared & Developed by Karthik LVN",
    },
)

# ── Bootstrap: ensure directories exist ───────────────────────────────────
from modules.utils import ensure_directories, setup_logging, get_current_financial_year, cleanup_temp_files
from modules.authentication import initialize_users, render_login_page
from modules.audit import initialize_audit_db, log_event
from modules.settings import initialize_settings, load_settings

ensure_directories()
initialize_users()
initialize_audit_db()
initialize_settings()

logger = setup_logging()

# ---------------------------------------------------------------------------
# Global CSS
# ---------------------------------------------------------------------------

GLOBAL_CSS = """
<style>
/* ── Font: system-ui fallback (no network dependency) ── */

/* ── Root Variables ── */
:root {
    --primary: #00D4FF;
    --secondary: #A78BFA;
    --success: #34D399;
    --warning: #FBBF24;
    --danger: #F87171;
    --dark-bg: #0A0A1A;
    --card-bg: rgba(26, 26, 46, 0.85);
    --border: rgba(0, 212, 255, 0.2);
    --text: #EAEAEA;
    --muted: #64748B;
}

/* ── App Background ── */
.stApp {
    background: linear-gradient(135deg, #0A0A1A 0%, #0D1B2A 40%, #0A0A1A 100%);
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0D1422 0%, #0A0F1E 100%) !important;
    border-right: 1px solid rgba(0, 212, 255, 0.15) !important;
}

[data-testid="stSidebar"] .stButton button {
    background: transparent;
    border: 1px solid rgba(0, 212, 255, 0.2);
    color: #EAEAEA;
    border-radius: 8px;
    padding: 8px 16px;
    font-family: 'Inter', sans-serif;
    transition: all 0.2s ease;
    text-align: left;
    width: 100%;
    margin: 2px 0;
}
[data-testid="stSidebar"] .stButton button:hover {
    background: rgba(0, 212, 255, 0.12);
    border-color: #00D4FF;
    color: #00D4FF;
    transform: translateX(4px);
}

/* ── Metric Cards ── */
[data-testid="metric-container"] {
    background: rgba(26, 26, 46, 0.8) !important;
    border: 1px solid rgba(0, 212, 255, 0.2) !important;
    border-radius: 12px !important;
    padding: 16px !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(0, 212, 255, 0.15);
    border-radius: 8px;
    overflow: hidden;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(13, 27, 42, 0.8);
    border-radius: 12px;
    padding: 4px;
    gap: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    color: #94A3B8;
    font-family: 'Inter', sans-serif;
    font-weight: 500;
    padding: 8px 16px;
}
.stTabs [aria-selected="true"] {
    background: rgba(0, 212, 255, 0.15) !important;
    color: #00D4FF !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
    border: 1px solid rgba(0, 212, 255, 0.15);
    border-radius: 12px;
    background: rgba(26, 26, 46, 0.5);
}

/* ── Buttons ── */
.stButton button[kind="primary"] {
    background: linear-gradient(135deg, #00D4FF, #0099BB) !important;
    color: #000 !important;
    border: none !important;
    font-weight: 700;
    border-radius: 8px;
    font-family: 'Inter', sans-serif;
    transition: all 0.2s ease;
}
.stButton button[kind="primary"]:hover {
    background: linear-gradient(135deg, #00EEFF, #00D4FF) !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 16px rgba(0, 212, 255, 0.4);
}

/* ── Inputs ── */
.stTextInput input, .stSelectbox select, .stTextArea textarea {
    background: rgba(13, 27, 42, 0.8) !important;
    border: 1px solid rgba(0, 212, 255, 0.25) !important;
    color: #EAEAEA !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
}
.stTextInput input:focus, .stSelectbox select:focus {
    border-color: #00D4FF !important;
    box-shadow: 0 0 0 2px rgba(0, 212, 255, 0.2) !important;
}

/* ── Progress bar ── */
.stProgress > div > div {
    background: linear-gradient(90deg, #00D4FF, #A78BFA) !important;
    border-radius: 4px !important;
}

/* ── Divider ── */
hr {
    border-color: rgba(0, 212, 255, 0.1) !important;
    margin: 20px 0 !important;
}

/* ── Success/Error/Warning boxes ── */
.stAlert {
    border-radius: 10px;
    border-left-width: 4px;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0A0A1A; }
::-webkit-scrollbar-thumb { background: rgba(0, 212, 255, 0.3); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #00D4FF; }

/* ── Hide Streamlit branding ── */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }
.stDeployButton { display: none; }
</style>
"""


# ---------------------------------------------------------------------------
# Splash Screen
# ---------------------------------------------------------------------------

def show_splash_screen() -> None:
    """Display animated splash screen for 3 seconds on first launch."""

    splash_placeholder = st.empty()

    splash_html = """
    <div style="
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 85vh;
        background: linear-gradient(135deg, #0A0A1A 0%, #0D1B2A 50%, #0A0A1A 100%);
    ">
        <div style="text-align: center; animation: fadeInUp 0.8s ease-out;">
            <div style="font-size: 5rem; margin-bottom: 16px; animation: pulse 2s infinite;">📊</div>

            <div style="
                font-size: 2.2rem;
                font-weight: 800;
                background: linear-gradient(90deg, #00D4FF, #A78BFA, #00D4FF);
                background-size: 200% auto;
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                animation: shimmer 2s linear infinite;
                margin-bottom: 8px;
            ">
                GST Input Reconciliation System
            </div>

            <div style="
                font-size: 1rem;
                color: #A78BFA;
                letter-spacing: 4px;
                text-transform: uppercase;
                margin-bottom: 6px;
            ">
                Enterprise Edition
            </div>

            <div style="color: #64748B; font-size: 0.85rem; margin-bottom: 32px;">
                Prepared & Developed by
                <span style="color: #00D4FF; font-weight: 700;">Karthik LVN</span>
            </div>

            <!-- Loading bar -->
            <div style="
                width: 320px;
                height: 4px;
                background: rgba(255,255,255,0.1);
                border-radius: 4px;
                overflow: hidden;
                margin: 0 auto 16px auto;
            ">
                <div style="
                    height: 100%;
                    background: linear-gradient(90deg, #00D4FF, #A78BFA);
                    border-radius: 4px;
                    animation: loadingBar 2.8s ease-in-out forwards;
                "></div>
            </div>

            <div style="color: #374151; font-size: 0.78rem; animation: blink 1.2s infinite;">
                Loading application…
            </div>
        </div>
    </div>

    <style>
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(30px); }
        to   { opacity: 1; transform: translateY(0); }
    }
    @keyframes shimmer {
        0%   { background-position: 0% center; }
        100% { background-position: 200% center; }
    }
    @keyframes pulse {
        0%, 100% { transform: scale(1); }
        50%       { transform: scale(1.08); }
    }
    @keyframes loadingBar {
        0%   { width: 0%; }
        30%  { width: 40%; }
        70%  { width: 80%; }
        100% { width: 100%; }
    }
    @keyframes blink {
        0%, 100% { opacity: 1; }
        50%       { opacity: 0.4; }
    }
    </style>
    """

    splash_placeholder.markdown(splash_html, unsafe_allow_html=True)
    time.sleep(3)
    splash_placeholder.empty()


# ---------------------------------------------------------------------------
# Session State Initialization
# ---------------------------------------------------------------------------

def init_session_state() -> None:
    """Initialize all required session state variables with defaults."""

    defaults = {
        "authenticated": False,
        "username": None,
        "role": None,
        "full_name": None,
        "session_id": None,
        "login_time": None,
        "current_page": "Dashboard",
        "splash_shown": False,
        "app_settings": None,
        # Data
        "pr_df": None,
        "gstr2b_df": None,
        "pr_mapped": None,
        "gstr2b_mapped": None,
        "recon_results": None,
        "master_df": None,
        # History
        "upload_history": [],
        # Mapping
        "auto_mapping_PR": None,
        "auto_mapping_GSTR2B": None,
        # Cleaning summaries
        "cleaning_summary_PR": None,
        "cleaning_summary_GSTR2B": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# ---------------------------------------------------------------------------
# Sidebar Navigation
# ---------------------------------------------------------------------------

# 5 Main workflow steps + admin
PAGES = {
    "Dashboard":       "🏠",
    "Upload Data":     "📤",
    "Column Mapping":  "🗂️",
    "Reconcile":       "⚙️",
    "Reports":         "📄",
    "Settings":        "⚙️",
}

ADMIN_ONLY_PAGES = {"User Management": "👥"}

# Step labels shown in sidebar
STEP_LABELS = [
    ("Dashboard",      "🏠", "Overview & KPIs"),
    ("Upload Data",    "📤", "Templates & Upload"),
    ("Column Mapping", "🗂️", "Map Columns"),
    ("Reconcile",      "⚙️", "Run Reconciliation"),
    ("Reports",        "📄", "View & Download"),
]


def render_sidebar() -> None:
    """Render the left sidebar with navigation, user info, and branding."""

    with st.sidebar:
        # ── Logo / Branding ──────────────────────────────────────────────
        st.markdown(
            """
            <div style="text-align:center; padding:16px 8px 8px 8px;
                 border-bottom:1px solid rgba(0,212,255,0.15); margin-bottom:12px;">
                <div style="font-size:1.8rem;">📊</div>
                <div style="font-size:0.85rem; font-weight:700; color:#00D4FF; line-height:1.3;">
                    GST Input<br>Reconciliation System
                </div>
                <div style="font-size:0.65rem; color:#A78BFA; margin-top:2px;">
                    Enterprise Edition v1.0
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ── User info ────────────────────────────────────────────────────
        username = st.session_state.get("username", "")
        full_name = st.session_state.get("full_name", username)
        role = st.session_state.get("role", "user")
        role_badge_color = "#00D4FF" if role == "admin" else "#A78BFA"

        st.markdown(
            f"""
            <div style="background:rgba(26,26,46,0.6); border-radius:10px;
                 padding:10px 12px; margin-bottom:12px;
                 border:1px solid rgba(0,212,255,0.1);">
                <div style="font-size:0.85rem; color:#EAEAEA; font-weight:600;">
                    👤 {full_name}
                </div>
                <div style="font-size:0.72rem; color:#64748B;">@{username}</div>
                <span style="background:{role_badge_color}22; color:{role_badge_color};
                      border:1px solid {role_badge_color}44; border-radius:4px;
                      padding:1px 8px; font-size:0.68rem; font-weight:600;">
                    {role.upper()}
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ── Navigation ───────────────────────────────────────────────────
        st.markdown(
            "<div style='font-size:0.68rem; color:#374151; "
            "letter-spacing:1px; margin-bottom:6px;'>NAVIGATION</div>",
            unsafe_allow_html=True,
        )

        current_page = st.session_state.get("current_page", "Dashboard")

        for page_name, icon in PAGES.items():
            is_active = current_page == page_name
            btn_style = (
                "border-color: #00D4FF !important; color: #00D4FF !important; "
                "background: rgba(0,212,255,0.1) !important;"
                if is_active else ""
            )
            label = f"{icon} {page_name}"
            if st.button(label, key=f"nav_{page_name}", use_container_width=True):
                st.session_state["current_page"] = page_name
                st.rerun()

        # Admin-only pages
        if role == "admin":
            st.markdown(
                "<div style='font-size:0.68rem; color:#374151; "
                "letter-spacing:1px; margin:12px 0 6px 0;'>ADMIN</div>",
                unsafe_allow_html=True,
            )
            for page_name, icon in ADMIN_ONLY_PAGES.items():
                if st.button(f"{icon} {page_name}", key=f"nav_{page_name}", use_container_width=True):
                    st.session_state["current_page"] = page_name
                    st.rerun()

        st.divider()

        # ── Logout ────────────────────────────────────────────────────────
        if st.button("🚪 Logout", use_container_width=True, key="logout_btn"):
            log_event("LOGOUT", f"User '{username}' logged out.")
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            # Clean temp files on logout
            try:
                cleanup_temp_files(older_than_hours=0)
            except Exception:
                pass
            st.rerun()

        # ── Sidebar Footer ─────────────────────────────────────────────────
        st.markdown(
            """
            <div style="text-align:center; color:#1E293B; font-size:0.65rem; margin-top:20px;">
                Prepared & Developed by<br>
                <span style="color:#374151;">Karthik LVN</span><br>
                © 2026 All Rights Reserved
            </div>
            """,
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# Page Footer
# ---------------------------------------------------------------------------

def render_footer() -> None:
    """Render the global page footer."""
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div style="border-top:1px solid rgba(0,212,255,0.1); padding-top:12px;
             text-align:center; color:#374151; font-size:0.74rem; margin-top:20px;">
            <strong style="color:#4B5563;">GST Input Reconciliation System</strong>
            &nbsp;·&nbsp; Enterprise Edition v1.0 &nbsp;·&nbsp;
            Prepared & Developed by <strong style="color:#00D4FF;">Karthik LVN</strong><br>
            © 2026 Karthik LVN &nbsp;·&nbsp; All Rights Reserved &nbsp;·&nbsp;
            Developed using Python &amp; Streamlit
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Page Router
# ---------------------------------------------------------------------------

def route_page(page: str) -> None:
    """Dispatch to the appropriate page renderer."""

    if page == "Dashboard":
        from modules.dashboard import render_home_dashboard
        render_home_dashboard()

    elif page == "Upload Data":
        render_upload_and_template_page()

    elif page == "Column Mapping":
        from modules.mapping import render_mapping_page
        render_mapping_page()

    elif page == "Reconcile":
        from modules.matching import render_reconciliation_page
        render_reconciliation_page()

    elif page == "Reports":
        from modules.reports import render_reports_page
        render_reports_page()

    elif page == "Settings":
        from modules.settings import render_settings_page
        render_settings_page()

    elif page == "User Management":
        from modules.authentication import render_user_management_page
        render_user_management_page()

    else:
        # Fallback to dashboard
        from modules.dashboard import render_home_dashboard
        render_home_dashboard()


# ---------------------------------------------------------------------------
# Upload + Template Download Page
# ---------------------------------------------------------------------------

def render_upload_and_template_page() -> None:
    """Upload page with downloadable Excel templates and close button."""
    import io
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    # ── Close / Done button ───────────────────────────────────────────
    hdr, close_col = st.columns([8, 1])
    hdr.markdown(
        "<h2 style='color:#00D4FF; margin:0;'>&#128228; Upload Data</h2>",
        unsafe_allow_html=True,
    )
    if close_col.button("✖ Close", key="close_upload"):
        st.session_state["current_page"] = "Dashboard"
        st.rerun()

    st.markdown(
        "<p style='color:#94A3B8;'>Step 1: Download the template &nbsp;|&nbsp; "
        "Step 2: Fill your data &nbsp;|&nbsp; Step 3: Upload both files</p>",
        unsafe_allow_html=True,
    )

    # ── Template Download Section ────────────────────────────────────
    st.markdown(
        "<h4 style='color:#A78BFA;'>&#11015; Download Templates</h4>",
        unsafe_allow_html=True,
    )

    PR_COLUMNS = [
        "vendor_name", "gstin", "invoice_number", "invoice_date",
        "taxable_value", "cgst", "sgst", "igst", "cess",
        "total_gst", "invoice_value",
    ]
    GSTR_COLUMNS = [
        "vendor_name", "gstin", "invoice_number", "invoice_date",
        "taxable_value", "cgst", "sgst", "igst", "cess",
        "total_gst", "invoice_value",
    ]

    def make_template(columns: list, sheet_name: str, sample_rows: int = 3) -> bytes:
        """Create a styled Excel template with sample rows."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

        # Header row styling
        header_fill = PatternFill("solid", fgColor="0D1B2A")
        header_font = Font(name="Calibri", bold=True, color="00D4FF", size=11)
        center = Alignment(horizontal="center", vertical="center")

        for ci, col in enumerate(columns, 1):
            cell = ws.cell(1, ci, col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(ci)
            ].width = max(len(col) + 4, 18)

        # Sample rows
        samples = {
            "vendor_name":   ["ABC Traders", "XYZ Pvt Ltd", "PQR Enterprises"],
            "gstin":         ["29ABCDE1234F1Z5", "27XYZPQ5678G1Z3", "33PQRST9012H1Z7"],
            "invoice_number":["INV-001", "INV-002", "INV-003"],
            "invoice_date":  ["01-04-2024", "05-04-2024", "10-04-2024"],
            "taxable_value": [100000, 250000, 75000],
            "cgst":          [9000, 22500, 6750],
            "sgst":          [9000, 22500, 6750],
            "igst":          [0, 0, 0],
            "cess":          [0, 0, 0],
            "total_gst":     [18000, 45000, 13500],
            "invoice_value": [118000, 295000, 88500],
        }
        row_fill = PatternFill("solid", fgColor="1A1A2E")
        row_font = Font(name="Calibri", color="EAEAEA", size=10)

        for ri in range(sample_rows):
            for ci, col in enumerate(columns, 1):
                val = samples.get(col, ["", "", ""])
                cell = ws.cell(ri + 2, ci, val[ri] if ri < len(val) else "")
                cell.font = row_font
                cell.fill = row_fill

        ws.row_dimensions[1].height = 22
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    tc1, tc2 = st.columns(2)
    with tc1:
        st.markdown(
            "<div style='background:rgba(0,212,255,0.05); border:1px solid #00D4FF33; "
            "border-radius:10px; padding:14px; text-align:center;'>"
            "<div style='color:#00D4FF; font-weight:700;'>&#128196; Purchase Register Template</div>"
            "<div style='color:#64748B; font-size:0.8rem; margin:6px 0;'>11 standard columns with sample data</div>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.download_button(
            label="⬇ Download PR Template",
            data=make_template(PR_COLUMNS, "Purchase Register"),
            file_name="Purchase_Register_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_pr_template",
        )

    with tc2:
        st.markdown(
            "<div style='background:rgba(167,139,250,0.05); border:1px solid #A78BFA33; "
            "border-radius:10px; padding:14px; text-align:center;'>"
            "<div style='color:#A78BFA; font-weight:700;'>&#128196; GSTR-2B Template</div>"
            "<div style='color:#64748B; font-size:0.8rem; margin:6px 0;'>11 standard columns with sample data</div>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.download_button(
            label="⬇ Download GSTR-2B Template",
            data=make_template(GSTR_COLUMNS, "GSTR-2B"),
            file_name="GSTR2B_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_gstr_template",
        )

    st.divider()

    # ── File Upload Section ───────────────────────────────────────────
    st.markdown(
        "<h4 style='color:#A78BFA;'>&#11014; Upload Your Files</h4>",
        unsafe_allow_html=True,
    )
    from modules.upload import render_upload_page
    render_upload_page(show_header=False)

    st.divider()
    # ── Proceed button ─────────────────────────────────────────────
    pr_done   = st.session_state.get("pr_df") is not None
    gstr_done = st.session_state.get("gstr2b_df") is not None
    if pr_done and gstr_done:
        st.success("✅ Both files uploaded! Ready to proceed.")
        if st.button("▶ Next: Column Mapping", type="primary", use_container_width=True, key="upload_next"):
            st.session_state["current_page"] = "Column Mapping"
            st.rerun()
    else:
        missing = []
        if not pr_done:   missing.append("Purchase Register")
        if not gstr_done: missing.append("GSTR-2B")
        st.info(f"Please upload: {', '.join(missing)}")



# ---------------------------------------------------------------------------
# About Page
# ---------------------------------------------------------------------------

def render_about_page() -> None:
    """Render the About page."""
    settings = st.session_state.get("app_settings", {})

    st.markdown(
        f"""
        <div style="text-align:center; padding:40px 0;">
            <div style="font-size:4rem; margin-bottom:16px;">📊</div>
            <div style="font-size:2rem; font-weight:800; color:#00D4FF;">
                GST Input Reconciliation System
            </div>
            <div style="font-size:1rem; color:#A78BFA; margin-top:6px; margin-bottom:4px;">
                Enterprise Edition v1.0
            </div>
            <div style="font-size:0.85rem; color:#64748B;">
                Prepared & Developed by
                <span style="color:#00D4FF; font-weight:700;">Karthik LVN</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.divider()

    a1, a2 = st.columns(2)

    with a1:
        st.markdown("### 🚀 Application Info")
        info = {
            "Version": "1.0 Enterprise Edition",
            "Framework": "Python + Streamlit",
            "Matching Engine": "5-Tier (Exact + Fuzzy)",
            "Max File Size": "500 MB",
            "Supported Formats": "Excel (.xlsx, .xls), CSV",
            "PDF Engine": "ReportLab",
            "Chart Engine": "Plotly",
            "Fuzzy Matching": "RapidFuzz",
            "Reports": "12-Sheet Excel + PDF + CSV",
        }
        for k, v in info.items():
            st.markdown(
                f"<div style='display:flex; justify-content:space-between; "
                f"padding:6px 0; border-bottom:1px solid rgba(0,212,255,0.1);'>"
                f"<span style='color:#94A3B8;'>{k}</span>"
                f"<span style='color:#EAEAEA; font-weight:600;'>{v}</span>"
                f"</div>",
                unsafe_allow_html=True,
            )

    with a2:
        st.markdown("### ✨ Features")
        features = [
            "5-tier intelligent matching engine",
            "Fuzzy matching with confidence scoring",
            "Automatic column detection & mapping",
            "Complete data cleaning pipeline",
            "GSTIN format + checksum validation",
            "12-sheet styled Excel workbook export",
            "Professional ReportLab PDF reports",
            "Interactive Plotly analytics dashboard",
            "Admin/User role-based access control",
            "Admin approval workflow for new users",
            "SQLite audit log for all actions",
            "Vendor & monthly summary analytics",
            "Global search & multi-filter support",
            "Supports 500K+ invoices efficiently",
        ]
        for f in features:
            st.markdown(
                f"<div style='color:#94A3B8; padding:4px 0;'>✅ {f}</div>",
                unsafe_allow_html=True,
            )

    st.divider()

    st.markdown(
        """
        <div style="text-align:center; color:#374151; font-size:0.8rem; padding:20px;">
            GST Input Reconciliation System &nbsp;·&nbsp; Enterprise Edition v1.0<br>
            Prepared & Developed by <strong style="color:#00D4FF;">Karthik LVN</strong><br>
            © 2026 Karthik LVN &nbsp;·&nbsp; All Rights Reserved<br>
            Developed using Python &amp; Streamlit &nbsp;·&nbsp; Confidential
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Main Application Entry
# ---------------------------------------------------------------------------

def main() -> None:
    """Main application entry point."""

    # ── Inject global styles ────────────────────────────────────────────────
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)

    # ── Initialize session ──────────────────────────────────────────────────
    init_session_state()

    # ── Splash Screen disabled for fast startup ──────────────────────────
    # (was causing 3+ second delay on first load)
    st.session_state["splash_shown"] = True

    # ── Load settings into session state ──────────────────────────────────
    if st.session_state.get("app_settings") is None:
        st.session_state["app_settings"] = load_settings()

    # ── Authentication gate (DISABLED — re-enable later) ──────────────────
    # if not st.session_state.get("authenticated", False):
    #     render_login_page()
    #     return

    # Auto-login as admin (no login required)
    if not st.session_state.get("authenticated", False):
        st.session_state["authenticated"] = True
        st.session_state["username"]      = "admin"
        st.session_state["role"]          = "admin"
        st.session_state["full_name"]     = "Karthik LVN"
        st.session_state["current_page"]  = "Dashboard"

    # ── Authenticated: render main app ─────────────────────────────────────
    render_sidebar()

    current_page = st.session_state.get("current_page", "Dashboard")
    route_page(current_page)

    render_footer()

    # ── Periodic cleanup (runs silently) ───────────────────────────────────
    if st.session_state.get("app_settings", {}).get("auto_cleanup", True):
        try:
            cleanup_temp_files(older_than_hours=24)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Entry
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    main()
