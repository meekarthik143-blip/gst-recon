"""
GST Input Reconciliation System – Enterprise Edition
Settings Module
Prepared & Developed by Karthik LVN

Provides:
  - Company Master management
  - Application preferences
  - Settings persistence via JSON
  - Streamlit settings page
"""

from pathlib import Path
from typing import Any

import streamlit as st

from modules.utils import (
    get_project_root,
    load_json_config,
    save_json_config,
    validate_gstin,
    get_financial_years,
    get_current_financial_year,
    setup_logging,
)
from modules.audit import log_event

logger = setup_logging()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SETTINGS_FILE: Path = get_project_root() / "data" / "settings.json"

FINANCIAL_YEARS = get_financial_years(2020, 2030)
CURRENT_FY = get_current_financial_year()

DEFAULT_SETTINGS: dict = {
    "company_name": "My Company Pvt Ltd",
    "gstin": "",
    "address": "",
    "financial_year": CURRENT_FY,
    "default_match_percentage": 85,
    "theme": "dark",
    "currency": "INR",
    "date_format": "DD-MM-YYYY",
    "export_format": "xlsx",
    "max_file_size_mb": 500,
    "auto_cleanup": True,
    "save_data": False,
    "decimal_places": 2,
    "gst_tolerance": 1.0,
    "show_branding": True,
    "rows_per_page": 100,
    # ── Email / SMTP Configuration ──────────────────────────────────────────
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_user": "meekarthik143@gmail.com",
    "smtp_password": "",           # Set your Gmail App Password here
    "smtp_from": "meekarthik143@gmail.com",
    "smtp_tls": True,
    "admin_email": "meekarthik143@gmail.com",
    "email_notifications": True,
}


# ---------------------------------------------------------------------------
# Settings I/O
# ---------------------------------------------------------------------------

def initialize_settings() -> None:
    """Create settings.json with defaults if it does not exist."""
    if not SETTINGS_FILE.exists():
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        save_json_config(SETTINGS_FILE, DEFAULT_SETTINGS)
        logger.info("Settings initialized with defaults.")


def load_settings() -> dict:
    """
    Load application settings from JSON, merging with defaults for any missing keys.

    Returns:
        Settings dictionary.
    """
    initialize_settings()
    stored = load_json_config(SETTINGS_FILE)
    if not isinstance(stored, dict):
        stored = {}

    # Merge: default values fill in any keys missing from the stored file
    merged = DEFAULT_SETTINGS.copy()
    merged.update(stored)
    return merged


def save_settings(settings: dict) -> None:
    """
    Persist settings to JSON and update Streamlit session state.

    Args:
        settings: Settings dictionary to save.
    """
    try:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        save_json_config(SETTINGS_FILE, settings)
        st.session_state["app_settings"] = settings
        log_event("SETTINGS", "Application settings updated.")
        logger.info("Settings saved.")
    except Exception as e:
        logger.error(f"Failed to save settings: {e}")


def get_setting(key: str, default: Any = None) -> Any:
    """
    Retrieve a single setting value from session state (or disk).

    Args:
        key:     Setting key.
        default: Default value if key is not found.

    Returns:
        Setting value.
    """
    settings = st.session_state.get("app_settings") or load_settings()
    return settings.get(key, default)


def get_company_name() -> str:
    """Shortcut: return the configured company name."""
    return get_setting("company_name", "My Company Pvt Ltd")


def get_financial_year() -> str:
    """Shortcut: return the configured financial year."""
    return get_setting("financial_year", CURRENT_FY)


# ---------------------------------------------------------------------------
# Streamlit Settings Page
# ---------------------------------------------------------------------------

def render_settings_page() -> None:
    """Render the application settings page."""

    # Admin-only guard for destructive settings
    is_admin = st.session_state.get("role") == "admin"

    st.markdown(
        "<h2 style='color:#00D4FF;'>⚙️ Settings</h2>",
        unsafe_allow_html=True,
    )

    current = st.session_state.get("app_settings") or load_settings()

    # ── Section 1: Company Master ─────────────────────────────────────────
    with st.expander("🏢 Company Master", expanded=True):
        c1, c2 = st.columns(2)

        company_name = c1.text_input(
            "Company Name *",
            value=current.get("company_name", ""),
            key="s_company_name",
        )

        gstin_input = c2.text_input(
            "Company GSTIN",
            value=current.get("gstin", ""),
            max_chars=15,
            key="s_gstin",
        )

        if gstin_input:
            if validate_gstin(gstin_input.upper()):
                c2.success("✅ Valid GSTIN format")
            else:
                c2.error("❌ Invalid GSTIN format")

        address = st.text_area(
            "Company Address",
            value=current.get("address", ""),
            height=80,
            key="s_address",
        )

        fy_options = FINANCIAL_YEARS
        fy_default = current.get("financial_year", CURRENT_FY)
        try:
            fy_idx = fy_options.index(fy_default)
        except ValueError:
            fy_idx = 0

        financial_year = st.selectbox(
            "Financial Year",
            fy_options,
            index=fy_idx,
            key="s_fy",
        )

    # ── Section 2: Reconciliation ─────────────────────────────────────────
    with st.expander("🔧 Reconciliation Settings", expanded=True):
        r1, r2 = st.columns(2)

        match_pct = r1.slider(
            "Default Match Percentage (%)",
            min_value=80,
            max_value=100,
            value=current.get("default_match_percentage", 85),
            step=5,
            key="s_match_pct",
            help="Minimum fuzzy match score to consider an invoice as matched.",
        )

        gst_tolerance = r2.number_input(
            "GST Difference Tolerance (₹)",
            min_value=0.0,
            max_value=100.0,
            value=float(current.get("gst_tolerance", 1.0)),
            step=0.5,
            key="s_gst_tolerance",
            help="GST amounts differing by less than this value are treated as matched.",
        )

    # ── Section 3: Display ─────────────────────────────────────────────────
    with st.expander("🎨 Display Settings"):
        d1, d2, d3 = st.columns(3)

        theme = d1.selectbox(
            "Theme",
            ["dark", "light"],
            index=0 if current.get("theme", "dark") == "dark" else 1,
            key="s_theme",
        )

        currency = d2.selectbox(
            "Currency",
            ["INR", "USD", "EUR"],
            index=["INR", "USD", "EUR"].index(current.get("currency", "INR")),
            key="s_currency",
        )

        date_format = d3.selectbox(
            "Date Format",
            ["DD-MM-YYYY", "MM-DD-YYYY", "YYYY-MM-DD"],
            index=["DD-MM-YYYY", "MM-DD-YYYY", "YYYY-MM-DD"].index(
                current.get("date_format", "DD-MM-YYYY")
            ),
            key="s_date_format",
        )

        decimal_places = st.slider(
            "Decimal Places",
            min_value=0,
            max_value=4,
            value=current.get("decimal_places", 2),
            key="s_decimal",
        )

    # ── Section 4: Export ─────────────────────────────────────────────────
    with st.expander("📤 Export Settings"):
        e1, e2 = st.columns(2)

        export_format = e1.selectbox(
            "Default Export Format",
            ["xlsx", "csv", "pdf"],
            index=["xlsx", "csv", "pdf"].index(current.get("export_format", "xlsx")),
            key="s_export_fmt",
        )

        rows_per_page = e2.selectbox(
            "Rows per Page",
            [50, 100, 250, 500, 1000],
            index=[50, 100, 250, 500, 1000].index(
                current.get("rows_per_page", 100)
            ),
            key="s_rows_per_page",
        )

    # ── Section 5: Data Management ────────────────────────────────────────
    with st.expander("🗄️ Data Management", expanded=False):
        auto_cleanup = st.toggle(
            "Auto-cleanup temporary files (after 24h)",
            value=current.get("auto_cleanup", True),
            key="s_auto_cleanup",
        )

        if is_admin:
            save_data = st.toggle(
                "Permanently save uploaded data",
                value=current.get("save_data", False),
                key="s_save_data",
            )
            if save_data:
                st.warning(
                    "⚠️ Enabling this will store uploaded invoice data on disk. "
                    "Ensure compliance with your data privacy policy."
                )
        else:
            save_data = current.get("save_data", False)
            st.info("🔒 Data management settings are restricted to Admin users.")

    # ── Save Button ────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    col_save, col_reset = st.columns([1, 1])

    if col_save.button("💾 Save Settings", type="primary", use_container_width=True):
        updated_settings = {
            "company_name": company_name,
            "gstin": gstin_input.upper() if gstin_input else "",
            "address": address,
            "financial_year": financial_year,
            "default_match_percentage": match_pct,
            "gst_tolerance": gst_tolerance,
            "theme": theme,
            "currency": currency,
            "date_format": date_format,
            "decimal_places": decimal_places,
            "export_format": export_format,
            "rows_per_page": rows_per_page,
            "auto_cleanup": auto_cleanup,
            "save_data": save_data if is_admin else current.get("save_data", False),
            "max_file_size_mb": current.get("max_file_size_mb", 500),
            "show_branding": True,
        }
        save_settings(updated_settings)
        st.success("✅ Settings saved successfully!")
        st.rerun()

    if col_reset.button("🔄 Reset to Defaults", use_container_width=True):
        save_settings(DEFAULT_SETTINGS.copy())
        st.warning("Settings reset to defaults.")
        st.rerun()

    # ── Current Settings Summary ──────────────────────────────────────────
    st.divider()
    st.subheader("Current Settings Summary")

    summary_data = {
        "Setting": [
            "Company Name", "GSTIN", "Financial Year",
            "Default Match %", "Currency", "Date Format",
            "Export Format", "Auto Cleanup", "Theme",
        ],
        "Value": [
            current.get("company_name", "-"),
            current.get("gstin", "Not Set"),
            current.get("financial_year", "-"),
            f"{current.get('default_match_percentage', 85)}%",
            current.get("currency", "INR"),
            current.get("date_format", "DD-MM-YYYY"),
            current.get("export_format", "xlsx").upper(),
            "Yes" if current.get("auto_cleanup", True) else "No",
            current.get("theme", "dark").capitalize(),
        ],
    }

    import pandas as pd
    st.dataframe(
        pd.DataFrame(summary_data),
        use_container_width=True,
        hide_index=True,
    )

    # ── Email / SMTP Configuration (Admin only) ───────────────────────────
    if is_admin:
        st.divider()
        st.markdown(
            "<h3 style='color:#00D4FF;'>📧 Email Configuration (Gmail SMTP)</h3>",
            unsafe_allow_html=True,
        )
        st.markdown(
            """
            <div style='background:rgba(0,212,255,0.05); border:1px solid #00D4FF33;
                 border-radius:10px; padding:14px 18px; margin-bottom:16px; font-size:0.83rem; color:#94A3B8;'>
            <strong style='color:#00D4FF;'>📌 Gmail Setup (one-time)</strong><br>
            1. Go to <strong>myaccount.google.com → Security → 2-Step Verification</strong> → Enable it<br>
            2. Then go to <strong>App Passwords</strong> → Select App: "Mail" → Generate<br>
            3. Copy the 16-character App Password and paste it below (not your regular Gmail password)<br>
            4. Save settings and click <strong>Test Email</strong> to verify
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("email_config_form"):
            ec1, ec2 = st.columns(2)
            smtp_host = ec1.text_input(
                "SMTP Host", value=current.get("smtp_host", "smtp.gmail.com"), key="smtp_host"
            )
            smtp_port = ec2.number_input(
                "SMTP Port", value=int(current.get("smtp_port", 587)),
                min_value=1, max_value=65535, key="smtp_port"
            )
            smtp_user = st.text_input(
                "📧 Gmail Address", value=current.get("smtp_user", "meekarthik143@gmail.com"), key="smtp_user"
            )
            smtp_pass = st.text_input(
                "🔑 App Password (16 chars)", type="password",
                value=current.get("smtp_password", ""),
                placeholder="xxxx xxxx xxxx xxxx",
                key="smtp_pass",
            )
            admin_email = st.text_input(
                "📬 Admin Alert Email (receives reset requests)",
                value=current.get("admin_email", "meekarthik143@gmail.com"),
                key="admin_email_input",
            )
            smtp_tls = st.checkbox(
                "Use TLS (recommended for Gmail port 587)",
                value=current.get("smtp_tls", True), key="smtp_tls"
            )

            em_save, em_test = st.columns(2)
            save_email = em_save.form_submit_button("💾 Save Email Config", type="primary", use_container_width=True)
            test_email = em_test.form_submit_button("📨 Send Test Email", use_container_width=True)

        if save_email:
            updated = load_settings()
            updated.update({
                "smtp_host": smtp_host,
                "smtp_port": int(smtp_port),
                "smtp_user": smtp_user,
                "smtp_password": smtp_pass,
                "smtp_from": smtp_user,
                "smtp_tls": smtp_tls,
                "admin_email": admin_email,
                "email_notifications": True,
            })
            save_settings(updated)
            st.success("✅ Email configuration saved!")
            st.rerun()

        if test_email:
            try:
                from modules.email_utils import send_email
                test_settings = load_settings()
                ok, msg = send_email(
                    smtp_host=test_settings.get("smtp_host", "smtp.gmail.com"),
                    smtp_port=int(test_settings.get("smtp_port", 587)),
                    smtp_user=test_settings.get("smtp_user", ""),
                    smtp_password=test_settings.get("smtp_password", ""),
                    from_addr=test_settings.get("smtp_from", test_settings.get("smtp_user", "")),
                    to_addr=test_settings.get("admin_email", test_settings.get("smtp_user", "")),
                    subject="[GST Recon] Test Email — SMTP Working!",
                    html_body="<h2 style='color:#00D4FF;'>✅ Email configuration is working correctly!</h2>"
                              "<p>Your GST Input Reconciliation System can send emails successfully.</p>"
                              "<p style='color:#64748B;'>Prepared & Developed by Karthik LVN</p>",
                    use_tls=test_settings.get("smtp_tls", True),
                )
                if ok:
                    st.success(f"✅ Test email sent to {test_settings.get('admin_email', '')}! Check your inbox.")
                else:
                    st.error(f"❌ {msg}")
            except Exception as e:
                st.error(f"❌ Error: {e}")
