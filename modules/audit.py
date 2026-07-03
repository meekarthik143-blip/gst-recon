"""
GST Input Reconciliation System – Enterprise Edition
Audit Log Module
Prepared & Developed by Karthik LVN

Provides:
  - SQLite-backed audit log
  - Event logging for all user actions
  - Audit log viewer page in Streamlit
"""

import sqlite3
import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from modules.utils import get_project_root, setup_logging

logger = setup_logging()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

AUDIT_DB: Path = get_project_root() / "data" / "audit.db"

EVENT_TYPES = [
    "LOGIN",
    "LOGOUT",
    "UPLOAD",
    "PROCESS",
    "EXPORT",
    "REPORT",
    "SETTINGS",
    "USER_MGMT",
    "ERROR",
    "INFO",
]

EVENT_COLORS = {
    "LOGIN": "#00D4FF",
    "LOGOUT": "#64748B",
    "UPLOAD": "#A78BFA",
    "PROCESS": "#34D399",
    "EXPORT": "#FBBF24",
    "REPORT": "#F97316",
    "SETTINGS": "#60A5FA",
    "USER_MGMT": "#F472B6",
    "ERROR": "#EF4444",
    "INFO": "#94A3B8",
}


# ---------------------------------------------------------------------------
# Database Initialization
# ---------------------------------------------------------------------------

def initialize_audit_db() -> None:
    """
    Create the SQLite audit database and table if they do not already exist.
    """
    try:
        AUDIT_DB.parent.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(str(AUDIT_DB)) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS audit_log (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp   TEXT    NOT NULL,
                    username    TEXT    DEFAULT 'system',
                    role        TEXT    DEFAULT 'system',
                    event_type  TEXT    NOT NULL,
                    description TEXT,
                    file_name   TEXT    DEFAULT '',
                    ip_address  TEXT    DEFAULT 'localhost',
                    session_id  TEXT    DEFAULT ''
                )
                """
            )
            conn.commit()
        logger.info("Audit DB initialized.")
    except Exception as e:
        logger.error(f"Failed to initialize audit DB: {e}")


# ---------------------------------------------------------------------------
# Event Logging
# ---------------------------------------------------------------------------

def log_event(
    event_type: str,
    description: str,
    username: Optional[str] = None,
    file_name: Optional[str] = None,
) -> None:
    """
    Insert an audit log event into the database.

    Automatically reads username and role from st.session_state when not
    explicitly provided.

    Args:
        event_type:   One of the EVENT_TYPES constants.
        description:  Human-readable description of the event.
        username:     Override username (default: from session_state).
        file_name:    Associated file name, if any.
    """
    try:
        initialize_audit_db()

        # Resolve username / role from session state if not supplied
        if username is None:
            username = st.session_state.get("username", "system")
        role = st.session_state.get("role", "system")
        session_id = st.session_state.get("session_id", "")

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with sqlite3.connect(str(AUDIT_DB)) as conn:
            conn.execute(
                """
                INSERT INTO audit_log
                    (timestamp, username, role, event_type, description, file_name, ip_address, session_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    timestamp,
                    username,
                    role,
                    event_type.upper(),
                    description,
                    file_name or "",
                    "localhost",
                    session_id,
                ),
            )
            conn.commit()
    except Exception as e:
        # Silently log to file – never crash the main app due to audit failure
        logger.error(f"Audit log failed: {e}")


# ---------------------------------------------------------------------------
# Audit Log Retrieval
# ---------------------------------------------------------------------------

def get_audit_log(
    limit: int = 500,
    username: Optional[str] = None,
    event_type: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
) -> pd.DataFrame:
    """
    Query the audit log with optional filters.

    Args:
        limit:       Maximum number of rows to return.
        username:    Filter by username.
        event_type:  Filter by event type.
        start_date:  Filter by date >= start_date (YYYY-MM-DD).
        end_date:    Filter by date <= end_date (YYYY-MM-DD).

    Returns:
        DataFrame of audit records.
    """
    try:
        initialize_audit_db()

        query = "SELECT * FROM audit_log WHERE 1=1"
        params: list = []

        if username and username != "All":
            query += " AND username = ?"
            params.append(username)
        if event_type and event_type != "All":
            query += " AND event_type = ?"
            params.append(event_type)
        if start_date:
            query += " AND timestamp >= ?"
            params.append(start_date)
        if end_date:
            query += " AND timestamp <= ?"
            params.append(end_date + " 23:59:59")

        query += " ORDER BY id DESC LIMIT ?"
        params.append(limit)

        with sqlite3.connect(str(AUDIT_DB)) as conn:
            df = pd.read_sql_query(query, conn, params=params)
        return df

    except Exception as e:
        logger.error(f"Failed to retrieve audit log: {e}")
        return pd.DataFrame()


def get_audit_summary() -> dict:
    """
    Return summary statistics for the audit log.

    Returns:
        Dict with total_events and events_by_type counts.
    """
    try:
        initialize_audit_db()
        with sqlite3.connect(str(AUDIT_DB)) as conn:
            total = conn.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0]
            rows = conn.execute(
                "SELECT event_type, COUNT(*) as cnt FROM audit_log GROUP BY event_type"
            ).fetchall()
        return {
            "total_events": total,
            "events_by_type": {r[0]: r[1] for r in rows},
        }
    except Exception as e:
        logger.error(f"Failed to get audit summary: {e}")
        return {"total_events": 0, "events_by_type": {}}


# ---------------------------------------------------------------------------
# Streamlit Page
# ---------------------------------------------------------------------------

def render_audit_log_page() -> None:
    """Render the Audit Log viewer page in Streamlit."""

    st.markdown(
        """
        <h2 style='color:#00D4FF; margin-bottom:4px;'>🔍 Audit Log</h2>
        <p style='color:#94A3B8; margin-top:0;'>Track all user activities, uploads, and system events</p>
        """,
        unsafe_allow_html=True,
    )

    # ── Summary cards ──────────────────────────────────────────────────────
    summary = get_audit_summary()
    total = summary.get("total_events", 0)
    by_type = summary.get("events_by_type", {})

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📋 Total Events", f"{total:,}")
    col2.metric("🔑 Logins", by_type.get("LOGIN", 0))
    col3.metric("📤 Uploads", by_type.get("UPLOAD", 0))
    col4.metric("📊 Exports", by_type.get("EXPORT", 0))

    st.divider()

    # ── Filters ────────────────────────────────────────────────────────────
    with st.expander("🔧 Filters", expanded=True):
        fc1, fc2, fc3, fc4 = st.columns(4)

        all_users = ["All"]
        try:
            with sqlite3.connect(str(AUDIT_DB)) as conn:
                users_result = conn.execute(
                    "SELECT DISTINCT username FROM audit_log ORDER BY username"
                ).fetchall()
            all_users += [r[0] for r in users_result]
        except Exception:
            pass

        filter_user = fc1.selectbox("👤 User", all_users, key="audit_user_filter")
        filter_event = fc2.selectbox(
            "📌 Event Type", ["All"] + EVENT_TYPES, key="audit_event_filter"
        )
        filter_start = fc3.date_input(
            "📅 From Date",
            value=datetime.date.today() - datetime.timedelta(days=30),
            key="audit_start",
        )
        filter_end = fc4.date_input(
            "📅 To Date", value=datetime.date.today(), key="audit_end"
        )
        limit = st.slider("Max Rows", 50, 2000, 500, 50, key="audit_limit")

    # ── Load data ──────────────────────────────────────────────────────────
    df = get_audit_log(
        limit=limit,
        username=filter_user if filter_user != "All" else None,
        event_type=filter_event if filter_event != "All" else None,
        start_date=str(filter_start),
        end_date=str(filter_end),
    )

    if df.empty:
        st.info("No audit records found for the selected filters.")
        return

    # ── Display dataframe ──────────────────────────────────────────────────
    st.markdown(f"**Showing {len(df):,} records**")

    display_cols = ["timestamp", "username", "role", "event_type", "description", "file_name"]
    display_df = df[display_cols] if all(c in df.columns for c in display_cols) else df

    # Apply background styling for event_type column
    def highlight_event(row: pd.Series) -> list[str]:
        color = EVENT_COLORS.get(row.get("event_type", ""), "#1A1A2E")
        return [
            f"background-color: {color}22; color: white" if col == "event_type" else ""
            for col in display_df.columns
        ]

    st.dataframe(
        display_df.style.apply(highlight_event, axis=1),
        use_container_width=True,
        height=450,
    )

    # ── Mini chart ─────────────────────────────────────────────────────────
    if by_type:
        import plotly.express as px

        chart_df = pd.DataFrame(
            list(by_type.items()), columns=["Event Type", "Count"]
        ).sort_values("Count", ascending=False)

        fig = px.bar(
            chart_df,
            x="Event Type",
            y="Count",
            title="Events by Type",
            color="Event Type",
            color_discrete_map={k: v for k, v in EVENT_COLORS.items()},
            template="plotly_dark",
        )
        fig.update_layout(
            showlegend=False,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            height=300,
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── Export ─────────────────────────────────────────────────────────────
    st.divider()
    csv_data = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Export Audit Log (CSV)",
        data=csv_data,
        file_name=f"audit_log_{datetime.date.today()}.csv",
        mime="text/csv",
        key="audit_export_btn",
    )
