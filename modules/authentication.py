"""
GST Input Reconciliation System – Enterprise Edition
Authentication Module
Prepared & Developed by Karthik LVN

Provides:
  - User registration with admin approval workflow
  - Admin user management (approve / reject / delete / reset password)
  - Forgot Password with admin reset + Gmail email notification
  - Secure login with SHA-256 hashed passwords
  - Streamlit login page and user management page
"""

import re
import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from modules.utils import (
    hash_password,
    verify_password,
    validate_email,
    load_json_config,
    save_json_config,
    get_project_root,
    setup_logging,
)
from modules.audit import log_event

logger = setup_logging()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

USERS_FILE: Path = get_project_root() / "data" / "users.json"

DEFAULT_USERS: list[dict] = [
    {
        "username": "admin",
        "password_hash": hash_password("Admin@2026"),
        "role": "admin",
        "approved": True,
        "full_name": "Administrator",
        "email": "admin@lvnp.com",
        "created_at": "2026-01-01",
        "last_login": None,
    }
]

PASSWORD_REGEX = re.compile(
    r"^(?=.*[A-Z])(?=.*[0-9])(?=.*[@$!%*#?&])[A-Za-z0-9@$!%*#?&]{8,}$"
)


# ---------------------------------------------------------------------------
# User Store Helpers
# ---------------------------------------------------------------------------

def initialize_users() -> None:
    """Create users.json with default admin account if it does not exist."""
    if not USERS_FILE.exists():
        USERS_FILE.parent.mkdir(parents=True, exist_ok=True)
        save_json_config(USERS_FILE, DEFAULT_USERS)
        logger.info("User store initialized with default admin.")


def load_users() -> list[dict]:
    """
    Load all users from the JSON user store.

    Returns:
        List of user dictionaries.
    """
    initialize_users()
    data = load_json_config(USERS_FILE)
    if isinstance(data, list):
        return data
    return []


def save_users(users: list[dict]) -> None:
    """
    Persist the users list to the JSON user store.

    Args:
        users: List of user dictionaries to save.
    """
    save_json_config(USERS_FILE, users)


def _find_user(username: str, users: list[dict]) -> Optional[dict]:
    """Return the user dict matching username, or None."""
    return next(
        (u for u in users if u.get("username", "").lower() == username.lower()), None
    )


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------

def authenticate(username: str, password: str) -> Optional[dict]:
    """
    Authenticate a user with username and password.

    Args:
        username: User's login name.
        password: Plain-text password.

    Returns:
        User dict on success, None on failure.
        Returns None with an error flag if account is pending approval.
    """
    users = load_users()
    user = _find_user(username, users)

    if user is None:
        return None

    if not verify_password(password, user.get("password_hash", "")):
        return None

    if not user.get("approved", False):
        # Return a special sentinel so the UI can show 'pending approval'
        return {"__pending__": True, "full_name": user.get("full_name", username)}

    # Update last_login timestamp
    for u in users:
        if u.get("username", "").lower() == username.lower():
            u["last_login"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            break
    save_users(users)
    log_event("LOGIN", f"User '{username}' logged in.", username=username)
    return user


# ---------------------------------------------------------------------------
# Registration
# ---------------------------------------------------------------------------

def _validate_password_strength(password: str) -> tuple[bool, str]:
    """Check password strength rules."""
    if len(password) < 8:
        return False, "Password must be at least 8 characters long."
    if not re.search(r"[A-Z]", password):
        return False, "Password must contain at least one uppercase letter."
    if not re.search(r"[0-9]", password):
        return False, "Password must contain at least one digit."
    if not re.search(r"[@$!%*#?&]", password):
        return False, "Password must contain at least one special character (@$!%*#?&)."
    return True, ""


def register_user(
    username: str,
    password: str,
    full_name: str,
    email: str,
) -> tuple[bool, str]:
    """
    Register a new user. Account requires admin approval before login.

    Args:
        username:  Desired username (must be unique).
        password:  Plain-text password.
        full_name: User's full name.
        email:     User's email address.

    Returns:
        (True, success_message) or (False, error_message).
    """
    # Validate inputs
    if not username or not username.strip():
        return False, "Username cannot be empty."
    if not validate_email(email):
        return False, "Invalid email address format."

    ok, msg = _validate_password_strength(password)
    if not ok:
        return False, msg

    users = load_users()

    # Check uniqueness
    if _find_user(username, users) is not None:
        return False, f"Username '{username}' is already taken. Please choose another."

    new_user: dict = {
        "username": username.strip().lower(),
        "password_hash": hash_password(password),
        "role": "user",
        "approved": False,
        "full_name": full_name.strip(),
        "email": email.strip().lower(),
        "created_at": datetime.date.today().strftime("%Y-%m-%d"),
        "last_login": None,
    }

    users.append(new_user)
    save_users(users)
    log_event(
        "USER_MGMT",
        f"New user registration: '{username}' – awaiting admin approval.",
        username=username,
    )
    return True, "Registration successful! Your account is pending admin approval."


# ---------------------------------------------------------------------------
# Admin Operations
# ---------------------------------------------------------------------------

def approve_user(username: str, approved_by: str) -> bool:
    """
    Approve a pending user account.

    Args:
        username:    Username to approve.
        approved_by: Admin username performing the action.

    Returns:
        True if approved successfully.
    """
    users = load_users()
    for u in users:
        if u.get("username", "").lower() == username.lower():
            u["approved"] = True
            save_users(users)
            log_event(
                "USER_MGMT",
                f"User '{username}' approved by '{approved_by}'.",
                username=approved_by,
            )
            return True
    return False


def reject_user(username: str, rejected_by: str) -> bool:
    """
    Reject and remove a pending user account.

    Args:
        username:    Username to reject.
        rejected_by: Admin username performing the action.

    Returns:
        True if rejected successfully.
    """
    users = load_users()
    original_count = len(users)
    users = [u for u in users if u.get("username", "").lower() != username.lower()]

    if len(users) == original_count:
        return False  # Not found

    save_users(users)
    log_event(
        "USER_MGMT",
        f"User '{username}' rejected by '{rejected_by}'.",
        username=rejected_by,
    )
    return True


def delete_user(username: str, deleted_by: str) -> tuple[bool, str]:
    """
    Delete a user account. Admin cannot delete their own account.

    Args:
        username:   Username to delete.
        deleted_by: Admin username performing the action.

    Returns:
        (True, message) or (False, error_message).
    """
    if username.lower() == deleted_by.lower():
        return False, "You cannot delete your own admin account."

    users = load_users()
    original_count = len(users)
    users = [u for u in users if u.get("username", "").lower() != username.lower()]

    if len(users) == original_count:
        return False, f"User '{username}' not found."

    save_users(users)
    log_event(
        "USER_MGMT",
        f"User '{username}' deleted by '{deleted_by}'.",
        username=deleted_by,
    )
    return True, f"User '{username}' has been deleted."


def change_password(
    username: str, old_password: str, new_password: str
) -> tuple[bool, str]:
    """
    Change a user's password after verifying the old password.

    Args:
        username:     Username.
        old_password: Current plain-text password.
        new_password: New plain-text password.

    Returns:
        (True, message) or (False, error_message).
    """
    users = load_users()
    user = _find_user(username, users)

    if user is None:
        return False, "User not found."
    if not verify_password(old_password, user.get("password_hash", "")):
        return False, "Current password is incorrect."

    ok, msg = _validate_password_strength(new_password)
    if not ok:
        return False, msg

    for u in users:
        if u.get("username", "").lower() == username.lower():
            u["password_hash"] = hash_password(new_password)
            break

    save_users(users)
    log_event("USER_MGMT", f"Password changed for '{username}'.", username=username)
    return True, "Password changed successfully."


def get_pending_users() -> list[dict]:
    """Return users whose accounts are pending admin approval."""
    return [u for u in load_users() if not u.get("approved", False)]


def get_all_users() -> list[dict]:
    """Return all user accounts (for admin view)."""
    return load_users()


# ---------------------------------------------------------------------------
# Forgot Password & Admin Reset
# ---------------------------------------------------------------------------

def submit_reset_request(username_or_email: str) -> tuple[bool, str, dict]:
    """
    Submit a password reset request for a user.
    Stores the request flag in users.json for admin to see.

    Args:
        username_or_email: Username OR registered email address.

    Returns:
        (success, message, user_dict_or_empty)
    """
    users = load_users()
    user = None

    # Try matching by username first, then by email
    for u in users:
        if (
            u.get("username", "").lower() == username_or_email.lower()
            or u.get("email", "").lower() == username_or_email.lower()
        ):
            user = u
            break

    if user is None:
        return False, "No account found with that username or email.", {}

    if not user.get("approved", False):
        return False, "Account is not yet approved. Contact the administrator.", {}

    # Mark reset request on the user record
    for u in users:
        if u.get("username") == user["username"]:
            u["reset_requested"] = True
            u["reset_requested_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            break

    save_users(users)
    log_event("USER_MGMT", f"Password reset requested for '{user['username']}'.")
    return True, "Reset request submitted. The administrator will be notified.", user


def admin_reset_password(
    target_username: str,
    new_password: str,
    admin_username: str,
) -> tuple[bool, str]:
    """
    Admin forcefully resets a user's password.

    Args:
        target_username: The user whose password to reset.
        new_password:    New plain-text password (must pass strength check).
        admin_username:  Admin performing the reset.

    Returns:
        (success, message)
    """
    ok, msg = _validate_password_strength(new_password)
    if not ok:
        return False, msg

    users = load_users()
    user = _find_user(target_username, users)

    if user is None:
        return False, "User not found."

    for u in users:
        if u.get("username", "").lower() == target_username.lower():
            u["password_hash"] = hash_password(new_password)
            u["reset_requested"] = False
            u["reset_requested_at"] = None
            u["password_reset_by"] = admin_username
            u["password_reset_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            break

    save_users(users)
    log_event(
        "USER_MGMT",
        f"Admin '{admin_username}' reset password for '{target_username}'.",
    )
    return True, f"Password for '{target_username}' has been reset successfully."


def get_reset_requests() -> list[dict]:
    """Return all users who have an active password reset request."""
    return [
        u for u in load_users()
        if u.get("reset_requested", False)
    ]


# ---------------------------------------------------------------------------
# Streamlit Login Page
# ---------------------------------------------------------------------------

_LOGIN_CSS = """
<style>
/* ── Body ── */
[data-testid="stApp"] {
    background: linear-gradient(135deg, #0A0A1A 0%, #0D1B2A 50%, #0A0A1A 100%);
}
.login-wrapper {
    display: flex;
    justify-content: center;
    align-items: center;
    min-height: 75vh;
}
.login-card {
    background: rgba(26, 26, 46, 0.85);
    backdrop-filter: blur(16px);
    border: 1px solid rgba(0, 212, 255, 0.25);
    border-radius: 20px;
    padding: 40px 48px;
    max-width: 480px;
    width: 100%;
    box-shadow: 0 8px 32px rgba(0, 212, 255, 0.12), 0 2px 8px rgba(0,0,0,0.4);
}
.login-title {
    font-size: 1.7rem;
    font-weight: 700;
    color: #00D4FF;
    text-align: center;
    margin-bottom: 2px;
}
.login-sub {
    font-size: 0.82rem;
    color: #94A3B8;
    text-align: center;
    margin-bottom: 28px;
}
.brand-badge {
    display: inline-block;
    background: linear-gradient(90deg, #00D4FF22, #A78BFA22);
    border: 1px solid #00D4FF44;
    border-radius: 8px;
    padding: 4px 14px;
    font-size: 0.78rem;
    color: #A78BFA;
    margin-bottom: 20px;
}
</style>
"""


def render_login_page() -> None:
    """Render the professional login / registration page."""

    st.markdown(_LOGIN_CSS, unsafe_allow_html=True)

    # ── Branding header ────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center; padding: 24px 0 8px 0;">
            <div style="font-size:2.4rem; margin-bottom:6px;">📊</div>
            <div style="font-size:1.6rem; font-weight:800; color:#00D4FF; letter-spacing:0.5px;">
                GST Input Reconciliation System
            </div>
            <div style="font-size:0.85rem; color:#A78BFA; margin-top:4px; margin-bottom:8px;">
                Enterprise Edition v1.0
            </div>
            <div style="font-size:0.78rem; color:#64748B;">
                Prepared & Developed by <span style="color:#00D4FF; font-weight:600;">Karthik LVN</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Tabs: Login / Register / Forgot Password ───────────────────────────
    tab_login, tab_register, tab_forgot = st.tabs(
        ["🔑 Login", "📝 Register", "🔓 Forgot Password"]
    )

    # ── LOGIN TAB ──────────────────────────────────────────────────────────
    with tab_login:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input(
                "👤 Username",
                placeholder="Enter your username",
                key="login_username",
            )
            password = st.text_input(
                "🔒 Password",
                type="password",
                placeholder="Enter your password",
                key="login_password",
            )

            col_rem, col_btn = st.columns([1, 1])
            remember_me = col_rem.checkbox("Remember Me", key="remember_me")

            submitted = st.form_submit_button(
                "🚀 Login", use_container_width=True, type="primary"
            )

        if submitted:
            if not username or not password:
                st.error("Please enter both username and password.")
            else:
                result = authenticate(username.strip(), password)

                if result is None:
                    st.error("❌ Invalid username or password.")
                elif result.get("__pending__"):
                    st.warning(
                        f"⏳ Account for **{result.get('full_name', username)}** is "
                        "pending admin approval. Please wait."
                    )
                else:
                    import uuid
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = result["username"]
                    st.session_state["role"] = result["role"]
                    st.session_state["full_name"] = result.get("full_name", username)
                    st.session_state["session_id"] = str(uuid.uuid4())
                    st.session_state["login_time"] = datetime.datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                    st.session_state["current_page"] = "Dashboard"
                    st.rerun()

    # ── REGISTER TAB ──────────────────────────────────────────────────────
    with tab_register:
        with st.form("register_form", clear_on_submit=True):
            reg_full_name = st.text_input(
                "👤 Full Name", placeholder="Your full name", key="reg_full_name"
            )
            reg_email = st.text_input(
                "📧 Email", placeholder="your@email.com", key="reg_email"
            )
            reg_username = st.text_input(
                "🆔 Username", placeholder="Choose a username", key="reg_username"
            )
            reg_password = st.text_input(
                "🔒 Password",
                type="password",
                placeholder="Min 8 chars, 1 upper, 1 digit, 1 special",
                key="reg_password",
            )
            reg_confirm = st.text_input(
                "🔒 Confirm Password",
                type="password",
                placeholder="Repeat your password",
                key="reg_confirm",
            )

            reg_submitted = st.form_submit_button(
                "📝 Submit Registration", use_container_width=True, type="primary"
            )

        if reg_submitted:
            if not all([reg_full_name, reg_email, reg_username, reg_password, reg_confirm]):
                st.error("All fields are required.")
            elif reg_password != reg_confirm:
                st.error("Passwords do not match.")
            else:
                ok, msg = register_user(
                    username=reg_username.strip(),
                    password=reg_password,
                    full_name=reg_full_name.strip(),
                    email=reg_email.strip(),
                )
                if ok:
                    st.success(f"✅ {msg}")
                    st.info(
                        "📬 Your account is awaiting approval from the Administrator. "
                        "You will be able to login once approved."
                    )
                else:
                    st.error(f"❌ {msg}")

        st.markdown(
            "<div style='color:#64748B; font-size:0.76rem; margin-top:12px;'>"
            "Password must be at least 8 characters and contain:<br>"
            "&bull; One uppercase letter &nbsp;&bull; One digit &nbsp;&bull; One special character (@$!%*#?&)"
            "</div>",
            unsafe_allow_html=True,
        )

    # ── FORGOT PASSWORD TAB ────────────────────────────────────────────────
    with tab_forgot:
        st.markdown(
            """
            <div style='background:rgba(251,146,60,0.08); border:1px solid rgba(251,146,60,0.3);
                 border-radius:10px; padding:14px 18px; margin-bottom:16px;'>
                <div style='color:#FB923C; font-weight:700; font-size:0.95rem;'>🔓 Forgot Your Password?</div>
                <div style='color:#94A3B8; font-size:0.82rem; margin-top:4px;'>
                    Enter your username or registered email below. The administrator
                    will be notified and will reset your password.
                    You will receive an email once your password is reset.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("forgot_pwd_form", clear_on_submit=True):
            fp_input = st.text_input(
                "👤 Username or Registered Email",
                placeholder="Enter your username or email address",
                key="fp_input",
            )
            fp_submitted = st.form_submit_button(
                "📩 Submit Reset Request", use_container_width=True, type="primary"
            )

        if fp_submitted:
            if not fp_input.strip():
                st.error("Please enter your username or email.")
            else:
                ok, msg, user_data = submit_reset_request(fp_input.strip())
                if ok:
                    # Notify admin via email if SMTP configured
                    try:
                    
                        from modules.settings import load_settings
                        from modules.email_utils import (
                            build_reset_request_email,
                            send_via_settings,
                        )
                        settings = load_settings()
                        admin_email = settings.get("admin_email", settings.get("smtp_user", ""))
                        if admin_email:
                            subject, html = build_reset_request_email(
                                admin_name="Administrator",
                                requesting_user=user_data.get("username", fp_input),
                                requesting_fullname=user_data.get("full_name", fp_input),
                            )
                            send_via_settings(admin_email, subject, html, settings)
                    except Exception:
                        pass  # email failure is non-critical

                    st.success(
                        "✅ Reset request submitted successfully!\n\n"
                        "The administrator has been notified. "
                        "Please wait for them to reset your password — "
                        "you will receive an email once it's done."
                    )
                else:
                    st.error(f"❌ {msg}")

        st.markdown(
            "<div style='color:#64748B; font-size:0.76rem; margin-top:16px; text-align:center;'>"
            "If you continue to have issues, contact: "
            "<strong style='color:#00D4FF;'>meekarthik143@gmail.com</strong>"
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown(
        "<hr style='border-color:#1A1A2E; margin-top:32px;'>"
        "<div style='text-align:center; color:#374151; font-size:0.74rem;'>"
        "© 2026 Karthik LVN &nbsp;·&nbsp; All Rights Reserved &nbsp;·&nbsp; "
        "Developed using Python & Streamlit"
        "</div>",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Admin – User Management Page
# ---------------------------------------------------------------------------

def render_user_management_page() -> None:
    """Render the admin-only user management page."""

    # Guard: admin only
    if st.session_state.get("role") != "admin":
        st.error("🚫 Access Denied. Admin privileges required.")
        return

    current_admin = st.session_state.get("username", "admin")

    st.markdown(
        "<h2 style='color:#00D4FF;'>👥 User Management</h2>",
        unsafe_allow_html=True,
    )

    # ── Reset Password Requests ─────────────────────────────────────────────
    reset_requests = get_reset_requests()
    if reset_requests:
        st.markdown(
            f"<div style='background:rgba(0,212,255,0.1); border:1px solid #00D4FF44; "
            f"border-radius:8px; padding:10px 16px; margin-bottom:12px;'>"
            f"<span style='color:#00D4FF; font-weight:700;'>🔓 {len(reset_requests)} Password Reset Request(s)</span>"
            f"</div>",
            unsafe_allow_html=True,
        )
        for ru in reset_requests:
            with st.container(border=True):
                r1, r2, r3 = st.columns([3, 3, 2])
                r1.markdown(
                    f"**{ru.get('full_name','?')}** (`{ru.get('username','?')}`)"
                )
                r2.markdown(
                    f"📧 {ru.get('email','-')}  "
                    f"\n🕐 {ru.get('reset_requested_at','-')}"
                )

                with r3:
                    with st.form(f"reset_form_{ru['username']}"):
                        new_pwd = st.text_input(
                            "New Temp Password",
                            type="password",
                            key=f"new_pwd_{ru['username']}",
                            placeholder="Min 8 chars + upper + digit + special",
                        )
                        if st.form_submit_button("🔑 Set Password", type="primary", use_container_width=True):
                            if not new_pwd.strip():
                                st.error("Enter a password.")
                            else:
                                ok, msg = admin_reset_password(
                                    target_username=ru["username"],
                                    new_password=new_pwd.strip(),
                                    admin_username=current_admin,
                                )
                                if ok:
                                    # Send email notification to user
                                    try:
                                        from modules.settings import load_settings
                                        from modules.email_utils import (
                                            build_password_reset_email,
                                            send_via_settings,
                                        )
                                        settings = load_settings()
                                        user_email = ru.get("email", "")
                                        if user_email:
                                            subj, html = build_password_reset_email(
                                                username=ru["username"],
                                                full_name=ru.get("full_name", ru["username"]),
                                                temp_password=new_pwd.strip(),
                                            )
                                            esok, emsg = send_via_settings(
                                                user_email, subj, html, settings
                                            )
                                            if esok:
                                                st.success(f"{msg} — Email sent to {user_email}")
                                            else:
                                                st.success(msg)
                                                st.warning(f"Email not sent: {emsg}")
                                        else:
                                            st.success(msg)
                                    except Exception as e:
                                        st.success(msg)
                                        st.warning(f"Email error: {e}")
                                    st.rerun()
                                else:
                                    st.error(msg)
        st.divider()

    # ── Pending Approvals ──────────────────────────────────────────────────
    pending = get_pending_users()

    if pending:
        st.markdown(
            f"<div style='background:#F97316; color:#fff; border-radius:8px; "
            f"padding:8px 16px; font-weight:600; margin-bottom:12px;'>"
            f"⚠️ {len(pending)} Pending Approval(s)</div>",
            unsafe_allow_html=True,
        )

        for u in pending:
            with st.container():
                c1, c2, c3, c4 = st.columns([3, 2, 1, 1])
                c1.write(f"**{u.get('full_name', '-')}** (`{u.get('username', '-')}`)")
                c2.write(u.get("email", "-"))
                c3.write(u.get("created_at", "-"))

                approve_key = f"approve_{u['username']}"
                reject_key  = f"reject_{u['username']}"

                if c4.button("✅ Approve", key=approve_key, type="primary"):
                    if approve_user(u["username"], current_admin):
                        # Send approval email
                        try:
                            from modules.settings import load_settings
                            from modules.email_utils import build_approval_email, send_via_settings
                            settings = load_settings()
                            subj, html = build_approval_email(
                                username=u["username"],
                                full_name=u.get("full_name", u["username"]),
                            )
                            send_via_settings(u.get("email", ""), subj, html, settings)
                        except Exception:
                            pass
                        st.success(f"User '{u['username']}' approved!")
                        st.rerun()

                if c4.button("❌ Reject", key=reject_key):
                    if reject_user(u["username"], current_admin):
                        st.warning(f"User '{u['username']}' rejected and removed.")
                        st.rerun()

        st.divider()
    else:
        st.success("✅ No pending approval requests.")

    # ── All Users Table ────────────────────────────────────────────────────
    st.subheader("All Users")
    all_users = get_all_users()

    if not all_users:
        st.info("No users found.")
        return

    users_df = pd.DataFrame(all_users)
    display_cols = ["full_name", "username", "email", "role", "approved", "reset_requested", "last_login", "created_at"]
    display_df = users_df[[c for c in display_cols if c in users_df.columns]].copy()
    st.dataframe(display_df, use_container_width=True, height=300)

    # ── Delete User ─────────────────────────────────────────────────────────
    st.divider()
    st.subheader("Delete User")

    deletable = [
        u["username"] for u in all_users
        if u.get("username", "").lower() != current_admin.lower()
    ]

    if not deletable:
        st.info("No other users to delete.")
        return

    del_col1, del_col2 = st.columns([2, 1])
    user_to_delete = del_col1.selectbox(
        "Select user to delete", deletable, key="del_user_select"
    )
    confirm = del_col1.checkbox(
        f"I confirm deletion of '{user_to_delete}'", key="del_confirm"
    )
    if del_col2.button("🗑️ Delete User", key="del_user_btn", type="primary", disabled=not confirm):
        ok, msg = delete_user(user_to_delete, current_admin)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    # ── Change My Password ─────────────────────────────────────────────────
    st.divider()
    st.subheader("Change My Password")
    with st.form("change_pwd_form"):
        cp_old     = st.text_input("Current Password",     type="password", key="cp_old")
        cp_new     = st.text_input("New Password",         type="password", key="cp_new")
        cp_confirm = st.text_input("Confirm New Password", type="password", key="cp_confirm")
        cp_submitted = st.form_submit_button("🔒 Change Password", type="primary")

    if cp_submitted:
        if cp_new != cp_confirm:
            st.error("New passwords do not match.")
        else:
            ok, msg = change_password(current_admin, cp_old, cp_new)
            if ok:
                st.success(msg)
            else:
                st.error(msg)


