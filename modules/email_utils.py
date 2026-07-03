"""
GST Input Reconciliation System – Enterprise Edition
Email Utilities Module
Prepared & Developed by Karthik LVN

Provides:
  - SMTP email sending via Gmail or custom server
  - Password reset notification emails
  - Admin alert emails
"""

import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional

from modules.utils import setup_logging

logger = setup_logging()

APP_NAME = "GST Input Reconciliation System"
BRAND   = "Karthik LVN"


# ---------------------------------------------------------------------------
# SMTP Send Helper
# ---------------------------------------------------------------------------

def send_email(
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    from_addr: str,
    to_addr: str,
    subject: str,
    html_body: str,
    use_tls: bool = True,
) -> tuple[bool, str]:
    """
    Send an HTML email via SMTP.

    Args:
        smtp_host:     SMTP server hostname (e.g. smtp.gmail.com)
        smtp_port:     SMTP port (587 for TLS, 465 for SSL)
        smtp_user:     Login username
        smtp_password: Login password / App Password
        from_addr:     Sender email address
        to_addr:       Recipient email address
        subject:       Email subject line
        html_body:     HTML content of the email
        use_tls:       Use STARTTLS (True for port 587)

    Returns:
        (success: bool, message: str)
    """
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = f"{APP_NAME} <{from_addr}>"
        msg["To"]      = to_addr

        msg.attach(MIMEText(html_body, "html", "utf-8"))

        if use_tls:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=15)
            server.ehlo()
            server.starttls()
            server.ehlo()
        else:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=15)

        server.login(smtp_user, smtp_password)
        server.sendmail(from_addr, to_addr, msg.as_string())
        server.quit()

        logger.info(f"Email sent to {to_addr}: {subject}")
        return True, "Email sent successfully."

    except smtplib.SMTPAuthenticationError:
        msg = "SMTP authentication failed. Check your email and App Password."
        logger.error(msg)
        return False, msg
    except smtplib.SMTPException as e:
        msg = f"SMTP error: {e}"
        logger.error(msg)
        return False, msg
    except Exception as e:
        msg = f"Email sending failed: {e}"
        logger.error(msg)
        return False, msg


# ---------------------------------------------------------------------------
# Email Templates
# ---------------------------------------------------------------------------

def _base_template(title: str, body_html: str) -> str:
    """Wrap content in the branded email template."""
    year = datetime.date.today().year
    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>
  body {{ margin:0; padding:0; background:#0A0A1A; font-family:'Segoe UI',sans-serif; }}
  .container {{ max-width:600px; margin:30px auto; background:#1A1A2E;
               border-radius:16px; overflow:hidden;
               border:1px solid rgba(0,212,255,0.2); }}
  .header {{ background:linear-gradient(135deg,#0D1B2A,#1A1A2E);
             padding:28px 32px; text-align:center;
             border-bottom:2px solid rgba(0,212,255,0.3); }}
  .header h1 {{ margin:0; font-size:22px; color:#00D4FF; font-weight:800; }}
  .header p  {{ margin:6px 0 0; font-size:13px; color:#A78BFA; }}
  .body {{ padding:28px 32px; color:#EAEAEA; line-height:1.7; }}
  .body h2 {{ color:#00D4FF; font-size:18px; margin-top:0; }}
  .body p  {{ color:#CBD5E1; font-size:14px; margin:10px 0; }}
  .btn {{ display:inline-block; background:linear-gradient(135deg,#00D4FF,#0099BB);
          color:#000 !important; font-weight:700; font-size:14px;
          padding:12px 28px; border-radius:8px; text-decoration:none;
          margin:16px 0; }}
  .otp-box {{ background:#0A0A1A; border:2px solid #00D4FF44;
              border-radius:12px; padding:20px; text-align:center; margin:20px 0; }}
  .otp-code {{ font-size:36px; font-weight:900; color:#00D4FF;
               letter-spacing:12px; font-family:monospace; }}
  .divider {{ border:none; border-top:1px solid #1E293B; margin:20px 0; }}
  .footer {{ background:#0A0A1A; padding:16px 32px; text-align:center;
             font-size:11px; color:#374151; border-top:1px solid #1E293B; }}
  .footer strong {{ color:#4B5563; }}
  .warning {{ background:#1A0F0A; border:1px solid #FB923C44;
              border-radius:8px; padding:12px 16px; color:#FB923C;
              font-size:13px; margin:12px 0; }}
</style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>&#128202; {APP_NAME}</h1>
      <p>Prepared &amp; Developed by <strong>{BRAND}</strong></p>
    </div>
    <div class="body">
      {body_html}
    </div>
    <div class="footer">
      <strong>{APP_NAME}</strong> &bull; Enterprise Edition v1.0<br>
      Prepared &amp; Developed by <strong>{BRAND}</strong><br>
      &copy; {year} {BRAND} &bull; All Rights Reserved<br>
      <em>This is an automated message. Do not reply.</em>
    </div>
  </div>
</body>
</html>
"""


def build_password_reset_email(username: str, full_name: str, temp_password: str) -> tuple[str, str]:
    """
    Build the password reset email sent by admin.

    Returns:
        (subject, html_body)
    """
    subject = f"[{APP_NAME}] Your Password Has Been Reset"

    body = f"""
<h2>&#128274; Password Reset</h2>
<p>Hello <strong>{full_name}</strong>,</p>
<p>Your password for the <strong>{APP_NAME}</strong> has been reset by the administrator.</p>

<div class="otp-box">
  <p style="color:#94A3B8; font-size:13px; margin:0 0 8px;">Your Temporary Password</p>
  <div class="otp-code">{temp_password}</div>
</div>

<div class="warning">
  &#9888; Please log in immediately and change your password from the Settings page.
  This temporary password should not be shared.
</div>

<p style="color:#94A3B8; font-size:13px;">
  Username: <strong style="color:#EAEAEA;">{username}</strong><br>
  Reset Time: <strong style="color:#EAEAEA;">{datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")}</strong>
</p>

<hr class="divider">
<p style="font-size:12px; color:#64748B;">
  If you did not request this password reset, please contact your administrator immediately.
</p>
"""
    return subject, _base_template(subject, body)


def build_approval_email(username: str, full_name: str) -> tuple[str, str]:
    """Build the account approval notification email."""
    subject = f"[{APP_NAME}] Your Account Has Been Approved"
    body = f"""
<h2>&#10003; Account Approved</h2>
<p>Hello <strong>{full_name}</strong>,</p>
<p>Great news! Your account for <strong>{APP_NAME}</strong> has been approved by the administrator.</p>
<p>You can now log in with your registered credentials:</p>
<p style="background:#0A0A1A; padding:12px; border-radius:8px; border:1px solid #00D4FF44;">
  <strong style="color:#00D4FF;">Username:</strong>
  <span style="color:#EAEAEA; font-family:monospace;">{username}</span>
</p>
<p style="color:#94A3B8; font-size:13px;">
  Use the password you set during registration.
</p>
<hr class="divider">
<p style="font-size:12px; color:#64748B;">
  If you did not register for this system, please contact your administrator.
</p>
"""
    return subject, _base_template(subject, body)


def build_reset_request_email(
    admin_name: str,
    requesting_user: str,
    requesting_fullname: str,
) -> tuple[str, str]:
    """Build the email sent to admin when a user requests a password reset."""
    subject = f"[{APP_NAME}] Password Reset Requested by {requesting_user}"
    body = f"""
<h2>&#128274; Password Reset Request</h2>
<p>Hello <strong>{admin_name}</strong>,</p>
<p>A user has requested a password reset:</p>
<p style="background:#0A0A1A; padding:12px; border-radius:8px; border:1px solid #FBBF2444;">
  <strong style="color:#FBBF24;">Username:</strong>
  <span style="color:#EAEAEA;">{requesting_user}</span><br>
  <strong style="color:#FBBF24;">Full Name:</strong>
  <span style="color:#EAEAEA;">{requesting_fullname}</span><br>
  <strong style="color:#FBBF24;">Requested At:</strong>
  <span style="color:#EAEAEA;">{datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")}</span>
</p>
<p>Please log in to the <strong>User Management</strong> page to set a temporary password for this user.</p>
<hr class="divider">
<p style="font-size:12px; color:#64748B;">
  This is an automated alert from {APP_NAME}.
</p>
"""
    return subject, _base_template(subject, body)


# ---------------------------------------------------------------------------
# Send via stored settings
# ---------------------------------------------------------------------------

def send_via_settings(
    to_addr: str,
    subject: str,
    html_body: str,
    settings: dict,
) -> tuple[bool, str]:
    """
    Send email using SMTP settings from the app settings dict.

    Args:
        to_addr:  Recipient address.
        subject:  Email subject.
        html_body: HTML body.
        settings: App settings dict (from modules.settings).

    Returns:
        (success, message)
    """
    smtp_host = settings.get("smtp_host", "")
    smtp_port = int(settings.get("smtp_port", 587))
    smtp_user = settings.get("smtp_user", "")
    smtp_pass = settings.get("smtp_password", "")
    from_addr = settings.get("smtp_from", smtp_user)
    use_tls   = settings.get("smtp_tls", True)

    if not smtp_host or not smtp_user or not smtp_pass:
        return False, "SMTP not configured. Please set up email in Settings → Email Configuration."

    return send_email(
        smtp_host=smtp_host,
        smtp_port=smtp_port,
        smtp_user=smtp_user,
        smtp_password=smtp_pass,
        from_addr=from_addr,
        to_addr=to_addr,
        subject=subject,
        html_body=html_body,
        use_tls=use_tls,
    )
