"""
GST Input Reconciliation System – Enterprise Edition
Utility Functions Module
Prepared & Developed by Karthik LVN

Provides shared helpers used across all other modules:
  - Logging setup
  - Path management
  - GSTIN validation
  - Date parsing / standardization
  - Currency formatting
  - String cleaning
  - File management
  - Password hashing
  - JSON config I/O
  - DataFrame utilities
"""

import os
import re
import json
import hashlib
import logging
import datetime
from pathlib import Path
from typing import Optional, Any

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Loguru logger – write to file AND stderr
# ---------------------------------------------------------------------------
try:
    from loguru import logger as _loguru_logger

    def setup_logging() -> Any:
        """Configure loguru logger writing to data/logs/app.log."""
        log_dir = get_project_root() / "data" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / "app.log"

        _loguru_logger.remove()  # remove default stderr handler
        _loguru_logger.add(
            log_file,
            rotation="10 MB",
            retention="30 days",
            level="INFO",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {module}:{line} | {message}",
        )
        _loguru_logger.add(
            lambda msg: print(msg, end=""),
            level="WARNING",
            colorize=True,
        )
        return _loguru_logger

    logger = _loguru_logger

except ImportError:
    # Fallback to standard logging if loguru not available
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)

    def setup_logging():  # type: ignore[misc]
        return logger


# ---------------------------------------------------------------------------
# Path Management
# ---------------------------------------------------------------------------

def get_project_root() -> Path:
    """Return the absolute path to the project root directory."""
    # This file lives at <root>/modules/utils.py  →  parent.parent = root
    return Path(__file__).resolve().parent.parent


def ensure_directories() -> None:
    """Create required data directories if they do not already exist."""
    root = get_project_root()
    dirs = [
        root / "data",
        root / "data" / "uploads",
        root / "data" / "temp",
        root / "data" / "reports",
        root / "data" / "logs",
        root / "assets",
        root / "sample_data",
    ]
    for directory in dirs:
        directory.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# GSTIN Validation
# ---------------------------------------------------------------------------

GSTIN_REGEX = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"

# Mapping for GSTIN checksum calculation
_GSTIN_CHARSET = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def validate_gstin(gstin: str) -> bool:
    """
    Validate GSTIN format using regex.

    Args:
        gstin: GSTIN string to validate.

    Returns:
        True if format is valid, False otherwise.
    """
    if not isinstance(gstin, str):
        return False
    gstin = gstin.strip().upper()
    return bool(re.match(GSTIN_REGEX, gstin))


def validate_gstin_checksum(gstin: str) -> bool:
    """
    Validate GSTIN with full checksum verification (MOD-36 algorithm).

    Args:
        gstin: GSTIN string to validate.

    Returns:
        True if GSTIN passes format AND checksum validation.
    """
    if not validate_gstin(gstin):
        return False

    gstin = gstin.strip().upper()
    factor = 2
    total = 0
    code_point_count = len(_GSTIN_CHARSET)

    for char in gstin[:-1]:
        if char not in _GSTIN_CHARSET:
            return False
        digit = _GSTIN_CHARSET.index(char) * factor
        factor = 1 if factor == 2 else 2
        digit = (digit // code_point_count) + (digit % code_point_count)
        total += digit

    check_code_point = (code_point_count - (total % code_point_count)) % code_point_count
    check_char = _GSTIN_CHARSET[check_code_point]
    return gstin[-1] == check_char


# ---------------------------------------------------------------------------
# Date Utilities
# ---------------------------------------------------------------------------

DATE_FORMATS = [
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%Y-%m-%d",
    "%m/%d/%Y",
    "%d %b %Y",
    "%d-%b-%Y",
    "%b %Y",
    "%Y%m%d",
    "%d.%m.%Y",
    "%d %B %Y",
    "%B %Y",
]


def parse_date(date_str: Any) -> Optional[datetime.date]:
    """
    Try to parse a date string using multiple known formats.

    Args:
        date_str: Input date value (string, datetime, or date).

    Returns:
        datetime.date object if parsed successfully, else None.
    """
    if date_str is None:
        return None
    if isinstance(date_str, datetime.datetime):
        return date_str.date()
    if isinstance(date_str, datetime.date):
        return date_str

    date_str = str(date_str).strip()
    if not date_str or date_str.lower() in ("nan", "none", "nat", ""):
        return None

    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue

    # Try python-dateutil as a last resort
    try:
        from dateutil import parser as du_parser
        return du_parser.parse(date_str, dayfirst=True).date()
    except Exception:
        return None


def standardize_date(date_str: Any) -> str:
    """
    Standardize a date to DD-MM-YYYY string format.

    Args:
        date_str: Raw date value.

    Returns:
        Standardized date string or original string if parsing fails.
    """
    parsed = parse_date(date_str)
    if parsed is not None:
        return parsed.strftime("%d-%m-%Y")
    return str(date_str) if date_str is not None else ""


# ---------------------------------------------------------------------------
# Currency & Number Utilities
# ---------------------------------------------------------------------------

def format_currency(amount: float, currency: str = "INR") -> str:
    """
    Format a float as Indian currency (₹1,23,456.78).

    Args:
        amount: Numeric amount.
        currency: Currency code ('INR', 'USD', 'EUR').

    Returns:
        Formatted currency string.
    """
    try:
        amount = float(amount)
    except (ValueError, TypeError):
        return "₹0.00"

    symbols = {"INR": "₹", "USD": "$", "EUR": "€"}
    symbol = symbols.get(currency, "₹")

    if currency == "INR":
        # Indian number formatting: X,XX,XX,XX,...
        abs_amount = abs(amount)
        integer_part = int(abs_amount)
        decimal_part = round(abs_amount - integer_part, 2)

        s = str(integer_part)
        if len(s) > 3:
            last3 = s[-3:]
            rest = s[:-3]
            groups = [rest[max(0, i - 2):i] for i in range(len(rest), 0, -2)]
            groups.reverse()
            s = ",".join(groups) + "," + last3
        formatted = f"{symbol}{'-' if amount < 0 else ''}{s}.{int(decimal_part * 100):02d}"
    else:
        formatted = f"{symbol}{amount:,.2f}"

    return formatted


def round_gst(value: Any) -> float:
    """
    Round a GST value to 2 decimal places safely.

    Args:
        value: Numeric value to round.

    Returns:
        Float rounded to 2 decimal places, or 0.0 on error.
    """
    try:
        return round(float(value), 2)
    except (ValueError, TypeError):
        return 0.0


def safe_float(val: Any) -> float:
    """
    Safely convert any value to float.

    Args:
        val: Input value.

    Returns:
        Float value or 0.0 if conversion fails.
    """
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0


# ---------------------------------------------------------------------------
# String Cleaning Utilities
# ---------------------------------------------------------------------------

_SPECIAL_CHARS_RE = re.compile(r"[/*\\\s]+")
_HIDDEN_CHARS_RE = re.compile(r"[\x00-\x1F\x7F-\x9F\u200b-\u200f\ufeff]")


def clean_invoice_number(inv: Any) -> str:
    """
    Normalize an invoice number: remove /, *, spaces, special chars, uppercase.

    Args:
        inv: Raw invoice number.

    Returns:
        Cleaned, uppercase invoice number string.
    """
    if inv is None:
        return ""
    s = str(inv).strip()
    s = _HIDDEN_CHARS_RE.sub("", s)           # remove hidden/control chars
    s = re.sub(r"[/*\\\s]", "", s)             # remove slashes, asterisks, spaces
    s = re.sub(r"[^\w\-]", "", s)              # keep word chars and hyphens
    return s.upper()


def clean_vendor_name(name: Any) -> str:
    """
    Clean a vendor/supplier name: strip whitespace, title-case, normalize spaces.

    Args:
        name: Raw vendor name.

    Returns:
        Cleaned vendor name string.
    """
    if name is None:
        return ""
    s = str(name).strip()
    s = _HIDDEN_CHARS_RE.sub("", s)
    s = re.sub(r"\s+", " ", s)                # collapse multiple spaces
    return s.strip().title()


def normalize_gstin(gstin: Any) -> str:
    """
    Normalize a GSTIN: strip whitespace, uppercase, remove spaces/dashes.

    Args:
        gstin: Raw GSTIN value.

    Returns:
        Normalized GSTIN string (15 chars if valid).
    """
    if gstin is None:
        return ""
    s = str(gstin).strip()
    s = _HIDDEN_CHARS_RE.sub("", s)
    s = re.sub(r"[\s\-]", "", s)
    return s.upper()


# ---------------------------------------------------------------------------
# File Management
# ---------------------------------------------------------------------------

def get_temp_path(filename: str) -> Path:
    """
    Return a path inside the data/temp directory with a timestamp prefix.

    Args:
        filename: Original filename.

    Returns:
        Full path for temp file.
    """
    ensure_directories()
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return get_project_root() / "data" / "temp" / f"{ts}_{filename}"


def get_report_path(filename: str) -> Path:
    """
    Return a path inside the data/reports directory.

    Args:
        filename: Report filename.

    Returns:
        Full path for report file.
    """
    ensure_directories()
    return get_project_root() / "data" / "reports" / filename


def cleanup_temp_files(older_than_hours: int = 24) -> None:
    """
    Delete temp files older than the specified number of hours.

    Args:
        older_than_hours: Age threshold in hours. Default is 24.
    """
    temp_dir = get_project_root() / "data" / "temp"
    if not temp_dir.exists():
        return

    cutoff = datetime.datetime.now() - datetime.timedelta(hours=older_than_hours)
    for f in temp_dir.iterdir():
        if f.is_file():
            mtime = datetime.datetime.fromtimestamp(f.stat().st_mtime)
            if mtime < cutoff:
                try:
                    f.unlink()
                    logger.info(f"Deleted temp file: {f.name}")
                except OSError as e:
                    logger.warning(f"Could not delete temp file {f.name}: {e}")


# ---------------------------------------------------------------------------
# JSON Config I/O
# ---------------------------------------------------------------------------

def load_json_config(path: Path) -> dict:
    """
    Load a JSON configuration file.

    Args:
        path: Path to the JSON file.

    Returns:
        Parsed dictionary or empty dict if file doesn't exist/is invalid.
    """
    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        logger.error(f"Failed to load config from {path}: {e}")
    return {}


def save_json_config(path: Path, data: dict) -> None:
    """
    Save a dictionary as a JSON configuration file.

    Args:
        path: Path to write the JSON file.
        data: Dictionary to serialize.
    """
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False, default=str)
    except OSError as e:
        logger.error(f"Failed to save config to {path}: {e}")


# ---------------------------------------------------------------------------
# Password Hashing
# ---------------------------------------------------------------------------

_SALT_PREFIX = "GST_LVNP_SALT_2026"


def hash_password(password: str) -> str:
    """
    Hash a password using SHA-256 with a fixed salt prefix.

    Args:
        password: Plain-text password.

    Returns:
        Hex-encoded SHA-256 hash string.
    """
    salted = f"{_SALT_PREFIX}_{password}"
    return hashlib.sha256(salted.encode("utf-8")).hexdigest()


def verify_password(password: str, hashed: str) -> bool:
    """
    Verify a plain-text password against a stored hash.

    Args:
        password: Plain-text password to check.
        hashed: Previously hashed password.

    Returns:
        True if password matches, False otherwise.
    """
    return hash_password(password) == hashed


# ---------------------------------------------------------------------------
# DataFrame Utilities
# ---------------------------------------------------------------------------

def safe_merge(
    left: pd.DataFrame,
    right: pd.DataFrame,
    **kwargs: Any,
) -> pd.DataFrame:
    """
    Perform a pandas merge with error handling.

    Args:
        left: Left DataFrame.
        right: Right DataFrame.
        **kwargs: Additional arguments passed to pd.merge.

    Returns:
        Merged DataFrame or empty DataFrame on failure.
    """
    try:
        return pd.merge(left, right, **kwargs)
    except Exception as e:
        logger.error(f"Merge failed: {e}")
        return pd.DataFrame()


def df_to_display(df: pd.DataFrame, max_rows: int = 10_000) -> pd.DataFrame:
    """
    Prepare a DataFrame for Streamlit display: cap rows, format numerics.

    Args:
        df: Input DataFrame.
        max_rows: Maximum rows to return.

    Returns:
        Display-ready DataFrame.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    display_df = df.head(max_rows).copy()

    # Format float columns to 2 decimal places for display
    for col in display_df.select_dtypes(include="float64").columns:
        display_df[col] = display_df[col].apply(lambda x: round(x, 2) if pd.notna(x) else 0.0)

    return display_df


# ---------------------------------------------------------------------------
# Email Validation
# ---------------------------------------------------------------------------

_EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")


def validate_email(email: str) -> bool:
    """
    Validate an email address format.

    Args:
        email: Email string to validate.

    Returns:
        True if valid email format.
    """
    if not isinstance(email, str):
        return False
    return bool(_EMAIL_RE.match(email.strip()))


# ---------------------------------------------------------------------------
# Financial Year Utilities
# ---------------------------------------------------------------------------

def get_financial_years(start: int = 2020, end: int = 2029) -> list[str]:
    """
    Generate a list of financial year strings (e.g., '2023-24').

    Args:
        start: Starting calendar year.
        end: Ending calendar year.

    Returns:
        List of FY strings.
    """
    return [f"{y}-{str(y + 1)[-2:]}" for y in range(start, end)]


def get_current_financial_year() -> str:
    """
    Determine the current Indian financial year.

    Returns:
        Current FY string, e.g., '2025-26'.
    """
    today = datetime.date.today()
    if today.month >= 4:
        return f"{today.year}-{str(today.year + 1)[-2:]}"
    else:
        return f"{today.year - 1}-{str(today.year)[-2:]}"


def get_month_name(month_num: int) -> str:
    """Return abbreviated month name for a month number (1-12)."""
    months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ]
    if 1 <= month_num <= 12:
        return months[month_num - 1]
    return str(month_num)
