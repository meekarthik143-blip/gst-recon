# GST Input Reconciliation System
## Enterprise Edition v1.0

> Prepared & Developed by **Karthik LVN**

---

## Overview

A production-ready, enterprise-grade GST Input Tax Credit (ITC) Reconciliation
web application built with **Python 3.12+** and **Streamlit**.

Automates the reconciliation of your **Purchase Register** against **GSTR-2B**
data with intelligent matching, fuzzy search, and professional reporting.

---

## Features

| Category | Feature |
|---|---|
| **Matching** | 5-tier engine: GSTIN+Invoice → Date → GST Amount → Taxable Value → Fuzzy |
| **Cleaning** | GSTIN validation, invoice normalization, date standardization, dedup |
| **Mapping** | Auto-detect + manual column mapping, save/load profiles |
| **Status** | 15+ status types (Perfect Match, Missing, GST Diff, Duplicate, etc.) |
| **Analytics** | Interactive Plotly charts: pie, bar, line, heatmap |
| **Reports** | 12-sheet Excel, PDF (ReportLab), CSV exports with full branding |
| **Auth** | Admin + User roles, admin approval workflow for new registrations |
| **Audit** | SQLite audit log: login, upload, process, export events |
| **Performance** | Supports 100K–500K invoices with vectorized pandas operations |
| **Security** | No permanent storage by default, temp file cleanup, GSTIN validation |

---

## Project Structure

```
GST-Recon/
├── app.py                          # Main entry point (splash, routing, sidebar)
├── requirements.txt
├── README.md
├── .streamlit/
│   └── config.toml                 # Dark theme configuration
├── modules/
│   ├── __init__.py
│   ├── utils.py                    # Shared helpers, validators, formatters
│   ├── authentication.py           # Login, registration, admin approval
│   ├── audit.py                    # SQLite audit log
│   ├── settings.py                 # Company master, app preferences
│   ├── upload.py                   # File upload (xlsx, csv, 500MB)
│   ├── cleaning.py                 # Data cleaning pipeline
│   ├── mapping.py                  # Column auto-mapping + manual override
│   ├── matching.py                 # 5-tier reconciliation engine
│   ├── reconciliation.py           # Status classification, KPIs, results page
│   ├── dashboard.py                # KPI cards, Plotly charts
│   └── reports.py                  # Excel (12 sheets), PDF, CSV exports
├── sample_data/
│   ├── generate_samples.py         # Generate test data
│   ├── sample_purchase_register.xlsx
│   └── sample_gstr2b.xlsx
└── data/                           # Auto-created at runtime
    ├── uploads/
    ├── temp/
    ├── reports/
    ├── logs/
    ├── users.json                  # User accounts (hashed passwords)
    ├── settings.json               # App settings
    ├── audit.db                    # SQLite audit database
    └── mapping_profile.json        # Saved column mappings
```

---

## Quick Start

### 1. Install Python 3.12+

Download from [python.org](https://www.python.org/downloads/)

### 2. Create a Virtual Environment

```powershell
cd "d:\LVNP\AI Innovation\GST-Recon"
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install Dependencies

```powershell
pip install -r requirements.txt
```

### 4. Generate Sample Data (Optional)

```powershell
python sample_data/generate_samples.py
```

### 5. Run the Application

```powershell
streamlit run app.py
```

Open your browser at **http://localhost:8501**

---

## Default Credentials

| Role | Username | Password |
|---|---|---|
| **Admin** | `admin` | `Admin@2026` |

> ⚠️ Change the admin password after first login via Settings → User Management.

### New User Registration
1. Click **Register** on the login page
2. Fill in details (password must have 1 uppercase, 1 digit, 1 special character)
3. Wait for Admin approval
4. Admin approves via **User Management** page
5. User can now log in

---

## Usage Workflow

```
1. Login (admin or approved user)
     ↓
2. Upload Purchase Register & GSTR-2B files
     ↓
3. Map columns to standard fields (auto-detected)
     ↓
4. Run Reconciliation (configure match threshold)
     ↓
5. Review Results (tabs: Perfect Match, Missing, GST Diff, Duplicates…)
     ↓
6. Export Reports (Excel 12-sheet / PDF / CSV)
```

---

## Password Policy

- Minimum 8 characters
- At least 1 uppercase letter (A–Z)
- At least 1 digit (0–9)
- At least 1 special character (@$!%*#?&)

---

## Reconciliation Statuses

| Status | Meaning | Action |
|---|---|---|
| Perfect Match | All fields match | No action required |
| Missing in Books | In GSTR-2B, not in PR | Book purchase entry |
| Missing in GSTR-2B | In PR, not in GSTR-2B | Follow up with vendor |
| GST Difference | GST amount mismatch | Request amendment |
| Taxable Difference | Taxable value mismatch | Verify invoice amount |
| Date Difference | Invoice date mismatch | Verify date with vendor |
| Duplicate | Duplicate invoice found | Remove duplicate |
| Fuzzy Match | Probable match (fuzzy) | Confirm manually |
| Manual Review | Needs human review | Manual verification |

---

## Export Reports

### Excel Workbook (12 Sheets)
1. Summary | 2. Matched | 3. Missing in Books | 4. Missing in GSTR-2B
5. GST Difference | 6. Taxable Difference | 7. Duplicate | 8. Manual Review
9. Vendor Summary | 10. Monthly Summary | 11. Dashboard Stats | 12. Audit Trail

All sheets include branded header (Karthik LVN) and footer.

### PDF Report (ReportLab)
- Multi-page professional report
- Header/footer on every page with branding
- Executive summary, KPI table, detailed sections

### CSV
- Master CSV, Missing in Books, Missing in GSTR-2B

---

## Deployment on Streamlit Cloud

1. Push the project to a **GitHub repository**
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → Select your repo
4. Set **Main file path**: `app.py`
5. Click **Deploy**

> **Note**: For Streamlit Cloud, add your secrets in `.streamlit/secrets.toml` and ensure all dependencies are in `requirements.txt`.

---

## Security Notes

- Passwords are hashed with SHA-256 (salted)
- No data is permanently stored unless "Save Data" is enabled in Settings
- Temp files auto-deleted after 24 hours
- All user actions logged to SQLite audit database

---

## Tech Stack

| Library | Purpose | Version |
|---|---|---|
| streamlit | Web UI framework | ≥1.35 |
| pandas | Data processing | ≥2.2 |
| numpy | Numerical operations | ≥1.26 |
| rapidfuzz | Fuzzy string matching | ≥3.9 |
| openpyxl | Excel read/write | ≥3.1 |
| reportlab | PDF generation | ≥4.1 |
| plotly | Interactive charts | ≥5.22 |
| loguru | Structured logging | ≥0.7 |
| python-dateutil | Date parsing | ≥2.9 |

---

## Branding

All reports, exports, and pages include:

> **GST Input Reconciliation System**
> Prepared & Developed by **Karthik LVN**
> © 2026 Karthik LVN · All Rights Reserved · Developed using Python & Streamlit

---

## License

© 2026 Karthik LVN · All Rights Reserved

This application is developed for commercial use. Unauthorized copying,
distribution, or modification is prohibited.
