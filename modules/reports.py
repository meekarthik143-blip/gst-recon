"""
GST Input Reconciliation System – Enterprise Edition
Reports Export Module
Prepared & Developed by Karthik LVN

Provides:
  - Excel export (12-sheet styled workbook with branding)
  - PDF export (ReportLab multi-page report with header/footer)
  - CSV export
  - Report header/footer with Karthik LVN branding
"""

import io
import datetime
from typing import Optional

import pandas as pd
import streamlit as st

from modules.utils import (
    format_currency,
    safe_float,
    setup_logging,
    get_report_path,
)
from modules.audit import log_event
from modules.reconciliation import (
    get_vendor_summary,
    get_monthly_summary,
    get_kpi_summary,
    STATUS_COLORS,
    STATUS_PERFECT_MATCH,
    STATUS_MISSING_BOOKS,
    STATUS_MISSING_GSTR2B,
    STATUS_GST_DIFF,
    STATUS_TAXABLE_DIFF,
    STATUS_DUPLICATE,
    STATUS_MANUAL_REVIEW,
    STATUS_FUZZY_MATCH,
)

logger = setup_logging()

# ---------------------------------------------------------------------------
# Branding
# ---------------------------------------------------------------------------

BRAND_NAME = "Karthik LVN"
APP_NAME = "GST Input Reconciliation System"
APP_VERSION = "1.0 Enterprise Edition"
FOOTER_TEXT = f"© 2026 {BRAND_NAME} · All Rights Reserved · Developed using Python & Streamlit"
CONFIDENTIAL_TEXT = "CONFIDENTIAL REPORT — For Internal Use Only"


def _get_report_metadata() -> dict:
    """Return common metadata used in report headers."""
    settings = st.session_state.get("app_settings", {})
    return {
        "company_name": settings.get("company_name", "My Company Pvt Ltd"),
        "gstin": settings.get("gstin", ""),
        "financial_year": settings.get("financial_year", "2025-26"),
        "report_date": datetime.date.today().strftime("%d-%m-%Y"),
        "generated_at": datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        "generated_by": st.session_state.get("full_name", "System"),
    }


# ---------------------------------------------------------------------------
# Excel Export  (openpyxl)
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color: str) -> tuple:
    """Convert hex color to RGB tuple."""
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def export_excel(
    master_df: pd.DataFrame,
    recon_results: Optional[dict] = None,
) -> bytes:
    """
    Generate a styled 12-sheet Excel workbook.

    Sheets:
      1  Summary          2  Matched (Perfect)
      3  Missing in Books 4  Missing in GSTR-2B
      5  GST Difference   6  Taxable Difference
      7  Duplicate        8  Manual Review
      9  Vendor Summary  10  Monthly Summary
      11 Dashboard Stats 12  Audit Trail

    Args:
        master_df:     Master reconciliation DataFrame.
        recon_results: Original reconciliation results dict.

    Returns:
        Excel file content as bytes.
    """
    from openpyxl import Workbook
    from openpyxl.styles import (
        PatternFill, Font, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference

    meta = _get_report_metadata()

    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    # ── Color palette ────────────────────────────────────────────────────
    DARK_BG = "0A0A1A"
    HEADER_BG = "1A1A2E"
    ACCENT = "00D4FF"
    GREEN = "34D399"
    RED = "F87171"
    ORANGE = "FB923C"
    YELLOW = "FBBF24"
    PURPLE = "A78BFA"
    WHITE = "EAEAEA"
    SUBHEADER = "1E293B"

    def _header_fill(color: str = HEADER_BG) -> PatternFill:
        return PatternFill("solid", fgColor=color)

    def _header_font(color: str = ACCENT, bold: bool = True, size: int = 11) -> Font:
        return Font(name="Calibri", bold=bold, color=color, size=size)

    def _cell_font(color: str = WHITE, bold: bool = False, size: int = 10) -> Font:
        return Font(name="Calibri", bold=bold, color=color, size=size)

    def _center_align() -> Alignment:
        return Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _border() -> Border:
        side = Side(style="thin", color="1A1A2E")
        return Border(left=side, right=side, top=side, bottom=side)

    def _write_brand_header(ws, title: str, span_cols: int = 10):
        """Write a branded header block at the top of a worksheet."""
        ws.merge_cells(f"A1:{get_column_letter(span_cols)}1")
        ws["A1"] = APP_NAME
        ws["A1"].font = Font(name="Calibri", bold=True, size=16, color=ACCENT)
        ws["A1"].alignment = _center_align()
        ws["A1"].fill = _header_fill(DARK_BG)

        ws.merge_cells(f"A2:{get_column_letter(span_cols)}2")
        ws["A2"] = f"Prepared & Developed by {BRAND_NAME}"
        ws["A2"].font = Font(name="Calibri", bold=False, size=11, color=PURPLE)
        ws["A2"].alignment = _center_align()
        ws["A2"].fill = _header_fill(DARK_BG)

        ws.merge_cells(f"A3:{get_column_letter(span_cols)}3")
        ws["A3"] = title
        ws["A3"].font = Font(name="Calibri", bold=True, size=13, color=WHITE)
        ws["A3"].alignment = _center_align()
        ws["A3"].fill = _header_fill(HEADER_BG)

        info_row = 4
        ws[f"A{info_row}"] = f"Company: {meta['company_name']}"
        ws[f"A{info_row}"].font = _cell_font(WHITE)
        ws[f"C{info_row}"] = f"FY: {meta['financial_year']}"
        ws[f"C{info_row}"].font = _cell_font(WHITE)
        ws[f"E{info_row}"] = f"Report Date: {meta['report_date']}"
        ws[f"E{info_row}"].font = _cell_font(WHITE)
        ws[f"G{info_row}"] = f"Generated: {meta['generated_at']}"
        ws[f"G{info_row}"].font = _cell_font(WHITE)

        for col in range(1, span_cols + 1):
            ws.cell(info_row, col).fill = _header_fill(SUBHEADER)

        return 5  # next data start row

    def _write_footer(ws, row: int, span_cols: int = 10):
        ws.merge_cells(f"A{row}:{get_column_letter(span_cols)}{row}")
        ws[f"A{row}"] = FOOTER_TEXT
        ws[f"A{row}"].font = Font(name="Calibri", size=9, italic=True, color="64748B")
        ws[f"A{row}"].alignment = _center_align()
        ws[f"A{row}"].fill = _header_fill(DARK_BG)

    def _df_to_sheet(
        ws,
        df: pd.DataFrame,
        start_row: int,
        header_color: str = ACCENT,
        row_colors: Optional[list[str]] = None,
    ):
        """Write a DataFrame to a worksheet starting at start_row."""
        if df.empty:
            ws.cell(start_row, 1).value = "No data available for this category."
            ws.cell(start_row, 1).font = _cell_font(ORANGE, bold=True)
            return start_row + 1

        cols = list(df.columns)

        # Header row
        for ci, col in enumerate(cols, 1):
            cell = ws.cell(start_row, ci)
            cell.value = str(col).replace("_", " ").title()
            cell.font = Font(name="Calibri", bold=True, size=10, color="000000")
            cell.fill = _header_fill(header_color)
            cell.alignment = _center_align()
            cell.border = _border()

        # Data rows
        even_fill = PatternFill("solid", fgColor="0F172A")
        odd_fill = PatternFill("solid", fgColor="1E293B")

        for ri, (_, row) in enumerate(df.iterrows(), 1):
            row_fill = even_fill if ri % 2 == 0 else odd_fill
            for ci, col in enumerate(cols, 1):
                cell = ws.cell(start_row + ri, ci)
                val = row[col]
                if isinstance(val, float):
                    cell.value = round(val, 2)
                    cell.number_format = "#,##0.00"
                else:
                    cell.value = val
                cell.font = _cell_font(WHITE)
                cell.fill = row_fill
                cell.alignment = Alignment(vertical="center")
                cell.border = _border()

        # Auto column width
        for ci, col in enumerate(cols, 1):
            max_len = max(
                len(str(col)),
                df[col].astype(str).str.len().max() if not df.empty else 0,
            )
            ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 35)

        ws.row_dimensions[start_row].height = 20
        return start_row + len(df) + 2

    def _get_status_df(status: str) -> pd.DataFrame:
        if master_df.empty or "status" not in master_df.columns:
            return pd.DataFrame()
        return master_df[master_df["status"] == status].copy()

    # ── Sheet 1: Summary ──────────────────────────────────────────────────
    ws1 = wb.create_sheet("1 Summary")
    ws1.sheet_properties.tabColor = ACCENT
    next_row = _write_brand_header(ws1, "GST ITC Reconciliation Summary", span_cols=6)
    ws1.row_dimensions[1].height = 30
    ws1.row_dimensions[2].height = 20

    kpis = get_kpi_summary(master_df) if not master_df.empty else {}
    summary_data = [
        ("Total Invoices (PR)", kpis.get("total_invoices", 0)),
        ("Matched – Perfect Match", kpis.get("matched_count", 0)),
        ("Missing in Books", kpis.get("missing_books_count", 0)),
        ("Missing in GSTR-2B", kpis.get("missing_gstr2b_count", 0)),
        ("GST Difference Records", master_df[master_df["status"] == STATUS_GST_DIFF].shape[0] if not master_df.empty and "status" in master_df.columns else 0),
        ("Duplicate Records", kpis.get("duplicate_count", 0)),
        ("Manual Review Required", kpis.get("manual_review_count", 0)),
        ("Match Rate (%)", f"{kpis.get('match_rate_percent', 0):.2f}%"),
        ("Total Purchase Value (₹)", format_currency(kpis.get("total_purchase_value", 0))),
        ("Total GST (₹)", format_currency(kpis.get("total_gst", 0))),
        ("Total GST Difference (₹)", format_currency(kpis.get("gst_difference_total", 0))),
        ("Report Generated By", meta["generated_by"]),
        ("Report Generated At", meta["generated_at"]),
        ("Financial Year", meta["financial_year"]),
        ("Company Name", meta["company_name"]),
    ]

    for si, (label, value) in enumerate(summary_data, next_row):
        ws1.cell(si, 1).value = label
        ws1.cell(si, 1).font = Font(name="Calibri", bold=True, size=10, color=ACCENT)
        ws1.cell(si, 1).fill = _header_fill(HEADER_BG)
        ws1.cell(si, 2).value = value
        ws1.cell(si, 2).font = _cell_font(WHITE)
        ws1.cell(si, 2).fill = _header_fill(SUBHEADER)
        for col in range(1, 7):
            ws1.cell(si, col).border = _border()

    ws1.column_dimensions["A"].width = 35
    ws1.column_dimensions["B"].width = 25
    _write_footer(ws1, next_row + len(summary_data) + 2, 6)

    # ── Sheets 2-8: Data tabs ──────────────────────────────────────────────
    def _status_df(status_val):
        if master_df.empty or "status" not in master_df.columns:
            return pd.DataFrame()
        return master_df[master_df["status"] == status_val].copy()

    def _multi_status_df(status_vals):
        if master_df.empty or "status" not in master_df.columns:
            return pd.DataFrame()
        return master_df[master_df["status"].isin(status_vals)].copy()

    sheet_configs = [
        ("2 Matched", _status_df(STATUS_PERFECT_MATCH), GREEN),
        ("3 Missing in Books", _status_df(STATUS_MISSING_BOOKS), RED),
        ("4 Missing in GSTR-2B", _status_df(STATUS_MISSING_GSTR2B), ORANGE),
        ("5 GST Difference", _status_df(STATUS_GST_DIFF), YELLOW),
        ("6 Taxable Difference", _status_df(STATUS_TAXABLE_DIFF), YELLOW),
        ("7 Duplicate", recon_results.get("pr_duplicates", pd.DataFrame()) if recon_results else pd.DataFrame(), PURPLE),
        ("8 Manual Review", _multi_status_df([STATUS_MANUAL_REVIEW, STATUS_FUZZY_MATCH]), ACCENT),
    ]

    for sname, sdf, scolor in sheet_configs:
        ws = wb.create_sheet(sname)
        start = _write_brand_header(ws, sname, span_cols=max(len(sdf.columns) if not sdf.empty else 1, 8))
        _df_to_sheet(ws, sdf.head(50000) if not sdf.empty else pd.DataFrame(), start, header_color=scolor)
        _write_footer(ws, start + (len(sdf) if not sdf.empty else 0) + 5, max(len(sdf.columns) if not sdf.empty else 1, 8))

    # ── Sheet 9: Vendor Summary ────────────────────────────────────────────
    ws9 = wb.create_sheet("9 Vendor Summary")
    vendor_df = get_vendor_summary(master_df) if not master_df.empty else pd.DataFrame()
    start9 = _write_brand_header(ws9, "Vendor-wise Summary", span_cols=10)
    _df_to_sheet(ws9, vendor_df.head(5000) if not vendor_df.empty else pd.DataFrame(), start9, ACCENT)
    _write_footer(ws9, start9 + (len(vendor_df) if not vendor_df.empty else 0) + 5, 10)

    # ── Sheet 10: Monthly Summary ──────────────────────────────────────────
    ws10 = wb.create_sheet("10 Monthly Summary")
    monthly_df = get_monthly_summary(master_df) if not master_df.empty else pd.DataFrame()
    start10 = _write_brand_header(ws10, "Monthly Summary", span_cols=10)
    _df_to_sheet(ws10, monthly_df if not monthly_df.empty else pd.DataFrame(), start10, PURPLE)
    _write_footer(ws10, start10 + (len(monthly_df) if not monthly_df.empty else 0) + 5, 10)

    # ── Sheet 11: Dashboard Stats ──────────────────────────────────────────
    ws11 = wb.create_sheet("11 Dashboard Stats")
    start11 = _write_brand_header(ws11, "Dashboard Statistics", span_cols=6)
    kpi_df = pd.DataFrame(list(kpis.items()), columns=["Metric", "Value"]) if kpis else pd.DataFrame()
    _df_to_sheet(ws11, kpi_df, start11, ACCENT)
    _write_footer(ws11, start11 + len(kpi_df) + 5, 6)

    # ── Sheet 12: Audit Trail ──────────────────────────────────────────────
    ws12 = wb.create_sheet("12 Audit Trail")
    start12 = _write_brand_header(ws12, "Audit Trail", span_cols=8)
    try:
        from modules.audit import get_audit_log
        audit_df = get_audit_log(limit=1000)
    except Exception:
        audit_df = pd.DataFrame()
    _df_to_sheet(ws12, audit_df if not audit_df.empty else pd.DataFrame(), start12, PURPLE)
    _write_footer(ws12, start12 + (len(audit_df) if not audit_df.empty else 0) + 5, 8)

    # ── Output ────────────────────────────────────────────────────────────
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    log_event("REPORT", "Excel report generated (12 sheets)")
    return output.getvalue()


# ---------------------------------------------------------------------------
# PDF Export (ReportLab)
# ---------------------------------------------------------------------------

def export_pdf(
    master_df: pd.DataFrame,
    recon_results: Optional[dict] = None,
) -> bytes:
    """
    Generate a professional multi-page PDF reconciliation report.

    Args:
        master_df:     Master reconciliation DataFrame.
        recon_results: Original reconciliation results dict.

    Returns:
        PDF file content as bytes.
    """
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm, mm
        from reportlab.lib import colors
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Table, TableStyle,
            Spacer, PageBreak, HRFlowable,
        )
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        from reportlab.pdfgen import canvas as rl_canvas
    except ImportError:
        raise RuntimeError(
            "ReportLab is required for PDF export. Install with: pip install reportlab"
        )

    meta = _get_report_metadata()
    kpis = get_kpi_summary(master_df) if not master_df.empty else {}

    # Colors
    C_DARK = colors.HexColor("#0A0A1A")
    C_HEADER = colors.HexColor("#1A1A2E")
    C_ACCENT = colors.HexColor("#00D4FF")
    C_GREEN = colors.HexColor("#34D399")
    C_RED = colors.HexColor("#F87171")
    C_ORANGE = colors.HexColor("#FB923C")
    C_YELLOW = colors.HexColor("#FBBF24")
    C_PURPLE = colors.HexColor("#A78BFA")
    C_WHITE = colors.HexColor("#EAEAEA")
    C_GRAY = colors.HexColor("#64748B")
    C_SUB = colors.HexColor("#1E293B")

    output_buffer = io.BytesIO()
    PAGE_W, PAGE_H = landscape(A4)

    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        "Title", parent=styles["Title"],
        fontSize=18, textColor=C_ACCENT, spaceAfter=4,
        alignment=TA_CENTER, fontName="Helvetica-Bold",
    )
    style_sub = ParagraphStyle(
        "Sub", parent=styles["Normal"],
        fontSize=10, textColor=C_PURPLE, spaceAfter=2,
        alignment=TA_CENTER, fontName="Helvetica",
    )
    style_h2 = ParagraphStyle(
        "H2", parent=styles["Heading2"],
        fontSize=13, textColor=C_ACCENT, spaceAfter=6,
        fontName="Helvetica-Bold",
    )
    style_normal = ParagraphStyle(
        "Normal2", parent=styles["Normal"],
        fontSize=9, textColor=C_WHITE, spaceAfter=2,
        fontName="Helvetica",
    )
    style_footer = ParagraphStyle(
        "Footer", parent=styles["Normal"],
        fontSize=8, textColor=C_GRAY, alignment=TA_CENTER,
        fontName="Helvetica-Oblique",
    )
    style_small = ParagraphStyle(
        "Small", parent=styles["Normal"],
        fontSize=8, textColor=C_GRAY, fontName="Helvetica",
    )

    def _header_footer(canvas_obj, doc):
        """Draw page header and footer on every page."""
        canvas_obj.saveState()
        w, h = canvas_obj._pagesize

        # Header bar
        canvas_obj.setFillColor(C_HEADER)
        canvas_obj.rect(0, h - 1.5 * cm, w, 1.5 * cm, fill=1, stroke=0)
        canvas_obj.setFillColor(C_ACCENT)
        canvas_obj.setFont("Helvetica-Bold", 12)
        canvas_obj.drawString(1 * cm, h - 0.9 * cm, APP_NAME)
        canvas_obj.setFillColor(C_PURPLE)
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.drawString(1 * cm, h - 1.3 * cm, f"Prepared & Developed by {BRAND_NAME}")
        canvas_obj.setFillColor(C_GRAY)
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.drawRightString(w - 1 * cm, h - 0.9 * cm, f"FY: {meta['financial_year']}")
        canvas_obj.drawRightString(w - 1 * cm, h - 1.3 * cm, f"Generated: {meta['generated_at']}")

        # Footer bar
        canvas_obj.setFillColor(C_HEADER)
        canvas_obj.rect(0, 0, w, 1.2 * cm, fill=1, stroke=0)
        canvas_obj.setFillColor(C_GRAY)
        canvas_obj.setFont("Helvetica-Oblique", 7)
        canvas_obj.drawString(1 * cm, 0.45 * cm, FOOTER_TEXT)
        canvas_obj.drawString(1 * cm, 0.2 * cm, CONFIDENTIAL_TEXT)
        canvas_obj.drawRightString(
            w - 1 * cm, 0.35 * cm,
            f"Page {canvas_obj.getPageNumber()}"
        )
        canvas_obj.restoreState()

    doc = SimpleDocTemplate(
        output_buffer,
        pagesize=landscape(A4),
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=2.2 * cm,
        bottomMargin=1.8 * cm,
        title=f"{APP_NAME} – Reconciliation Report",
        author=BRAND_NAME,
    )

    story = []

    # ── Title Page ────────────────────────────────────────────────────────
    story.append(Spacer(1, 1 * cm))
    story.append(Paragraph(APP_NAME, style_title))
    story.append(Paragraph(f"Prepared & Developed by {BRAND_NAME}", style_sub))
    story.append(Paragraph(APP_VERSION, style_sub))
    story.append(Spacer(1, 0.5 * cm))
    story.append(HRFlowable(width="100%", thickness=1, color=C_ACCENT))
    story.append(Spacer(1, 0.3 * cm))

    meta_table_data = [
        ["Company", meta["company_name"], "GSTIN", meta.get("gstin", "-")],
        ["Financial Year", meta["financial_year"], "Report Date", meta["report_date"]],
        ["Generated By", meta["generated_by"], "Generated At", meta["generated_at"]],
    ]
    meta_style = TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), C_HEADER),
        ("BACKGROUND", (2, 0), (2, -1), C_HEADER),
        ("TEXTCOLOR", (0, 0), (0, -1), C_ACCENT),
        ("TEXTCOLOR", (2, 0), (2, -1), C_ACCENT),
        ("TEXTCOLOR", (1, 0), (1, -1), C_WHITE),
        ("TEXTCOLOR", (3, 0), (3, -1), C_WHITE),
        ("FONT", (0, 0), (-1, -1), "Helvetica", 9),
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 9),
        ("FONT", (2, 0), (2, -1), "Helvetica-Bold", 9),
        ("GRID", (0, 0), (-1, -1), 0.5, C_GRAY),
        ("ROWBACKGROUND", (0, 0), (-1, -1), [C_SUB, C_HEADER]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ])
    mt = Table(meta_table_data, colWidths=[3.5 * cm, 7 * cm, 3.5 * cm, 7 * cm])
    mt.setStyle(meta_style)
    story.append(mt)

    # ── KPI Summary Table ─────────────────────────────────────────────────
    story.append(Spacer(1, 0.8 * cm))
    story.append(Paragraph("Executive Summary – KPI Overview", style_h2))

    kpi_table_data = [
        ["Metric", "Value", "Metric", "Value"],
        ["Total Invoices", f"{kpis.get('total_invoices', 0):,}", "Match Rate", f"{kpis.get('match_rate_percent', 0):.2f}%"],
        ["Matched (Perfect)", f"{kpis.get('matched_count', 0):,}", "Pending Invoices", f"{kpis.get('pending_count', 0):,}"],
        ["Missing in Books", f"{kpis.get('missing_books_count', 0):,}", "Missing in GSTR-2B", f"{kpis.get('missing_gstr2b_count', 0):,}"],
        ["Total Purchase Value", format_currency(kpis.get("total_purchase_value", 0)), "Total GST", format_currency(kpis.get("total_gst", 0))],
        ["GST Difference (₹)", format_currency(kpis.get("gst_difference_total", 0)), "Duplicates", f"{kpis.get('duplicate_count', 0):,}"],
        ["Manual Review", f"{kpis.get('manual_review_count', 0):,}", "Fuzzy Matches", f"{kpis.get('fuzzy_match_count', 0):,}"],
    ]
    kpi_style = TableStyle([
        ("BACKGROUND", (0, 0), (3, 0), C_ACCENT),
        ("TEXTCOLOR", (0, 0), (3, 0), C_DARK),
        ("FONT", (0, 0), (3, 0), "Helvetica-Bold", 9),
        ("BACKGROUND", (0, 1), (0, -1), C_HEADER),
        ("BACKGROUND", (2, 1), (2, -1), C_HEADER),
        ("TEXTCOLOR", (0, 1), (0, -1), C_ACCENT),
        ("TEXTCOLOR", (2, 1), (2, -1), C_ACCENT),
        ("TEXTCOLOR", (1, 1), (1, -1), C_WHITE),
        ("TEXTCOLOR", (3, 1), (3, -1), C_WHITE),
        ("ROWBACKGROUND", (0, 1), (-1, -1), [C_SUB, C_HEADER]),
        ("FONT", (0, 1), (-1, -1), "Helvetica", 9),
        ("FONT", (0, 1), (0, -1), "Helvetica-Bold", 9),
        ("FONT", (2, 1), (2, -1), "Helvetica-Bold", 9),
        ("GRID", (0, 0), (-1, -1), 0.5, C_GRAY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ])
    kt = Table(kpi_table_data, colWidths=[5 * cm, 7 * cm, 5 * cm, 7 * cm])
    kt.setStyle(kpi_style)
    story.append(kt)

    # ── Detailed sections for key categories ──────────────────────────────
    def _df_to_pdf_table(df: pd.DataFrame, max_rows: int = 100) -> Optional[Table]:
        if df is None or df.empty:
            return None
        display_cols = [c for c in df.columns if not c.startswith("_")][:12]  # max 12 cols
        sample = df[display_cols].head(max_rows)

        header = [str(c).replace("_", " ").title() for c in display_cols]
        rows = [header]
        for _, row in sample.iterrows():
            rows.append([str(row[c])[:30] if pd.notna(row[c]) else "" for c in display_cols])

        col_width = (PAGE_W - 3 * cm) / len(display_cols)
        tbl = Table(rows, colWidths=[col_width] * len(display_cols))
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), C_ACCENT),
            ("TEXTCOLOR", (0, 0), (-1, 0), C_DARK),
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7),
            ("ROWBACKGROUND", (0, 1), (-1, -1), [C_SUB, C_HEADER]),
            ("TEXTCOLOR", (0, 1), (-1, -1), C_WHITE),
            ("FONT", (0, 1), (-1, -1), "Helvetica", 7),
            ("GRID", (0, 0), (-1, -1), 0.3, C_GRAY),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ]))
        return tbl

    sections = []
    if not master_df.empty and "status" in master_df.columns:
        sections = [
            ("Missing in Books", master_df[master_df["status"] == STATUS_MISSING_BOOKS], C_RED),
            ("Missing in GSTR-2B", master_df[master_df["status"] == STATUS_MISSING_GSTR2B], C_ORANGE),
            ("GST Difference", master_df[master_df["status"] == STATUS_GST_DIFF], C_YELLOW),
            ("Vendor Summary", get_vendor_summary(master_df), C_ACCENT),
        ]

    for sec_title, sec_df, _ in sections:
        story.append(PageBreak())
        story.append(Paragraph(sec_title, style_h2))
        story.append(Spacer(1, 0.3 * cm))
        if sec_df is not None and not sec_df.empty:
            tbl = _df_to_pdf_table(sec_df, max_rows=50)
            if tbl:
                story.append(tbl)
                story.append(Spacer(1, 0.3 * cm))
                story.append(
                    Paragraph(
                        f"Showing first 50 of {len(sec_df)} records. Export full data as CSV/Excel.",
                        style_small,
                    )
                )
        else:
            story.append(Paragraph("No records in this category.", style_normal))

    # ── Final footer page ─────────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Spacer(1, 2 * cm))
    story.append(Paragraph(f"— End of Report —", style_title))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(FOOTER_TEXT, style_footer))
    story.append(Paragraph(CONFIDENTIAL_TEXT, style_footer))

    doc.build(story, onFirstPage=_header_footer, onLaterPages=_header_footer)
    output_buffer.seek(0)

    log_event("REPORT", "PDF report generated")
    return output_buffer.getvalue()


# ---------------------------------------------------------------------------
# CSV Export
# ---------------------------------------------------------------------------

def export_csv(df: pd.DataFrame, filename_hint: str = "reconciliation") -> bytes:
    """
    Export a DataFrame to CSV bytes.

    Args:
        df:            DataFrame to export.
        filename_hint: Used for audit logging.

    Returns:
        CSV content as bytes.
    """
    log_event("EXPORT", f"CSV export: {filename_hint} ({len(df)} rows)")
    return df.to_csv(index=False).encode("utf-8")


# ---------------------------------------------------------------------------
# Streamlit Reports Page
# ---------------------------------------------------------------------------

def render_reports_page() -> None:
    """Render the Reports export page."""

    st.markdown(
        "<h2 style='color:#00D4FF;'>📄 Reports</h2>",
        unsafe_allow_html=True,
    )

    master_df = st.session_state.get("master_df")
    recon_results = st.session_state.get("recon_results")

    if master_df is None or master_df.empty:
        st.warning(
            "⚠️ No reconciliation data available. "
            "Please run Reconciliation first."
        )
        return

    meta = _get_report_metadata()
    kpis = get_kpi_summary(master_df)

    # ── Report header preview ──────────────────────────────────────────────
    st.markdown(
        f"""
        <div style="background:rgba(26,26,46,0.9); border:1px solid #00D4FF44;
             border-radius:12px; padding:20px 24px; margin-bottom:20px;">
            <div style="font-size:1.2rem; font-weight:800; color:#00D4FF; text-align:center;">
                GST INPUT RECONCILIATION REPORT
            </div>
            <div style="font-size:0.85rem; color:#A78BFA; text-align:center; margin-top:4px;">
                Prepared & Developed by <strong>Karthik LVN</strong>
            </div>
            <hr style="border-color:#1A1A2E; margin:12px 0;">
            <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;
                 font-size:0.82rem; color:#94A3B8;">
                <div>🏢 <strong style="color:#EAEAEA;">{meta['company_name']}</strong></div>
                <div>📅 FY: <strong style="color:#EAEAEA;">{meta['financial_year']}</strong></div>
                <div>⏰ <strong style="color:#EAEAEA;">{meta['generated_at']}</strong></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Quick stats ────────────────────────────────────────────────────────
    qs1, qs2, qs3, qs4 = st.columns(4)
    qs1.metric("Total Invoices", f"{kpis.get('total_invoices', 0):,}")
    qs2.metric("Matched", f"{kpis.get('matched_count', 0):,}")
    qs3.metric("Match Rate", f"{kpis.get('match_rate_percent', 0):.1f}%")
    qs4.metric("GST Difference", format_currency(kpis.get("gst_difference_total", 0)))

    st.divider()

    # ── Export options ─────────────────────────────────────────────────────
    st.subheader("📤 Export Options")

    ec1, ec2, ec3 = st.columns(3)

    with ec1:
        st.markdown("### 📊 Excel Report")
        st.markdown("12-sheet workbook with styled headers, branding, and audit trail.")
        if st.button("Generate Excel Report", type="primary", use_container_width=True, key="gen_excel"):
            with st.spinner("Generating Excel… this may take a moment for large datasets."):
                try:
                    excel_bytes = export_excel(master_df, recon_results)
                    fname = f"GST_Recon_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
                    st.download_button(
                        "⬇️ Download Excel",
                        data=excel_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_excel_reports_page",
                    )
                    st.success("✅ Excel report ready!")
                except Exception as e:
                    st.error(f"Excel generation failed: {e}")

    with ec2:
        st.markdown("### 📄 PDF Report")
        st.markdown("Professional multi-page PDF with header, footer, and branding on every page.")
        if st.button("Generate PDF Report", type="primary", use_container_width=True, key="gen_pdf"):
            with st.spinner("Generating PDF…"):
                try:
                    pdf_bytes = export_pdf(master_df, recon_results)
                    fname = f"GST_Recon_{datetime.date.today().strftime('%Y%m%d')}.pdf"
                    st.download_button(
                        "⬇️ Download PDF",
                        data=pdf_bytes,
                        file_name=fname,
                        mime="application/pdf",
                        key="dl_pdf_reports_page",
                    )
                    st.success("✅ PDF report ready!")
                except Exception as e:
                    st.error(f"PDF generation failed: {e}")

    with ec3:
        st.markdown("### 📋 CSV Exports")
        st.markdown("Individual CSV files for each category.")

        if st.button("Master CSV", use_container_width=True, key="gen_csv_master"):
            csv = export_csv(master_df, "master")
            st.download_button(
                "⬇️ Master CSV",
                data=csv,
                file_name="master_reconciliation.csv",
                mime="text/csv",
                key="dl_master_csv",
            )

        if "status" in master_df.columns:
            missing_books = master_df[master_df["status"] == STATUS_MISSING_BOOKS]
            if not missing_books.empty:
                if st.button("Missing in Books CSV", use_container_width=True, key="gen_csv_mb"):
                    csv = export_csv(missing_books, "missing_books")
                    st.download_button(
                        "⬇️ Missing in Books CSV",
                        data=csv,
                        file_name="missing_in_books.csv",
                        mime="text/csv",
                        key="dl_mb_csv",
                    )

            missing_gstr = master_df[master_df["status"] == STATUS_MISSING_GSTR2B]
            if not missing_gstr.empty:
                if st.button("Missing in GSTR-2B CSV", use_container_width=True, key="gen_csv_mg"):
                    csv = export_csv(missing_gstr, "missing_gstr2b")
                    st.download_button(
                        "⬇️ Missing in GSTR-2B CSV",
                        data=csv,
                        file_name="missing_in_gstr2b.csv",
                        mime="text/csv",
                        key="dl_mg_csv",
                    )

    st.divider()

    # ── Report footer preview ──────────────────────────────────────────────
    st.markdown(
        f"""
        <div style="text-align:center; color:#374151; font-size:0.78rem; margin-top:16px;">
            Prepared & Developed by <strong>Karthik LVN</strong> &nbsp;·&nbsp;
            © 2026 All Rights Reserved &nbsp;·&nbsp; Confidential Report
        </div>
        """,
        unsafe_allow_html=True,
    )
