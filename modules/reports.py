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

    # ── Sheets 1-9: Data tabs + summaries (built below) ──────────────────
    def _status_df(status_val):
        if master_df.empty or "status" not in master_df.columns:
            return pd.DataFrame()
        return master_df[master_df["status"] == status_val].copy()

    def _multi_status_df(status_vals):
        if master_df.empty or "status" not in master_df.columns:
            return pd.DataFrame()
        return master_df[master_df["status"].isin(status_vals)].copy()

    sheet_configs = [
        # Sheet name                               DataFrame                               Tab color
        ("1 Fully Reconciled",    _status_df(STATUS_PERFECT_MATCH),                        GREEN),
        ("2 Not Available in 2B", _status_df(STATUS_MISSING_GSTR2B),                       ORANGE),
        ("3 Not Accounted in Books", _status_df(STATUS_MISSING_BOOKS),                     RED),
        ("4 GST Difference",      _status_df(STATUS_GST_DIFF),                             YELLOW),
        ("5 Taxable Difference",  _status_df(STATUS_TAXABLE_DIFF),                         YELLOW),
        ("6 Fuzzy Match Review",  _multi_status_df([STATUS_MANUAL_REVIEW, STATUS_FUZZY_MATCH]), ACCENT),
        ("7 Duplicates",          recon_results.get("pr_duplicates", pd.DataFrame()) if recon_results else pd.DataFrame(), PURPLE),
    ]

    for sname, sdf, scolor in sheet_configs:
        ws = wb.create_sheet(sname)
        ws.sheet_properties.tabColor = scolor
        span = max(len(sdf.columns) if not sdf.empty else 1, 8)
        start = _write_brand_header(ws, sname, span_cols=span)
        _df_to_sheet(ws, sdf.head(50000) if not sdf.empty else pd.DataFrame(), start, header_color=scolor)
        _write_footer(ws, start + (len(sdf) if not sdf.empty else 0) + 5, span)

    # ── Sheet 8: Vendor Summary ────────────────────────────────────────────
    ws9 = wb.create_sheet("8 Vendor Summary")
    ws9.sheet_properties.tabColor = ACCENT
    vendor_df = get_vendor_summary(master_df) if not master_df.empty else pd.DataFrame()
    start9 = _write_brand_header(ws9, "Vendor-wise Summary", span_cols=10)
    _df_to_sheet(ws9, vendor_df.head(5000) if not vendor_df.empty else pd.DataFrame(), start9, ACCENT)
    _write_footer(ws9, start9 + (len(vendor_df) if not vendor_df.empty else 0) + 5, 10)

    # ── Sheet 9: Monthly Summary ───────────────────────────────────────────
    ws10 = wb.create_sheet("9 Monthly Summary")
    ws10.sheet_properties.tabColor = PURPLE
    monthly_df = get_monthly_summary(master_df) if not master_df.empty else pd.DataFrame()
    start10 = _write_brand_header(ws10, "Monthly Summary", span_cols=10)
    _df_to_sheet(ws10, monthly_df if not monthly_df.empty else pd.DataFrame(), start10, PURPLE)
    _write_footer(ws10, start10 + (len(monthly_df) if not monthly_df.empty else 0) + 5, 10)

    # ── Sheet 10: Overall Summary ──────────────────────────────────────────
    ws1 = wb.create_sheet("10 Summary", 0)   # insert at position 0 (first sheet)
    ws1.sheet_properties.tabColor = ACCENT
    next_row = _write_brand_header(ws1, "GST ITC Reconciliation Summary", span_cols=6)
    ws1.row_dimensions[1].height = 30
    ws1.row_dimensions[2].height = 20

    kpis = get_kpi_summary(master_df) if not master_df.empty else {}
    summary_data = [
        ("Total Invoices (PR)",          kpis.get("total_invoices", 0)),
        ("Fully Reconciled",             kpis.get("matched_count", 0)),
        ("Not Available in GSTR-2B",     master_df[master_df["status"] == STATUS_MISSING_GSTR2B].shape[0] if not master_df.empty and "status" in master_df.columns else 0),
        ("Not Accounted in Books",       kpis.get("missing_books_count", 0)),
        ("GST Difference Records",       master_df[master_df["status"] == STATUS_GST_DIFF].shape[0] if not master_df.empty and "status" in master_df.columns else 0),
        ("Duplicate Records",            kpis.get("duplicate_count", 0)),
        ("Fuzzy / Manual Review",        kpis.get("manual_review_count", 0)),
        ("Match Rate (%)",               f"{kpis.get('match_rate_percent', 0):.2f}%"),
        ("Total Purchase Value (INR)",   format_currency(kpis.get("total_purchase_value", 0))),
        ("Total GST (INR)",              format_currency(kpis.get("total_gst", 0))),
        ("Total GST Difference (INR)",   format_currency(kpis.get("gst_difference_total", 0))),
        ("Report Generated By",          meta["generated_by"]),
        ("Report Generated At",          meta["generated_at"]),
        ("Financial Year",               meta["financial_year"]),
        ("Company Name",                 meta["company_name"]),
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
# Simple Sheet-wise Excel Builder (user-requested format)
# ---------------------------------------------------------------------------

def _build_sheet_excel(recon_results: dict) -> bytes:
    """
    Multi-sheet Excel report.
    Sheets: Summary | Matched | Missing in Books | Missing in GSTR-2B |
            Near Match - Review (full PR+2B side-by-side + reason cols) |
            Duplicates | Unreconciled | All Records |
            As Per Books | As Per GSTR-2B
    """
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import datetime as dt

    # ── Styles ─────────────────────────────────────────────────────────────
    WHITE_BG   = PatternFill("solid", fgColor="FFFFFF")
    ALT_BG     = PatternFill("solid", fgColor="F8FAFC")
    BORDER_CLR = "D1D5DB"
    DATA_FONT  = Font(name="Calibri", size=9, color="111827")
    CTR  = Alignment(horizontal="center", vertical="center")
    LEFT = Alignment(horizontal="left",   vertical="center")

    HDR_CONFIGS = {
        "matched": ("1F5E40", "FFFFFF"), "books":   ("7F1D1D", "FFFFFF"),
        "gstr":    ("78350F", "FFFFFF"), "near":    ("78350F", "FBBF24"),
        "dup":     ("3B0764", "FFFFFF"), "unrecon": ("1E293B", "FFFFFF"),
        "all":     ("0F172A", "00D4FF"), "pr_view": ("1E3A8A", "FFFFFF"),
        "gb_view": ("14532D", "FFFFFF"),
    }

    def thin_border():
        s = Side(style="thin", color=BORDER_CLR)
        return Border(left=s, right=s, top=s, bottom=s)

    def _fmt_val(val):
        try:
            if pd.isna(val):
                return ""
        except Exception:
            pass
        if hasattr(val, "strftime"):
            return val.strftime("%d-%m-%Y")
        if isinstance(val, str) and val.endswith(" 00:00:00"):
            return val[:10]
        if isinstance(val, float):
            return round(val, 2)
        return val

    def _col_label(c):
        c = c.replace("_pr", " (Books)").replace("_gstr2b", " (2B)")
        return c.replace("_", " ").strip().title()

    def _write_rows(ws, df_out, hdr_fill, hdr_font):
        col_labels = [_col_label(c) for c in df_out.columns]
        for ci, label in enumerate(col_labels, 1):
            cell = ws.cell(1, ci, label)
            cell.font = hdr_font; cell.fill = hdr_fill
            cell.alignment = CTR; cell.border = thin_border()
        ws.row_dimensions[1].height = 20
        ws.freeze_panes = "A2"
        for ri, row in enumerate(df_out.itertuples(index=False), 2):
            fill = WHITE_BG if ri % 2 == 0 else ALT_BG
            for ci, val in enumerate(row, 1):
                c = ws.cell(ri, ci)
                c.value = _fmt_val(val); c.font = DATA_FONT
                c.fill = fill; c.border = thin_border()
                c.alignment = CTR if isinstance(val, (int, float)) else LEFT
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    c.number_format = "#,##0.00"
        # Auto-fit
        for ci, label in enumerate(col_labels, 1):
            vals = [str(ws.cell(r, ci).value or "") for r in range(1, min(len(df_out) + 2, 52))]
            ws.column_dimensions[get_column_letter(ci)].width = min(max((max(len(v) for v in vals) if vals else 10) + 2, 12), 42)

    # Standard sheet (keeps only PR or plain invoice cols)
    def write_sheet(wb, title, df, tab_color, fill_key):
        ws = wb.create_sheet(title)
        ws.sheet_properties.tabColor = tab_color
        if df is None or (hasattr(df, "empty") and df.empty):
            ws["A1"] = f"No {title} records."
            ws["A1"].font = Font(name="Calibri", color="94A3B8", italic=True, size=10)
            return
        hdr_bg, hdr_fg = HDR_CONFIGS.get(fill_key, ("0F172A", "FFFFFF"))
        WANT = ["vendor_name", "gstin", "invoice_number", "invoice_date",
                "taxable_value", "cgst", "sgst", "igst", "cess", "total_gst", "invoice_value"]
        pr_cols = [c + "_pr" for c in WANT if c + "_pr" in df.columns]
        plain   = [c for c in WANT if c in df.columns and c + "_pr" not in df.columns]
        use_cols = pr_cols if pr_cols else (plain if plain else [c for c in df.columns if not c.startswith("_")][:14])
        df_out = df[use_cols].copy()
        _write_rows(ws, df_out, PatternFill("solid", fgColor=hdr_bg), Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # Near Match sheet — FULL side-by-side PR + GSTR-2B + reason cols first
    def write_near_match_sheet(wb, df, tab_color):
        ws = wb.create_sheet("Near Match - Review")
        ws.sheet_properties.tabColor = tab_color
        if df is None or (hasattr(df, "empty") and df.empty):
            ws["A1"] = "No Near Match records."
            ws["A1"].font = Font(name="Calibri", color="94A3B8", italic=True, size=10)
            return
        hdr_bg, hdr_fg = HDR_CONFIGS["near"]
        # Reason/score cols first, then all PR cols, then all GSTR-2B cols
        REASON_COLS = ["Match Reason", "Similarity %", "fuzzy_score", "confidence", "match_tier", "match_key"]
        WANT = ["vendor_name", "gstin", "invoice_number", "invoice_date",
                "taxable_value", "cgst", "sgst", "igst", "cess", "total_gst", "invoice_value"]
        pr_cols  = [c + "_pr"     for c in WANT if c + "_pr"     in df.columns]
        gst_cols = [c + "_gstr2b" for c in WANT if c + "_gstr2b" in df.columns]
        reason_cols = [c for c in REASON_COLS if c in df.columns]
        # Also include source cols if present
        extra_source = [c for c in ["source_pr"] if c in df.columns]
        full_cols = reason_cols + pr_cols + extra_source + gst_cols
        if not full_cols:
            full_cols = [c for c in df.columns if not c.startswith("_")]
        df_out = df[full_cols].copy()
        _write_rows(ws, df_out, PatternFill("solid", fgColor=hdr_bg), Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # "As Per Books" — all PR-side records with status
    def write_per_books(wb, matched, books, gstr, fuzzy):
        ws = wb.create_sheet("As Per Books")
        ws.sheet_properties.tabColor = "1E40AF"
        hdr_bg, hdr_fg = HDR_CONFIGS["pr_view"]
        WANT = ["vendor_name", "gstin", "invoice_number", "invoice_date",
                "taxable_value", "cgst", "sgst", "igst", "cess", "total_gst", "invoice_value"]

        def _extract(df, status_label, suffix="_pr"):
            if df is None or (hasattr(df, "empty") and df.empty):
                return pd.DataFrame()
            cols = [c + suffix for c in WANT if c + suffix in df.columns]
            if not cols:
                cols = [c for c in WANT if c in df.columns]
                suffix = ""
            if not cols:
                return pd.DataFrame()
            out = df[cols].copy()
            out.columns = [c.replace(suffix, "") for c in out.columns]
            out.insert(0, "Status", status_label)
            return out

        # Books records = PR side of matched + PR only unmatched (missing_in_gstr2b) + PR side of fuzzy
        frames = []
        frames.append(_extract(matched, "Matched"))
        # missing_in_gstr2b is already PR-perspective (no suffix)
        if gstr is not None and not gstr.empty:
            cols = [c for c in WANT if c in gstr.columns]
            if cols:
                out = gstr[cols].copy(); out.insert(0, "Status", "Not in GSTR-2B"); frames.append(out)
        frames.append(_extract(fuzzy, "Near Match (Verify)", suffix="_pr"))
        all_pr = pd.concat([f for f in frames if not f.empty], ignore_index=True)
        if all_pr.empty:
            ws["A1"] = "No Books records available."
            return
        _write_rows(ws, all_pr, PatternFill("solid", fgColor=hdr_bg), Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # "As Per GSTR-2B" — all GSTR-2B-side records with status
    def write_per_gstr2b(wb, matched, books, gstr, fuzzy):
        ws = wb.create_sheet("As Per GSTR-2B")
        ws.sheet_properties.tabColor = "166534"
        hdr_bg, hdr_fg = HDR_CONFIGS["gb_view"]
        WANT = ["vendor_name", "gstin", "invoice_number", "invoice_date",
                "taxable_value", "cgst", "sgst", "igst", "cess", "total_gst", "invoice_value"]

        def _extract(df, status_label, suffix="_gstr2b"):
            if df is None or (hasattr(df, "empty") and df.empty):
                return pd.DataFrame()
            cols = [c + suffix for c in WANT if c + suffix in df.columns]
            if not cols:
                cols = [c for c in WANT if c in df.columns]
                suffix = ""
            if not cols:
                return pd.DataFrame()
            out = df[cols].copy()
            out.columns = [c.replace(suffix, "") for c in out.columns]
            out.insert(0, "Status", status_label)
            return out

        frames = []
        frames.append(_extract(matched, "Matched"))
        # missing_in_books is already GSTR-2B-perspective (no suffix)
        if books is not None and not books.empty:
            cols = [c for c in WANT if c in books.columns]
            if cols:
                out = books[cols].copy(); out.insert(0, "Status", "Not in Books"); frames.append(out)
        frames.append(_extract(fuzzy, "Near Match (Verify)", suffix="_gstr2b"))
        all_gb = pd.concat([f for f in frames if not f.empty], ignore_index=True)
        if all_gb.empty:
            ws["A1"] = "No GSTR-2B records available."
            return
        _write_rows(ws, all_gb, PatternFill("solid", fgColor=hdr_bg), Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # ── Data ───────────────────────────────────────────────────────────────
    matched = recon_results.get("matched",           pd.DataFrame())
    books   = recon_results.get("missing_in_books",  pd.DataFrame())
    gstr    = recon_results.get("missing_in_gstr2b", pd.DataFrame())
    fuzzy   = recon_results.get("fuzzy_candidates",  pd.DataFrame())
    pr_dup  = recon_results.get("pr_duplicates",     pd.DataFrame())
    gst_dup = recon_results.get("gstr2b_duplicates", pd.DataFrame())
    stats   = recon_results.get("stats", {})

    dup_all = pd.concat([(pr_dup  if isinstance(pr_dup,  pd.DataFrame) else pd.DataFrame()),
                         (gst_dup if isinstance(gst_dup, pd.DataFrame) else pd.DataFrame())],
                        ignore_index=True)
    unrecon_frames = [d for d in [books, gstr, fuzzy] if isinstance(d, pd.DataFrame) and not d.empty]
    unrecon = pd.concat(unrecon_frames, ignore_index=True) if unrecon_frames else pd.DataFrame()
    all_frames = ([matched] if isinstance(matched, pd.DataFrame) and not matched.empty else []) + unrecon_frames
    all_records = pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()

    # ── Workbook ───────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    # Summary sheet
    ws_sum = wb.create_sheet("Summary")
    ws_sum.sheet_properties.tabColor = "00D4FF"
    ws_sum.column_dimensions["A"].width = 38
    ws_sum.column_dimensions["B"].width = 18
    ws_sum["A1"] = "GST Input Reconciliation Report"
    ws_sum["A1"].font = Font(name="Calibri", bold=True, size=16, color="0F172A")
    ws_sum["A1"].fill = PatternFill("solid", fgColor="E0F7FA")
    ws_sum["A2"] = "Prepared & Developed by Karthik LVN"
    ws_sum["A2"].font = Font(name="Calibri", size=10, color="7C3AED", italic=True)
    ws_sum["A2"].fill = PatternFill("solid", fgColor="F3E8FF")
    ws_sum["A3"] = f"Generated: {dt.date.today().strftime('%d-%m-%Y')}"
    ws_sum["A3"].font = Font(name="Calibri", size=9, color="6B7280")
    for r, label, val, bg, fg, is_hdr in [
        (5, "Category",                           "Count",                                      "1E40AF", "FFFFFF", True),
        (6, "Matched (Fully Reconciled)",          stats.get("total_matched", 0),                "ECFDF5", "065F46", False),
        (7, "Missing in Books (In 2B, not in PR)", stats.get("missing_in_books_count", 0),       "FEF2F2", "991B1B", False),
        (8, "Missing in GSTR-2B (In PR, not 2B)", stats.get("missing_in_gstr2b_count", 0),      "FFF7ED", "92400E", False),
        (9, "Near Match (Review Required)",        stats.get("fuzzy_candidates_count", 0),       "FFFBEB", "B45309", False),
        (10,"PR Duplicates",                       stats.get("pr_duplicates_count", 0),          "F5F3FF", "4C1D95", False),
        (11,"GSTR-2B Duplicates",                  stats.get("gstr2b_duplicates_count", 0),      "F5F3FF", "4C1D95", False),
        (12,"Match Rate (%)",                      f"{stats.get('match_rate', 0):.2f}%",         "EFF6FF", "1E3A8A", False),
    ]:
        lc = ws_sum.cell(r, 1, label)
        vc = ws_sum.cell(r, 2, val)
        for cell in [lc, vc]:
            cell.fill   = PatternFill("solid", fgColor=bg)
            cell.font   = Font(name="Calibri", bold=is_hdr, size=10, color=fg)
            cell.border = thin_border()
            cell.alignment = LEFT
        vc.alignment = CTR
        ws_sum.row_dimensions[r].height = 18

    # Data sheets
    write_sheet(wb, "Matched",            matched,    "34D399", "matched")
    write_sheet(wb, "Missing in Books",   books,      "F87171", "books")
    write_sheet(wb, "Missing in GSTR-2B", gstr,       "FB923C", "gstr")
    write_near_match_sheet(wb, fuzzy,     "FBBF24")
    write_sheet(wb, "Duplicates",         dup_all,    "A78BFA", "dup")
    write_sheet(wb, "Unreconciled",       unrecon,    "94A3B8", "unrecon")
    write_sheet(wb, "All Records",        all_records,"00D4FF", "all")
    write_per_books(wb, matched, books, gstr, fuzzy)
    write_per_gstr2b(wb, matched, books, gstr, fuzzy)

    buf = io.BytesIO()
    wb.save(buf)
    log_event("EXPORT", "Sheet-wise Excel report generated")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Streamlit Reports Page
# ---------------------------------------------------------------------------

def render_reports_page() -> None:
    """Render the Reports export page."""
    import datetime

    # Header + Back to Dashboard
    hdr_col, dash_col = st.columns([8, 2])
    hdr_col.markdown("<h2 style='color:#00D4FF; margin:0;'>Reports</h2>", unsafe_allow_html=True)
    if dash_col.button("Back to Dashboard", key="rep_goto_dash", use_container_width=True):
        st.session_state["current_page"] = "Dashboard"
        st.rerun()

    recon_results = st.session_state.get("recon_results")

    if recon_results is None:
        st.markdown(
            "<div style='background:rgba(248,113,113,0.08); border:1px solid #F8717133; "
            "border-radius:10px; padding:24px; text-align:center; margin-top:16px;'>"
            "<div style='color:#F87171; font-size:1.1rem; font-weight:700;'>No reconciliation data yet</div>"
            "<div style='color:#94A3B8; margin-top:8px;'>Complete the workflow: "
            "<strong>Upload Data</strong> - <strong>Column Mapping</strong> - "
            "<strong>Reconcile</strong> - then come back here.</div></div>",
            unsafe_allow_html=True,
        )
        if st.button("Go to Upload Data", type="primary", use_container_width=True, key="rep_goto_upload"):
            st.session_state["current_page"] = "Upload Data"
            st.rerun()
        return

    stats = recon_results.get("stats", {})

    # Summary KPIs
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Matched",            f"{stats.get('total_matched', 0):,}")
    c2.metric("Missing in Books",   f"{stats.get('missing_in_books_count', 0):,}")
    c3.metric("Missing in GSTR-2B", f"{stats.get('missing_in_gstr2b_count', 0):,}")
    c4.metric("Near Match",         f"{stats.get('fuzzy_candidates_count', 0):,}")
    c5.metric("Match Rate",         f"{stats.get('match_rate', 0):.1f}%")

    st.divider()

    # Excel Download
    st.markdown(
        "<div style='font-size:1rem; font-weight:700; color:#A78BFA; margin-bottom:4px;'>"
        "Download Full Reconciliation Report (Excel)</div>",
        unsafe_allow_html=True,
    )
    st.caption(
        "9 sheets: Summary | Matched | Missing in Books | Missing in GSTR-2B | "
        "Near Match (full side-by-side + Reason/Score) | Duplicates | Unreconciled | "
        "All Records | As Per Books | As Per GSTR-2B"
    )
    try:
        excel_bytes = _build_sheet_excel(recon_results)
        fname = f"GST_Recon_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="Download Excel Report",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_excel_direct",
            type="primary",
        )
    except Exception as e:
        st.error(f"Could not prepare Excel: {e}")

    st.divider()

    # Preview tabs
    WANT = ["vendor_name", "gstin", "invoice_number", "invoice_date",
            "taxable_value", "cgst", "sgst", "igst", "cess", "total_gst", "invoice_value"]

    def _clean_show(df, label):
        if df is None or (hasattr(df, "empty") and df.empty):
            st.info(f"No {label} records.")
            return
        # Keep PR cols (or plain) for display
        pr_cols = [c + "_pr" for c in WANT if c + "_pr" in df.columns]
        plain   = [c for c in WANT if c in df.columns and c + "_pr" not in df.columns]
        use = pr_cols if pr_cols else plain
        out = df[use].copy() if use else df.copy()
        out.columns = [c.replace("_pr", "").replace("_", " ").title() for c in out.columns]
        st.dataframe(out, use_container_width=True, hide_index=True)
        st.caption(f"{len(df):,} records")

    tab_m, tab_b, tab_g, tab_nm, tab_dup = st.tabs([
        f"Matched ({stats.get('total_matched', 0):,})",
        f"Missing in Books ({stats.get('missing_in_books_count', 0):,})",
        f"Missing in GSTR-2B ({stats.get('missing_in_gstr2b_count', 0):,})",
        f"Near Match - Review ({stats.get('fuzzy_candidates_count', 0):,})",
        f"Duplicates ({stats.get('pr_duplicates_count', 0) + stats.get('gstr2b_duplicates_count', 0):,})",
    ])

    with tab_m:
        st.caption("Invoices that exist in BOTH Purchase Register and GSTR-2B with matching details.")
        _clean_show(recon_results.get("matched"), "Matched")

    with tab_b:
        st.caption("Invoices present in GSTR-2B but NOT found in your Purchase Register (Books).")
        _clean_show(recon_results.get("missing_in_books"), "Missing in Books")

    with tab_g:
        st.caption("Invoices present in your Purchase Register (Books) but NOT found in GSTR-2B.")
        _clean_show(recon_results.get("missing_in_gstr2b"), "Missing in GSTR-2B")

    with tab_nm:
        st.markdown(
            "<div style='background:rgba(251,191,36,0.08); border:1px solid #FBBF2433; "
            "border-radius:8px; padding:10px 14px; margin-bottom:10px;'>"
            "<strong style='color:#FBBF24;'>Near Match (Review Required)</strong>"
            "<div style='color:#94A3B8; font-size:0.83rem; margin-top:4px;'>"
            "Found in both files with slight differences (e.g. invoice format INV-001 vs INV/001). "
            "<b>Match Reason</b> and <b>Similarity %</b> explain the difference. "
            "PR data shown first (Books), then GSTR-2B data. Verify and decide manually."
            "</div></div>",
            unsafe_allow_html=True,
        )
        nm_df = recon_results.get("fuzzy_candidates")
        if nm_df is None or (hasattr(nm_df, "empty") and nm_df.empty):
            st.info("No Near Match records.")
        else:
            # Full side-by-side: reason cols + PR cols + GSTR-2B cols
            REASON = ["Match Reason", "Similarity %", "fuzzy_score", "confidence", "match_tier"]
            pr_cols  = [c + "_pr"     for c in WANT if c + "_pr"     in nm_df.columns]
            gst_cols = [c + "_gstr2b" for c in WANT if c + "_gstr2b" in nm_df.columns]
            reason_cols = [c for c in REASON if c in nm_df.columns]
            show_cols = reason_cols + pr_cols + gst_cols
            if not show_cols:
                show_cols = [c for c in nm_df.columns if not c.startswith("_")]
            disp = nm_df[show_cols].copy()
            disp.columns = [
                c.replace("_pr", " (Books)").replace("_gstr2b", " (2B)")
                 .replace("_", " ").strip().title()
                for c in disp.columns
            ]
            st.dataframe(disp, use_container_width=True, hide_index=True)
            st.caption(f"{len(nm_df):,} near match records — verify each row manually")

    with tab_dup:
        st.caption("Duplicate invoice numbers detected within the same file.")
        pr_dup  = recon_results.get("pr_duplicates",     pd.DataFrame())
        gst_dup = recon_results.get("gstr2b_duplicates", pd.DataFrame())
        if isinstance(pr_dup, pd.DataFrame) and not pr_dup.empty:
            st.markdown("**Purchase Register Duplicates**")
            _clean_show(pr_dup, "PR Duplicates")
        if isinstance(gst_dup, pd.DataFrame) and not gst_dup.empty:
            st.markdown("**GSTR-2B Duplicates**")
            _clean_show(gst_dup, "GSTR-2B Duplicates")

    st.markdown(
        "<div style='text-align:center; color:#374151; font-size:0.75rem; margin-top:24px;'>"
        "Prepared &amp; Developed by <strong>Karthik LVN</strong> &nbsp;&#183;&nbsp; "
        "2026 All Rights Reserved</div>",
        unsafe_allow_html=True,
    )
