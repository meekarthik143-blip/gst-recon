"""
GST Reconciliation Report Builder
Prepared & Developed by Karthik LVN | 9849270702
"""

import io
import datetime

import pandas as pd
import streamlit as st

from modules.utils import setup_logging
from modules.audit import log_event

logger = setup_logging()


def _build_sheet_excel(recon_results: dict) -> bytes:
    """
    Multi-sheet Excel report.
    Sheets: Summary | Matched | Missing in Books | Missing in GSTR-2B |
            Near Match - Review | Duplicates | Unreconciled | All Records |
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
        # Auto-fit columns
        for ci, label in enumerate(col_labels, 1):
            vals = [str(ws.cell(r, ci).value or "") for r in range(1, min(len(df_out) + 2, 52))]
            ws.column_dimensions[get_column_letter(ci)].width = min(
                max((max(len(v) for v in vals) if vals else 10) + 2, 12), 42
            )

    WANT = ["vendor_name", "gstin", "invoice_number", "invoice_date",
            "taxable_value", "cgst", "sgst", "igst", "cess", "total_gst", "invoice_value"]

    # ── Helper: write two-section sheet (Books ↑, empty row, GSTR-2B ↓) ──────
    def _write_two_section(ws, df, hdr_bg, hdr_fg, pr_cols, gst_cols,
                           pr_label="📚  AS PER BOOKS", gst_label="📋  AS PER GSTR-2B",
                           reason_cols=None):
        """
        Writes two stacked sections in one worksheet:
          • Section 1: label row + header + pr_cols data
          • Empty separator row
          • Section 2: label row + header + gst_cols data
        """
        hdr_fill = PatternFill("solid", fgColor=hdr_bg)
        hdr_font = Font(name="Calibri", bold=True, color=hdr_fg, size=9)
        sec_fill = PatternFill("solid", fgColor="0F172A")
        sec_font = Font(name="Calibri", bold=True, color="00D4FF", size=10)

        def _write_section(start_row, sec_label, columns):
            """Write one section starting at start_row. Returns next free row."""
            # Section label row (spans width of columns)
            lbl = ws.cell(start_row, 1, sec_label)
            lbl.font = sec_font; lbl.fill = sec_fill
            lbl.alignment = LEFT
            try:
                ws.merge_cells(start_row=start_row, start_column=1,
                               end_row=start_row, end_column=max(len(columns), 1))
            except Exception:
                pass
            ws.row_dimensions[start_row].height = 18
            cur = start_row + 1

            if not columns:
                ws.cell(cur, 1, "No data."); return cur + 1

            # Column header row
            for ci, col in enumerate(columns, 1):
                c = ws.cell(cur, ci, _col_label(col))
                c.font = hdr_font; c.fill = hdr_fill
                c.alignment = CTR; c.border = thin_border()
            ws.row_dimensions[cur].height = 18
            cur += 1

            # Data rows
            data = df[columns].copy()
            for ri, row in enumerate(data.itertuples(index=False)):
                fill = WHITE_BG if ri % 2 == 0 else ALT_BG
                for ci, val in enumerate(row, 1):
                    cell = ws.cell(cur, ci)
                    cell.value = _fmt_val(val); cell.font = DATA_FONT
                    cell.fill = fill; cell.border = thin_border()
                    cell.alignment = CTR if isinstance(val, (int, float)) else LEFT
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        cell.number_format = "#,##0.00"
                cur += 1

            # Auto-fit columns for this section
            for ci, col in enumerate(columns, 1):
                vals = [str(ws.cell(r, ci).value or "") for r in range(start_row, min(cur, start_row + 52))]
                ws.column_dimensions[get_column_letter(ci)].width = min(
                    max((max(len(v) for v in vals) if vals else 10) + 2, 12), 42
                )
            return cur

        cur_row = 1
        # Section 1: Books (PR side) — optionally prepend reason columns
        s1_cols = (reason_cols or []) + pr_cols
        if s1_cols:
            ws.freeze_panes = "A3"
            cur_row = _write_section(cur_row, pr_label, s1_cols)
        # Empty separator row
        cur_row += 1
        # Section 2: GSTR-2B side
        if gst_cols:
            _write_section(cur_row, gst_label, gst_cols)

    # Standard sheet — single-source DFs (plain columns, no _pr/_gstr2b suffix)
    def write_sheet(wb, title, df, tab_color, fill_key):
        ws = wb.create_sheet(title)
        ws.sheet_properties.tabColor = tab_color
        if df is None or (hasattr(df, "empty") and df.empty):
            ws["A1"] = f"No {title} records."
            ws["A1"].font = Font(name="Calibri", color="94A3B8", italic=True, size=10)
            return
        hdr_bg, hdr_fg = HDR_CONFIGS.get(fill_key, ("0F172A", "FFFFFF"))
        # Prefer PR-suffixed cols; fall back to plain cols
        pr_cols  = [c + "_pr" for c in WANT if c + "_pr" in df.columns]
        gst_cols = [c + "_gstr2b" for c in WANT if c + "_gstr2b" in df.columns]
        plain    = [c for c in WANT if c in df.columns and c + "_pr" not in df.columns]

        if pr_cols and gst_cols:
            # Merged DF — show two sections
            _write_two_section(ws, df, hdr_bg, hdr_fg, pr_cols, gst_cols)
        else:
            use_cols = pr_cols or plain or [c for c in df.columns if not c.startswith("_")][:14]
            df_out = df[use_cols].copy()
            _write_rows(ws, df_out,
                        PatternFill("solid", fgColor=hdr_bg),
                        Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # Near Match — two sections: Books (with reason cols) + GSTR-2B
    def write_near_match_sheet(wb, df, tab_color):
        ws = wb.create_sheet("Near Match - Review")
        ws.sheet_properties.tabColor = tab_color
        if df is None or (hasattr(df, "empty") and df.empty):
            ws["A1"] = "No Near Match records."
            ws["A1"].font = Font(name="Calibri", color="94A3B8", italic=True, size=10)
            return
        hdr_bg, hdr_fg = HDR_CONFIGS["near"]
        REASON_COLS = ["Match Reason", "Similarity %", "fuzzy_score", "confidence", "match_tier"]
        pr_cols  = [c + "_pr"     for c in WANT if c + "_pr"     in df.columns]
        gst_cols = [c + "_gstr2b" for c in WANT if c + "_gstr2b" in df.columns]
        reason_cols = [c for c in REASON_COLS if c in df.columns]
        _write_two_section(
            ws, df, hdr_bg, hdr_fg, pr_cols, gst_cols,
            pr_label="📚  AS PER BOOKS  (Near Match — Verify)",
            gst_label="📋  AS PER GSTR-2B  (Near Match — Verify)",
            reason_cols=reason_cols,
        )

    # ── "As Per Books" — every PR record exactly once ──────────────────────
    def write_per_books(wb, matched, missing_in_gstr2b, fuzzy, pr_dup):
        """
        PR perspective: one row per PR record.
        Sources:
          - matched        → has _pr and _gstr2b columns → extract _pr side
          - fuzzy          → has _pr and _gstr2b columns → extract _pr side
          - missing_in_gstr2b → plain PR columns (no suffix)
          - pr_dup         → plain PR columns (no suffix)
        Deduplicate on (gstin, invoice_number) to ensure 1 row per PR invoice.
        """
        ws = wb.create_sheet("As Per Books")
        ws.sheet_properties.tabColor = "1E40AF"
        hdr_bg, hdr_fg = HDR_CONFIGS["pr_view"]

        def _pr_side(df, status):
            if df is None or (hasattr(df, "empty") and df.empty):
                return pd.DataFrame()
            pr_cols = [c + "_pr" for c in WANT if c + "_pr" in df.columns]
            if not pr_cols:
                return pd.DataFrame()
            out = df[pr_cols].copy()
            out.columns = [c.replace("_pr", "") for c in out.columns]
            out.insert(0, "Status", status)
            return out

        def _plain(df, status):
            if df is None or (hasattr(df, "empty") and df.empty):
                return pd.DataFrame()
            cols = [c for c in WANT if c in df.columns]
            if not cols:
                return pd.DataFrame()
            out = df[cols].copy()
            out.insert(0, "Status", status)
            return out

        frames = [
            _pr_side(matched,            "Matched"),
            _plain(missing_in_gstr2b,   "Not in GSTR-2B"),
            _pr_side(fuzzy,              "Near Match (Verify)"),
            _plain(pr_dup,              "Duplicate (PR)"),
        ]
        all_pr = pd.concat(
            [f for f in frames if isinstance(f, pd.DataFrame) and not f.empty],
            ignore_index=True
        )
        if all_pr.empty:
            ws["A1"] = "No Books records available."
            return
        # Deduplicate: keep first occurrence per GSTIN + Invoice Number
        key_cols = [c for c in ["gstin", "invoice_number"] if c in all_pr.columns]
        if key_cols:
            all_pr = all_pr.drop_duplicates(subset=key_cols, keep="first")
        _write_rows(ws, all_pr,
                    PatternFill("solid", fgColor=hdr_bg),
                    Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # ── "As Per GSTR-2B" — every GSTR-2B record exactly once ──────────────
    def write_per_gstr2b(wb, matched, missing_in_books, fuzzy, gst_dup):
        """
        GSTR-2B perspective: one row per GSTR-2B record.
        Sources:
          - matched        → has _pr and _gstr2b columns → extract _gstr2b side
          - fuzzy          → has _pr and _gstr2b columns → extract _gstr2b side
          - missing_in_books → plain GSTR-2B columns (no suffix)
          - gst_dup        → plain GSTR-2B columns (no suffix)
        Deduplicate on (gstin, invoice_number) to ensure 1 row per GSTR-2B invoice.
        """
        ws = wb.create_sheet("As Per GSTR-2B")
        ws.sheet_properties.tabColor = "166534"
        hdr_bg, hdr_fg = HDR_CONFIGS["gb_view"]

        def _gstr_side(df, status):
            if df is None or (hasattr(df, "empty") and df.empty):
                return pd.DataFrame()
            g_cols = [c + "_gstr2b" for c in WANT if c + "_gstr2b" in df.columns]
            if not g_cols:
                return pd.DataFrame()
            out = df[g_cols].copy()
            out.columns = [c.replace("_gstr2b", "") for c in out.columns]
            out.insert(0, "Status", status)
            return out

        def _plain(df, status):
            if df is None or (hasattr(df, "empty") and df.empty):
                return pd.DataFrame()
            cols = [c for c in WANT if c in df.columns]
            if not cols:
                return pd.DataFrame()
            out = df[cols].copy()
            out.insert(0, "Status", status)
            return out

        frames = [
            _gstr_side(matched,         "Matched"),
            _plain(missing_in_books,   "Not in Books"),
            _gstr_side(fuzzy,           "Near Match (Verify)"),
            _plain(gst_dup,            "Duplicate (2B)"),
        ]
        all_gb = pd.concat(
            [f for f in frames if isinstance(f, pd.DataFrame) and not f.empty],
            ignore_index=True
        )
        if all_gb.empty:
            ws["A1"] = "No GSTR-2B records available."
            return
        # Deduplicate: keep first occurrence per GSTIN + Invoice Number
        key_cols = [c for c in ["gstin", "invoice_number"] if c in all_gb.columns]
        if key_cols:
            all_gb = all_gb.drop_duplicates(subset=key_cols, keep="first")
        _write_rows(ws, all_gb,
                    PatternFill("solid", fgColor=hdr_bg),
                    Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # ── Extract result sets ────────────────────────────────────────────────
    matched = recon_results.get("matched",           pd.DataFrame())
    books   = recon_results.get("missing_in_books",  pd.DataFrame())
    gstr    = recon_results.get("missing_in_gstr2b", pd.DataFrame())
    fuzzy   = recon_results.get("fuzzy_candidates",  pd.DataFrame())
    pr_dup  = recon_results.get("pr_duplicates",     pd.DataFrame())
    gst_dup = recon_results.get("gstr2b_duplicates", pd.DataFrame())
    stats   = recon_results.get("stats", {})

    dup_all = pd.concat(
        [(pr_dup  if isinstance(pr_dup,  pd.DataFrame) else pd.DataFrame()),
         (gst_dup if isinstance(gst_dup, pd.DataFrame) else pd.DataFrame())],
        ignore_index=True
    )
    unrecon_frames = [d for d in [books, gstr, fuzzy] if isinstance(d, pd.DataFrame) and not d.empty]
    unrecon = pd.concat(unrecon_frames, ignore_index=True) if unrecon_frames else pd.DataFrame()
    all_frames = ([matched] if isinstance(matched, pd.DataFrame) and not matched.empty else []) + unrecon_frames
    all_records = pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()

    # ── Build Workbook ─────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    # Summary sheet
    ws_sum = wb.create_sheet("Summary")
    ws_sum.sheet_properties.tabColor = "00D4FF"
    ws_sum.column_dimensions["A"].width = 40
    ws_sum.column_dimensions["B"].width = 18
    ws_sum["A1"] = "GST Input Reconciliation Report"
    ws_sum["A1"].font = Font(name="Calibri", bold=True, size=16, color="0F172A")
    ws_sum["A1"].fill = PatternFill("solid", fgColor="E0F7FA")
    ws_sum["A2"] = "Prepared & Developed by Karthik LVN  |  9849270702"
    ws_sum["A2"].font = Font(name="Calibri", size=10, color="7C3AED", italic=True)
    ws_sum["A2"].fill = PatternFill("solid", fgColor="F3E8FF")
    ws_sum["A3"] = f"Generated: {dt.date.today().strftime('%d-%m-%Y')}"
    ws_sum["A3"].font = Font(name="Calibri", size=9, color="6B7280")
    for r, label, val, bg, fg, is_hdr in [
        (5,  "Category",                           "Count",                                    "1E40AF", "FFFFFF", True),
        (6,  "Total GSTR-2B Records",              stats.get("total_gstr2b", 0),               "EFF6FF", "1E3A8A", False),
        (7,  "Total Purchase Register Records",    stats.get("total_pr", 0),                   "EFF6FF", "1E3A8A", False),
        (8,  "Matched (Fully Reconciled)",         stats.get("total_matched", 0),              "ECFDF5", "065F46", False),
        (9,  "Missing in Books (In 2B, not in PR)",stats.get("missing_in_books_count", 0),    "FEF2F2", "991B1B", False),
        (10, "Missing in GSTR-2B (In PR, not 2B)",stats.get("missing_in_gstr2b_count", 0),   "FFF7ED", "92400E", False),
        (11, "Near Match (Review Required)",       stats.get("fuzzy_candidates_count", 0),    "FFFBEB", "B45309", False),
        (12, "PR Duplicates",                      stats.get("pr_duplicates_count", 0),       "F5F3FF", "4C1D95", False),
        (13, "GSTR-2B Duplicates",                 stats.get("gstr2b_duplicates_count", 0),   "F5F3FF", "4C1D95", False),
        (14, "Match Rate (%)",                     f"{stats.get('match_rate', 0):.2f}%",      "ECFDF5", "065F46", False),
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
    write_sheet(wb, "Matched",            matched,     "34D399", "matched")
    write_sheet(wb, "Missing in Books",   books,       "F87171", "books")
    write_sheet(wb, "Missing in GSTR-2B", gstr,        "FB923C", "gstr")
    write_near_match_sheet(wb, fuzzy,                  "FBBF24")
    write_sheet(wb, "Duplicates",         dup_all,     "A78BFA", "dup")
    write_sheet(wb, "Unreconciled",       unrecon,     "94A3B8", "unrecon")
    write_sheet(wb, "All Records",        all_records, "00D4FF", "all")
    write_per_books(wb,  matched, gstr,  fuzzy, pr_dup)
    write_per_gstr2b(wb, matched, books, fuzzy, gst_dup)

    buf = io.BytesIO()
    wb.save(buf)
    log_event("EXPORT", "Sheet-wise Excel report generated")
    return buf.getvalue()
