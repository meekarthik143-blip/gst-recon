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
    9-sheet Excel report:
      1  Summary          2  Matched (Books+2B side-by-side, like Near Match)
      3  Missing in Books 4  Missing in GSTR-2B
      5  Near Match-Review 6  Duplicates
      7  Unreconciled     8  All Records
      9  As Per Books     10 As Per GSTR-2B
    """
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import datetime as dt

    WHITE_BG  = PatternFill("solid", fgColor="FFFFFF")
    ALT_BG    = PatternFill("solid", fgColor="F8FAFC")
    DATA_FONT = Font(name="Calibri", size=9, color="111827")
    CTR  = Alignment(horizontal="center", vertical="center")
    LEFT = Alignment(horizontal="left",   vertical="center")

    HDR = {
        "matched": ("1F5E40","FFFFFF"), "books":   ("7F1D1D","FFFFFF"),
        "gstr":    ("78350F","FFFFFF"), "near":    ("1E3A5F","FBBF24"),
        "dup":     ("3B0764","FFFFFF"), "unrecon": ("1E293B","FFFFFF"),
        "all":     ("0F172A","00D4FF"), "pr_view": ("1E3A8A","FFFFFF"),
        "gb_view": ("14532D","FFFFFF"),
    }

    def _border():
        s = Side(style="thin", color="D1D5DB")
        return Border(left=s, right=s, top=s, bottom=s)

    def _fmt(val):
        try:
            if pd.isna(val): return ""
        except Exception:
            pass
        if hasattr(val, "strftime"): return val.strftime("%d-%m-%Y")
        if isinstance(val, str) and val.endswith(" 00:00:00"): return val[:10]
        if isinstance(val, float): return round(val, 2)
        return val

    def _lbl(c):
        return c.replace("_pr"," (Books)").replace("_gstr2b"," (2B)").replace("_"," ").strip().title()

    WANT = ["vendor_name","gstin","invoice_number","invoice_date",
            "taxable_value","cgst","sgst","igst","cess","total_gst","invoice_value"]

    # ── Generic single-section writer ──────────────────────────────────────
    def _write_rows(ws, df_out, hdr_fill, hdr_font, start_row=1):
        """Write header + data rows starting at start_row. Returns next free row."""
        cols = list(df_out.columns)
        for ci, col in enumerate(cols, 1):
            c = ws.cell(start_row, ci, _lbl(col))
            c.font = hdr_font; c.fill = hdr_fill
            c.alignment = CTR; c.border = _border()
        ws.row_dimensions[start_row].height = 20
        cur = start_row + 1
        for ri, row in enumerate(df_out.itertuples(index=False)):
            fill = WHITE_BG if ri % 2 == 0 else ALT_BG
            for ci, val in enumerate(row, 1):
                cell = ws.cell(cur, ci)
                cell.value = _fmt(val); cell.font = DATA_FONT
                cell.fill = fill; cell.border = _border()
                cell.alignment = CTR if isinstance(val,(int,float)) else LEFT
                if isinstance(val,(int,float)) and not isinstance(val,bool):
                    cell.number_format = "#,##0.00"
            cur += 1
        # Auto-fit
        for ci, col in enumerate(cols, 1):
            vals = [str(ws.cell(r,ci).value or "") for r in range(start_row, min(cur, start_row+52))]
            ws.column_dimensions[get_column_letter(ci)].width = min(
                max((max(len(v) for v in vals) if vals else 10)+2, 12), 42
            )
        return cur

    # ── Comparative sheet: Books cols + GSTR-2B cols in SAME ROW ───────────
    def _write_comparative(ws, df, hdr_bg, hdr_fg, reason_cols=None):
        """
        Shows both PR side and GSTR-2B side in same row — like Near Match Review.
        Layout: [Reason cols] [Books cols] [GSTR-2B cols]
        All matched in one wide table. freeze at A2.
        """
        hdr_fill = PatternFill("solid", fgColor=hdr_bg)
        hdr_font = Font(name="Calibri", bold=True, color=hdr_fg, size=9)

        pr_cols  = [c+"_pr"     for c in WANT if c+"_pr"     in df.columns]
        gst_cols = [c+"_gstr2b" for c in WANT if c+"_gstr2b" in df.columns]
        rcols    = [c for c in (reason_cols or []) if c in df.columns]

        all_cols = rcols + pr_cols + gst_cols
        if not all_cols:
            all_cols = [c for c in df.columns if not c.startswith("_")]

        df_out = df[all_cols].copy()
        ws.freeze_panes = "A2"
        _write_rows(ws, df_out, hdr_fill, hdr_font, start_row=1)

    # ── Plain single-source sheet ───────────────────────────────────────────
    def write_sheet(wb, title, df, tab_color, fill_key):
        ws = wb.create_sheet(title)
        ws.sheet_properties.tabColor = tab_color
        if df is None or (hasattr(df,"empty") and df.empty):
            ws["A1"] = f"No {title} records."
            ws["A1"].font = Font(name="Calibri", color="94A3B8", italic=True, size=10)
            return
        hdr_bg, hdr_fg = HDR.get(fill_key, ("0F172A","FFFFFF"))
        # Plain columns (no _pr/_gstr2b suffix)
        plain = [c for c in WANT if c in df.columns]
        use   = plain or [c for c in df.columns if not c.startswith("_")][:14]
        df_out = df[use].copy()
        ws.freeze_panes = "A2"
        _write_rows(ws, df_out,
                    PatternFill("solid", fgColor=hdr_bg),
                    Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # ── Matched sheet: Books (PR) columns only — clean readable format ────────
    def write_matched_sheet(wb, df):
        ws = wb.create_sheet("Matched")
        ws.sheet_properties.tabColor = "34D399"
        if df is None or (hasattr(df,"empty") and df.empty):
            ws["A1"] = "No Matched records."
            ws["A1"].font = Font(name="Calibri", color="94A3B8", italic=True, size=10)
            return
        hdr_bg, hdr_fg = HDR["matched"]
        hdr_fill = PatternFill("solid", fgColor=hdr_bg)
        hdr_font = Font(name="Calibri", bold=True, color=hdr_fg, size=9)

        # Extract PR side columns with (Books) suffix labels
        pr_cols = [c + "_pr" for c in WANT if c + "_pr" in df.columns]
        if not pr_cols:
            # Fallback: plain columns
            pr_cols = [c for c in WANT if c in df.columns]
        if not pr_cols:
            ws["A1"] = "No Matched records."; return

        df_out = df[pr_cols].copy()
        # Rename: vendor_name_pr → Vendor Name (Books), gstin_pr → Gstin (Books), etc.
        df_out.columns = [
            c.replace("_pr", "").replace("_", " ").title() + " (Books)"
            for c in df_out.columns
        ]
        ws.freeze_panes = "A2"
        _write_rows(ws, df_out, hdr_fill, hdr_font, start_row=1)



    # ── Near Match sheet: comparative with reason cols ──────────────────────
    def write_near_match_sheet(wb, df):
        ws = wb.create_sheet("Near Match - Review")
        ws.sheet_properties.tabColor = "FBBF24"
        if df is None or (hasattr(df,"empty") and df.empty):
            ws["A1"] = "No Near Match records."
            ws["A1"].font = Font(name="Calibri", color="94A3B8", italic=True, size=10)
            return
        hdr_bg, hdr_fg = HDR["near"]
        REASON = ["Match Reason","Similarity %","fuzzy_score","confidence","match_tier"]
        _write_comparative(ws, df, hdr_bg, hdr_fg, reason_cols=REASON)

    # ── As Per Books ────────────────────────────────────────────────────────
    def write_per_books(wb, matched, missing_in_gstr2b, fuzzy, pr_dup):
        ws = wb.create_sheet("As Per Books")
        ws.sheet_properties.tabColor = "1E40AF"
        hdr_bg, hdr_fg = HDR["pr_view"]

        def _pr_side(df, status):
            if df is None or (hasattr(df,"empty") and df.empty): return pd.DataFrame()
            cols = [c+"_pr" for c in WANT if c+"_pr" in df.columns]
            if not cols: return pd.DataFrame()
            out = df[cols].copy(); out.columns = [c.replace("_pr","") for c in out.columns]
            out.insert(0,"Status",status); return out

        def _plain(df, status):
            if df is None or (hasattr(df,"empty") and df.empty): return pd.DataFrame()
            cols = [c for c in WANT if c in df.columns]
            if not cols: return pd.DataFrame()
            out = df[cols].copy(); out.insert(0,"Status",status); return out

        frames = [
            _pr_side(matched,           "Matched"),
            _plain(missing_in_gstr2b,  "Not in GSTR-2B"),
            _pr_side(fuzzy,             "Near Match (Verify)"),
            _plain(pr_dup,             "Duplicate (PR)"),
        ]
        all_pr = pd.concat([f for f in frames if isinstance(f,pd.DataFrame) and not f.empty],
                           ignore_index=True)
        if all_pr.empty:
            ws["A1"] = "No Books records available."; return
        key = [c for c in ["gstin","invoice_number"] if c in all_pr.columns]
        if key: all_pr = all_pr.drop_duplicates(subset=key, keep="first")
        ws.freeze_panes = "A2"
        _write_rows(ws, all_pr,
                    PatternFill("solid", fgColor=hdr_bg),
                    Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # ── As Per GSTR-2B ─────────────────────────────────────────────────────
    def write_per_gstr2b(wb, matched, missing_in_books, fuzzy, gst_dup):
        ws = wb.create_sheet("As Per GSTR-2B")
        ws.sheet_properties.tabColor = "166534"
        hdr_bg, hdr_fg = HDR["gb_view"]

        def _gstr_side(df, status):
            if df is None or (hasattr(df,"empty") and df.empty): return pd.DataFrame()
            cols = [c+"_gstr2b" for c in WANT if c+"_gstr2b" in df.columns]
            if not cols: return pd.DataFrame()
            out = df[cols].copy(); out.columns = [c.replace("_gstr2b","") for c in out.columns]
            out.insert(0,"Status",status); return out

        def _plain(df, status):
            if df is None or (hasattr(df,"empty") and df.empty): return pd.DataFrame()
            cols = [c for c in WANT if c in df.columns]
            if not cols: return pd.DataFrame()
            out = df[cols].copy(); out.insert(0,"Status",status); return out

        frames = [
            _gstr_side(matched,        "Matched"),
            _plain(missing_in_books,  "Not in Books"),
            _gstr_side(fuzzy,          "Near Match (Verify)"),
            _plain(gst_dup,           "Duplicate (2B)"),
        ]
        all_gb = pd.concat([f for f in frames if isinstance(f,pd.DataFrame) and not f.empty],
                           ignore_index=True)
        if all_gb.empty:
            ws["A1"] = "No GSTR-2B records available."; return
        key = [c for c in ["gstin","invoice_number"] if c in all_gb.columns]
        if key: all_gb = all_gb.drop_duplicates(subset=key, keep="first")
        ws.freeze_panes = "A2"
        _write_rows(ws, all_gb,
                    PatternFill("solid", fgColor=hdr_bg),
                    Font(name="Calibri", bold=True, color=hdr_fg, size=9))

    # ── Extract result sets ─────────────────────────────────────────────────
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
    unrecon_frames = [d for d in [books,gstr,fuzzy] if isinstance(d,pd.DataFrame) and not d.empty]
    unrecon = pd.concat(unrecon_frames, ignore_index=True) if unrecon_frames else pd.DataFrame()

    # ── Build Workbook ──────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    # 1. Summary
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

    for r, label, val, bg, fg, bold in [
        (5,  "Category",                            "Count",                                   "1E40AF","FFFFFF",True),
        (6,  "Total GSTR-2B Records",               stats.get("total_gstr2b",0),               "EFF6FF","1E3A8A",False),
        (7,  "Total Purchase Register Records",     stats.get("total_pr",0),                   "EFF6FF","1E3A8A",False),
        (8,  "Matched (Fully Reconciled)",          stats.get("total_matched",0),              "ECFDF5","065F46",False),
        (9,  "Missing in Books (In 2B, not in PR)", stats.get("missing_in_books_count",0),     "FEF2F2","991B1B",False),
        (10, "Missing in GSTR-2B (In PR, not 2B)", stats.get("missing_in_gstr2b_count",0),    "FFF7ED","92400E",False),
        (11, "Near Match (Review Required)",        stats.get("fuzzy_candidates_count",0),     "FFFBEB","B45309",False),
        (12, "PR Duplicates",                       stats.get("pr_duplicates_count",0),        "F5F3FF","4C1D95",False),
        (13, "GSTR-2B Duplicates",                  stats.get("gstr2b_duplicates_count",0),    "F5F3FF","4C1D95",False),
        (14, "Match Rate (%)",                      f"{stats.get('match_rate',0):.2f}%",       "ECFDF5","065F46",False),
    ]:
        lc = ws_sum.cell(r,1,label); vc = ws_sum.cell(r,2,val)
        for cell in [lc,vc]:
            cell.fill  = PatternFill("solid", fgColor=bg)
            cell.font  = Font(name="Calibri", bold=bold, size=10, color=fg)
            cell.border = _border(); cell.alignment = LEFT
        vc.alignment = CTR
        ws_sum.row_dimensions[r].height = 18

    # 2. Matched — Books + 2B columns side by side (comparative)
    write_matched_sheet(wb, matched)

    # 3. Missing in Books
    write_sheet(wb, "Missing in Books",   books,   "F87171", "books")

    # 4. Missing in GSTR-2B
    write_sheet(wb, "Missing in GSTR-2B", gstr,   "FB923C", "gstr")

    # 5. Near Match - Review (comparative with reason cols)
    write_near_match_sheet(wb, fuzzy)

    # 6. Duplicates
    write_sheet(wb, "Duplicates",         dup_all, "A78BFA", "dup")

    # 7. Unreconciled
    write_sheet(wb, "Unreconciled",       unrecon, "94A3B8", "unrecon")

    # 8. All Records — plain extract (PR side if merged, else plain)
    ws_all = wb.create_sheet("All Records")
    ws_all.sheet_properties.tabColor = "00D4FF"
    if not (isinstance(matched, pd.DataFrame) and not matched.empty):
        all_records = unrecon
    else:
        # Extract PR side of matched + all unreconciled
        pr_cols = [c+"_pr" for c in WANT if c+"_pr" in matched.columns]
        if pr_cols:
            m_pr = matched[pr_cols].copy()
            m_pr.columns = [c.replace("_pr","") for c in m_pr.columns]
            m_pr.insert(0,"Status","Matched")
        else:
            m_pr = pd.DataFrame()
        all_records = pd.concat(
            [f for f in [m_pr]+unrecon_frames if isinstance(f,pd.DataFrame) and not f.empty],
            ignore_index=True
        )
    if all_records is not None and not all_records.empty:
        plain = [c for c in WANT if c in all_records.columns]
        cols_out = (["Status"] if "Status" in all_records.columns else []) + plain
        df_all = all_records[cols_out].copy()
        ws_all.freeze_panes = "A2"
        _write_rows(ws_all, df_all,
                    PatternFill("solid", fgColor="0F172A"),
                    Font(name="Calibri", bold=True, color="00D4FF", size=9))
    else:
        ws_all["A1"] = "No records."

    # 9. As Per Books / As Per GSTR-2B
    write_per_books(wb,  matched, gstr,  fuzzy, pr_dup)
    write_per_gstr2b(wb, matched, books, fuzzy, gst_dup)

    buf = io.BytesIO()
    wb.save(buf)
    log_event("EXPORT", "9-sheet Excel report generated")
    return buf.getvalue()
