"""
GST Input Reconciliation System – Enterprise Edition
Sample Data Generator
Prepared & Developed by Karthik LVN

Generates realistic sample data for testing the reconciliation system:
  - sample_purchase_register.xlsx  (100 rows)
  - sample_gstr2b.xlsx             (100 rows with intentional mismatches)

Usage:
    python sample_data/generate_samples.py
"""

import random
import datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

NUM_ROWS = 100
MATCH_RATE = 0.60       # 60% perfect matches
MISSING_BOOKS = 0.15    # 15% in GSTR-2B but not in PR
MISSING_GSTR = 0.15     # 15% in PR but not in GSTR-2B
GST_DIFF = 0.05         # 5% GST differences
DATE_DIFF = 0.05        # 5% date differences

OUTPUT_DIR = Path(__file__).parent

# ---------------------------------------------------------------------------
# Master Data
# ---------------------------------------------------------------------------

VENDORS = [
    ("Reliance Industries Ltd", "27AAACR5055K1ZK"),
    ("Tata Motors Limited", "27AAACT2727Q1ZW"),
    ("Infosys Limited", "29AACCI1681G1ZM"),
    ("Wipro Limited", "29AABCW0788K1ZO"),
    ("HDFC Bank Limited", "24AAACH5964A1ZG"),
    ("ITC Limited", "32AAACI3407R1ZC"),
    ("Larsen & Toubro Ltd", "27AAACL0260L1ZM"),
    ("HCL Technologies Ltd", "29AAACH2892K1ZH"),
    ("Bajaj Auto Limited", "27AAACB0671N1ZV"),
    ("Asian Paints Limited", "24AAACA1840G1Z9"),
    ("Maruti Suzuki India Ltd", "07AAACM3025H1ZD"),
    ("Sun Pharmaceutical Ind", "24AAACS1840A1ZJ"),
    ("Tech Mahindra Limited", "27AAACT8309R1ZB"),
    ("Hindustan Unilever Ltd", "27AAACH8678H1Z2"),
    ("Adani Enterprises Ltd", "24AAACA9928L1ZX"),
    ("Godrej Consumer Products", "27AAACG1565E1ZB"),
    ("Britannia Industries Ltd", "33AAACB0680R1ZU"),
    ("Dabur India Limited", "07AAACD5483H1ZE"),
    ("Marico Limited", "27AAACM3025H1ZT"),
    ("Hero MotoCorp Limited", "07AAACH0717L1ZK"),
]

INVOICE_PREFIXES = ["INV", "BILL", "GST", "TX", "VCH", "DOC"]

FY_START = datetime.date(2025, 4, 1)
FY_END = datetime.date(2026, 3, 31)


def random_date() -> datetime.date:
    delta = (FY_END - FY_START).days
    return FY_START + datetime.timedelta(days=random.randint(0, delta))


def format_date(d: datetime.date) -> str:
    return d.strftime("%d-%m-%Y")


def random_invoice_number() -> str:
    prefix = random.choice(INVOICE_PREFIXES)
    num = random.randint(1000, 99999)
    suffix = random.choice(["", "/25-26", "-2026", ""])
    return f"{prefix}{num}{suffix}"


def random_amounts() -> dict:
    taxable = round(random.uniform(5000, 500000), 2)
    gst_rate = random.choice([0.05, 0.12, 0.18, 0.28])

    if random.random() < 0.4:  # IGST (inter-state)
        igst = round(taxable * gst_rate, 2)
        cgst = sgst = 0.0
    else:  # CGST + SGST (intra-state)
        igst = 0.0
        cgst = round(taxable * gst_rate / 2, 2)
        sgst = cgst

    cess = round(taxable * 0.01, 2) if random.random() < 0.1 else 0.0
    total_gst = round(igst + cgst + sgst + cess, 2)
    invoice_value = round(taxable + total_gst, 2)

    return {
        "taxable_value": taxable,
        "cgst": cgst,
        "sgst": sgst,
        "igst": igst,
        "cess": cess,
        "total_gst": total_gst,
        "invoice_value": invoice_value,
    }


# ---------------------------------------------------------------------------
# Data Generation
# ---------------------------------------------------------------------------

def generate_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Generate Purchase Register and GSTR-2B DataFrames.

    Returns:
        (pr_df, gstr2b_df)
    """
    pr_rows = []
    gstr_rows = []

    # Build a pool of base invoices
    all_invoices = []
    for _ in range(NUM_ROWS):
        vendor, gstin = random.choice(VENDORS)
        inv_no = random_invoice_number()
        inv_date = random_date()
        amounts = random_amounts()

        all_invoices.append({
            "vendor_name": vendor,
            "gstin": gstin,
            "invoice_number": inv_no,
            "invoice_date": format_date(inv_date),
            **amounts,
        })

    # Shuffle
    random.shuffle(all_invoices)

    n = len(all_invoices)
    idx = 0

    # Perfect matches (60%)
    n_match = int(n * MATCH_RATE)
    for i in range(n_match):
        row = all_invoices[idx].copy()
        pr_rows.append(row.copy())
        gstr_rows.append(row.copy())
        idx += 1

    # GST differences (5%)
    n_gst_diff = int(n * GST_DIFF)
    for i in range(n_gst_diff):
        row = all_invoices[idx].copy()
        pr_rows.append(row.copy())
        gstr_row = row.copy()
        # Introduce a small GST error
        gstr_row["total_gst"] = round(gstr_row["total_gst"] + random.uniform(-50, 50), 2)
        gstr_row["invoice_value"] = round(gstr_row["taxable_value"] + gstr_row["total_gst"], 2)
        gstr_rows.append(gstr_row)
        idx += 1

    # Date differences (5%)
    n_date_diff = int(n * DATE_DIFF)
    for i in range(n_date_diff):
        row = all_invoices[idx].copy()
        pr_rows.append(row.copy())
        gstr_row = row.copy()
        # Introduce a date difference (1-3 days off)
        d = datetime.datetime.strptime(row["invoice_date"], "%d-%m-%Y").date()
        d2 = d + datetime.timedelta(days=random.choice([1, 2, 3, -1, -2]))
        gstr_row["invoice_date"] = format_date(d2)
        gstr_rows.append(gstr_row)
        idx += 1

    # Missing in books (15%) — only in GSTR-2B
    n_miss_books = int(n * MISSING_BOOKS)
    for i in range(n_miss_books):
        if idx < len(all_invoices):
            gstr_rows.append(all_invoices[idx].copy())
            idx += 1

    # Missing in GSTR-2B (15%) — only in PR
    n_miss_gstr = int(n * MISSING_GSTR)
    for i in range(n_miss_gstr):
        if idx < len(all_invoices):
            pr_rows.append(all_invoices[idx].copy())
            idx += 1

    # Add a few duplicates to PR
    if len(pr_rows) > 5:
        duplicates = random.sample(pr_rows[:20], min(3, len(pr_rows)))
        pr_rows.extend(duplicates)

    random.shuffle(pr_rows)
    random.shuffle(gstr_rows)

    pr_df = pd.DataFrame(pr_rows)
    gstr_df = pd.DataFrame(gstr_rows)

    # Add PR-specific columns (typical real-world Purchase Register)
    pr_df["narration"] = pr_df.apply(
        lambda r: f"Purchase from {r['vendor_name'][:20]}", axis=1
    )
    pr_df["ledger"] = "Purchase A/c"
    pr_df["place_of_supply"] = pr_df["gstin"].str[:2]
    pr_df["is_rcm"] = pr_df.apply(lambda _: random.choice([False, False, False, True]), axis=1)

    # Add GSTR-2B-specific columns
    gstr_df["supply_type"] = gstr_df["igst"].apply(
        lambda x: "INTER" if x > 0 else "INTRA"
    )
    gstr_df["itc_available"] = gstr_df.apply(
        lambda _: random.choice([True, True, True, False]), axis=1
    )
    gstr_df["reverse_charge"] = gstr_df.apply(
        lambda _: random.choice([False, False, False, True]), axis=1
    )
    gstr_df["filing_period"] = gstr_df["invoice_date"].apply(
        lambda d: f"{d[3:5]}/{d[6:]}" if len(d) >= 10 else ""
    )

    return pr_df, gstr_df


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("[*] Generating sample data...")
    random.seed(42)

    pr_df, gstr_df = generate_data()

    pr_path = OUTPUT_DIR / "sample_purchase_register.xlsx"
    gstr_path = OUTPUT_DIR / "sample_gstr2b.xlsx"

    pr_df.to_excel(pr_path, index=False, engine="openpyxl")
    gstr_df.to_excel(gstr_path, index=False, engine="openpyxl")

    print(f"[OK] Purchase Register : {pr_path} ({len(pr_df)} rows)")
    print(f"[OK] GSTR-2B           : {gstr_path} ({len(gstr_df)} rows)")
    print()
    print("Match breakdown:")
    print(f"  Perfect matches     : ~{int(100 * 0.60)} records")
    print(f"  GST differences     : ~{int(100 * 0.05)} records")
    print(f"  Date differences    : ~{int(100 * 0.05)} records")
    print(f"  Missing in Books    : ~{int(100 * 0.15)} records (GSTR-2B only)")
    print(f"  Missing in GSTR-2B  : ~{int(100 * 0.15)} records (PR only)")
    print(f"  Duplicates in PR    : 3 records")
    print()
    print("Use these files to test the reconciliation system!")


if __name__ == "__main__":
    main()
