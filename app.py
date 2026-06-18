import streamlit as st
import pandas as pd
import re
import os
from pypdf import PdfReader

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="D365 Accounting Journal Generator", layout="wide")
st.title("📊 D365 Zoho, Stripe & BOA Journal Generator")
st.write(
    "Upload your daily processing packages below to build your D365 upload templates. "
    "BOA-only expense days are supported."
)

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Fixed D365 Settings")
company_id = st.sidebar.text_input("Company", value="bwa")
offset_account = st.sidebar.text_input("Offset Account", value="B1000002")
debit_ledger_acct = st.sidebar.text_input(
    "Debit Line Account (Ledger)",
    value="43170111-U26C05001-B735350-UOA003",
)

# Repo-level lookup filenames (must be committed in the same GitHub directory as app.py)
MASTER_FILE_NAME = "Customer Master Account File.xlsx"
CASH_CODE_FILE = "Cash Code Masterlist.xlsx"

# ─── FILE UPLOADERS ──────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    gateway_file = st.file_uploader(
        "1. Upload Zoho Payout Export (CSV/XLSX) or Stripe PDF",
        type=["csv", "xlsx", "pdf"],
    )
with col2:
    invoice_files = st.file_uploader(
        "2. Upload Invoice PDF(s) — all invoices for this payout batch",
        type=["pdf", "csv", "xlsx", "txt"],
        accept_multiple_files=True,
    )
with col3:
    boa_file = st.file_uploader(
        "3. Upload Bank of America Statement (CSV/XLSX)",
        type=["csv", "xlsx"],
    )


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf(f) -> str:
    try:
        f.seek(0)
        return "\n".join(p.extract_text() or "" for p in PdfReader(f).pages)
    except Exception:
        return ""


def safe_float(v) -> float:
    try:
        return float(str(v).replace(",", "").replace("$", "").strip())
    except Exception:
        return 0.0


def clean_for_match(text) -> str:
    if pd.isna(text) or text is None:
        return ""
    t = str(text).lower()
    t = re.sub(r'\S+@\S+', '', t)
    t = re.sub(r'\b(llc|pllc|inc|corp|co|incorporated|limited|llp|dba)\b', ' ', t)
    t = re.sub(r'[^a-z0-9\s]', ' ', t)
    return " ".join(t.split())


def lookup_customer(name, cust_df, acct_col, name_col):
    key = clean_for_match(name)
    if not key:
