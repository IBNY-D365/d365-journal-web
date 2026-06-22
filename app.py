import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
from datetime import datetime
import re
import io

# =====================================================================
# 1. HARDCODED CONFIGURATIONS & MAPPINGS (From Specification Document)
# =====================================================================
CASH_CODE_MAPPING = {
    "due-on-receipt": ("AR001", "AR Collection_AP"),
    "monthly": ("AR002", "AR Collection_MPP"),
    "financing": ("AR003", "AR Collection_Financing"),
    "leasing": ("AR004", "AR Collection_Leasing"),
    "net 1 day": ("AR005", "AR Collection_Net_1Day"),
    "net 10 days": ("AR006", "AR Collection_Net_10Days"),
    "net 25 days": ("AR007", "AR Collection_Net_25Days"),
    "net 30 days": ("AR008", "AR Collection_Net_30Days"),
    "net 40 days": ("AR009", "AR Collection_Net_40Days"),
    "net 45 days": ("AR010", "AR Collection_Net_45Days"),
    "net 60 days": ("AR011", "AR Collection_Net_60Days"),
    "fallback": ("AR012", "AR Collection_Other")
}

OFFSET_ACCOUNT_ROUTING = {
    "3371": "B1000002",
    "3924": "B1000003",
    "3384": "B1000001"
}

D365_TEMPLATE_COLUMNS = [
    "Date", "Voucher", "Account name", "Company", "Account type", "Account",
    "Posting Profile", "Cash code", "Description", "Debit", "Credit",
    "Item sales tax group", "Sales tax code", "Offset company", "Bank Account Type",
    "Offset account", "Offset transaction text", "Currency", "Exchange rate",
    "Item sales tax group2", "Sales group", "Withholding tax group",
    "Release date", "Reversing entry", "Reversing date"
]

# =====================================================================
# 2. DATA UTILITIES & MODELS
# =====================================================================
class BOARecord(BaseModel):
    date: datetime.date
    description: str
    net_amount: float
    source_account: str

class ZohoRecord(BaseModel):
    customer_name: Optional[str] = None
    gross_amount: float
    merchant_fee: float
    invoice_number: Optional[str] = None

class AccountMasterItem(BaseModel):
    account_number: str
    account_name: str
    payment_term: str

def clean_numeric_value(val: Any) -> float:
    """Removes currency symbols, commas, and whitespace to safely parse numbers."""
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    cleaned_str = str(val).strip().replace('$', '').replace(',', '')
    try:
        return float(cleaned_str)
    except ValueError:
        return 0.0

# =====================================================================
# 3. COMPONENT PARSERS
# =====================================================================
def parse_invoice_pdf(pdf_file) -> Optional[str]:
    """Extracts customer name from the 'Bill to' section of an invoice PDF."""
    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
        
        match = re.search(r"Bill\s+to[:]?\s*(.*)", full_text, re.IGNORECASE)
        if match:
            return match.group(1).split('\n')[0].strip()
    except Exception as e:
        st.error(f"Error parsing PDF invoice: {e}")
    return None

# =====================================================================
# 4. STREAMLIT INTERFACE SETUP
# =====================================================================
st.set_page_config(page_title="D365 General Journal Automation", layout="wide")
st.title("D365 General Journal Automation Engine")
st.subheader("Automate Zoho Payment Reconciliations into D365 Templates")

# Keep all uploader elements organized globally in the sidebar
st.sidebar.header("Data Sources Upload")
masterlist_file = st.sidebar.file_uploader("1. Account Masterlist (Excel)", type=["xlsx"])
boa_file = st.sidebar.file_uploader("2. Bank of America Report (CSV)", type=["csv"])
zoho_file = st.sidebar.file_uploader("3. Zoho Transaction Summary (CSV/XLSX)", type=["csv", "xlsx"])
uploaded_invoices = st.sidebar.file_uploader("4. Backup Customer Invoices (PDFs)", type=["pdf"], accept_multiple_files=True)

if not (masterlist_file and boa_file and zoho_file):
    st.info("💡 Complete the layout configuration requirements by dropping your Source Masterlist, Bank Reports, and Zoho Summaries into the sidebar panel.")
else:
    # -----------------------------------------------------------------
    # STEP A: LOAD MASTERLIST
    # -----------------------------------------------------------------
    master_df = pd.read_excel(masterlist_file)
    master_lookup: Dict[str, AccountMasterItem] = {}
    for _, row in master_df.iterrows():
        name_key = str(row['Account Name']).strip().lower()
        master_lookup[name_key] = AccountMasterItem(
            account_number=str(row['Account #']),
            account_name=str(row['Account Name']),
            payment_term=str(row.get('Payment Term', 'due-on-receipt')).strip().lower()
        )

    # -----------------------------------------------------------------
    # STEP B: LOAD INVOICE REPOSITORY (If any names are missing)
    # -----------------------------------------------------------------
    invoice_cache = {}
    if uploaded_invoices:
        for inv in uploaded_invoices:
            inv_id = inv.name.replace(".pdf", "")
            extracted_name = parse_invoice_pdf(inv)
            if extracted_name:
                invoice_cache[inv_id] = extracted_name

    # -----------------------------------------------------------------
