import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
from datetime import datetime
import re
import io
import os

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
    date: Any
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

def parse_zoho_pdf(pdf_file) -> List[ZohoRecord]:
    """Fallback extraction regex engine to read Zoho Gross and Fees from raw PDF strings safely."""
    records = []
    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
        
        # Broaden regex selectors to look for common numerical distributions safely
        gross_matches = re.findall(r"(?:Gross Amount|Total Paid|Amount|Gross)[:]?\s*([\d\.\,\$]+)", full_text, re.IGNORECASE)
        fee_matches = re.findall(r"(?:Merchant Fee|Fee|Charges)[:]?\s*([\d\.\,\$]+)", full_text, re.IGNORECASE)
        inv_matches = re.findall(r"(?:Invoice Number|Invoice #|Inv #)[:]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
        cust_matches = re.findall(r"(?:Customer Name|Customer|Bill To)[:]?\s*(.*)", full_text, re.IGNORECASE)
        
        # Safe length guard to protect against IndexError loops
        num_records = max(len(gross_matches), len(fee_matches), 1)
        
        for i in range(num_records):
            gross = clean_numeric_value(gross_matches[i]) if i < len(gross_matches) else 0.0
            fee = clean_numeric_value(fee_matches[i]) if i < len(fee_matches) else 0.0
            inv = inv_matches[i].strip() if i < len(inv_matches) else None
            cust = cust_matches[i].split('\n')[0].strip() if i < len(cust_matches) else None
            
            # Only add if we actually extracted meaningful amounts
            if gross > 0 or fee > 0:
                records.append(ZohoRecord(customer_name=cust, gross_amount=gross, merchant_fee=fee, invoice_number=inv))
                
        # Fallback structural check: if regex logic pulled zero lines, initialize an aggregate placeholder row
        if not records:
            records.append(ZohoRecord(customer_name=None, gross_amount=0.0, merchant_fee=0.0, invoice_number=None))
            
    except Exception as e:
        st.error(f"Error parsing Zoho PDF Summary: {e}")
    return records

# =====================================================================
# 4. STREAMLIT INTERFACE SETUP
# =====================================================================
st.set_page_config(page_title="D365 General Journal Automation", layout="wide")
st.title("D365 General Journal Automation Engine")
st.subheader("Daily Operational Reconciliations Matrix")

# Locate Permanent Masterlist inside GitHub Repository Root Folder Workspace
MASTERLIST_PATH = "Account_Masterlist.xlsx"

if not os.path.exists(MASTERLIST_PATH):
    st.error(f"❌ Core configuration file `{MASTERLIST_PATH}` missing from your GitHub repository root folder. Please commit it to your repository.")
    st.stop()

# Interactive User Panel Upload Areas
st.sidebar.header("📅 Daily Variable Inputs")
boa_file = st.sidebar.file_uploader("1. Bank of America Report (Excel/CSV)", type=["xlsx", "csv"])
zoho_file = st.sidebar.file_uploader("2. Zoho Transaction Summary (PDF/Excel/CSV)", type=["pdf", "xlsx", "csv"])
uploaded_invoices = st.sidebar.file_uploader("3. Customer Invoices (PDFs)", type=["pdf"], accept_multiple_files=True)

if not (boa_file and zoho_file):
    st.info("💡 Staging required: Please drop today's Bank of America report and matching Zoho summary sheet into the sidebar container panel.")
else:
    # -----------------------------------------------------------------
    # STEP A: LOAD PERMANENT MASTERLIST FROM GITHUB WORKSPACE
    # -----------------------------------------------------------------
    master_df = pd.read_excel(MASTERLIST_PATH)
    master_lookup: Dict[str, AccountMasterItem] = {}
    for _, row in master_df.iterrows():
        name_key = str(row['Account Name']).strip().lower()
        master_lookup[name_key] = AccountMasterItem(
            account_number=str(row['Account #']),
            account_name=str(row['Account Name']),
            payment_term=str(row.get('Payment Term', 'due-on-receipt')).strip().lower()
        )

    # -----------------------------------------------------------------
    # STEP B: LOAD INVOICE REPOSITORY (If names are missing)
    # -----------------------------------------------------------------
    invoice_cache = {}
    if uploaded_invoices:
        for inv in uploaded_invoices:
            inv_id = inv.name.replace(".pdf", "")
            extracted_name = parse_invoice_pdf(inv)
            if extracted_name:
                invoice_cache[inv_id] = extracted_name

    # -----------------------------------------------------------------
    # STEP C: DYNAMIC PARSING FOR BANK OF AMERICA (EXCEL & CSV SUPPORT)
    # -----------------------------------------------------------------
    if boa_file.name.endswith('.csv'):
        boa_df = pd.read_csv(boa_file)
    else:
        boa_df = pd.read_excel(boa_file)
    
    boa_df.columns = [str(col).strip().lower() for col in boa_df.columns]
    
    desc_target = next((c for c in ['description', 'transaction description', 'payee', 'memo'] if c in boa_df.columns), None)
    date_target = next((c for c in ['posting date', 'date', 'transaction date'] if c in boa_df.columns), None)
    amount_target = next((c for c in ['net amount', 'amount', 'net_amount'] if c in boa_df
