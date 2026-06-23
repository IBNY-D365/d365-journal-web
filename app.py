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
    "Item sales tax group2", "Sales tax group", "Withholding tax group",
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
    """Rule 3.2: Extracts 'Bill to' customer name from invoice backup PDFs."""
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

def parse_zoho_pdf_safely(pdf_file) -> List[ZohoRecord]:
    """Layout-agnostic parser designed to pull individual line entries out of Zoho text blocks."""
    records = []
    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
            
        # Clean whitespaces and line breaks to normalize the stream layout
        full_text_clean = " ".join(full_text.split())
        
        # Pull clean groupings of financial decimals and identifiers sequentially
        # Pattern handles looking for blocks like "INV-060926000350 $1,000.00 $30.00"
        item_patterns = re.findall(r"(INV-\d+|[A-Za-z0-9\-]+)\s+\$?([\d\.,]+)\s+\$?([\d\.,]+)", full_text_clean)
        
        for match in item_patterns:
            inv_num = match[0].strip()
            gross = clean_numeric_value(match[1])
            fee = clean_numeric_value(match[2])
            
            if gross > 0:
                records.append(ZohoRecord(
                    customer_name=None, # To be resolved downstream via Invoice Master list routing rules
                    gross_amount=gross,
                    merchant_fee=fee,
                    invoice_number=inv_num
                ))
                
        # Fallback structural block: if sequential string grouping patterns yield 0 rows, check global number distributions
        if not records:
            all_decimals = [clean_numeric_value(n) for n in re.findall(r"\b\d+(?:[\.,]\d{2})+\b", full_text_clean)]
            unique_decimals = sorted(list(set([num for num in all_decimals if num > 0])), reverse=True)
            
            inv_match = re.search(r"(?i)(?:invoice\s*number|invoice\s*#|inv\s*#)[:]?\s*([A-Za-z0-9\-]+)", full_text_clean)
            inv_num = inv_match.group(1).strip() if inv_match else None
            
            if len(unique_decimals) >= 2:
                records.append(ZohoRecord(
                    customer_name=None,
                    gross_amount=unique_decimals[0],
                    merchant_fee=unique_decimals[-1],
                    invoice_number=inv_num
                ))
    except Exception as e:
        st.error(f"Error executing safe PDF parser engine: {e}")
        
    return records

# =====================================================================
# 4. STREAMLIT INTERFACE SETUP
# =====================================================================
st.set_page_config(page_title="D365 General Journal Automation", layout="wide")
st.title("D365 General Journal Automation Engine")
st.subheader("Daily Operational Reconciliations Matrix")

MASTERLIST_PATH = "Account Masterlist.xlsx"

if not os.path.exists(MASTERLIST_PATH):
    st.error(f"❌ Core configuration file `{MASTERLIST_PATH}` missing from your GitHub repository root folder. Please commit it to your repository.")
    st.stop()

st.sidebar.header("📅 Daily Variable Inputs")
boa_file = st.sidebar.file_uploader("1. Bank of America Report (Excel/CSV)", type=["xlsx", "csv"])
zoho_file = st.sidebar.file_uploader("2. Zoho Transaction Summary (PDF/Excel/CSV)", type=["pdf", "xlsx", "csv"])
uploaded_invoices = st.sidebar.file_uploader("3. Customer Invoices (PDFs)", type=["pdf"], accept_multiple_files=True)

if not (boa_file and zoho_file):
    st.info("💡 Staging required: Please drop today's Bank of America report and matching Zoho summary sheet into the sidebar container panel.")
else:
    # STEP A: LOAD MASTERLIST
    master_df = pd.read_excel(MASTERLIST_PATH)
    master_lookup: Dict[str, AccountMasterItem] = {}
    for _, row in master_df.iterrows():
        name_key = str(row['Account Name']).strip().lower()
        master_lookup[name_key] = AccountMasterItem(
            account_number=str(row['Account #']),
            account_name=str(row['Account Name']),
            payment_term=str(row.get('Payment Term', 'due-on-receipt')).strip().lower()
        )

    # STEP B: LOAD INVOICE REPOSITORY
    invoice_cache = {}
    if uploaded_invoices:
        for inv in uploaded_invoices:
            inv_id = inv.name.replace(".pdf", "")
            extracted_name = parse_invoice_pdf(inv)
            if extracted_name:
                invoice_cache[inv_id] = extracted_name

    # STEP C: PARSE BANK OF AMERICA
    if boa_file.name.endswith('.csv'):
        boa_df = pd.read_csv(boa_file)
    else:
        boa_df = pd.read_excel(boa_file)
    
    boa_df.columns = [str(col).strip().lower() for col in boa_df.columns]
    
    desc_target = next((c for c in ['description', 'transaction description', 'payee', 'memo'] if c in boa_df.columns), None)
    date_target = next((c for c in ['posting date', 'date', 'transaction date'] if c in boa_df.columns), None)
    amount_target = next((c for c in ['net amount', 'amount', 'net_amount'] if c in boa_df.columns), None)
    account_target = next((c for c in ['source account', 'account', 'account number', 'account_number'] if c in boa_df.columns), None)
            
    if desc_target is None:
        st.error("❌ Could not find a transaction 'Description' column variant in your Bank of America report.")
        st.stop()

    boa_records: List[BOARecord] = []
    for _, row in boa_df.iterrows():
        row_description = str(row.get(desc_target, ''))
        if "ZOHO" in row_description.upper():
            parsed_date = datetime.today().date()
            if date_target and pd.notna(row[date_target]):
                try:
                    parsed_date = pd.to_datetime(row[date_target]).date()
                except Exception:
                    pass
                    
            boa_records.append(BOARecord(
                date=parsed_date,
                description=row_description,
                net_amount=clean_numeric_value(row.get(amount_target, 0.0)),
                source_account=str(row.get(account_target, '')).strip() if account_target else ""
            ))

    # STEP D: PARSE ZOHO SUMMARY
    zoho_records: List[ZohoRecord] = []
    if zoho_file.name.endswith('.pdf'):
        zoho_records = parse_zoho_pdf_safely(zoho_file)
    else:
        if zoho_file.name.endswith('.csv'):
            zoho_df = pd.read_csv(zoho_file)
        else:
            zoho_df = pd.read_excel(zoho_file)
            
        zoho_df.columns = [str(c).strip() for c in zoho_df.columns]
        for _, row in zoho_df.iterrows():
            cust_name = str(row['Customer']).strip() if pd.notna(row.get('Customer')) else None
            inv_num = str(row['Invoice Number']).strip() if pd.notna(row.get('Invoice Number')) else None
            zoho_records.append(ZohoRecord(
                customer_name=cust_name,
                gross_amount=clean_numeric_value(row.get('Gross Amount', 0.0)),
                merchant_fee=clean_numeric_value(row.get('Merchant Fee', 0.0)),
                invoice_number=inv_num
            ))

    # Apply manual invoice mapping overrides globally to
