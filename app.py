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
    fallback_personal_name: Optional[str] = None  # Holds individual's name for documentation trace

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
# 3. ADVANCED EXTRACTION ENGINE
# =====================================================================
def extract_invoice_metadata_intelligent(pdf_file) -> Dict[str, Any]:
    """Scans invoice structure to isolate business entities from line descriptions when needed."""
    result = {"customer_name": None, "invoice_number": None, "gross_amount": 0.0, "fallback_personal_name": None}
    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
            
        full_text_clean = " ".join(full_text.split())
        
        # 1. Capture basic reference tags
        inv_num = pdf_file.name.replace(".pdf", "")
        inv_match = re.search(r"(INV-\d+)", full_text_clean, re.IGNORECASE)
        result["invoice_number"] = inv_match.group(1).strip() if inv_match else inv_num
        
        all_decimals = [clean_numeric_value(n) for n in re.findall(r"\b\d+(?:[\.,]\d{2})+\b", full_text_clean)]
        result["gross_amount"] = max(all_decimals) if all_decimals else 0.0
        
        # 2. Capture 'Bill to' baseline string value
        bill_to_name = None
        bill_to_match = re.search(r"Bill\s+to[:]?\s*([A-Za-z0-9\s\.\,\_\-]+)", full_text, re.IGNORECASE)
        if bill_to_match:
            bill_to_name = bill_to_match.group(1).split('\n')[0].strip()
            
        # 3. Apply business entity contextual filtering matrix
        lines = full_text.split("\n")
        detected_business = None
        
        for line in lines:
            line_lower = line.lower()
            # If line is part of line items/descriptions but is not our company prefix
            if "-" in line and ("item" not in line_lower and "description" not in line_lower):
                # Isolate sub-entities wrapped inside the transaction context strings
                biz_match = re.search(r"-\s*([A-Za-z0-9\s\.]+)(?:\s*-|$)", line)
                if biz_match:
                    candidate = biz_match.group(1).strip()
                    if "inbody" not in candidate.lower() and candidate:
                        detected_business = candidate
                        break
                        
        if detected_business:
            result["customer_name"] = detected_business
        else:
            # If no corporate name variant exists, mark as fallback unrecorded item context
            result["customer_name"] = None
            result["fallback_personal_name"] = bill_to_name
            
    except Exception as e:
        st.error(f"Error executing intelligent metadata capture: {e}")
    return result

def parse_zoho_summary_pdf_bulletproof(pdf_file) -> List[ZohoRecord]:
    records = []
    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
            
        text_stream = " ".join(full_text.split())
        text_tokens = text_stream.split(" ")
        
        for idx, token in enumerate(text_tokens):
            if "INV-" in token.upper() or (token.startswith("06") and len(token) >= 10):
                inv_id = re.sub(r'[^A-Za-z0-9\-]', '', token)
                
                forward_pool = []
                for step in range(1, 15):
                    if idx + step < len(text_tokens):
                        potential_num = text_tokens[idx + step]
                        if re.search(r"\d+\.\d{2}", potential_num):
                            forward_pool.append(clean_numeric_value(potential_num))
                
                if len(forward_pool) >= 1:
                    gross = forward_pool[0]
                    fee = forward_pool[1] if len(forward_pool) >= 2 else 0.0
                    
                    if gross > 0 and not any(r.invoice_number == inv_id for r in records):
                        records.append(ZohoRecord(
                            customer_name=None,
                            gross_amount=gross,
                            merchant_fee=fee,
                            invoice_number=inv_id
                        ))
    except Exception as e:
        st.error(f"Error executing token stream layout parser: {e}")
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
zoho_file = st.sidebar.file_uploader("2. Zoho Transaction Summary or Direct Invoices (PDF/Excel/CSV)", type=["pdf", "xlsx", "csv"])
uploaded_invoices = st.sidebar.file_uploader("3. Extra Customer Invoices (PDFs) [Optional]", type=["pdf"], accept_multiple_files=True)

if not (boa_file and zoho_file):
    st.info("💡 Staging required: Please drop today's Bank of America report and matching Zoho summary sheet into the sidebar container panel.")
else:
    # STEP A: LOAD MASTERLIST
    master_df = pd.read_excel(MASTERLIST_PATH)
    master_df.columns = [str(col).strip() for col in master_df.columns]
    master_headers_lower = {str(col).lower(): str(col) for col in master_df.columns}
    
    ml_name_col = next((master_headers_lower[k] for k in ['account name', 'name', 'customer name'] if k in master_headers_lower), None)
    ml_num_col = next((master_headers_lower[k] for k in ['account #', 'account number', 'account no', 'account'] if k in master_headers_lower), None)
    ml_term_col = next((master_headers_lower[k] for k in ['payment term', 'payment terms', 'terms'] if k in master_headers_lower), None)
    
    if not ml_name_col or not ml_num_col:
        st.error("❌ Could not identify definitive baseline 'Account Name' or 'Account #' tracking headers inside Masterlist spreadsheet.")
        st.stop()
        
    master_lookup: Dict[str, AccountMasterItem] = {}
    for _, row in master_df.iterrows():
        name_val = str(row[ml_name_col]).strip()
        name_key = name_val.lower()
        num_val = str(row[ml_num_col]).strip()
        term_val = str(row.get(ml_term_col, 'due-on-receipt')).strip().lower() if ml_term_col else 'due-on-receipt'
        
        master_lookup[name_key] = AccountMasterItem(
            account_number=num_val,
            account_name=name_val,
            payment_term=term_val
        )

    zoho_records: List
