import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pydantic import BaseModel, Field
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
    
    # Strip spaces, $, and commas
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
    # STEP C: PARSE OPERATIONAL FILES WITH EXTRACTION PROTECTION
    # -----------------------------------------------------------------
    boa_df = pd.read_csv(boa_file)
    
    # Strip spaces and normalize headers to lowercase to match bulletproof criteria
    normalized_headers = {str(col).strip().lower(): str(col) for col in boa_df.columns}
    
    # Define our lowercase target check groups
    desc_target = next((normalized_headers[k] for k in ['description', 'transaction description', 'payee', 'memo'] if k in normalized_headers), None)
    date_target = next((normalized_headers[k] for k in ['posting date', 'date', 'transaction date'] if k in normalized_headers), None)
    amount_target = next((normalized_headers[k] for k in ['net amount', 'amount', 'net_amount'] if k in normalized_headers), None)
    account_target = next((normalized_headers[k] for k in ['source account', 'account', 'account number', 'account_number'] if k in normalized_headers), None)
            
    if desc_target is None:
        st.error("❌ Could not find a transaction 'Description' column variant in your Bank of America CSV. Please review file headers.")
        st.stop()

    boa_records: List[BOARecord] = []
    for _, row in boa_df.iterrows():
        raw_desc = str(row.get(desc_target, ''))
        
        if "ZOHO" in raw_desc.upper():  # Rule 3.1 Filtering Logic
            parsed_date = datetime.today().date()
            if date_target and pd.notna(row[date_target]):
                try:
                    parsed_date = pd.to_datetime(row[date_target]).date()
                except Exception:
                    pass
                    
            boa_records.append(BOARecord(
                date=parsed_date,
                description=raw_desc,
                net_amount=clean_numeric_value(row.get(amount_target, 0.0)),
                source_account=str(row.get(account_target, '')).strip() if account_target else ""
            ))

    # Zoho Summary Parsing Block
    if zoho_file.name.endswith('.csv'):
        zoho_df = pd.read_csv(zoho_file)
    else:
        zoho_df = pd.read_excel(zoho_file)
        
    zoho_records: List[ZohoRecord] = []
    for _, row in zoho_df.iterrows():
        cust_name = str(row['Customer']).strip() if pd.notna(row.get('Customer')) else None
        inv_num = str(row['Invoice Number']).strip() if pd.notna(row.get('Invoice Number')) else None
        
        # Rule 3.2: Fallback context mapping lookup via invoice parsing engine cache
        if not cust_name and inv_num in invoice_cache:
            cust_name = invoice_cache[inv_num]
            
        zoho_records.append(ZohoRecord(
            customer_name=cust_name,
            gross_amount=clean_numeric_value(row.get('Gross Amount', 0.0)),
            merchant_fee=clean_numeric_value(row.get('Merchant Fee', 0.0)),
            invoice_number=inv_num
        ))

    # -----------------------------------------------------------------
    # STEP D: PROCESSING EXECUTION & VALIDATION MATRIX
    # -----------------------------------------------------------------
    all_journal_lines = []
    validation_errors = []

    for boa_rec in boa_records:
        matched_zoho = zoho_records  
        
        # Rule 3.3 Balance Invariant Validation
        total_gross = sum(z.gross_amount for z in matched_zoho)
        total_fees = sum(z.merchant_fee for z in matched_zoho)
        calculated_net = total_gross - total_fees
        
        if abs(calculated_net - boa_rec.net_amount) > 0.01:
            validation_errors.append(
                f"🚨 **Mathematical Balance Discrepancy!** Bank Net: ${boa_rec.net_amount:.2f} | "
                f"Calculated Target: ${calculated_net:.2f} (Gross Payments: ${total_gross:.2f}, Zoho Fees: ${total_fees:.2f})"
            )
            continue

        offset_acct = OFFSET_ACCOUNT_ROUTING.get(boa_rec.source_account)
        if not offset_acct:
            validation_errors.append(f"❌ Unmapped Bank of America Routing Target Base Account: {boa_rec.source_account}")
            continue

        processed_accounts = []
        
        # 1. Generate Credit Lines
        for z_rec in matched_zoho:
            if not z_rec.customer_name:
                validation_errors.append(f"❌ Missing Customer Profile Reference for Invoice Tracker ID: {z_rec.invoice_number}")
                continue
                
            name_key = z_rec.customer_name.lower()
            if name_key not in master_lookup:
                validation_errors.append(f"❌ Unregistered Ledger Entity: '{z_rec.customer_name}' absent from database lookup.")
                continue
                
            master_item = master_lookup[name_key]
            processed_accounts.append(master_item)
            
            cash_code, _ = CASH_CODE_MAPPING.get(master_item.payment_term, CASH_CODE_MAPPING['fallback'])
            prefix = "MPP " if cash_code == "AR002" else ""
            desc = f"{prefix}{master_item.account_number} {master_item.account_name}_{boa_rec.description}"
            
            all_journal_lines.append({
                "Date": boa_rec.date, "Voucher": "", "Account name": master_item.account_name,
                "Company": "bwa", "Account type": "Customer", "Account": master_item.account_number,
                "Posting Profile": "AutoPost", "Cash code": cash_code, "Description": desc,
                "Debit": "", "Credit": z_rec.gross_amount, "Item sales tax group": "", "Sales tax code": "",
                "Offset company": "bwa", "Bank Account Type": "Bank", "Offset account": offset_acct,
                "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                "Item sales tax group2": "", "Sales group": "AVATAX", "Withholding tax group": "",
                "Release date": "", "Reversing entry": "No", "Reversing date": ""
            })

        # 2. Generate Grouped Debit Fee Line (Rule 3.3 Multiple Customer Payments)
        if total_fees > 0 and len(processed_accounts) > 0:
            if len(processed_accounts) == 1:
                acc = processed_accounts[0]
                fee_desc = f"Zoho Merchant Fee {acc.account_number} {acc.account_name}_{boa_rec.description}"
            else:
                account_strings = ", ".join(
