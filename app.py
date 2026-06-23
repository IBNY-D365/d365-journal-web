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
    amount_target = next((c for c in ['net amount', 'amount', 'net_amount'] if c in boa_df.columns), None)
    account_target = next((c for c in ['source account', 'account', 'account number', 'account_number'] if c in boa_df.columns), None)
            
    if desc_target is None:
        st.error("❌ Could not find a transaction 'Description' column variant in your Bank of America report. Please check file headers.")
        st.stop()

    boa_records: List[BOARecord] = []
    for _, row in boa_df.iterrows():
        row_description = str(row.get(desc_target, ''))
        
        if "ZOHO" in row_description.upper():  # Rule 3.1 Filtering Logic
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

    # -----------------------------------------------------------------
    # STEP D: DYNAMIC PARSING FOR ZOHO SUMMARY (PDF, EXCEL & CSV SUPPORT)
    # -----------------------------------------------------------------
    zoho_records: List[ZohoRecord] = []
    
    if zoho_file.name.endswith('.pdf'):
        zoho_records = parse_zoho_pdf(zoho_file)
    else:
        if zoho_file.name.endswith('.csv'):
            zoho_df = pd.read_csv(zoho_file)
        else:
            zoho_df = pd.read_excel(zoho_file)
            
        zoho_df.columns = [str(c).strip() for c in zoho_df.columns]
        
        for _, row in zoho_df.iterrows():
            cust_name = str(row['Customer']).strip() if pd.notna(row.get('Customer')) else None
            inv_num = str(row['Invoice Number']).strip() if pd.notna(row.get('Invoice Number')) else None
            
            if not cust_name and inv_num in invoice_cache:
                cust_name = invoice_cache[inv_num]
                
            zoho_records.append(ZohoRecord(
                customer_name=cust_name,
                gross_amount=clean_numeric_value(row.get('Gross Amount', 0.0)),
                merchant_fee=clean_numeric_value(row.get('Merchant Fee', 0.0)),
                invoice_number=inv_num
            ))

    # -----------------------------------------------------------------
    # STEP E: PROCESSING EXECUTION & VALIDATION MATRIX
    # -----------------------------------------------------------------
    all_journal_lines = []
    validation_errors = []

    for boa_rec in boa_records:
        matched_zoho = zoho_records  
        
        total_gross = sum(z.gross_amount for z in matched_zoho)
        total_fees = sum(z.merchant_fee for z in matched_zoho)
        calculated_net = total_gross - total_fees
        
        # If the text parsing pulled a row with zeroes, bypass verification and prompt for file checking
        if total_gross == 0 and total_fees == 0:
            validation_errors.append("⚠️ **Data Ingestion Alert:** No financial entries found in Zoho PDF. Please check the format or upload customer invoices to verify names.")
            continue
            
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
                validation_errors.append(f"❌ Unregistered Entity: '{z_rec.customer_name}' absent from Masterlist database.")
                continue
                
            master_item = master_lookup[name_key]
            processed_accounts.append(master_item)
            
            term_info = CASH_CODE_MAPPING.get(master_item.payment_term, CASH_CODE_MAPPING['fallback'])
            cash_code = term_info[0] # Fixed Tuple index access mapping pattern safely
            prefix = "MPP " if cash_code == "AR002" else ""
            
            current_boa_description = str(boa_rec.description)
            desc = f"{prefix}{master_item.account_number} {master_item.account_name}_{current_boa_description}"
            
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
            current_boa_description = str(boa_rec.description)
            if len(processed_accounts) == 1:
                acc = processed_accounts[0]
                fee_desc = f"Zoho Merchant Fee {acc.account_number} {acc.account_name}_{current_boa_description}"
            else:
                account_strings = ", ".join([f"{a.account_number} {a.account_name}" for a in processed_accounts])
                fee_desc = f"Zoho Merchant Fee {account_strings}_{current_boa_description}"

            all_journal_lines.append({
                "Date": boa_rec.date, "Voucher": "", "Account name": "Outside Service (Finance)",
                "Company": "bwa", "Account type": "Ledger", "Account": "43170111-U26C05001-B735350-UOA003",
                "Posting Profile": "", "Cash code": "OSF005", "Description": fee_desc,
                "Debit": total_fees, "Credit": "", "Item sales tax group": "", "Sales tax code": "",
                "Offset company": "bwa", "Bank Account Type": "Bank", "Offset account": offset_acct,
                "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                "Item sales tax group2": "", "Sales group": "AVATAX", "Withholding tax group": "",
                "Release date": "", "Reversing entry": "No", "Reversing date": ""
            })

    # -----------------------------------------------------------------
    # STEP F: DISPLAY INTERACTIVE METRICS & EXPORT DOWNLOADS
    # -----------------------------------------------------------------
    if validation_errors:
        st.error("### Pipeline Validation Discrepancies Checked")
        for error in validation_errors:
            st.markdown(error)

    if all_journal_lines:
        st.success(f"### Transformed {len(all_journal_lines)} Journal Lines Successfully!")
        output_df = pd.DataFrame(all_journal_lines, columns=D365_TEMPLATE_COLUMNS)
        st.dataframe(output_df)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name="Journal Lines")
        
        st.download_button(
            label="📥 Download Generated D365 Journal Import Sheet",
            data=buffer.getvalue(),
            file_name="D365_General_Journal_Import.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
