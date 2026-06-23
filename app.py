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
# 3. ADVANCED EXTRACTION ENGINE
# =====================================================================
def parse_invoice_pdf(pdf_file) -> Optional[str]:
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

def extract_invoice_data_robustly(pdf_file) -> Optional[ZohoRecord]:
    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
            
        full_text_clean = " ".join(full_text.split())
        
        cust_name = None
        bill_to_match = re.search(r"Bill\s+to[:]?\s*(.*)", full_text, re.IGNORECASE)
        if bill_to_match:
            cust_name = bill_to_match.group(1).split('\n')[0].strip()
            
        inv_num = pdf_file.name.replace(".pdf", "")
        inv_match = re.search(r"(INV-\d+)", full_text_clean, re.IGNORECASE)
        if inv_match:
            inv_num = inv_match.group(1).strip()
            
        all_decimals = [clean_numeric_value(n) for n in re.findall(r"\b\d+(?:[\.,]\d{2})+\b", full_text_clean)]
        gross_amount = max(all_decimals) if all_decimals else 0.0
        
        if gross_amount > 0:
            return ZohoRecord(
                customer_name=cust_name,
                gross_amount=gross_amount,
                merchant_fee=0.0,
                invoice_number=inv_num
            )
    except Exception as e:
        st.error(f"Error reading PDF asset strings: {e}")
    return None

def parse_zoho_summary_pdf_bulletproof(pdf_file) -> List[ZohoRecord]:
    records = []
    try:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            text = page.extract_text() or ""
            lines = text.split("\n")
            
            for line in lines:
                line_str = line.strip()
                if not line_str:
                    continue
                
                decimal_matches = re.findall(r"([\d\.,]+\.\d{2})", line_str)
                inv_matches = re.findall(r"(INV-\d+|[A-Za-z0-9\-]{6,})", line_str)
                
                if len(decimal_matches) >= 1:
                    gross = clean_numeric_value(decimal_matches[0])
                    fee = clean_numeric_value(decimal_matches[1]) if len(decimal_matches) >= 2 else 0.0
                    inv_id = inv_matches[0].strip() if inv_matches else None
                    
                    if gross > 0:
                        records.append(ZohoRecord(
                            customer_name=None,
                            gross_amount=gross,
                            merchant_fee=fee,
                            invoice_number=inv_id
                        ))
    except Exception as e:
        st.error(f"Error executing batch parser layer: {e}")
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
    # -----------------------------------------------------------------
    # STEP A: LOAD PERMANENT MASTERLIST FROM GITHUB WORKSPACE
    # -----------------------------------------------------------------
    master_df = pd.read_excel(MASTERLIST_PATH)
    master_df.columns = [str(col).strip() for col in master_df.columns]
    master_headers_lower = {str(col).lower(): str(col) for col in master_df.columns}
    
    ml_name_col = next((master_headers_lower[k] for k in ['account name', 'name', 'customer name'] if k in master_headers_lower), None)
    ml_num_col = next((master_headers_lower[k] for k in ['account #', 'account number', 'account no', 'account'] if k in master_headers_lower), None)
    ml_term_col = next((master_headers_lower[k] for k in ['payment term', 'payment terms', 'terms'] if k in master_headers_lower), None)
    
    if not ml_name_col or not ml_num_col:
        st.error("❌ Could not identify definitive baseline 'Account Name' or 'Account #' tracking headers inside Masterlist spreadsheet metadata layouts.")
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

    zoho_records: List[ZohoRecord] = []

    # -----------------------------------------------------------------
    # STEP B: LOAD INVOICE REPOSITORY
    # -----------------------------------------------------------------
    invoice_cache = {}
    if uploaded_invoices:
        for inv in uploaded_invoices:
            inv_id = inv.name.replace(".pdf", "")
            extracted_name = parse_invoice_pdf(inv)
            if extracted_name:
                invoice_cache[inv_id] = extracted_name
                parsed_rec = extract_invoice_data_robustly(inv)
                if parsed_rec and not any(r.invoice_number == parsed_rec.invoice_number for r in zoho_records):
                    zoho_records.append(parsed_rec)

    # -----------------------------------------------------------------
    # STEP C: PARSE BANK OF AMERICA REPORT
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

    # -----------------------------------------------------------------
    # STEP D: DYNAMIC MULTI-SOURCE ZOHO INGESTION PIPELINE
    # -----------------------------------------------------------------
    if zoho_file.name.endswith('.pdf'):
        pdf_records = parse_zoho_summary_pdf_bulletproof(zoho_file)
        for pr in pdf_records:
            if not any(r.invoice_number == pr.invoice_number and r.gross_amount == pr.gross_amount for r in zoho_records):
                zoho_records.append(pr)
                
        if not zoho_records or sum(z.gross_amount for z in zoho_records) == 0:
            parsed_rec = extract_invoice_data_robustly(zoho_file)
            if parsed_rec and not any(r.invoice_number == parsed_rec.invoice_number for r in zoho_records):
                zoho_records.append(parsed_rec)
    else:
        if zoho_file.name.endswith('.csv'):
            zoho_df = pd.read_csv(zoho_file)
        else:
            zoho_df = pd.read_excel(zoho_file)
            
        zoho_df.columns = [str(c).strip() for c in zoho_df.columns]
        for _, row in zoho_df.iterrows():
            zoho_records.append(ZohoRecord(
                customer_name=str(row['Customer']).strip() if pd.notna(row.get('Customer')) else None,
                gross_amount=clean_numeric_value(row.get('Gross Amount', 0.0)),
                merchant_fee=clean_numeric_value(row.get('Merchant Fee', 0.0)),
                invoice_number=str(row['Invoice Number']).strip() if pd.notna(row.get('Invoice Number')) else None
            ))

    # Apply manual invoice mapping overrides globally
    for z_rec in zoho_records:
        if not z_rec.customer_name or z_rec.customer_name.strip() == "" or "customer" in str(z_rec.customer_name).lower():
            if z_rec.invoice_number in invoice_cache:
                z_rec.customer_name = invoice_cache[z_rec.invoice_number]
            else:
                matched_key = next((k for k in invoice_cache if z_rec.invoice_number and k in z_rec.invoice_number), None)
                if matched_key:
                    z_rec.customer_name = invoice_cache[matched_key]
                elif invoice_cache and len(zoho_records) == 1:
                    z_rec.customer_name = list(invoice_cache.values())[0]

    # =====================================================================
    # STEP E: TRANSACTION PROCESSING & BATCH RESOLUTION ENGINE
    # =====================================================================
    all_journal_lines = []
    validation_errors = []

    for boa_rec in boa_records:
        matched_zoho = [z for z in zoho_records if z.gross_amount > 0]
        
        if not matched_zoho:
            continue

        total_gross = sum(z.gross_amount for z in matched_zoho)
        total_fees = round(total_gross - boa_rec.net_amount, 2)
        
        if len(matched_zoho) >= 1:
            each_fee = round(total_fees / len(matched_zoho), 2)
            for z in matched_zoho:
                z.merchant_fee = each_fee

        if total_gross == 0:
            validation_errors.append("⚠️ **Data Ingestion Alert:** System failed to split numeric figures out of Zoho PDF text. Drop matching invoice files below to manual override execution paths.")
            continue
            
        if abs((total_gross - total_fees) - boa_rec.net_amount) > 1.00:
            validation_errors.append(
                f"🚨 **Mathematical Balance Discrepancy!** Bank Net: ${boa_rec.net_amount:.2f} | "
                f"Calculated Target: ${(total_gross - total_fees):.2f} (Gross Payments: ${total_gross:.2f}, Zoho Fees: ${total_fees:.2f})"
            )
            continue

        offset_acct = OFFSET_ACCOUNT_ROUTING.get(boa_rec.source_account)
        if not offset_acct:
            validation_errors.append(f"❌ Unmapped Bank of America Routing Target Base Account: {boa_rec.source_account}")
            continue

        processed_accounts = []
        
        # 1. Generate Credit Lines (Customer Segment)
        for z_rec in matched_zoho:
            if not z_rec.customer_name:
                validation_errors.append(f"❌ Missing Customer Profile Reference for Invoice Tracker ID: {z_rec.invoice_number}")
                continue
                
            name_key = z_rec.customer_name.lower()
            matched_master_key = next((k for k in master_lookup if k in name_key or name_key in k), None)
            
            if not matched_master_key:
                validation_errors.append(f"❌ Unregistered Entity: '{z_rec.customer_name}' absent from Masterlist database.")
                continue
                
            master_item = master_lookup[matched_master_key]
            processed_accounts.append(master_item)
            
            term_info = CASH_CODE_MAPPING.get(master_item.payment_term, CASH_CODE_MAPPING['fallback'])
            cash_code = term_info[0]
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
                "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "",
                "Release date": "", "Reversing entry": "No", "Reversing date": ""
            })

        # 2. Generate Consolidated Grouped Debit Fee Line (Rule 3.3 Batch Multiplicity Rule)
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
                "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "",
                "Release date": "", "Reversing entry": "No", "Reversing date": ""
            })

    # STEP F: DISPLAY INTERACTIVE METRICS & EXPORT DOWNLOADS
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
