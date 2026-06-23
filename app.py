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
# 1. HARDCODED CONFIGURATIONS & MAPPINGS
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
    fallback_personal_name: Optional[str] = None

class AccountMasterItem(BaseModel):
    account_number: str
    account_name: str
    payment_term: str
    cs_ps_ticket: Optional[str] = None

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

def normalize_entity_name(name: str) -> str:
    """Strips suffixes like LLC, INC, punctuation, and normalizes spacing for Masterlist cross-referencing."""
    if not name or pd.isna(name):
        return ""
    name = str(name).lower()
    name = re.sub(r'[.,\-\'"&]', '', name)
    name = re.sub(r'\b(llc|inc|ltd|corp|co|pllc|pc)\b', '', name)
    return " ".join(name.split())

# =====================================================================
# 3. ADVANCED EXTRACTION ENGINE WITH MASTERLIST NER
# =====================================================================
def extract_invoice_metadata_intelligent(pdf_file, known_entities: List[str]) -> Dict[str, Any]:
    """Scans invoice text against the exact Masterlist database to guarantee correct names."""
    result = {"customer_name": None, "invoice_number": None, "gross_amount": 0.0, "fallback_personal_name": None}
    try:
        pdf_file.seek(0)
        reader = PdfReader(pdf_file)
        full_text = " ".join(page.extract_text() or "" for page in reader.pages)
        full_text_clean = " ".join(full_text.split())
        
        # Search space formatted identically to the masterlist keys
        search_space = normalize_entity_name(full_text_clean)
        
        # Extract Invoice Number
        inv_num = pdf_file.name.replace(".pdf", "")
        inv_match = re.search(r"(INV-\d+)", full_text_clean, re.IGNORECASE)
        result["invoice_number"] = inv_match.group(1).upper() if inv_match else inv_num
        
        # Extract Gross Amount
        all_decimals = [clean_numeric_value(n) for n in re.findall(r"\b\d+(?:[\.,]\d{2})+\b", full_text_clean)]
        result["gross_amount"] = max(all_decimals) if all_decimals else 0.0
        
        # NER Customer Matching: Find the masterlist name inside the invoice text
        valid_entities = [e for e in known_entities if len(e) > 3]
        valid_entities.sort(key=len, reverse=True) # Sort so "Underground Gym" hits before "Gym"
        
        for entity in valid_entities:
            # Exact whole-word matching inside the cleaned document text
            if f" {entity} " in f" {search_space} ":
                result["customer_name"] = entity
                break
                
        # Last resort fallback if masterlist name is completely missing from document
        if not result["customer_name"]:
            bill_to_match = re.search(r"Bill\s+To[:\s]*([A-Za-z0-9\s\.\,\-\&]+?)(?:\s+Ship|\s+Date|\s+INV-|$)", full_text_clean, re.IGNORECASE)
            if bill_to_match:
                result["fallback_personal_name"] = bill_to_match.group(1).strip()
                
    except Exception as e:
        st.error(f"Error executing intelligent metadata capture on {pdf_file.name}: {e}")
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
    ml_num_col = next((master_headers_lower[k] for k in ['account', 'account #', 'account number', 'account no'] if k in master_headers_lower), None)
    ml_term_col = next((master_headers_lower[k] for k in ['payment term', 'payment terms', 'terms'] if k in master_headers_lower), None)
    ml_ticket_col = next((master_headers_lower[k] for k in ['cs/ps ticket', 'ticket'] if k in master_headers_lower), None)
    
    if not ml_name_col or not ml_num_col:
        st.error(f"❌ Could not identify definitive baseline 'Account Name' or 'Account' tracking headers.")
        st.stop()
        
    master_lookup: Dict[str, AccountMasterItem] = {}
    ticket_lookup: Dict[str, AccountMasterItem] = {}
    
    for _, row in master_df.iterrows():
        name_val = str(row[ml_name_col]).strip()
        name_key = normalize_entity_name(name_val)
        
        num_val = str(row[ml_num_col]).strip()
        term_val = str(row.get(ml_term_col, 'due-on-receipt')).strip().lower() if ml_term_col else 'due-on-receipt'
        
        ticket_val = str(row[ml_ticket_col]).strip() if ml_ticket_col and pd.notna(row[ml_ticket_col]) else ""
        
        item = AccountMasterItem(
            account_number=num_val,
            account_name=name_val,
            payment_term=term_val,
            cs_ps_ticket=ticket_val
        )
        
        if name_key:
            master_lookup[name_key] = item
            
        if ticket_val:
            ticket_key = normalize_entity_name(ticket_val)
            if ticket_key:
                ticket_lookup[ticket_key] = item

    # Compile the known entities from masterlist to cross-reference against the PDFs
    known_entity_keys = list(master_lookup.keys()) + list(ticket_lookup.keys())

    # STEP B: PARSE UPLOADED SUPPORTING INVOICES
    invoice_sources_list = []
    if uploaded_invoices:
        for inv in uploaded_invoices:
            meta = extract_invoice_metadata_intelligent(inv, known_entity_keys)
            if meta["gross_amount"] > 0:
                invoice_sources_list.append(ZohoRecord(
                    customer_name=meta["customer_name"],
                    gross_amount=meta["gross_amount"],
                    merchant_fee=0.0,
                    invoice_number=meta["invoice_number"],
                    fallback_personal_name=meta["fallback_personal_name"]
                ))

    # STEP C: PARSE BANK OF AMERICA REPORT
    if boa_file.name.endswith('.csv'):
        raw_bytes = boa_file.read()
        lines = raw_bytes.decode('utf-8').splitlines()
        boa_file.seek(0)
        
        skip_count = 0
        for idx, line in enumerate(lines):
            if "date" in line.lower() and "description" in line.lower():
                skip_count = idx
                break
        boa_df = pd.read_csv(boa_file, skiprows=skip_count)
    else:
        boa_df = pd.read_excel(boa_file)
    
    boa_df.columns = [str(col).strip().lower() for col in boa_df.columns]
    desc_target = next((c for c in ['description', 'transaction description', 'payee', 'memo'] if c in boa_df.columns), None)
    date_target = next((c for c in ['posting date', 'date', 'transaction date'] if c in boa_df.columns), None)
    amount_target = next((c for c in ['net amount', 'amount', 'net_amount'] if c in boa_df.columns), None)
    account_target = next((c for c in ['source account', 'account', 'account number', 'account_number'] if c in boa_df.columns), None)

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
                source_account=str(row.get(account_target, '')).strip() if account_target else "3371"
            ))

    # STEP D: PARSE PRIMARY ZOHO SUMMARY SHEET & RESOLVE CONFLICTS
    raw_zoho_pool: List[ZohoRecord] = []
    if zoho_file.name.endswith('.pdf'):
        raw_zoho_pool = parse_zoho_summary_pdf_bulletproof(zoho_file)
    else:
        if zoho_file.name.endswith('.csv'):
            zoho_df = pd.read_csv(zoho_file)
        else:
            zoho_df = pd.read_excel(zoho_file)
        zoho_df.columns = [str(c).strip() for c in zoho_df.columns]
        for _, row in zoho_df.iterrows():
            raw_zoho_pool.append(ZohoRecord(
                customer_name=str(row['Customer']).strip() if pd.notna(row.get('Customer')) else None,
                gross_amount=clean_numeric_value(row.get('Gross Amount', 0.0)),
                merchant_fee=clean_numeric_value(row.get('Merchant Fee', 0.0)),
                invoice_number=str(row['Invoice Number']).strip() if pd.notna(row.get('Invoice Number')) else None
            ))

    # --- STRICT OVERRIDE LOGIC ---
    if uploaded_invoices and len(invoice_sources_list) > 0:
        # If user uploaded supporting invoices, bypass the noisy Zoho PDF and strictly use the clean invoices.
        zoho_records = invoice_sources_list
    else:
        # Deduplicate raw_zoho_pool if no invoices were provided.
        zoho_records = []
        seen_invoices = set()
        for r in raw_zoho_pool:
            if r.invoice_number and r.invoice_number not in seen_invoices:
                zoho_records.append(r)
                seen_invoices.add(r.invoice_number)

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
            validation_errors.append("⚠️ **Data Ingestion Alert:** System failed to split numeric figures out of Zoho PDF text.")
            continue
            
        if abs((total_gross - total_fees) - boa_rec.net_amount) > 1.00:
            validation_errors.append(
                f"🚨 **Mathematical Balance Discrepancy!** Bank Net: ${boa_rec.net_amount:.2f} | Calculated Target: ${(total_gross - total_fees):.2f}"
            )
            continue

        offset_acct = OFFSET_ACCOUNT_ROUTING.get(boa_rec.source_account, "B1000002")
        processed_accounts = []
        
        # 1. Generate Credit Lines (Customer Segment)
        for z_rec in matched_zoho:
            current_boa_description = str(boa_rec.description)
            master_item = None
            
            # Lookup name against normalized master dictionary
            if z_rec.customer_name:
                query_key = normalize_entity_name(z_rec.customer_name)
                if query_key in master_lookup:
                    master_item = master_lookup[query_key]
                elif query_key in ticket_lookup:
                    master_item = ticket_lookup[query_key]
                else:
                    matched_key = next((k for k in master_lookup if k in query_key or query_key in k), None)
                    if matched_key:
                        master_item = master_lookup[matched_key]
            
            if not master_item and z_rec.fallback_personal_name:
                fallback_key = normalize_entity_name(z_rec.fallback_personal_name)
                matched_ticket_key = next((k for k in ticket_lookup if fallback_key in k or k in fallback_key), None)
                if matched_ticket_key:
                    master_item = ticket_lookup[matched_ticket_key]

            if not master_item:
                account_num = "21040102-B1000002"
                account_type = "Ledger"
                account_name = "Temporary Receipt"
                cash_code = "AR012"
                
                display_label = z_rec.fallback_personal_name if z_rec.fallback_personal_name else (z_rec.customer_name if z_rec.customer_name else "Unknown")
                desc = f"{display_label} (UNRECORDED ENTITY)_{current_boa_description}"
            else:
                processed_accounts.append(master_item)
                term_info = CASH_CODE_MAPPING.get(master_item.payment_term, CASH_CODE_MAPPING['fallback'])
                cash_code = term_info[0]
                prefix = "MPP " if cash_code == "AR002" else ""
                
                account_num = master_item.account_number
                account_type = "Customer"
                account_name = master_item.account_name
                desc = f"{prefix}{account_num} {account_name}_{current_boa_description}"
            
            all_journal_lines.append({
                "Date": boa_rec.date, "Voucher": "", "Account name": account_name,
                "Company": "bwa", "Account type": account_type, "Account": account_num,
                "Posting Profile": "AutoPost" if account_type == "Customer" else "", "Cash code": cash_code, "Description": desc,
                "Debit": "", "Credit": z_rec.gross_amount, "Item sales tax group": "", "Sales tax code": "",
                "Offset company": "bwa", "Bank Account Type": "Bank", "Offset account": offset_acct,
                "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "",
                "Release date": "", "Reversing entry": "No", "Reversing date": ""
            })

        # 2. Generate Consolidated Grouped Debit Fee Line
        if total_fees > 0:
            current_boa_description = str(boa_rec.description)
            if len(processed_accounts) == 1:
                acc = processed_accounts[0]
                fee_desc = f"Zoho Merchant Fee {acc.account_number} {acc.account_name}_{current_boa_description}"
            elif len(processed_accounts) > 1:
                account_strings = ", ".join([f"{a.account_number} {a.account_name}" for a in processed_accounts])
                fee_desc = f"Zoho Merchant Fee {account_strings}_{current_boa_description}"
            else:
                fee_desc = f"Zoho Merchant Fee (Unresolved Suspense Pool Batch)_{current_boa_description}"

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
        
        buffer = io.
