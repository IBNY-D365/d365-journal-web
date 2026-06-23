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

# GitHub Repository File Paths
MASTERLIST_PATH = "Account Masterlist.xlsx"
MASTERLIST_PATH_CSV = "Account Masterlist.csv"
FORM_DB_PATH = "Form Master DB.xlsx"
FORM_DB_PATH_CSV = "Form Master DB.csv"

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
    norm_name: str
    norm_ticket: str

def clean_numeric_value(val: Any) -> float:
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return abs(float(val))
    cleaned_str = str(val).strip().replace('$', '').replace(',', '').replace('-', '')
    try:
        return float(cleaned_str)
    except ValueError:
        return 0.0

def normalize_name(name: str) -> str:
    if not name or pd.isna(name):
        return ""
    n = str(name).lower()
    n = re.sub(r'[,.\-&]', ' ', n)
    n = re.sub(r'\b(inc|llc|corp|ltd|incorporated|company|co|pllc)\b', '', n)
    n = re.sub(r'\s+', '', n)
    return n

def map_form_term_to_cash_code(term_str: str) -> str:
    t = str(term_str).lower().strip()
    if "ap" in t or "due on receipt" in t: return "due-on-receipt"
    if "mpp" in t or "monthly" in t: return "monthly"
    if "financ" in t: return "financing"
    if "leas" in t: return "leasing"
    if "net 1 day" in t: return "net 1 day"
    if "net 10" in t: return "net 10 days"
    if "net 25" in t: return "net 25 days"
    if "net 30" in t: return "net 30 days"
    if "net 40" in t: return "net 40 days"
    if "net 45" in t: return "net 45 days"
    if "net 60" in t: return "net 60 days"
    return "due-on-receipt"

# =====================================================================
# 3. ADVANCED EXTRACTION ENGINE
# =====================================================================
def extract_invoice_metadata_intelligent(pdf_file) -> Dict[str, Any]:
    result = {"customer_name": None, "invoice_number": None, "gross_amount": 0.0, "fallback_personal_name": None}
    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
            
        full_text_clean = " ".join(full_text.split())
        
        inv_num = pdf_file.name.replace(".pdf", "")
        inv_match = re.search(r"(INV-\d+)", full_text_clean, re.IGNORECASE)
        result["invoice_number"] = inv_match.group(1).strip() if inv_match else inv_num
        
        pm_match = re.search(r"Payment\s*Made[^\d\$]*\$?([0-9,]+\.\d{2})", full_text_clean, re.IGNORECASE)
        if pm_match:
            result["gross_amount"] = clean_numeric_value(pm_match.group(1))
        else:
            totals = re.findall(r"Total[^\d\$]*\$?([0-9,]+\.\d{2})", full_text_clean, re.IGNORECASE)
            if totals:
                result["gross_amount"] = clean_numeric_value(totals[-1])
            else:
                all_decimals = [clean_numeric_value(n) for n in re.findall(r"\b\d+(?:,\d{3})*\.\d{2}\b", full_text_clean)]
                result["gross_amount"] = max(all_decimals) if all_decimals else 0.0
        
        bill_to_match = re.search(r"Bill\s+To\s*([A-Za-z0-9\s\.\,\-]+?)(?:\s*\d|\s*Ship\s*To|$)", full_text_clean, re.IGNORECASE)
        if bill_to_match:
            result["fallback_personal_name"] = bill_to_match.group(1).strip()
            
        biz_matches = re.findall(r"InBody\d*\s*-\s*[^-]+?-\s*([A-Za-z0-9\s\.\,\&]+)", full_text_clean, re.IGNORECASE)
        for match in biz_matches:
            candidate = re.sub(r'\d+\.\d{2}.*', '', match).strip()
            if candidate and not any(k in candidate.lower() for k in ["malfunction", "check required", "sku", "labor", "board", "cable", "loaner"]):
                result["customer_name"] = candidate
                return result
                    
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
            
        if "All Transactions" in full_text:
            transactions_text = full_text.split("All Transactions")[-1]
        elif "Export" in full_text:
            transactions_text = full_text.split("Export")[-1]
        else:
            transactions_text = full_text
            
        transactions_text = " ".join(transactions_text.split())
        
        # Surgical split using time chunks
        chunks = re.split(r"\d{1,2}:\d{2}\s*(?:AM|PM)", transactions_text, flags=re.IGNORECASE)
        
        for chunk in chunks[1:]: 
            amounts = re.findall(r"[-]?\$?[0-9,]+\.\d{2}", chunk)
            if len(amounts) >= 3:
                gross = clean_numeric_value(amounts[0])
                fee = clean_numeric_value(amounts[1])
                
                desc_match = re.search(r"^(.*?)(?:-\$|\$|-[ ]?\$|[0-9])", chunk)
                desc = desc_match.group(1).strip() if desc_match else ""
                
                inv_match = re.search(r'(INV-\d+)', desc, re.IGNORECASE)
                inv_id = inv_match.group(1).strip() if inv_match else None
                
                cust_name = re.sub(r'INV-\d+.*', '', desc, flags=re.IGNORECASE).strip() if inv_id else desc.strip()
                cust_name = re.sub(r'\.\.\.$', '', cust_name).strip() 
                if cust_name and len(cust_name) < 2: cust_name = None
                
                if gross > 0 and not any((r.invoice_number == inv_id and inv_id is not None) or (r.gross_amount == gross and r.customer_name == cust_name) for r in records):
                    records.append(ZohoRecord(
                        customer_name=cust_name,
                        gross_amount=gross,
                        merchant_fee=fee,
                        invoice_number=inv_id
                    ))
    except Exception as e:
        st.error(f"Error executing summary parser: {e}")
    return records

# =====================================================================
# 4. STREAMLIT INTERFACE SETUP
# =====================================================================
st.set_page_config(page_title="D365 General Journal Automation", layout="wide")
st.title("D365 General Journal Automation Engine")
st.subheader("Daily Operational Reconciliations Matrix")

target_masterlist = MASTERLIST_PATH if os.path.exists(MASTERLIST_PATH) else (MASTERLIST_PATH_CSV if os.path.exists(MASTERLIST_PATH_CSV) else None)
if not target_masterlist:
    st.error("❌ Core configuration file `Account Masterlist` missing from your GitHub repository root folder.")
    st.stop()

st.sidebar.header("📅 Daily Variable Inputs")
boa_file = st.sidebar.file_uploader("1. Bank of America Report (Excel/CSV)", type=["xlsx", "csv"])
zoho_file = st.sidebar.file_uploader("2. Zoho Transaction Summary (PDF/Excel/CSV)", type=["pdf", "xlsx", "csv"])
uploaded_invoices = st.sidebar.file_uploader("3. Extra Customer Invoices (PDFs) [Optional]", type=["pdf"], accept_multiple_files=True)

if not (boa_file and zoho_file):
    st.info("💡 Staging required: Please drop today's Bank of America report and matching Zoho summary sheet into the sidebar container panel.")
else:
    # -----------------------------------------------------------------
    # STEP A: LOAD MASTERLIST
    # -----------------------------------------------------------------
    if target_masterlist.endswith('.csv'):
        master_df = pd.read_csv(target_masterlist)
    else:
        master_df = pd.read_excel(target_masterlist)
        
    master_df.columns = [str(col).strip() for col in master_df.columns]
    master_headers_lower = {str(col).lower(): str(col) for col in master_df.columns}
    
    ml_name_col = next((master_headers_lower[k] for k in ['account name', 'name', 'customer name'] if k in master_headers_lower), None)
    ml_num_col = next((master_headers_lower[k] for k in ['account #', 'account number', 'account no', 'account'] if k in master_headers_lower), None)
    ml_term_col = next((master_headers_lower[k] for k in ['payment term', 'payment terms', 'terms'] if k in master_headers_lower), None)
    ml_ticket_col = next((master_headers_lower[k] for k in ['cs/ps ticket', 'ticket', 'cs/ps'] if k in master_headers_lower), None)
    
    master_lookup: Dict[str, AccountMasterItem] = {}
    for _, row in master_df.iterrows():
        name_val = str(row[ml_name_col]).strip()
        num_val = str(row[ml_num_col]).strip()
        term_val = str(row.get(ml_term_col, 'due-on-receipt')).strip().lower() if ml_term_col else 'due-on-receipt'
        ticket_val = str(row.get(ml_ticket_col, '')).strip() if ml_ticket_col else ''
        
        master_lookup[name_val] = AccountMasterItem(
            account_number=num_val,
            account_name=name_val,
            payment_term=term_val,
            norm_name=normalize_name(name_val),
            norm_ticket=normalize_name(ticket_val)
        )

    # -----------------------------------------------------------------
    # STEP B: LOAD FORM MASTER DB FROM GITHUB
    # -----------------------------------------------------------------
    form_db_lookup = {}
    target_form_db = FORM_DB_PATH if os.path.exists(FORM_DB_PATH) else (FORM_DB_PATH_CSV if os.path.exists(FORM_DB_PATH_CSV) else None)
    
    if target_form_db:
        if target_form_db.endswith('.csv'):
            fdf = pd.read_csv(target_form_db)
        else:
            fdf = pd.read_excel(target_form_db)
        
        fdf.columns = [str(c).strip().lower() for c in fdf.columns]
        biz_col = next((c for c in fdf.columns if 'business name' in c or 'customer' in c), None)
        acct_col = next((c for c in fdf.columns if 'customer account' in c or 'account number' in c or 'account' in c), None)
        term_col = next((c for c in fdf.columns if 'invoice sent' in c or 'payment term' in c or 'term' in c), None)
        
        if biz_col and term_col:
            for _, row in fdf.iterrows():
                b_name = str(row[biz_col]).strip()
                t_val = str(row[term_col]).strip()
                a_val = str(row[acct_col]).strip() if acct_col and pd.notna(row[acct_col]) else ""
                
                if b_name and b_name.lower() != 'nan':
                    norm_b = normalize_name(b_name)
                    form_db_lookup[norm_b] = {
                        "term": t_val,
                        "account": a_val,
                        "raw_name": b_name
                    }
    else:
        st.warning("⚠️ Form Master DB not found in repository root. Please ensure 'Form Master DB.xlsx' or 'Form Master DB.csv' is committed to GitHub.")

    # -----------------------------------------------------------------
    # STEP C: EXTRACT ALL UPLOADED INVOICES
    # -----------------------------------------------------------------
    invoice_cache = {}
    invoice_sources_list = []
    
    if uploaded_invoices:
        for inv in uploaded_invoices:
            meta = extract_invoice_metadata_intelligent(inv)
            if meta["invoice_number"] or meta["customer_name"]:
                inv_key = meta["invoice_number"] if meta["invoice_number"] else meta["customer_name"]
                invoice_cache[inv_key] = {
                    "resolved_name": meta["customer_name"],
                    "fallback_personal_name": meta["fallback_personal_name"]
                }
                invoice_sources_list.append(ZohoRecord(
                    customer_name=meta["customer_name"],
                    gross_amount=meta["gross_amount"],
                    merchant_fee=0.0,
                    invoice_number=meta["invoice_number"],
                    fallback_personal_name=meta["fallback_personal_name"]
                ))

    # -----------------------------------------------------------------
    # STEP D: PARSE BANK OF AMERICA REPORT
    # -----------------------------------------------------------------
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
        row_net_amount = clean_numeric_value(row.get(amount_target, 0.0))
        
        if "ZOHO PAYMENTS" in row_description.upper() and row_net_amount > 0:
            parsed_date = datetime.today().date()
            if date_target and pd.notna(row[date_target]):
                try:
                    parsed_date = pd.to_datetime(row[date_target]).date()
                except Exception:
                    pass
            boa_records.append(BOARecord(
                date=parsed_date,
                description=row_description,
                net_amount=row_net_amount,
                source_account=str(row.get(account_target, '')).strip() if account_target else "3371"
            ))

    # -----------------------------------------------------------------
    # STEP E: PARSE PRIMARY ZOHO SUMMARY SHEET & DEDUPLICATE
    # -----------------------------------------------------------------
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

    if not raw_zoho_pool or sum(r.gross_amount for r in raw_zoho_pool) == 0:
        raw_zoho_pool = invoice_sources_list

    zoho_deduped_dict = {}
    for r in raw_zoho_pool:
        key = r.invoice_number if r.invoice_number else f"{r.customer_name}_{r.gross_amount}"
        if key not in zoho_deduped_dict:
            zoho_deduped_dict[key] = r
            
    zoho_records = list(zoho_deduped_dict.values())

    for z_rec in zoho_records:
        if z_rec.invoice_number in invoice_cache:
            cache_hit = invoice_cache[z_rec.invoice_number]
            z_rec.customer_name = cache_hit["resolved_name"] if cache_hit["resolved_name"] else z_rec.customer_name
            z_rec.fallback_personal_name = cache_hit["fallback_personal_name"]

    # =====================================================================
    # STEP F: TRANSACTION PROCESSING & ONE-TO-ONE BOA ANCHORING
    # =====================================================================
    all_journal_lines = []
    validation_errors = []

    z_total_gross = sum(z.gross_amount for z in zoho_records)
    z_total_fees = sum(z.merchant_fee for z in zoho_records)
    z_net = round(z_total_gross - z_total_fees, 2)

    if z_total_gross == 0:
        validation_errors.append("⚠️ **Data Ingestion Alert:** System failed to extract gross amounts.")
    else:
        matched_boa = None
        for boa in boa_records:
            if abs(boa.net_amount - z_net) <= 1.00:
                matched_boa = boa
                break

        if not matched_boa:
            validation_errors.append(f"🚨 **Mathematical Balance Discrepancy!** No Bank of America deposit matches the calculated Zoho Net (${z_net:.2f}). Total Gross: ${z_total_gross:.2f}, Total Fees: ${z_total_fees:.2f}")
        else:
            offset_acct = OFFSET_ACCOUNT_ROUTING.get(matched_boa.source_account, "B1000002")
            processed_accounts = []
            
            for z_rec in zoho_records:
                current_boa_description = str(matched_boa.description)
                
                norm_biz = normalize_name(z_rec.customer_name)
                norm_per = normalize_name(z_rec.fallback_personal_name)
                
                matched_master_item = None
                form_match = None
                
                # Check Masterlist First
                for item in master_lookup.values():
                    if norm_biz and len(norm_biz) >= 4 and (norm_biz in item.norm_name or item.norm_name in norm_biz or item.norm_name.startswith(norm_biz)):
                        matched_master_item = item
                        break
                    if norm_per and len(norm_per) >= 4 and (norm_per in item.norm_name or item.norm_name in norm_per or item.norm_name.startswith(norm_per)):
                        matched_master_item = item
                        break
                    if item.norm_ticket:
                        if norm_per and len(norm_per) >= 4 and (norm_per in item.norm_ticket or item.norm_ticket in norm_per or item.norm_ticket.startswith(norm_per)):
                            matched_master_item = item
                            break
                        if norm_biz and len(norm_biz) >= 4 and (norm_biz in item.norm_ticket or item.norm_ticket in norm_biz or item.norm_ticket.startswith(norm_biz)):
                            matched_master_item = item
                            break

                # Check Form DB Second for Missing Masterlist Entries
                if form_db_lookup:
                    for q in [norm_biz, norm_per]:
                        if q and len(q) >= 4:
                            for k, v in form_db_lookup.items():
                                if q in k or k in q or k.startswith(q):
                                    form_match = v
                                    break
                        if form_match: break

                # ASSIGNMENT EXECUTION
                if matched_master_item:
                    final_term = matched_master_item.payment_term
                    if form_match and form_match["term"] and str(form_match["term"]).lower() != 'nan':
                        final_term = form_match["term"]
                        
                    term_info = CASH_CODE_MAPPING.get(map_form_term_to_cash_code(final_term), CASH_CODE_MAPPING['fallback'])
                    cash_code = term_info[0]
                    prefix = "MPP " if cash_code == "AR002" else ""
                    
                    account_num = matched_master_item.account_number
                    account_type = "Customer"
                    account_name = matched_master_item.account_name
                    desc = f"{prefix}{account_num} {account_name}_{current_boa_description}"
                    
                    processed_accounts.append(matched_master_item)

                elif form_match and form_match["account"]:
                    final_term = form_match["term"]
                    term_info = CASH_CODE_MAPPING.get(map_form_term_to_cash_code(final_term), CASH_CODE_MAPPING['fallback'])
                    cash_code = term_info[0]
                    prefix = "MPP " if cash_code == "AR002" else ""
                    
                    account_num = form_match["account"]
                    account_type = "Customer"
                    account_name = form_match["raw_name"]
                    desc = f"{prefix}{account_num} {account_name}_{current_boa_description}"
                    
                    processed_accounts.append(AccountMasterItem(account_number=account_num, account_name=account_name, payment_term=final_term, norm_name="", norm_ticket=""))

                else:
                    account_num = "21040102-B1000002"
                    account_type = "Ledger"
                    account_name = "Temporary Receipt"
                    
                    final_term = form_match["term"] if form_match else "due-on-receipt"
                    term_info = CASH_CODE_MAPPING.get(map_form_term_to_cash_code(final_term), CASH_CODE_MAPPING['fallback'])
                    cash_code = term_info[0]
                    
                    display_label = z_rec.customer_name if z_rec.customer_name else (z_rec.fallback_personal_name if z_rec.fallback_personal_name else "Unknown")
                    desc = f"{display_label} (UNRECORDED ENTITY)_{current_boa_description}"

                all_journal_lines.append({
                    "Date": matched_boa.date, "Voucher": "", "Account name": account_name,
                    "Company": "bwa", "Account type": account_type, "Account": account_num,
                    "Posting Profile": "AutoPost" if account_type == "Customer" else "", "Cash code": cash_code, "Description": desc,
                    "Debit": "", "Credit": z_rec.gross_amount, "Item sales tax group": "", "Sales tax code": "",
                    "Offset company": "bwa", "Bank Account Type": "Bank", "Offset account": offset_acct,
                    "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                    "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "",
                    "Release date": "", "Reversing entry": "No", "Reversing date": ""
                })

            if z_total_fees > 0:
                if len(processed_accounts) == 1:
                    acc = processed_accounts[0]
                    fee_desc = f"Zoho Merchant Fee {acc.account_number} {acc.account_name}_{str(matched_boa.description)}"
                elif len(processed_accounts) > 1:
                    account_strings = ", ".join([f"{a.account_number} {a.account_name}" for a in processed_accounts])
                    fee_desc = f"Zoho Merchant Fee {account_strings}_{str(matched_boa.description)}"
                else:
                    fee_desc = f"Zoho Merchant Fee (Unresolved Suspense Pool Batch)_{str(matched_boa.description)}"

                all_journal_lines.append({
                    "Date": matched_boa.date, "Voucher": "", "Account name": "Outside Service (Finance)",
                    "Company": "bwa", "Account type": "Ledger", "Account": "43170111-U26C05001-B735350-UOA003",
                    "Posting Profile": "", "Cash code": "OSF005", "Description": fee_desc,
                    "Debit": z_total_fees, "Credit": "", "Item sales tax group": "", "Sales tax code": "",
                    "Offset company": "bwa", "Bank Account Type": "Bank", "Offset account": offset_acct,
                    "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                    "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "",
                    "Release date": "", "Reversing entry": "No", "Reversing date": ""
                })

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
