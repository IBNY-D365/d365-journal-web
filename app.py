import streamlit as st
import pandas as pd
import re
import os
from pypdf import PdfReader

# Set up page layout
st.set_page_config(page_title="D365 Zoho & BOA Journal Generator", layout="wide")
st.title("📊 D365 Zoho, Invoice & BOA Journal Generator")
st.write("Upload your daily processing files below to build your flawless 25-column D365 upload package.")

# 1. SIDEBAR FIXED CONFIGURATION
st.sidebar.header("⚙️ Fixed D365 Settings")
company_id = st.sidebar.text_input("Company", value="bwa")
offset_account = st.sidebar.text_input("Offset Account", value="B1000002")
debit_ledger_acct = st.sidebar.text_input("Debit Line Account", value="43170111-U26C05001-B735350-UOA003")

# 2. THREE-FILE UPLOADERS
col1, col2, col3 = st.columns(3)
with col1:
    zoho_file = st.file_uploader("1. Upload Zoho Payments File", type=["csv", "xlsx"])
with col2:
    invoice_file = st.file_uploader("2. Upload Invoice File (PDF, Excel, or CSV)", type=["pdf", "csv", "xlsx", "txt"])
with col3:
    boa_file = st.file_uploader("3. Upload Bank of America Statement", type=["csv", "xlsx"])

# Target master filename inside your GitHub repository
def locate_master_file():
    target_base = "customer master account file"
    for f in os.listdir('.'):
        if target_base in f.lower():
            return f
    return None

MASTER_FILE_NAME = locate_master_file()

def extract_text_from_pdf(uploaded_pdf):
    try:
        reader = PdfReader(uploaded_pdf)
        full_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
        return full_text
    except Exception:
        return ""

# Comprehensive standardization helper
def super_clean_string(text):
    if pd.isna(text) or text is None:
        return ""
    # Strip common punctuation and convert to lowercase
    cleaned = re.sub(r'[.,\-_()\[\]]', ' ', str(text).lower())
    # Remove extra spaces
    return " ".join(cleaned.split())

# 3. AUTOMATED PARSING AND BALANCING PIPELINE
if zoho_file and invoice_file and boa_file:
    st.subheader("4. Review & Generate")
    
    try:
        if not MASTER_FILE_NAME:
            available_files = os.listdir('.')
            st.error(f"❌ Error: Could not locate your Master Excel sheet. Current files found in repo: {available_files}")
        else:
            # A. Load Master Reference Excel
            cust_df = pd.read_excel(MASTER_FILE_NAME, engine='openpyxl')
            cust_df.columns = [str(col).strip() for col in cust_df.columns]
            
            # Explicit master column targets
            name_col = "Account Name" if "Account Name" in cust_df.columns else cust_df.columns[2]
            acct_col = "Account" if "Account" in cust_df.columns else cust_df.columns[1]
            term_col = "Terms" if "Terms" in cust_df.columns else None
            
            # Pre-calculate normalized search targets cleanly
            cust_df['Account Name Clean'] = cust_df[name_col].apply(super_clean_string)
            
            # B. Extract Terms Information from Invoice File
            invoice_terms = "receipt" # Default fallback
            pdf_text_raw = ""
            if invoice_file.name.endswith('.pdf'):
                pdf_text_raw = extract_text_from_pdf(invoice_file)
                pdf_text_clean = super_clean_string(pdf_text_raw)
                if 'monthly' in pdf_text_clean or 'mpp' in pdf_text_clean:
                    invoice_terms = "monthly"
            else:
                if invoice_file.name.endswith('.csv'):
                    inv_df = pd.read_csv(invoice_file)
                else:
                    inv_df = pd.read_excel(invoice_file, engine='openpyxl')
                inv_df.columns = [str(col).strip() for col in inv_df.columns]
                inv_term_col = next((c for c in inv_df.columns if 'term' in c.lower()), None)
                if inv_term_col and not inv_df.empty:
                    invoice_terms = str(inv_df.iloc[0][inv_term_col]).lower()

            # C. Load Daily Zoho File
            if zoho_file.name.endswith('.csv'):
                zoho_df = pd.read_csv(zoho_file)
            else:
                zoho_df = pd.read_excel(zoho_file, engine='openpyxl')
            zoho_df.columns = [str(col).strip() for col in zoho_df.columns]
            
            # Lock explicit positions based on your target file layout
            zoho_gross_col = "Amount" if "Amount" in zoho_df.columns else zoho_df.columns[3]
            zoho_fee_col = "Fee" if "Fee" in zoho_df.columns else zoho_df.columns[4]
            zoho_cust_col = "CustomerName" if "CustomerName" in zoho_df.columns else zoho_df.columns[9]
            zoho_desc_col = "Description" if "Description" in zoho_df.columns else zoho_df.columns[10]
            zoho_type_col = "TransactionType" if "TransactionType" in zoho_df.columns else None

            # D. Load Bank of America File
            if boa_file.name.endswith('.csv'):
                boa_lines = boa_file.getvalue().decode('utf-8', errors='ignore').splitlines()
                skip_count = 0
                for line in boa_lines:
                    if 'Date,Description,Amount' in line or 'date,description,amount' in line.lower():
                        break
                    skip_count += 1
                boa_file.seek(0)
                boa_df = pd.read_csv(boa_file, skiprows=skip_count)
            else:
                boa_df = pd.read_excel(boa_file, engine='openpyxl')
                
            boa_df.columns = [str(col).strip() for col in boa_df.columns]
            
            boa_match_row = boa_df[boa_df['Description'].astype(str).str.contains('ZOHO PAYMENTS', case=False, na=False)]
            boa_date = ""
            boa_reference_desc = "NO BOA MATCH FOUND"
            
            if not boa_match_row.empty:
                boa_date_col = next((c for c in boa_df.columns if 'date' in c.lower() or 'post' in c.lower()), boa_df.columns[0])
                boa_desc_col = next((c for c in boa_df.columns if 'desc' in c.lower() or 'text' in c.lower()), boa_df.columns[1])
                boa_date = str(boa_match_row.iloc[0][boa_date_col]).strip()
                boa_reference_desc = str(boa_match_row.iloc[0][boa_desc_col]).strip()
            
            journal_rows = []
            
            # E. Core Processing Loop
            for idx, row in zoho_df.iterrows():
                # Filter out raw refund lines to avoid empty master mapping lookups
                if zoho_type_col and str(row[zoho_type_col]).strip().lower() == 'refund':
                    continue
                    
                cust_name_raw = str(row[zoho_cust_col]).strip()
                
                # EDITED LOGIC: Fallback handling when Zoho provides an empty CustomerName column
                if not cust_name_raw or cust_name_raw == "nan" or cust_name_raw == "":
                    # Attempt Extraction from the Zoho Description Text (e.g., "InBody New York - INV-...")
                    if zoho_desc_col in zoho_df.columns and str(row[zoho_desc_col]).strip():
                        desc_text = str(row[zoho_desc_col]).split('-')[0].strip()
                        if desc_text and desc_text.lower() != "nan":
                            cust_name_raw = desc_text
                    
                    # Secondary Attempt: If Description extraction failed, extract name from the Invoice PDF text lines
                    if (not cust_name_raw or cust_name_raw == "nan" or cust_name_raw == "") and pdf_text_raw:
                        lines = [line.strip() for line in pdf_text_raw.split('\n') if line.strip()]
                        if lines:
                            cust_name_raw = lines[0] # Fallback to first text element of the invoice layout
                
                # Final safety pass to ensure loop stability if fields are corrupt
                if not cust_name_raw or cust_name_raw == "nan" or cust_name_raw == "":
                    continue
                    
                gross_amt = abs(float(str(row[zoho_gross_col]).replace(',', '')))
                fee_amt = abs(float(str(row[zoho_fee_col]).replace(',', '')))
                
                customer_account_num = "MISSING_ACCT"
                payment_term = "receipt"
                final_account_name = cust_name_raw # Fallback
                
                if 'monthly' in invoice_terms or 'mpp' in invoice_terms:
                    payment_term = "monthly"
                
                # Normalize clean key directly from the uncorrupted Zoho column
                search_key = super_clean_string(cust_name_raw)
                
                if search_key:
                    # Execute exact string mapping
                    match_cust = cust_df[cust_df['Account Name Clean'] == search_key]
                    
                    if match_cust.empty:
                        match_cust = cust_df[cust_df['Account Name Clean'].str.contains(search_key, na=False) | 
                                             (search_key in cust_df['Account Name Clean'].to_string())]
                    
                    if not match_cust.empty:
                        customer_account_num = str(match_cust.iloc[0][acct_col]).strip()
                        # Always override output with the official clean Master Database corporate text string
                        final_account_name = str(match_cust.iloc[0][name_col]).strip()
                        if term_col and term_col in match_cust.columns:
                            term_check = str(match_cust.iloc[0][term_col]).lower()
                            if 'monthly' in term_check or 'mpp' in term_check:
                                payment_term = "monthly"
                
                # Build descriptions per requirements using the official master file name layout
                if payment_term == "monthly":
                    cash_code = "AR002"
                    credit_desc = f"MPP {customer_account_num} {final_account_name}_{boa_reference_desc}"
                else:
                    cash_code = "AR001"
                    credit_desc = f"{customer_account_num} {final_account_name}_{boa_reference_desc}"
                
                # ROW 1: THE D365 CREDIT LINE (CUSTOMER)
                journal_rows.append({
                    "Date": boa_date, "Voucher": "", "Account name": final_account_name, "Company": company_id,
                    "Account type": "Customer", "Account": customer_account_num, "Posting profile": "AutoPost",
                    "Cash code": cash_code, "Description": credit_desc, "Debit": "", "Credit": gross_amt,
                    "Item sales tax group": "", "Sales tax code": "", "Offset company": company_id, "Offset account type": "Bank",
                    "Offset account": offset_account, "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                    "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "", "
