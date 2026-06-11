import streamlit as st
import pandas as pd
import re
import os
from pypdf import PdfReader

# Set up page layout
st.set_set_config = st.set_page_config(page_title="D365 Zoho & BOA Journal Generator", layout="wide")
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

# Standardizes string checks by reducing characters to simple space-separated lowercase tokens
def super_clean_string(text):
    if pd.isna(text) or text is None:
        return ""
    cleaned = re.sub(r'[.,\-_()\[\]]', ' ', str(text).lower())
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
            
            # Detect dynamic columns
            name_col = next((c for c in cust_df.columns if 'name' in c.lower() or 'customer' in c.lower() or 'business' in c.lower()), cust_df.columns[0])
            acct_col = next((c for c in cust_df.columns if 'account' in c.lower() or 'acct' in c.lower() or 'code' in c.lower() or 'id' in c.lower()), cust_df.columns[1] if len(cust_df.columns) > 1 else cust_df.columns[0])
            term_col = next((c for c in cust_df.columns if 'term' in c.lower() or 'pay' in c.lower()), None)
            
            # Pre-calculate normalized search targets safely
            cust_df['Account Name Clean'] = cust_df[name_col].apply(super_clean_string)
            
            # B. Extract Information from Invoice File
            invoice_customer_name = ""
            invoice_terms = ""
            
            if invoice_file.name.endswith('.pdf'):
                pdf_text = extract_text_from_pdf(invoice_file)
                pdf_text_clean = super_clean_string(pdf_text)
                
                if 'monthly' in pdf_text_clean or 'mpp' in pdf_text_clean:
                    invoice_terms = "monthly"
                if 'due on receipt' in pdf_text_clean or 'dueonreceipt' in pdf_text_clean:
                    invoice_terms = "receipt"
                
                # Check cross-referenced normalized strings against the full text of the PDF
                for _, row_cust in cust_df.iterrows():
                    master_name_clean = row_cust['Account Name Clean']
                    if not master_name_clean or not pdf_text_clean:
                        continue
                    
                    # Split master name by 'dba' safely to isolate corporate identifier
                    core_part = master_name_clean.split(' dba ')[0].strip() if ' dba ' in master_name_clean else master_name_clean
                    
                    if len(core_part) > 4 and (core_part in pdf_text_clean or master_name_clean in pdf_text_clean):
                        invoice_customer_name = str(row_cust[name_col]).strip()
                        break
            else:
                if invoice_file.name.endswith('.csv'):
                    inv_df = pd.read_csv(invoice_file)
                else:
                    inv_df = pd.read_excel(invoice_file, engine='openpyxl')
                inv_df.columns = [str(col).strip() for col in inv_df.columns]
                inv_name_col = next((c for c in inv_df.columns if 'customer' in c.lower() or 'name' in c.lower()), None)
                inv_term_col = next((c for c in inv_df.columns if 'term' in c.lower()), None)
                if inv_name_col and not inv_df.empty:
                    invoice_customer_name = str(inv_df.iloc[0][inv_name_col]).strip()
                if inv_term_col and not inv_df.empty:
                    invoice_terms = str(inv_df.iloc[0][inv_term_col]).lower()

            # C. Load Daily Zoho File
            if zoho_file.name.endswith('.csv'):
                zoho_df = pd.read_csv(zoho_file)
            else:
                zoho_df = pd.read_excel(zoho_file, engine='openpyxl')
            zoho_df.columns = [str(col).strip() for col in zoho_df.columns]
            
            zoho_gross_col = next((c for c in zoho_df.columns if 'gross' in c.lower() or 'amount' in c.lower() or 'total' in c.lower()), zoho_df.columns[3] if len(zoho_df.columns) > 3 else zoho_df.columns[0])
            zoho_fee_col = next((c for c in zoho_df.columns if 'fee' in c.lower() or 'processing' in c.lower()), zoho_df.columns[4] if len(zoho_df.columns) > 4 else zoho_df.columns[0])
            zoho_cust_col = next((c for c in zoho_df.columns if 'customer' in c.lower() or 'name' in c.lower()), zoho_df.columns[9] if len(zoho_df.columns) > 9 else zoho_df.columns[0])
            zoho_type_col = next((c for c in zoho_df.columns if 'type' in c.lower()), None)

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
                if zoho_type_col and str(row[zoho_type_col]).strip().lower() == 'refund':
                    continue
                    
                zoho_cust_name = str(row[zoho_cust_col]).strip() if zoho_cust_col else ""
                cust_name_raw = zoho_cust_name if (zoho_cust_name and zoho_cust_name != "nan") else invoice_customer_name
                
                if not cust_name_raw or cust_name_raw == "nan":
                    continue
                    
                gross_amt = abs(float(str(row[zoho_gross_col]).replace(',', ''))) if zoho_gross_col else 0.0
                fee_amt = abs(float(str(row[zoho_fee_col]).replace(',', ''))) if zoho_fee_col else 0.0
                
                customer_account_num = "MISSING_ACCT"
                payment_term = "receipt"
                final_account_name = cust_name_raw # Fallback if no master row matches
                
                if 'monthly' in invoice_terms or 'mpp' in invoice_terms:
                    payment_term = "monthly"
                
                # Standardize comparison search variables cleanly with single spaces intact
                search_key = super_clean_string(cust_name_raw)
                
                if search_key:
                    # Look up master matching options using structural containment checks
                    match_cust = cust_df[
                        cust_df['Account Name Clean'].str.contains(search_key, na=False) |
                        (search_key in cust_df['Account Name Clean'].to_string())
                    ]
                    
                    if match_cust.empty:
                        for idx_c, row_c in cust_df.iterrows():
                            m_name = row_c['Account Name Clean']
                            if not m_name:
                                continue
                            
                            # Safe isolation fallback check that avoids attribute errors
                            m_core = m_name.split(' dba ')[0].strip() if ' dba ' in m_name else m_name
                            if search_key in m_core or m_core in search_key:
                                match_cust = cust_df.iloc[[idx_c]]
                                break
                    
                    if not match_cust.empty:
                        customer_account_num = str(match_cust.iloc[0][acct_col]).strip()
                        # RULE ENFORCED: Always force final_account_name to use the official corporate MASTER label record
                        final_account_name = str(match_cust.iloc[0][name_col]).strip()
                        if term_col:
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
                
                # ROW 1: THE D
