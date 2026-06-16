import streamlit as st
import pandas as pd
import re
import os
from pypdf import PdfReader

# Set up page layout
st.set_page_config(page_title="D365 Accounting Journal Generator", layout="wide")
st.title("📊 D365 Zoho, Stripe & BOA Journal Generator")
st.write("Upload your daily processing packages below to build your flawless 25-column D365 upload templates.")

# 1. SIDEBAR FIXED CONFIGURATION
st.sidebar.header("⚙️ Fixed D365 Settings")
company_id = st.sidebar.text_input("Company", value="bwa")
offset_account = st.sidebar.text_input("Offset Account", value="B1000002")
debit_ledger_acct = st.sidebar.text_input("Debit Line Account", value="43170111-U26C05001-B735350-UOA003")

# Hardcoded exact repository filenames visible in GitHub
MASTER_FILE_NAME = "Customer Master Account File.xlsx"
CASH_CODE_FILE = "Cash Code Masterlist.xlsx"

# 2. THREE-FILE UPLOADERS
col1, col2, col3 = st.columns(3)
with col1:
    gateway_file = st.file_uploader("1. Upload Processing File (Zoho CSV or Stripe PDF)", type=["csv", "xlsx", "pdf"])
with col2:
    invoice_file = st.file_uploader("2. Upload Invoice PDF (Required for Zoho Only)", type=["pdf", "csv", "xlsx", "txt"])
with col3:
    boa_file = st.file_uploader("3. Upload Bank of America Statement", type=["csv", "xlsx"])

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

def super_clean_string(text):
    if pd.isna(text) or text is None:
        return ""
    cleaned = re.sub(r'[.,\-_()\[\]]', ' ', str(text).lower())
    return " ".join(cleaned.split())

# 3. AUTOMATED PARSING AND BALANCING PIPELINE
if gateway_file and boa_file:
    st.subheader("4. Review & Generate")
    
    try:
        if not os.path.exists(MASTER_FILE_NAME) or not os.path.exists(CASH_CODE_FILE):
            st.error(f"❌ Error: Missing repository lookup master files. Please verify GitHub sync for '{MASTER_FILE_NAME}' and '{CASH_CODE_FILE}'.")
        else:
            # A. Load Customer Master File
            cust_df = pd.read_excel(MASTER_FILE_NAME, engine='openpyxl')
            cust_df.columns = [str(col).strip() for col in cust_df.columns]
            name_col = "Account Name" if "Account Name" in cust_df.columns else cust_df.columns[2]
            acct_col = "Account" if "Account" in cust_df.columns else cust_df.columns[1]
            cust_df['Account Name Clean'] = cust_df[name_col].apply(super_clean_string)
            
            # B. Load Cash Code Masterlist File
            cc_df = pd.read_excel(CASH_CODE_FILE, engine='openpyxl')
            cc_df.columns = [str(col).strip() for col in cc_df.columns]
            cc_term_col = next((c for c in cc_df.columns if 'term' in c.lower() or 'desc' in c.lower() or 'name' in c.lower()), cc_df.columns[0])
            cc_code_col = next((c for c in cc_df.columns if 'code' in c.lower()), cc_df.columns[1])
            
            # Helper to safely lookup cash codes from master sheet dynamically
            def dynamic_cash_code_lookup(term_string, fallback):
                clean_term = super_clean_string(term_string)
                match = cc_df[cc_df[cc_term_col].apply(super_clean_string).str.contains(clean_term, na=False)]
                if not match.empty:
                    return str(match.iloc[0][cc_code_col]).strip()
                return fallback

            # C. Load Bank of America File & Detect Settlement Engine
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
            
            boa_desc_col = next((c for c in boa_df.columns if 'desc' in c.lower() or 'text' in c.lower()), boa_df.columns[1])
            boa_date_col_name = next((c for c in boa_df.columns if 'date' in c.lower() or 'post' in c.lower()), boa_df.columns[0])
            
            is_stripe = boa_df[boa_df[boa_desc_col].astype(str).str.contains('STRIPE', case=False, na=False)]
            is_zoho = boa_df[boa_df[boa_desc_col].astype(str).str.contains('ZOHO PAYMENTS', case=False, na=False)]
            
            boa_date = ""
            boa_reference_desc = "PROCESSING CLEARANCE SETTLEMENT"
            
            if not is_stripe.empty:
                engine_mode = "STRIPE"
                boa_date = str(is_stripe.iloc[0][boa_date_col_name]).strip()
                boa_reference_desc = str(is_stripe.iloc[0][boa_desc_col]).strip()
            elif not is_zoho.empty:
                engine_mode = "ZOHO"
                boa_date = str(is_zoho.iloc[0][boa_date_col_name]).strip()
                boa_reference_desc = str(is_zoho.iloc[0][boa_desc_col]).strip()
            else:
                engine_mode = "ZOHO"
                if not boa_df.empty:
                    boa_date = str(boa_df.iloc[0][boa_date_col_name]).strip()
