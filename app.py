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
            
            name_col = "Account Name" if "Account Name" in cust_df.columns else cust_df.columns[2]
            acct_col = "Account" if "Account" in cust_df.columns else cust_df.columns[1]
            term_col = "Terms" if "Terms" in cust_df.columns else None
            
            cust_df['Account Name Clean'] = cust_df[name_col].apply(super_clean_string)
            
            # B. Extract Terms Information from Invoice File
            invoice_terms = "receipt" 
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
            
            zoho_gross_col = "Amount" if "Amount" in zoho_df.columns else zoho_df.columns[3]
            zoho_fee_col = "Fee" if "Fee" in zoho_df.columns else zoho_df.columns[4]
            zoho_cust_col = "CustomerName" if "CustomerName" in zoho_df.columns else zoho_df.columns[9]
            zoho_desc_col = "Description" if "Description" in zoho_df.columns else zoho_df.columns[10]
            zoho_type_col = "TransactionType" if "TransactionType" in zoho_df.columns else None
            zoho_date_col = "TransactionTime" if "TransactionTime" in zoho_df.columns else zoho_df.columns[6]

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
            
            # FIX: Widen containing search from 'ZOHO PAYMENTS' to just 'ZOHO' to avoid text variant failure
            boa_match_row = boa_df[boa_df['Description'].astype(str).str.contains('ZOHO', case=False, na=False)]
            boa_date = ""
            boa_reference_desc = "ZOHO PAYMENTS SETTLEMENT"
            
            if not boa_match_row.empty
