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

# Enhanced normalization function to strip punctuation and handle periods like "E." vs "E"
def super_clean_string(text):
    if pd.isna(text):
        return ""
    cleaned = re.sub(r'[.,\-_()\[\]]', ' ', str(text).lower())
    return "".join(cleaned.split())

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
            
            # Pre-calculate normalized search targets
            cust_df['Account Name Clean'] = cust_df[name_col].apply(super_clean_string)
            
            # B. Extract Information from Invoice File
            invoice_customer_name = ""
            invoice_terms = ""
            
            if invoice_file.name.endswith('.pdf'):
                pdf_text = extract_text_from_pdf(invoice_file)
                pdf_text_clean = super_clean_string(pdf_text)
                
                if 'monthly' in pdf_text_clean or 'mpp' in pdf_text_clean:
                    invoice_terms = "monthly"
                if 'dueonreceipt' in pdf_text_clean:
                    invoice_terms = "receipt"
                
                # Check cross-referenced normalized strings against the full text of the PDF
                for _, row_cust in cust_df.iterrows():
                    master_name_clean = row_cust['Account Name Clean']
                    core_part = master_name_clean.split('dba')[0].strip()
                    
                    if len(core_part) > 4 and (core_part in pdf_text_clean or master_name_clean in pdf_text_clean):
                        invoice_customer_name = str(row_cust[name_col]).strip()
                        break
            else:
                if invoice_file.name.endswith('.csv'):
                    inv_df = pd.read_csv(invoice_file)
                else:
                    inv_df = pd.read_excel(invoice_file, engine='openpyxl')
                inv_df.columns = [str(col).strip() for col in inv_df.columns]
                inv_name_col = next((c for c in inv_df.columns
