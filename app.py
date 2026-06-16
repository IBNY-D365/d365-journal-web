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
    # Strip non-alphanumeric artifacts, punctuation, and hyphens completely
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', ' ', str(text).lower())
    # Strip common corporate designations to isolate core matching names
    cleaned = re.sub(r'\b(llc|pllc|inc|corp|co|incorporated|limited|llp)\b', ' ', cleaned)
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
                    boa_reference_desc = str(boa_df.iloc[0][boa_desc_col]).strip()

            journal_rows = []

            # ==========================================
            # ENGINE MODE 1: STRIPE TRUE PARSER (7+1 DATA LAYOUT)
            # ==========================================
            if engine_mode == "STRIPE" or gateway_file.name.endswith('.pdf'):
                st.info("⚙️ Running Engine Mode: Stripe Factual PDF Extractor")
                
                pdf_text = extract_text_from_pdf(gateway_file)
                charge_blocks = []
                stripe_fee_accumulator = 0.0
                
                lines = [l.strip() for l in pdf_text.split('\n') if l.strip()]
                for line in lines:
                    line_lower = line.lower()
                    if "stripe fee" in line_lower:
                        fee_matches = [float(amt.replace('$', '').replace('USD', '').strip()) for amt in re.findall(r'-\s*\d+\.\d{2}', line)]
                        if fee_matches:
                            stripe_fee_accumulator += sum(fee_matches)
                    
                    elif "charge" in line_lower and ("plan" in line_lower or "agreement" in line_lower):
                        amounts = [float(amt.replace(',', '')) for amt in re.findall(r'\d+(?:,\d{3})*(?:\.\d{2})', line)]
                        
                        if amounts:
                            gross_amt = amounts[0]
                            row_fee = abs(amounts[1]) if len(amounts) > 1 else 0.0
                            stripe_fee_accumulator += row_fee
                            
                            # Cleanly extract true customer string by splitting out metadata
                            extracted_name = line.split('-')[-2].strip() if '-' in line else "Unknown Customer"
                            if "@" in extracted_name or ".com" in extracted_name:
                                extracted_name = line.split('-')[-3].strip()
                                
                            charge_blocks.append({
                                "name": extracted_name,
                                "gross": gross_amt,
                                "is_installment": "installment" in line_lower
                            })

                # Step 2: Generate entries with enhanced word-intersection matching
                for charge in charge_blocks:
                    gross_val = charge["gross"]
                    is_inst = charge["is_installment"]
                    
                    cash_code_label = "Installment" if is_inst else "Receipt"
                    cash_code = dynamic_cash_code_lookup(cash_code_label, "AR002" if is_inst else "AR001")
                    
                    customer_account_num = "MISSING_ACCT"
                    final_account_name = charge["name"].strip(" -")
                    
                    search_key = super_clean_string(final_account_name)
                    if search_key:
                        # Direct containment or intersection checking to bypass string artifact matching failures
                        match_cust = cust_df[cust_df['Account Name Clean'].apply(lambda x: search_key in str(x) or str(x) in search_key)]
                        
                        if match_cust.empty:
                            # Advanced token check: look for first two primary words overlapping
                            search_tokens = search_key.split()
                            if len(search_tokens) >= 2:
                                token_key = " ".join(search_tokens[:2])
                                match_cust = cust_df[cust_df['Account Name Clean'].str.contains(token_key, na=False)]

                        if not match_cust.empty:
                            customer_account_num = str(match_cust.iloc[0][acct_col]).strip()
                            final_account_name = str(match_cust.iloc[0][name_col]).strip()
                    
                    credit_desc = f"MPP {customer_account_num} {final_account_name}_{boa_reference_desc}" if is_inst else f"{customer_account_num} {final_account_name}_{boa_reference_desc}"
                    
                    journal_rows.append({
                        "Date": boa_date, "Voucher": "", "Account name": final_account_name, "Company": company_id,
                        "Account type": "Customer", "Account": customer_account_num, "Posting profile": "AutoPost",
                        "Cash code": cash_code, "Description": credit_desc
