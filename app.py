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
BOA_EXPENSES_FILE = "BOA3371 Expenses List.xlsx - BOA3371.csv"

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
    txt = str(text).lower()
    txt = re.sub(r'\S+@\S+', '', txt)
    txt = re.sub(r'\b(com|org|net|edu|gov)\b', '', txt)
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', ' ', txt)
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

            # C. Load BOA Expenses Mapping Guide
            if os.path.exists(BOA_EXPENSES_FILE):
                exp_df = pd.read_csv(BOA_EXPENSES_FILE, skiprows=1)
                exp_df.columns = [str(col).strip() for col in exp_df.columns]
            else:
                exp_df = pd.DataFrame()

            # D. Load Bank of America File & Clean Headings
            if boa_file.name.endswith('.csv'):
                boa_lines = boa_file.getvalue().decode('utf-8', errors='ignore').splitlines()
                skip_count = 0
                for line in boa_lines:
                    if 'Date,Description,Amount' in line or 'date,description,amount' in line.lower():
                        break
                    skip_count += 1
                boa_file.seek(0)
                boa_all_df = pd.read_csv(boa_file, skiprows=skip_count)
            else:
                boa_all_df = pd.read_excel(boa_file, engine='openpyxl')
            
            boa_all_df.columns = [str(col).strip() for col in boa_all_df.columns]
            boa_desc_col = next((c for c in boa_all_df.columns if 'desc' in c.lower() or 'text' in c.lower()), boa_all_df.columns[1])
            boa_date_col_name = next((c for c in boa_all_df.columns if 'date' in c.lower() or 'post' in c.lower()), boa_all_df.columns[0])
            boa_amount_col = next((c for c in boa_all_df.columns if 'amount' in c.lower() or 'amt' in c.lower()), boa_all_df.columns[2])

            # Isolate the explicit Engine Row for Stripe/Zoho Settlements
            is_stripe = boa_all_df[boa_all_df[boa_desc_col].astype(str).str.contains('STRIPE', case=False, na=False)]
            is_zoho = boa_all_df[boa_all_df[boa_desc_col].astype(str).str.contains('ZOHO PAYMENTS', case=False, na=False)]
            
            boa_date = ""
            boa_reference_desc = "PROCESSING CLEARANCE SETTLEMENT"
            boa_net_deposit = 0.0
            
            if not is_stripe.empty:
                engine_mode = "STRIPE"
                boa_date = str(is_stripe.iloc[0][boa_date_col_name]).strip()
                boa_reference_desc = str(is_stripe.iloc[0][boa_desc_col]).strip()
                boa_net_deposit = abs(float(str(is_stripe.iloc[0][boa_amount_col]).replace(',', '').replace('$', '')))
            elif not is_zoho.empty:
                engine_mode = "ZOHO"
                boa_date = str(is_zoho.iloc[0][boa_date_col_name]).strip()
                boa_reference_desc = str(is_zoho.iloc[0][boa_desc_col]).strip()
                boa_net_deposit = abs(float(str(is_zoho.iloc[0][boa_amount_col]).replace(',', '').replace('$', '')))
            else:
                engine_mode = "ZOHO"
                if not boa_all_df.empty:
                    boa_date = str(boa_all_df.iloc[0][boa_date_col_name]).strip()
                    boa_reference_desc = str(boa_all_df.iloc[0][boa_desc_col]).strip()
                    boa_net_deposit = abs(float(str(boa_all_df.iloc[0][boa_amount_col]).replace(',', '').replace('$', '')))

            journal_rows = []

            # ==========================================
            # ENGINE MODE 1: STRIPE TRUE PARSER (7+1 DATA LAYOUT)
            # ==========================================
            if engine_mode == "STRIPE" or gateway_file.name.endswith('.pdf'):
                st.info("⚙️ Running Engine Mode: Stripe Factual PDF Extractor")
                
                pdf_text = extract_text_from_pdf(gateway_file)
                charge_blocks = []
                total_extracted_gross = 0.0
                
                lines = [l.strip() for l in pdf_text.split('\n') if l.strip()]
                for idx, line in enumerate(lines):
                    line_lower = line.lower()
                    
                    if "charge" in line_lower and ("plan" in line_lower or "agreement" in line_lower):
                        amounts = [float(amt.replace(',', '')) for amt in re.findall(r'\d+(?:,\d{3})*(?:\.\d{2})', line)]
                        
                        if amounts:
                            gross_amt = amounts[0]
                            total_extracted_gross += gross_amt
                            
                            extracted_name = "Unknown Customer"
                            if "agreement -" in line_lower:
                                extracted_name = line.split("Agreement -")[-1].split(" -")[0].strip()
                            elif "plan -" in line_lower:
                                extracted_name = line.split("Plan -")[-1].split(" -")[0].strip()
                            
                            extracted_name = re.sub(r'\S+@\S+', '', extracted_name).split('@')[0].strip()
                            extracted_name = extracted_name.strip(" -")
                                
                            charge_blocks.append({
                                "name": extracted_name,
                                "gross": gross_amt,
                                "is_installment": "installment" in line_lower
                            })

                for charge in charge_blocks:
                    gross_val = charge["gross"]
                    is_inst = charge["is_installment"]
                    
                    cash_code_label = "Installment" if is_inst else "Receipt"
                    cash_code = dynamic_cash_code_lookup(cash_code_label, "AR002" if is_inst else "AR001")
                    
                    customer_account_num = "MISSING_ACCT"
                    final_account_name = charge["name"]
                    
                    search_key = super_clean_string(final_account_name)
                    if search_key:
                        match_cust = cust_df[cust_df['Account Name Clean'].apply(lambda x: search_key in str(x) or str(x) in search_key)]
                        
                        if match_cust.empty:
                            search_tokens = set(search_key.split())
                            best_match_row = None
                            max_overlap = 0
                            
                            for m_idx, m_row in cust_df.iterrows():
                                master_tokens = set(str(m_row['Account Name Clean']).split())
                                overlap = len(search_tokens.intersection(master_tokens))
                                
                                if overlap >= 2 and overlap > max_overlap:
                                    max_overlap = overlap
                                    best_match_row = m_row
                            
                            if best_match_row is not None:
                                customer_account_num = str(best_match_row[acct_col]).strip()
                                final_account_name = str(best_match_row[name_col]).strip()
                        else:
                            customer_account_num = str(match_cust.iloc[0][acct_col]).strip()
                            final_account_name = str(match_cust.iloc[0][name_col]).strip()
                    
                    credit_desc = f"MPP {customer_account_num} {final_account_name}_{boa_reference_desc}" if is_inst else f"{customer_account_num} {final_account_name}_{boa_reference_desc}"
                    
                    journal_rows.append({
                        "Date": boa_date, "Voucher": "", "Account name": final_account_name, "Company": company_id,
                        "Account type": "Customer", "Account": customer_account_num, "Posting profile": "AutoPost",
                        "Cash code": cash_code, "Description": credit_desc, "Debit": "", "Credit": gross_val,
                        "Item sales tax group": "", "Sales tax code": "", "Offset company": company_id, "Offset account type": "Bank",
                        "Offset account": offset_account, "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                        "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "", "Release date": "",
                        "Reversing entry": "No", "Reversing date": ""
                    })
                
                calculated_stripe_fee = round(total_extracted_gross - boa_net_deposit, 2)
                
                if calculated_stripe_fee > 0:
                    merchant_cash_code = dynamic_cash_code_lookup("Stripe Merchant Fee", "OSF005")
                    debit_desc = f"Stripe Merchant Fee_{boa_reference_desc}"
                    journal_rows.append({
                        "Date": boa_date, "Voucher": "", "Account name": "Outside Service (Finance)", "Company": company_id,
                        "Account type": "Ledger", "Account": debit_ledger_acct, "Posting profile": "",
                        "Cash code": merchant_cash_code, "Description": debit_desc, "Debit": calculated_stripe_fee, "Credit": "",
                        "Item sales tax group": "", "Sales tax code": "", "Offset company": company_id, "Offset account type": "Bank",
                        "Offset account": offset_account, "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                        "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "", "Release date": "",
                        "Reversing entry": "No", "Reversing date": ""
                    })

            # ==========================================
            # ENGINE MODE 2: ZOHO RECONCILIATION ENGINE
            # ==========================================
            else:
                st.info("⚙️ Running Engine Mode: Zoho Corporate Payment Pipeline")
                invoice_terms = "receipt"
                extracted_payer_from_invoice = ""
                
                if invoice_file and invoice_file.name.endswith('.pdf'):
                    pdf_text_raw = extract_text_from_pdf(invoice_file)
                    pdf_text_clean = super_clean_string(pdf_text_raw)
                    if 'monthly' in pdf_text_clean or 'mpp' in pdf_text_clean:
                        invoice_terms = "monthly"
                    
                    lines = [l.strip() for l in pdf_text_raw.split('\n') if l.strip()]
                    for idx, line in enumerate(lines):
                        if "bill to" in line.lower() or "invoice to" in line.lower():
                            if idx + 1 < len(lines):
                                extracted_payer_from_invoice = lines[idx + 1].strip()
                                break
                
                if gateway_file.name.endswith('.csv'):
                    zoho_df = pd.read_csv(gateway_file)
                else:
                    zoho_df = pd.read_excel(gateway_file, engine='openpyxl')
                zoho_df.columns = [str(col).strip() for col in zoho_df.columns]
                
                zoho_gross_col = "Amount" if "Amount" in zoho_df.columns else zoho_df.columns[3]
                zoho_fee_col = "Fee" if "Fee" in zoho_df.columns else zoho_df.columns[4]
                zoho_cust_col = "CustomerName" if "CustomerName" in zoho_df.columns else None
                
                for idx, row in zoho_df.iterrows():
                    payer_name = str(row[zoho_cust_col]).strip() if zoho_cust_col and zoho_cust_col in zoho_df.columns else ""
                    if not payer_name or payer_name == "nan" or "inbody" in payer_name.lower():
                        payer_name = extracted_payer_from_invoice if extracted_payer_from_invoice else "Unknown Payer"
                    
                    customer_account_num = "MISSING_ACCT"
                    final_account_name = payer_name
                    
                    search_key = super_clean_string(final_account_name)
                    if search_key:
                        match_cust = cust_df[cust_df['Account Name Clean'].str.contains(search_key, na=False)]
                        if not match_cust.empty:
                            customer_account_num = str(match_cust.iloc[0][acct_col]).strip()
                            final_account_name = str(match_cust.iloc[0][name_col]).strip()
                    
                    gross_amt = abs(float(str(row[zoho_gross_col]).replace(',', '')))
                    fee_amt = abs(float(str(row[zoho_fee_col]).replace(',', '')))
                    
                    cash_code = dynamic_cash_code_lookup("Installment" if invoice_terms == "monthly" else "Receipt", "AR002" if invoice_terms == "monthly" else "AR001")
                    credit_desc = f"MPP {customer_account_num} {final_account_name}_{boa_reference_desc}" if invoice_terms == "monthly" else f"{customer_account_num} {final_account_name}_{boa_reference_desc}"
                    
                    journal_rows.append({
                        "Date": boa_date, "Voucher": "", "Account name": final_account_name, "Company": company_id,
                        "Account type": "Customer", "Account": customer_account_num, "Posting profile": "AutoPost",
                        "Cash code": cash_code, "Description": credit_desc, "Debit": "", "Credit": gross_amt,
                        "Item sales tax group": "", "Sales tax code": "", "Offset company": company_id, "Offset account type": "Bank",
                        "Offset account": offset_account, "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                        "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "", "Release date": "",
                        "Reversing entry": "No", "Reversing date": ""
                    })
                    
                    if fee_amt > 0:
                        debit_desc = f"Zoho Merchant Fee {customer_account_num}_{final_account_name}_{boa_reference_desc}"
                        journal_rows.append({
                            "Date": boa_date, "Voucher": "", "Account name": "Outside Service (Finance)", "Company": company_id,
                            "Account type": "Ledger", "Account": debit_ledger_acct, "Posting profile": "",
                            "Cash code": "OSF005", "Description": debit_desc, "Debit": fee_amt, "Credit": "",
                            "Item sales tax group": "", "Sales tax code": "", "Offset company": company_id, "Offset account type": "Bank",
                            "Offset account": offset_account, "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                            "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "", "Release date": "",
                            "Reversing entry": "No", "Reversing date": ""
                        })

            # ==========================================
            # STEP E: AUTOMATED BANK STATEMENT EXPENSE INGESTION
            # ==========================================
            if not exp_df.empty:
                for idx, r_row in boa_all_df.iterrows():
                    raw_desc = str(r_row[boa_desc_col]).strip()
                    raw_amt_str = str(r_row[boa_amount_col]).replace(',', '').replace('$', '').strip()
                    
                    try:
                        raw_amt = float(raw_amt_str)
                    except ValueError:
                        continue
                        
                    # Process lines where corporate funds were debited (spent)
                    if raw_amt < 0:
                        debit_value = abs(raw_amt)
                        
                        # Check statement descriptions against your Expense Matrix mapping guide
                        for e_idx, e_row in exp_df.iterrows():
                            lookup_keyword = str(e_row['Bank Transaction Description']).split('*')[0].split()[0].strip().lower()
                            
                            if lookup_keyword and lookup_keyword in raw_desc.lower():
                                map_name = str(e_row['Account name']).strip()
                                map_type = str(e_row['Account type']).strip()
                                map_acct = str(e_row['Account']).strip()
                                map_cc = str(e_row['Cash code']).strip() if 'Cash code' in e_row else "SP001"
                                map_desc_base = str(e_row['Description']).strip()
                                
                                # Profile defaults: Blank out for Ledger entries, mark 'AutoPost' for Vendors
                                profile_flag = "AutoPost" if map_type.lower() == "vendor" else ""
                                combined_desc = f"{map_desc_base}_{raw_desc}"
                                t_date = str(r_row[boa_date_col_name]).strip()
                                
                                journal_rows.append({
                                    "Date": t_date, "Voucher": "", "Account name": map_name, "Company": company_id,
                                    "Account type": map_type, "Account": map_acct, "Posting profile": profile_flag,
                                    "Cash code": map_cc, "Description": combined_desc, "Debit": debit_value, "Credit": "",
                                    "Item sales tax group": "", "Sales tax code": "", "Offset company": company_id, "Offset account type": "Bank",
                                    "Offset account": offset_account, "Offset transaction text": "", "Currency": "USD", "Exchange rate": 1.00,
                                    "Item sales tax group2": "", "Sales tax group": "AVATAX", "Withholding tax group": "", "Release date": "",
                                    "Reversing entry": "No", "Reversing date": ""
                                })
                                break

            # Create final structured 25-column template layout output
            columns_25 = [
                "Date", "Voucher", "Account name", "Company", "Account type", "Account",
                "Posting profile", "Cash code", "Description", "Debit", "Credit",
                "Item sales tax group", "Sales tax code", "Offset company", "Offset account type", "Offset account",
                "Offset transaction text", "Currency", "Exchange rate", "Item sales tax group2",
                "Sales tax group", "Withholding tax group", "Release date", "Reversing entry", "Reversing date"
            ]
            
            final_df = pd.DataFrame(journal_rows)
            if not final_df.empty:
                final_df = final_df.reindex(columns=columns_25).fillna("")
                st.success("🎉 Cross-matched and balanced journal entries created seamlessly!")
                st.dataframe(final_df)
                
                csv_data = final_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Perfect D365 Upload CSV",
                    data=csv_data,
                    file_name="D365_Reconciliation_Journal.csv",
                    mime="text/csv"
                )
            else:
                st.warning("⚠️ No valid transaction entries found to process.")
            
    except Exception as e:
        st.error(f"❌ Automation mapping process failed: {str(e)}")
else:
    st.info("💡 Please upload your Gateway File (Zoho or Stripe) and Bank of America statement to activate the automated alignment engine.")
