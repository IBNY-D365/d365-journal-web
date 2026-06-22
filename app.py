import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re

# =========================================================================
# CONSTANTS & CONFIGURATION
# =========================================================================
D365_COLUMNS = [
    "Date", "Voucher", "Account name", "Company", "Account type", "Account",
    "Posting Profile", "Cash code", "Description", "Debit", "Credit",
    "Item sales tax group", "Sales tax code", "Offset company", "Bank Account Type",
    "Offset account", "Offset transaction text", "Currency", "Exchange rate",
    "Item sales tax group2", "Sales group", "Withholding tax group",
    "Release date", "Reversing entry", "Reversing date"
]

CASH_CODE_MAP = {
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
}

# Default paths for permanent master files
MASTER_ACCOUNT_PATH = "Customer Master Account File.xlsx"
FORM_DB_PATH = "Form_Master_DB.xlsx"

# =========================================================================
# APP INTERFACE
# =========================================================================
st.set_page_config(page_title="D365 Zoho Journal Automation", layout="wide")
st.title("📊 D365 General Journal Automation Engine")
st.subheader("Zoho Payments Conversion Portal")

st.sidebar.header("📁 Upload Transaction Data")

# 1. Bank of America Report - Accepts CSV and XLSX
boa_file = st.sidebar.file_uploader("1. Bank of America Report (Excel/CSV)", type=["xlsx", "csv"])

# 2. Zoho Records File - Accepts PDF, CSV, and XLSX
zoho_file = st.sidebar.file_uploader("2. Zoho Records File (PDF/CSV/Excel)", type=["pdf", "csv", "xlsx"])

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Master Settings")

# Check for local constant files so they don't need to be uploaded every time
master_loaded = False
form_db_loaded = False
df_master = None
df_form_db = None

if os.path.exists(MASTER_ACCOUNT_PATH):
    try:
        df_master = pd.read_excel(MASTER_ACCOUNT_PATH)
        master_loaded = True
        st.sidebar.success(f"✔️ Loaded '{MASTER_ACCOUNT_PATH}' dynamically.")
    except Exception as e:
        st.sidebar.error(f"Error loading local master account file: {e}")

if os.path.exists(FORM_DB_PATH):
    try:
        df_form_db = pd.read_excel(FORM_DB_PATH, sheet_name="Sales_PRF")
        form_db_loaded = True
        st.sidebar.success(f"✔️ Loaded '{FORM_DB_PATH}' (Sales_PRF) dynamically.")
    except Exception as e:
        st.sidebar.error(f"Error loading local Form Master DB: {e}")

# Optional backup uploaders if local system storage fails
if not master_loaded:
    uploaded_master = st.sidebar.file_uploader("Upload Customer Master Account File (Fallback)", type=["xlsx"])
    if uploaded_master:
        df_master = pd.read_excel(uploaded_master)
        master_loaded = True

if not form_db_loaded:
    uploaded_form_db = st.sidebar.file_uploader("Upload Form Master DB File (Fallback)", type=["xlsx"])
    if uploaded_form_db:
        try:
            df_form_db = pd.read_excel(uploaded_form_db, sheet_name="Sales_PRF")
            form_db_loaded = True
        except Exception:
            st.sidebar.error("Could not find 'Sales_PRF' tab in uploaded file.")

# =========================================================================
# CORE PROCESSING
# =========================================================================
if boa_file and zoho_file and master_loaded and form_db_loaded:
    try:
        # Load Bank of America File
        if boa_file.name.endswith('.csv'):
            df_boa = pd.read_csv(boa_file)
        else:
            df_boa = pd.read_excel(boa_file)
        
        # Load Zoho File
        if zoho_file.name.endswith('.csv'):
            df_zoho = pd.read_csv(zoho_file)
        elif zoho_file.name.endswith('.pdf'):
            st.warning("⚠️ PDF processing frame triggered. Verify PDF text parsing library layout matches format context.")
            df_zoho = pd.DataFrame(columns=['Customer Name', 'Gross Amount', 'Merchant Fee'])
        else:
            df_zoho = pd.read_excel(zoho_file)

        st.success("⚡ Data pipelines aligned successfully!")

        if st.button("Generate D365 General Journal", type="primary"):
            journal_lines = []
            
            # Identify core column markers
            desc_col = [c for c in df_boa.columns if 'desc' in c.lower() or 'ref' in c.lower()][0]
            date_col = [c for c in df_boa.columns if 'date' in c.lower()][0]
            acct_col = [c for c in df_boa.columns if 'account' in c.lower()][0]
            net_col = [c for c in df_boa.columns if 'net' in c.lower() or 'amount' in c.lower()][0]

            # Standard Transaction Filtering Logic
            df_boa_zoho = df_boa[df_boa[desc_col].astype(str).str.upper().str.contains("ZOHO")]

            for _, boa_row in df_boa_zoho.iterrows():
                boa_desc = str(boa_row[desc_col])
                boa_date = boa_row[date_col]
                boa_acct = str(boa_row[acct_col])
                boa_net = float(boa_row[net_col])

                # Routing Based on Source BOA Account
                offset_acct = "B1000001"
                if "3371" in boa_acct: offset_acct = "B1000002"
                elif "3924" in boa_acct: offset_acct = "B1000003"
                elif "3384" in boa_acct: offset_acct = "B1000001"

                matched_zoho = df_zoho.copy()

                if matched_zoho.empty and zoho_file.name.endswith('.pdf'):
                    continue

                batch_credits = []
                total_fees = 0.0
                
                for _, zoho_row in matched_zoho.iterrows():
                    raw_name = zoho_row.get('Customer Name') or zoho_row.get('Bill To') or ""
                    
                    # Master Name Lookup and Normalization
                    master_lookup = df_master[df_master['Account Name'].astype(str).str.lower() == str(raw_name).lower()]
                    if not master_lookup.empty:
                        account_num = master_lookup.iloc[0]['Account #']
                        account_name = master_lookup.iloc[0]['Account Name']
                    else:
                        account_num = "PENDING_LOOKUP"
                        account_name = raw_name

                    # Cash Code Parameters Rule
                    form_lookup = df_form_db[df_form_db['Customer Account'].astype(str) == str(account_num)]
                    cash_code = "AR001"
                    is_mpp = False
                    
                    if not form_lookup.empty:
                        term_string = str(form_lookup.iloc[0]['Invoice Sent']).lower()
                        matched_term = False
                        for term_key, (code, _) in CASH_CODE_MAP.items():
                            if term_key in term_string:
                                cash_code = code
                                if term_key == "monthly":
                                    is_mpp = True
                                matched_term = True
                                break
                        if not matched_term:
                            cash_code = "AR012"

                    prefix = "MPP " if is_mpp else ""
                    credit_desc = f"{prefix}{account_num} {account_name}_{boa_desc}"
                    
                    gross_amt = float(zoho_row.get('Gross Amount', 0.0))
                    fee_amt = float(zoho_row.get('Merchant Fee', 0.0))
                    total_fees += fee_amt

                    batch_credits.append({
                        "Date": boa_date, "Voucher": "", "Account name": account_name, "Company": "bwa",
                        "Account type": "Customer", "Account": account_num, "Posting Profile": "AutoPost",
                        "Cash code": cash_code, "Description": credit_desc, "Debit": np.nan, "Credit": gross_amt,
                        "Item sales tax group": "", "Sales tax code": "", "Offset company": "bwa",
                        "Bank Account Type": "Bank", "Offset account": offset_acct, "Offset transaction text": "",
                        "Currency": "USD", "Exchange rate": 1.00, "Item sales tax group2": "",
                        "Sales group": "AVATAX", "Withholding tax group": "", "Release date": "",
                        "Reversing entry": "No", "Reversing date": ""
                    })

                # Write customer payment lines
                for cred in batch_credits:
                    journal_lines.append(cred)

                # Append matching consolidated merchant fee segment
                if batch_credits:
                    if len(batch_credits) == 1:
                        debit_desc = f"Zoho Merchant Fee {batch_credits[0]['Description']}"
                    else:
                        acct_summary = ", ".join([f"{c['Account']} {c['Account name']}" for c in batch_credits])
                        debit_desc = f"Zoho Merchant Fee {acct_summary}_{boa_desc}"

                    journal_lines.append({
                        "Date": boa_date, "Voucher": "", "Account name": "Outside Service (Finance)", "Company": "bwa",
                        "Account type": "Ledger", "Account": "43170111-U26C05001-B735350-UOA003", "Posting Profile": "",
                        "Cash code": "OSF005", "Description": debit_desc, "Debit": total_fees, "Credit": np.nan,
                        "Item sales tax group": "", "Sales tax code": "", "Offset company": "bwa",
                        "Bank Account Type": "Bank", "Offset account": offset_acct, "Offset transaction text": "",
                        "Currency": "USD", "Exchange rate": 1.00, "Item sales tax group2": "",
                        "Sales group": "AVATAX", "Withholding tax group": "", "Release date": "",
                        "Reversing entry": "No", "Reversing date": ""
                    })

            if journal_lines:
                df_output = pd.DataFrame(journal_lines, columns=D365_COLUMNS)
                st.dataframe(df_output)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_output.to_excel(writer, index=False, sheet_name='D365_Upload')
                
                st.download_button(
                    label="📥 Download D365 General Journal Excel File",
                    data=buffer.getvalue(),
                    file_name="D365_Zoho_Payments_Journal.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("No processing lines generated based on active filtering definitions.")

    except Exception as e:
        st.error(f"Execution Error during generation: {e}")
else:
    st.info("💡 Ready. Drop the active Bank of America file and Zoho file into the sidebar to compute.")
