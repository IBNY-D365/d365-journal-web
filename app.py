import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# =========================================================================
# CONSTANTS & CONFIGURATION
# =========================================================================
D365_COLUMNS = [
    "Date", "Voucher", "Account name", "Company", "Account type", "Account",
    "Posting Profile", "Cash code", "Description", "Debit", "Credit",
    "Item sales tax group", "Sales tax code", "Offset company", "Bank Account Type",
    "Offset account", "Offset transaction text", "Currency", "Exchange rate",
    "Item sales tax group2", "Sales tax group", "Withholding tax group",
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

# =========================================================================
# APP INTERFACE
# =========================================================================
st.set_page_config(page_title="D365 Zoho Journal Automation", layout="wide")
st.title("📊 D365 General Journal Automation Engine")
st.subheader("Zoho Payments Conversion Portal")

st.sidebar.header("📁 Upload Source Data Files")
boa_file = st.sidebar.file_uploader("1. Bank of America Report (Excel)", type=["xlsx"])
zoho_file = st.sidebar.file_uploader("2. Zoho Records File (CSV/Excel)", type=["csv", "xlsx"])
master_file = st.sidebar.file_uploader("3. Customer Master Account File (Excel)", type=["xlsx"])
form_db_file = st.sidebar.file_uploader("4. Form Master DB File (Excel)", type=["xlsx"])

if boa_file and zoho_file and master_file and form_db_file:
    try:
        # Load Reference Data
        df_master = pd.read_excel(master_file)
        df_form_db = pd.read_excel(form_db_file, sheet_name="Sales_PRF")
        
        # Load Primary Reports
        df_boa = pd.read_excel(boa_file)
        if zoho_file.name.endswith('.csv'):
            df_zoho = pd.read_csv(zoho_file)
        else:
            df_zoho = pd.read_excel(zoho_file)

        st.success("⚡ All configuration and source data files loaded successfully!")

        # Processing Engine Triggers
        if st.button("Generate D365 General Journal", type="primary"):
            journal_lines = []
            
            # Filter Zoho Payments out from BOA Report
            # Rule: Source == "Bank of America" AND Description CONTAINS "ZOHO"
            # (Assuming Description/Reference columns vary, we check all string components or a designated column)
            desc_col = [c for c in df_boa.columns if 'desc' in c.lower() or 'ref' in c.lower()][0]
            date_col = [c for c in df_boa.columns if 'date' in c.lower()][0]
            acct_col = [c for c in df_boa.columns if 'account' in c.lower()][0]
            net_col = [c for c in df_boa.columns if 'net' in c.lower() or 'amount' in c.lower()][0]

            df_boa_zoho = df_boa[df_boa[desc_col].astype(str).str.upper().str.contains("ZOHO")]

            for _, boa_row in df_boa_zoho.iterrows():
                boa_desc = str(boa_row[desc_col])
                boa_date = boa_row[date_col]
                boa_acct = str(boa_row[acct_col])
                boa_net = float(boa_row[net_col])

                # Determine conditional Offset Account routing based on BOA source account
                offset_acct = "B1000001"  # Default / Fallback fallback
                if "3371" in boa_acct: offset_acct = "B1000002"
                elif "3924" in boa_acct: offset_acct = "B1000003"
                elif "3384" in boa_acct: offset_acct = "B1000001"

                # Filter matching records in Zoho file for this settlement sequence
                # (For implementation, mapping relies on correlating transaction reference or date matching)
                # Here we fetch matching records from Zoho dataset
                matched_zoho = df_zoho.copy() # In practice, filter matched_zoho by unique key or transaction ID

                if matched_zoho.empty:
                    continue

                batch_credits = []
                total_fees = 0.0
                
                for _, zoho_row in matched_zoho.iterrows():
                    # Name Resolution & Normalization Pipeline
                    raw_name = zoho_row.get('Customer Name') or zoho_row.get('Bill To') or ""
                    
                    # Master Lookup Rule
                    master_lookup = df_master[df_master['Account Name'].astype(str).str.lower() == str(raw_name).lower()]
                    if not master_lookup.empty:
                        account_num = master_lookup.iloc[0]['Account #']
                        account_name = master_lookup.iloc[0]['Account Name']
                    else:
                        account_num = "PENDING_LOOKUP"
                        account_name = raw_name

                    # Cash Code & Payment Term Parameters Rule
                    form_lookup = df_form_db[df_form_db['Customer Account'].astype(str) == str(account_num)]
                    cash_code = "AR001"  # Default
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
                            cash_code = "AR012" # Fallback mapping

                    # Construct Credit Description String
                    prefix = "MPP " if is_mpp else ""
                    credit_desc = f"{prefix}{account_num} {account_name}_{boa_desc}"
                    
                    gross_amt = float(zoho_row.get('Gross Amount', 0.0))
                    fee_amt = float(zoho_row.get('Merchant Fee', 0.0))
                    total_fees += fee_amt

                    # Save intermediate Credit Line items
                    batch_credits.append({
                        "Date": boa_date, "Voucher": "", "Account name": account_name, "Company": "bwa",
                        "Account type": "Customer", "Account": account_num, "Posting Profile": "AutoPost",
                        "Cash code": cash_code, "Description": credit_desc, "Debit": np.nan, "Credit": gross_amt,
                        "Item sales tax group": "", "Sales tax code": "", "Offset company": "bwa",
                        "Bank Account Type": "Bank", "Offset account": offset_acct, "Offset transaction text": "",
                        "Currency": "USD", "Exchange rate": 1.00, "Item sales tax group2": "",
                        "Sales tax group": "AVATAX", "Withholding tax group": "", "Release date": "",
                        "Reversing entry": "No", "Reversing date": ""
                    })

                # Process Multiplicity and Batch Resolution rules
                # Output credits sequentially
                for cred in batch_credits:
                    journal_lines.append(cred)

                # Append single unified Debit Fee Line per batch settlement block
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
                    "Sales tax group": "AVATAX", "Withholding tax group": "", "Release date": "",
                    "Reversing entry": "No", "Reversing date": ""
                })

            # Build and Render Result Output Dataframe
            df_output = pd.DataFrame(journal_lines, columns=D365_COLUMNS)
            st.dataframe(df_output)

            # Generate Downloadable Buffer Array
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_output.to_excel(writer, index=False, sheet_name='D365_Upload')
            
            st.download_button(
                label="📥 Download D365 General Journal Excel File",
                data=buffer.getvalue(),
                file_name="D365_Zoho_Payments_Journal.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Execution Error: Missing expected headers or formatting anomalies encountered. Details: {e}")
else:
    st.info("💡 Please upload all 4 required Excel/CSV ecosystem dependencies in the sidebar menu to launch processing.")
