import streamlit as st
import pandas as pd
import re
import os

# Set up page layout
st.set_page_config(page_title="D365 Zoho Journal Generator", layout="centered")
st.title("📊 D365 Zoho Journal Generator")
st.write("Upload your daily Zoho Payments Export (CSV or Excel) to generate a perfectly balanced D365 import CSV.")

# 1. SIDEBAR SETTINGS (Matches D365 General Journal Setup)
st.sidebar.header("⚙️ D365 Account Settings")
bank_account = st.sidebar.text_input("Bank Clearing Account", value="110100")
fee_account = st.sidebar.text_input("Merchant Fee Account", value="615000")
journal_name = st.sidebar.text_input("Journal Name", value="GEN-JRN")

# 2. FILE UPLOADER (Accepts both CSV and XLSX now!)
st.subheader("1. Upload Daily File")
zoho_file = st.file_uploader("Drop today's Zoho Payments Export", type=["csv", "xlsx"])

# Target filename inside GitHub backend
CUSTOMER_MASTER_PATH = "Customer Master Account File.csv"

def extract_invoice_number(description):
    if pd.isna(description):
        return None
    match = re.search(r'INV-\d+', str(description))
    return match.group(0) if match else None

# 3. PROCESSING AUTOMATION PIPELINE
if zoho_file:
    st.subheader("2. Review & Generate")
    
    try:
        # Check for permanent master mapping sheet
        if not os.path.exists(CUSTOMER_MASTER_PATH):
            st.error(f"❌ Error: Could not find your master reference file '{CUSTOMER_MASTER_PATH}' in your GitHub repository. Please ensure it is uploaded.")
        else:
            # Load Customer Master Lookup
            cust_df = pd.read_csv(CUSTOMER_MASTER_PATH)
            cust_df['Account Name'] = cust_df['Account Name'].astype(str).str.strip().str.lower()
            
            # Load incoming daily transaction file safely depending on extension
            if zoho_file.name.endswith('.csv'):
                zoho_df = pd.read_csv(zoho_file)
            else:
                zoho_df = pd.read_excel(zoho_file)
            
            # Standardize column names to remove accidental hidden spaces
            zoho_df.columns = [str(col).strip() for col in zoho_df.columns]
            
            journal_lines = []
            line_number = 1
            voucher_counter = 1
            
            for idx, row in zoho_df.iterrows():
                # Formats clean sequence numbering for your vouchers (VOU-001, VOU-002, etc.)
                voucher_id = f"VOU-{voucher_counter:03d}"
                
                raw_desc = row.get('Description', '')
                gross_amt = float(row.get('Gross Amount', 0))
                fee_amt = float(row.get('Merchant Fee', 0))
                net_amt = float(row.get('Net Amount', 0))
                
                # Extract embedded invoice parameters safely
                inv_num = extract_invoice_number(raw_desc)
                
                customer_id = "MISSING_CUST_ID"
                cust_name_clean = str(row.get('Customer Name', '')).strip().lower()
                
                # Match company records dynamically
                match = cust_df[cust_df['Account Name'] == cust_name_clean]
                if not match.empty:
                    customer_id = match.iloc[0]['Account']
                
                # ==========================================
                # LINE 1: CUSTOMER REVENUE LINE (GROSS DEBIT)
                # ==========================================
                cust_description = f"Gross Rec - Invoice {inv_num}" if inv_num else f"Gross Rec - {raw_desc}"
                journal_lines.append({
                    "JournalName": journal_name,
                    "LineNumber": line_number,
                    "Voucher": voucher_id,
                    "AccountType": "Customer",
                    "LedgerDimension": customer_id,
                    "Description": cust_description,
                    "Debit": gross_amt,
                    "Credit": "",
                    "CurrencyCode": "USD"
                })
                line_number += 1
                
                # ==========================================
                # LINE 2: MERCHANT EXPENSE LINE (FEE DEBIT)
                # ==========================================
                if fee_amt > 0:
                    fee_description = f"Zoho Merchant Fee - {inv_num}" if inv_num else "Zoho Merchant Fee Deduction"
                    journal_lines.append({
                        "JournalName": journal_name,
                        "LineNumber": line_number,
                        "Voucher": voucher_id,
                        "AccountType": "Ledger",
                        "LedgerDimension": fee_account,
                        "Description": fee_description,
                        "Debit": fee_amt,
                        "Credit": "",
                        "CurrencyCode": "USD"
                    })
                    line_number += 1
                    
                # ==========================================
                # LINE 3: SETTLEMENT CASH LINE (NET CREDIT)
                # ==========================================
                bank_description = f"Net Settlement Deposit - Zoho Payments {inv_num}" if inv_num else "Net Settlement Deposit - Zoho Payments"
                journal_lines.append({
                    "JournalName": journal_name,
                    "LineNumber": line_number,
                    "Voucher": voucher_id,
                    "AccountType": "Ledger",
                    "LedgerDimension": bank_account,
                    "Description": bank_description,
                    "Debit": "",
                    "Credit": net_amt,
                    "CurrencyCode": "USD"
                })
                line_number += 1
                voucher_counter += 1
                
            # Put data into the final DataFrame matching your D365 template layout
            final_df = pd.DataFrame(journal_lines)
            
            st.success("✅ Workflow files matched and formatted successfully! Previewing your D365 Journal below:")
            st.dataframe(final_df.head(15))
            
            # Export function matching the D365 framework requirements
            csv_data = final_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download D365 Upload CSV",
                data=csv_data,
                file_name="D365_Ready_General_Journal.csv",
                mime="text/csv"
            )
            
    except Exception as e:
        st.error(f"❌ An error occurred during matching processing: {str(e)}")
else:
    st.info("💡 Drop your Zoho export file (CSV or XLSX) above to verify formatting and generate your import package.")
