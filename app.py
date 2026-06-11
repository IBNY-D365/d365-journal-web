import streamlit as st
import pandas as pd
import re

# Set up page layout
st.set_page_config(page_title="D365 Zoho Journal Generator", layout="centered")
st.title("📊 D365 Zoho Journal Generator")
st.write("Upload your daily files below to generate a perfectly balanced D365 import CSV.")

# 1. SIDEBAR SETTINGS (For Accounting Configuration)
st.sidebar.header("⚙️ D365 Account Settings")
bank_account = st.sidebar.text_input("Bank Clearing Account", value="110100")
fee_account = st.sidebar.text_input("Merchant Fee Account", value="615000")
journal_name = st.sidebar.text_input("Journal Name", value="GEN-JRN")

# 2. FILE UPLOADERS
st.subheader("1. Upload Daily Files")
zoho_file = st.file_uploader("Drop today's Zoho Payments Export (CSV)", type=["csv"])
customer_file = st.file_uploader("Drop the Customer Master Account File (CSV/XLSX)", type=["csv", "xlsx"])

def extract_invoice_number(description):
    if pd.isna(description):
        return None
    match = re.search(r'INV-\d+', str(description))
    return match.group(0) if match else None

# 3. PROCESSING LOGIC
if zoho_file and customer_file:
    st.subheader("2. Review & Generate")
    
    try:
        # Load Customer Master
        if customer_file.name.endswith('.csv'):
            cust_df = pd.read_csv(customer_file)
        else:
            cust_df = pd.read_excel(customer_file)
            
        cust_df['Account Name'] = cust_df['Account Name'].astype(str).str.strip().str.lower()
        
        # Load Zoho Data
        zoho_df = pd.read_csv(zoho_file)
        
        journal_lines = []
        line_number = 1
        voucher_counter = 1
        
        for idx, row in zoho_df.iterrows():
            voucher_id = f"VOU-{voucher_counter:03d}"
            raw_desc = row.get('Description', '')
            gross_amt = float(row.get('Gross Amount', 0))
            fee_amt = float(row.get('Merchant Fee', 0))
            net_amt = float(row.get('Net Amount', 0))
            
            inv_num = extract_invoice_number(raw_desc)
            customer_id = "MISSING_CUST_ID"
            cust_name_clean = str(row.get('Customer Name', '')).strip().lower()
            
            match = cust_df[cust_df['Account Name'] == cust_name_clean]
            if not match.empty:
                customer_id = match.iloc[0]['Account']
            elif inv_num:
                raw_desc = f"{raw_desc} (Inv: {inv_num})"
                
            # Line 1: Customer Debit
            journal_lines.append({
                "JournalName": journal_name, "LineNumber": line_number, "Voucher": voucher_id,
                "AccountType": "Customer", "LedgerDimension": customer_id,
                "Description": f"Gross Rec - {raw_desc}", "Debit": gross_amt, "Credit": None, "CurrencyCode": "USD"
            })
            line_number += 1
            
            # Line 2: Merchant Fee Debit
            if fee_amt > 0:
                journal_lines.append({
                    "JournalName": journal_name, "LineNumber": line_number, "Voucher": voucher_id,
                    "AccountType": "Ledger", "LedgerDimension": fee_account,
                    "Description": f"Zoho Card Processing Fee - {inv_num if inv_num else ''}", "Debit": fee_amt, "Credit": None, "CurrencyCode": "USD"
                })
                line_number += 1
                
            # Line 3: Bank Clearing Credit
            journal_lines.append({
                "JournalName": journal_name, "LineNumber": line_number, "Voucher": voucher_id,
                "AccountType": "Ledger", "LedgerDimension": bank_account,
                "Description": "Net Settlement Deposit - Zoho Payments", "Debit": None, "Credit": net_amt, "CurrencyCode": "USD"
            })
            line_number += 1
            voucher_counter += 1
            
        final_df = pd.DataFrame(journal_lines)
        
        # Show data preview on screen
        st.success("✅ Files mapped successfully! Previewing your D365 Journal below:")
        st.dataframe(final_df.head(10))
        
        # 4. DOWNLOAD BUTTON
        csv_data = final_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download D365 Upload CSV",
            data=csv_data,
            file_name="D365_Ready_General_Journal.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"❌ An error occurred while parsing the data: {str(e)}")
else:
    st.info("💡 Please upload both files above to unlock the D365 output generator.")
