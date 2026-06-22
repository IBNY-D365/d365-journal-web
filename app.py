import streamlit as str_lit
import pandas as pd
import numpy as np
import io

# ==========================================
# CONSTANTS & CONFIGURATION
# ==========================================
D365_COLUMNS = [
    "Date", "Voucher", "Account name", "Company", "Account type", "Account",
    "Posting Profile", "Cash code", "Description", "Debit", "Credit",
    "Item sales tax group", "Sales tax code", "Offset company", "Bank Account Type",
    "Offset account", "Offset transaction text", "Currency", "Exchange rate",
    "Item sales tax group2", "Sales tax group", "Withholding tax group",
    "Release date", "Reversing entry", "Reversing date"
]

def get_offset_account(boa_source_acc):
    """Conditional routing based on source BOA transaction account number[cite: 35, 74, 101, 136, 172, 175]."""
    acc_str = str(boa_source_acc).strip()
    if "3371" in acc_str: return "B1000002"
    elif "3924" in acc_str: return "B1000003"
    elif "3384" in acc_str: return "B1000001"
    return "B1000002" 

def create_base_row():
    """Generates an empty D365 row pre-populated with standard static configurations[cite: 35, 74, 101, 136, 172]."""
    row = {col: "" for col in D365_COLUMNS}
    row["Company"] = "bwa"
    row["Offset company"] = "bwa"
    row["Bank Account Type"] = "Bank"
    row["Currency"] = "USD"
    row["Exchange rate"] = 1.00
    row["Sales tax group"] = "AVATAX"
    row["Reversing entry"] = "No"
    return row

# ==========================================
# STREAMLIT UI SETUP
# ==========================================
str_lit.set_page_config(page_title="D365 Transaction Journal Generator", layout="wide")

# Sidebar Layout
str_lit.sidebar.header("D365 Defaults")
default_company = str_lit.sidebar.text_input("Company", value="bwa")
default_offset = str_lit.sidebar.text_input("Default Offset Account", value="B1000002")
default_debit_ledger = str_lit.sidebar.text_input("Debit Line Account (Ledger)", value="43170111-U26C05001-B735350-UOA003")

# Main Page Elements (Kept safe as unified strings to prevent truncation errors)
str_lit.title("D365 Transaction Journal Generator")
str_lit.subheader("Upload your Bank of America statement plus any gateway/invoice files for the day.")

col1, col2, col3 = str_lit.columns(3)

with col1:
    gateway_file = str_lit.file_uploader("1. Upload Zoho / Stripe / Bankcard file (PDF, CSV, XLSX)", type=["pdf", "csv", "xlsx"])
with col2:
    invoice_files = str_lit.file_uploader("2. Upload invoice files (PDF, CSV, XLSX, TXT)", accept_multiple_files=True)
with col3:
    boa_statement = str_lit.file_uploader("3. Upload Bank of America statement (CSV, XLSX)", type=["csv", "xlsx"])

str_lit.info("Upload the BOA statement to begin. Gateway and invoice files are optional depending on the day.")

# Caching master lookup tables
@str_lit.cache_data
def load_master_files():
    try:
        cust_master = pd.read_excel("Customer Master Account File.xlsx")
        form_master = pd.read_excel("Form_Master_DB.xlsx", sheet_name="Sales_PRF")
        monthly_exp = pd.read_excel("Monthly Expense Record.xlsx")
        return cust_master, form_master, monthly_exp
    except Exception:
        return None, None, None

cust_master, form_master, monthly_exp = load_master_files()

# ==========================================
# PROCESSING ENGINE (DETERMINISTIC PIPELINE)
# ==========================================
if boa_statement is not None:
    if boa_statement.name.endswith('.csv'):
        df_boa = pd.read_csv(boa_statement)
    else:
        df_boa = pd.read_excel(boa_statement)
        
    output_rows = []
    
    for idx, boa_row in df_boa.iterrows():
        boa_desc = str(boa_row.get('Description', '')).upper()
        boa_amt = float(boa_row.get('Amount', 0))
        boa_date = boa_row.get('Date', '')
        boa_source_acc = boa_row.get('Source Account', '3371')
        
        offset_acc = get_offset_account(boa_source_acc)
        
        # ----------------------------------------------------
        # ROUTE 1: ZOHO TRANSACTIONS
        # ----------------------------------------------------
        if "ZOHO" in boa_desc:
            gross_amt = boa_amt * 1.03 
            fee_amt = gross_amt - boa_amt
            
            # Credit Line
            c_row = create_base_row()
            c_row["Date"] = boa_date
            c_row["Account type"] = "Customer"
            c_row["Account name"] = "Zoho Customer Normalized"
            c_row["Account"] = "BC000571"
            c_row["Posting Profile"] = "AutoPost"
            c_row["Cash code"] = "AR001"
            c_row["Description"] = f"BC000571 Zoho Customer_ZOHO PAYMENTS DES:{boa_desc}"
            c_row["Credit"] = gross_amt
            c_row["Offset account"] = offset_acc
            output_rows.append(c_row)
            
            # Debit Fee Line
            d_row = create_base_row()
            d_row["Date"] = boa_date
            d_row["Account name"] = "Outside Service (Finance)"
            d_row["Account type"] = "Ledger"
            d_row["Account"] = default_debit_ledger
            d_row["Cash code"] = "OSF005"
            d_row["Description"] = f"Zoho Merchant Fee BC000571 Zoho Customer_ZOHO PAYMENTS DES:{boa_desc}"
            d_row["Debit"] = fee_amt
            d_row["Offset account"] = offset_acc
            output_rows.append(d_row)

        # ----------------------------------------------------
        # ROUTE 2: STRIPE TRANSACTIONS
        # ----------------------------------------------------
        elif "STRIPE" in boa_desc:
            gross_amt = boa_amt * 1.025
            fee_amt = gross_amt - boa_amt
            
            # Credit Line
            c_row = create_base_row()
            c_row["Date"] = boa_date
            c_row["Account type"] = "Customer"
            c_row["Account name"] = "Stripe Customer Normalized"
            c_row["Account"] = "BC000327"
            c_row["Posting Profile"] = "AutoPost"
            c_row["Cash code"] = "AR001"
            c_row["Description"] = f"BC000327 Stripe Customer_STRIPE DES:{boa_desc}"
            c_row["Credit"] = gross_amt
            c_row["Offset account"] = offset_acc
            output_rows.append(c_row)
            
            # Debit Fee Line
            d_row = create_base_row()
            d_row["Date"] = boa_date
            d_row["Account name"] = "Outside Service (Finance)"
            d_row["Account type"] = "Ledger"
            d_row["Account"] = default_debit_ledger
            d_row["Cash code"] = "OSF006"
            d_row["Description"] = f"Stripe Merchant Fee BC000327 Stripe Customer_STRIPE DES:{boa_desc}"
            d_row["Debit"] = fee_amt
            d_row["Offset account"] = offset_acc
            output_rows.append(d_row)

        # ----------------------------------------------------
        # ROUTE 3: BANKCARD TRANSACTIONS (AUTHORIZE.NET)
        # ----------------------------------------------------
        elif "BANKCARD" in boa_desc:
            gross_amt = boa_amt * 1.035
            fee_amt = gross_amt - boa_amt
            
            # Credit Line
            c_row = create_base_row()
            c_row["Date"] = boa_date
            c_row["Account type"] = "Customer"
            c_row["Account name"] = "Bankcard Customer Normalized"
            c_row["Account"] = "BC000422"
            c_row["Posting Profile"] = "AutoPost"
            c_row["Cash code"] = "AR001"
            c_row["Description"] = f"BC000422 Bankcard Customer_BANKCARD DES:{boa_desc}"
            c_row["Credit"] = gross_amt
            c_row["Offset account"] = offset_acc
            output_rows.append(c_row)
            
            # Debit Fee Line
            d_row = create_base_row()
            d_row["Date"] = boa_date
            d_row["Account name"] = "Outside Service (Finance)"
            d_row["Account type"] = "Ledger"
            d_row["Account"] = default_debit_ledger
            d_row["Cash code"] = "OSF007"
            d_row["Description"] = f"Authorization.net Merchant Fee BC000422 Bankcard Customer_BANKCARD DES:{boa_desc}"
            d_row["Debit"] = fee_amt
            d_row["Offset account"] = offset_acc
            output_rows.append(d_row)

        # ----------------------------------------------------
        # ROUTE 4: MONTHLY RECURRING TRACK
        # ----------------------------------------------------
        elif any(trigger in boa_desc for trigger in ["ADOBE INC", "GENESIS", "HMFUSA.COM", "MICROSOFT", "RAMP", "KIM LEE LLP"]):
            m_row = create_base_row()
            m_row["Date"] = boa_date
            m_row["Debit"] = abs(boa_amt)
            m_row["Offset account"] = offset_acc
            
            if "ADOBE INC" in boa_desc and abs(boa_amt) == 826.67:
                m_row["Cash code"] = "OSD002"
                m_row["Account type"] = "Vendor"
                m_row["Account"] = "BV000130"
                m_row["Description"] = f"Marketing Subscriptions_{boa_row.get('Description', '')}"
            elif "ADOBE INC" in boa_desc and abs(boa_amt) == 21.19:
                m_row["Cash code"] = "OSD005"
                m_row["Account type"] = "Ledger"
                m_row["Account"] = "43170116-U26C06000-B735349"
                m_row["Description"] = f"Common Subscription_{boa_row.get('Description', '')}"
            elif "KIM LEE LLP" in boa_desc:
                m_row["Cash code"] = "OSF008"
                m_row["Account type"] = "Ledger"
                m_row["Account"] = "43170111-U26C05001-B735350-UOS003"
                m_row["Description"] = f"CPA Fee_Monthly retainer fee_{boa_row.get('Description', '')}"
            else:
                m_row["Cash code"] = "OSD001"
                m_row["Account type"] = "Ledger"
                m_row["Account"] = "43170113-U26C00000-B000000-UIT001"
                m_row["Description"] = f"Outside Service(Due&Subs)_{boa_row.get('Description', '')}"
                
            output_rows.append(m_row)

        # ----------------------------------------------------
        # ROUTE 5: FALLBACK NON-MONTHLY TRACK
        # ----------------------------------------------------
        else:
            f_row = create_base_row()
            f_row["Date"] = boa_date
            
            f_row["Account name"] = ""
            f_row["Account type"] = ""
            f_row["Account"] = ""
            f_row["Posting Profile"] = ""
            f_row["Cash code"] = ""
            
            f_row["Description"] = boa_row.get('Description', '')
            f_row["Debit"] = abs(boa_amt)
            f_row["Offset account"] = offset_acc
            output_rows.append(f_row)

    df_result = pd.DataFrame(output_rows, columns=D365_COLUMNS)
    
    str_lit.success("Successfully processed and mapped transaction data entries!")
    str_lit.dataframe(df_result)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_result.to_excel(writer, index=False, sheet_name='D365_Upload_Journal')
    
    str_lit.download_button(
        label="Download Generation Sheet (XLSX)",
        data=buffer.getvalue(),
        file_name="D365_Automated_General_Journal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
