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
    """Conditional routing based on source BOA transaction account number."""
    acc_str = str(boa_source_acc).strip()
    if "3371" in acc_str: return "B1000002"
    elif "3924" in acc_str: return "B1000003"
    elif "3384" in acc_str: return "B1000001"
    return "B1000002" 

def create_base_row():
    """Generates an empty D365 row pre-populated with standard static configurations."""
    row = {col: "" for col in D365_COLUMNS}
    row["Company"] = "bwa"
    row["Offset company"] = "bwa"
    row["Bank Account Type"] = "Bank"
    row["Currency"] = "USD"
    row["Exchange rate"] = 1.00
    row["Sales tax group"] = "AVATAX"
    row["Reversing entry"] = "No"
    return row

def robust_read_boa_csv(file_io):
    """
    Robustly scans BOA statement exports to skip metadata headers,
    standardizes columns case-insensitively, and handles jagged lines safely.
    """
    lines = [line.decode('utf-8', errors='ignore') for line in file_io.readlines()]
    file_io.seek(0)
    
    skip_rows = 0
    for idx, line in enumerate(lines):
        upper_line = line.upper()
        if "DESCRIPTION" in upper_line or "AMOUNT" in upper_line:
            skip_rows = idx
            break
            
    file_io.seek(0)
    # Added engine='python' and on_bad_lines='skip' to absorb erratic CSV lines seamlessly
    df = pd.read_csv(
        file_io, 
        skiprows=skip_rows, 
        engine='python', 
        on_bad_lines='skip'
    )
    
    # Standardize column headers to avoid string matching case/whitespace mismatches
    df.columns = df.columns.str.strip().str.title()
    return df

# ==========================================
# STREAMLIT UI SETUP
# ==========================================
str_lit.set_page_config(page_title="D365 Transaction Journal Generator", layout="wide")

str_lit.sidebar.header("D365 Defaults")
default_company = str_lit.sidebar.text_input("Company", value="bwa")
default_offset = str_lit.sidebar.text_input("Default Offset Account", value="B1000002")
default_debit_ledger = str_lit.sidebar.text_input("Debit Line Account (Ledger)", value="43170111-U26C05001-B735350-UOA003")

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
        df_boa = robust_read_boa_csv(boa_statement)
    else:
        df_boa = pd.read_excel(boa_statement)
        df_boa.columns = df_boa.columns.str.strip().str.title()
        
    output
