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
    """Conditional routing based on source Bank of America transaction account number."""
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

def bypass_and_read_boa(file_io):
    """
    Reads the file stream cleanly, handles carriage return layout anomalies, 
    and returns a standardized column dataframe with stripped column headers.
    """
    raw_bytes = file_io.read()
    file_io.seek(0)
    
    text_content = raw_bytes.decode('utf-8-sig', errors='ignore').replace('\r', '')
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    
    start_idx = 0
    for idx, line in enumerate(lines):
        up_line = line.upper()
        if "DESC" in up_line or "AMOUNT" in up_line or "POSTING" in up_line:
            start_idx = idx
            break
            
    clean_block = "\n".join(lines[start_idx:])
    df = pd.read_csv(io.StringIO(clean_block), sep=',', engine='python', on_bad_lines='skip')
    df.columns = df.columns.str.strip()
    return df

# ==========================================
# STREAMLIT UI SETUP
# ==========================================
str_lit.set_page_config(page_title="D365 Transaction Journal Generator", layout="wide")

str_lit.sidebar.header("D365 Defaults")
default_company = str_lit.sidebar.text_input("Company", value="bwa")
default_offset = str_lit.sidebar.text_input("Default Offset Account", value="B1000002")
default_debit_ledger = str_lit.sidebar.text_input("Debit Line Account (Ledger)", value="43170111-U26C05001-B735350-UOA003")

str_lit.title("""D365 Transaction Journal Generator""")
str_lit.subheader("""Upload your Bank of America statement plus any gateway/invoice files for the day.""")

col1, col2, col3 = str_lit.columns(3)

with col1:
    gateway_file = str_lit.file_uploader("""1. Upload Zoho / Stripe / Bankcard file (PDF, CSV, XLSX)""", type=["pdf", "csv", "xlsx"])
with col2:
    invoice_files = str_lit.file_uploader("""2. Upload invoice files (PDF, CSV, XLSX, TXT)""", accept_multiple_files=True)
with col3:
    boa_statement = str_lit.file_uploader("""3. Upload Bank of America statement (CSV, XLSX)""", type=["csv", "xlsx"])

str_lit.info("""Upload the BOA statement to begin. Gateway and invoice files are optional depending on the day.""")

# ==========================================
# PROCESSING ENGINE (DETERMINISTIC PIPELINE)
# ==========================================
if boa_statement is not None:
    df_boa = bypass_and_read_boa(boa_statement)
    output_rows = []
    
    col_names = list(df_boa.columns)
    desc_col = next((c for c in col_names if "DESC" in c.upper()), None)
    amt_col = next((c for c in col_names if "AMT" in c.upper() or "AMOUNT" in c.upper()), None)
    date_col = next((c for c in col_names if "DATE" in c.upper()), None)
    src_col = next((c for c in col_names if "ACC" in c.upper() or "SOURCE" in c.upper()), None)
    
    if desc_col is None and len(col_names) > 1: desc_col = col_names[1]
    if amt_col is None and len(col_names) > 2: amt_col = col_names[2]
    if date_col is None and len(col_names) > 0: date_col = col_names[0]
    if src_col is None and len(col_names) > 3: src_col = col_names[3]

    if desc_col and amt_col:
        for idx, boa_row in df_boa.iterrows():
            boa_desc = str(boa_row.get(desc_col, '')).upper().strip()
            
            # Filter out systemic summary macro rows natively
            if any(ignored in boa_desc for ignored in ["BEGINNING BALANCE", "TOTAL CREDITS", "TOTAL DEBITS", "ENDING BALANCE"]):
                continue
                
            if not boa_desc or boa_desc == 'NAN':
