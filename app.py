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
    if "3371" in acc_str: 
        return "B1000002"
    elif "3924" in acc_str: 
        return "B1000003"
    elif "3384" in acc_str: 
        return "B1000001"
    return "B1000002" 

def create_base_row():
    """Generates a base row formatted with required constant strings and values."""
    row = {col: "" for col in D365_COLUMNS}
    row["Company"] = "bwa"
    row["Offset company"] = "bwa"
    row["Bank Account Type"] = "Bank"
    row["Currency"] = "USD"
    row["Exchange rate"] = 1.00
    row["Sales tax group"] = "AVATAX"
    row["Reversing entry"] = "No"
    return row

def hyper_robust_boa_read(file_io):
    """
    Completely eliminates string header blocking. Reads the file stream natively,
    cleans up hidden carriage characters, and applies a multi-tiered header finding matrix.
    """
    raw_bytes = file_io.read()
    file_io.seek(0)
    
    # Clean out sneaky Byte Order Marks (BOM) and carriage returns natively
    text_content = raw_bytes.decode('utf-8-sig', errors='ignore').replace('\r', '')
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    
    # Fallback Tier 1: Look for any line containing DESC or DATE or AMOUNT
    start_row_idx = 0
    for idx, line in enumerate(lines):
        up_line = line.upper()
        if "DESC" in up_line or "AMOUNT" in up_line or "POSTING" in up_line:
            start_row_idx = idx
            break
            
    # Compile a clean CSV string block from the true data grid start position
    clean_csv_block = "\n".join(lines[start_row_idx:])
    
    # Read using python engine with adaptive delimiter sniffer fallbacks
    try:
        df = pd.read_csv(io.StringIO(clean_csv_block), sep=None, engine='python', on_bad_lines='skip')
    except Exception:
        df = pd.read_csv(io.StringIO(clean_csv_block), sep=',', engine='python', on_bad_lines='skip')
        
    # Standardize remaining headers completely to upper case to remove matching friction
    df.columns = df.columns.str.strip().str.upper()
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
    gateway_file = str_lit.file_uploader("""1. Upload Zoho / Stripe / Bankcard file
