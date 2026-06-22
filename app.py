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
    if "3371" in acc_str: 
        return "B1000002"
    elif "3924" in acc_str: 
        return "B1000003"
    elif "3384" in acc_str: 
        return "B1000001"
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
    df = pd.read_csv(
        file_io, 
        skiprows=skip_rows, 
        engine='python', 
        on_bad_lines='skip'
    )
    
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

str_lit.title("D
