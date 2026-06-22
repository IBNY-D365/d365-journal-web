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

def custom_parse_boa_csv(file_io):
    """
    Direct string block scanner that completely bypasses pandas column dependency issues.
    Extracts date, description, and amount natively using pure positional data streaming.
    """
    raw_bytes = file_io.read()
    file_io.seek(0)
    text = raw_bytes.decode('utf-8', errors='ignore')
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    parsed_records = []
    headers = []
    header_idx = -1
    
    # 1. Locate data grid entry point line-by-line
    for idx, line in enumerate(lines):
        clean_line = line.upper()
        if "DESCRIPTION" in clean_line and ("AMOUNT" in clean_line or "DEBIT" in clean_line or "CREDIT" in clean_line):
            headers = [h.strip().replace('"', '') for h in line.split(',')]
            header_idx = idx
            break
            
    if header_idx == -1:
        # Fallback if statement lacks formal header grid line
        headers = ["DATE", "DESCRIPTION", "AMOUNT", "SOURCE ACCOUNT"]
        header_idx = -1 
        
    # 2. Process data lines sequentially below header row
    for line in lines[header_idx + 1:]:
        # Handle commas inside quoted description text lines safely
        parts = []
        current_part = []
        in_quotes = False
        for char in line:
            if char == '"':
                in_quotes = not in_quotes
            elif char == ',' and not in_quotes:
                parts.append("".join(current_part).strip().replace('"', ''))
                current_part = []
            else:
                current_part.append(char)
        parts.append("".join(current_part).strip().replace('"', ''))
        
        if len(parts) < 2:
            continue
            
        # Map indices dynamically based on found positions
        row_dict = {}
        for h_i, h_name in enumerate(headers):
            if h_i < len(parts):
                row_dict[h_name.upper()] = parts[h_i]
                
        parsed_records.append(row_dict)
        
    return pd.DataFrame(parsed_records)

# ==========================================
# STREAMLIT UI SETUP
# ==========================================
str_lit.set_page_config(page_title="D365 Transaction Journal Generator", layout="wide")

str_lit.sidebar.header("D365 Defaults")
default_company = str_lit.sidebar.text_input("Company", value="bwa")
default_offset = str_lit.sidebar.text_input("Default Offset Account", value="B1000002")
default_debit
