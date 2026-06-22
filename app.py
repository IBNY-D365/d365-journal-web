import streamlit as str_lit
import pandas as pd
import numpy as np
import io
import re

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

def hyper_robust_line_split(line):
    """Safely splits CSV lines by commas while ignoring commas wrapped in quotation marks."""
    return re.findall(r'(?:[^,"]|"(?:\\.|[^"])*")+', line)

def read_boa_to_clean_dict(file_io):
    """
    Completely replaces pandas indexing loops. Scans lines natively and builds
    clean lookup dictionaries based on verified word footprints.
    """
    raw_content = file_io.read().decode('utf-8-sig', errors='ignore')
    file_io.seek(0)
    lines = [line.strip() for line in raw_content.split('\n') if line.strip()]
    
    header_idx = -1
    headers = []
    for idx, line in enumerate(lines):
        up_line = line.upper()
        if "DESC" in up_line or "AMOUNT" in up_line or "POSTING" in up_line:
            headers = [h.strip().replace('"', '') for h in hyper_robust_line_split(line)]
            header_idx = idx
            break
            
    if header_idx == -1:
        return []
        
    structured_data = []
    for line in lines[header_idx + 1:]:
        parts = [p.strip().replace('"', '') for p in hyper_robust_line_split(line)]
        if len(parts) < len(headers):
            continue
        row_map = {headers[i].upper(): parts[i] for i in range(len(headers))}
        structured_data.append(row_map)
        
    return structured_data

# ==========================================
# STREAMLIT UI SETUP
# ==========================================
str_
