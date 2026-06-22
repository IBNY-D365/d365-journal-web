import streamlit as str_lit
import pandas as pd
import numpy as np
import io
import csv

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

def fully_universal_read_boa(file_io):
    """
    Decodes the uploaded file stream cleanly, handles dialet/delimiter sniffing automatically,
    bypasses metadata summaries, and extracts transactions via standard dictionary structures.
    """
    # Read and decode lines while cleaning up byte anomalies
    raw_content = file_io.read().decode('utf-8', errors='ignore')
    file_io.seek(0)
    
    raw_lines = [line.strip() for line in raw_content.splitlines() if line.strip()]
    
    # Locate where the actual transaction layout table grid starts
    header_line_idx = -1
    for idx, line in enumerate(raw_lines):
        up_line = line.upper()
        if "DESC" in up_line and ("AMT" in up_line or "AMOUNT" in up_line or "DEBIT" in up_line or "CREDIT" in up_line):
            header_line_idx = idx
            break
            
    # Fallback to zero if layout cannot be sniffed explicitly
    if header_line_idx == -1:
        header_line_idx = 0
        
    data_content = "\n".join(raw_lines[header_line_idx:])
    
    # Sniff dialect to cleanly handle comma vs tab vs semicolon variations
    try:
        dialect = csv.Sniffer().sniff(raw_lines[header_line_idx])
        delimiter = dialect.delimiter
    except Exception:
        delimiter = "," # Safe default standard fallback
        
    # Read via native Python CSV engine to keep data structure safe
    reader = csv.Dict
