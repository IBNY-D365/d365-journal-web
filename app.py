import streamlit as str_lit
import pandas as pd
import numpy as np
import io
import pypdf

# ==========================================
# 1. CORE ARCHITECTURAL CONSTANTS
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

def extract_text_from_pdf(uploaded_file):
    """Extracts raw text blocks from uploaded invoice or gateway PDFs."""
    if uploaded_file is None:
        return ""
    try:
        pdf_reader = pypdf.PdfReader(uploaded_file)
        text_content = []
        for page in pdf_reader.pages:
            text_content.append(page.extract_text() or "")
        return "\n".join(text_content)
    except Exception:
        return ""

def hyper_robust_boa_read(file_io):
    """Bypasses bank layout formatting and safely reads rows based on index positions."""
    raw_bytes = file_io.read()
    file_io.seek(0)
    
    text_content = raw_bytes.decode('utf-8-sig', errors='ignore').replace('\r', '')
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    
    start_row_idx = 0
    for idx, line in enumerate(lines):
        up_line = line.upper()
        if "DESC" in up_line or "AMOUNT" in up_line or "POSTING" in up_line:
            start_row_idx = idx
            break
            
    clean_csv_block = "\n".join(lines[start_row_idx:])
    df = pd.read_csv(io.StringIO(clean_csv_block), sep=',', engine='python', on_bad_lines='skip')
    df.columns = df.columns.str.strip().str.upper()
    return df

# ==========================================
# 2. STREAMLIT UI LAYOUT
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

# Load reference files from local path natively
cust_master, form_master, monthly_exp = None, None, None
try:
    cust_master = pd.read_excel("Customer Master Account File.xlsx")
    form_master = pd.read_excel("Form_Master_DB.xlsx", sheet_name="Sales_PRF")
    monthly_exp = pd.read_excel("Monthly Expense Record.xlsx")
except Exception:
    pass

# ==========================================
# 3. AUTOMATION ROUTING ENGINE
# ==========================================
if boa_statement is not None:
    df_boa = hyper_robust_boa_read(boa_statement)
    output_rows = []
    
    # Extract structural text strings out of cross-referenced uploaded PDFs
    gateway_text = extract_text_from_pdf(gateway_file)
    
    desc_key = next((k for k in df_boa.columns if "DESC" in k), None)
    amt_key = next((k for k in df_boa.columns if "AMT" in k or "AMOUNT" in k or "DEBIT" in k or "CREDIT" in k), None)
    date_key = next((k for k in df_boa.columns if "DATE" in k), None
