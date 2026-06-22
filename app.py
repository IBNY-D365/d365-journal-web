import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re

# =========================================================================
# CONSTANTS & CONFIGURATION
# =========================================================================
D365_COLUMNS = [
    "Date", "Voucher", "Account name", "Company", "Account type", "Account",
    "Posting Profile", "Cash code", "Description", "Debit", "Credit",
    "Item sales tax group", "Sales tax code", "Offset company", "Bank Account Type",
    "Offset account", "Offset transaction text", "Currency", "Exchange rate",
    "Item sales tax group2", "Sales group", "Withholding tax group",
    "Release date", "Reversing entry", "Reversing date"
]

CASH_CODE_MAP = {
    "due-on-receipt": ("AR001", "AR Collection_AP"),
    "monthly": ("AR002", "AR Collection_MPP"),
    "financing": ("AR003", "AR Collection_Financing"),
    "leasing": ("AR004", "AR Collection_Leasing"),
    "net 1 day": ("AR005", "AR Collection_Net_1Day"),
    "net 10 days": ("AR006", "AR Collection_Net_10Days"),
    "net 25 days": ("AR007", "AR Collection_Net_25Days"),
    "net 30 days": ("AR008", "AR Collection_Net_30Days"),
    "net 40 days": ("AR009", "AR Collection_Net_40Days"),
    "net 45 days": ("AR010", "AR Collection_Net_45Days"),
    "net 60 days": ("AR011", "AR Collection_Net_60Days"),
}

MASTER_ACCOUNT_PATH = "Customer Master Account File.xlsx"
FORM_DB_PATH = "Form_Master_DB.xlsx"

# =========================================================================
# APP INTERFACE
# =========================================================================
st.set_page_config(page_title="D365 Zoho Journal Automation", layout="wide")
st.title("📊 D365 General Journal Automation Engine")
st.subheader("Zoho Payments Conversion Portal")

st.sidebar.header("📁 Upload Transaction Data")

boa_file = st.sidebar.file_uploader("1. Bank of America Report (Excel/CSV)", type=["xlsx", "csv"])
zoho_file = st.sidebar.file_uploader("2. Zoho Records File (PDF/CSV/Excel)", type=["pdf", "csv", "xlsx"])

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Master Settings")

master_loaded = False
form_db_loaded = False
df_master = None
df_form_db = None

if os.path.exists(MASTER_ACCOUNT_PATH):
    try:
        df_master = pd.read_excel(MASTER_ACCOUNT_PATH)
        master_loaded = True
        st.sidebar.success(f"✔️ Loaded '{MASTER_ACCOUNT_PATH}' dynamically.")
    except Exception as e:
        st.sidebar.error(f"Error loading local master account file: {e}")

if os.path.exists(FORM_DB_PATH):
    try:
        df_form_db = pd.read_excel(FORM_DB_PATH, sheet_name="Sales_PRF")
        form_db_loaded = True
        st.sidebar.success(f"✔️ Loaded '{FORM_DB_PATH}' (Sales_PRF) dynamically.")
    except Exception as e:
        st.sidebar.error(f"Error loading local Form Master DB: {e}")

if not master_loaded:
    uploaded_master = st.sidebar.file_uploader("Upload Customer Master Account File (Fallback)", type=["xlsx"])
    if uploaded_master:
        df_master = pd.read_excel(uploaded_master)
        master_loaded = True

if not form_db_loaded:
    uploaded_form_db = st.sidebar.file_uploader("Upload Form Master DB File (Fallback)", type=["xlsx"])
    if uploaded_form_db:
        try:
            df_form_db = pd.read_excel(uploaded_form_db, sheet_name="Sales_PRF")
            form_db_loaded = True
        except Exception:
            st.sidebar.error("Could not find 'Sales_PRF' tab in uploaded file.")

# =========================================================================
# CORE PROCESSING
# =========================================================================
if boa_file and zoho_file and master_loaded and form_db_loaded:
    try:
        # Robust loading of Bank of America File
