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
    """Conditional routing based on source Bank of America transaction account number[cite: 35, 38]."""
    acc_str = str(boa_source_acc).strip()
    if "3371" in acc_str: 
        return "B1000002"
    elif "3924" in acc_str: 
        return "B1000003"
    elif "3384" in acc_str: 
        return "B1000001"
    return "B1000002" 

def create_base_row():
    """Generates an empty D365 row pre-populated with standard static configurations[cite: 35, 38]."""
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
    and returns a standardized column dataframe[cite: 7].
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
    
    # Strip whitespace formatting out of column text fields natively
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
            # Crucial Fix: Standardize description tracking to uppercase to eliminate structural drop exceptions
            boa_desc = str(boa_row.get(desc_col, '')).upper().strip()
            
            if any(ignored in boa_desc for ignored in ["BEGINNING BALANCE", "TOTAL CREDITS", "TOTAL DEBITS", "ENDING BALANCE"]):
                continue
                
            if not boa_desc or boa_desc == 'NAN':
                continue
                
            try:
                raw_val = str(boa_row.get(amt_col, '0')).replace('$', '').replace(',', '').strip()
                boa_amt = float(raw_val)
            except ValueError:
                boa_amt = 0.0
                
            boa_date = str(boa_row.get(date_col, '')) if date_col else ''
            boa_source_acc = str(boa_row.get(src_col, '3371')) if src_col else '3371'
            
            offset_acc = get_offset_account(boa_source_acc)
            
            # ----------------------------------------------------
            # ROUTE 1: ZOHO PAYMENTS
            # ----------------------------------------------------
            if "ZOHO" in boa_desc:
                gross_amt = boa_amt * 1.03 
                fee_amt = gross_amt - boa_amt
                
                # Credit Line [cite: 35]
                c_row = create_base_row()
                c_row["Date"] = boa_date
                c_row["Account type"] = "Customer"
                c_row["Account name"] = "Page Fit Inc. DBA Intoxx Fitness" [cite: 36]
                c_row["Account"] = "BC000571" [cite: 36]
                c_row["Posting Profile"] = "AutoPost"
                c_row["Cash code"] = "AR001" [cite: 32]
                c_row["Description"] = f"BC000571 Page Fit Inc. DBA Intoxx Fitness_ZOHO PAYMENTS DES:{boa_desc}" [cite: 36]
                c_row["Credit"] = gross_amt
                c_row["Offset account"] = offset_acc
                output_rows.append(c_row)
                
                # Debit Fee Line [cite: 38]
                d_row = create_base_row()
                d_row["Date"] = boa_date
                d_row["Account name"] = "Outside Service (Finance)"
                d_row["Account type"] = "Ledger"
                d_row["Account"] = default_debit_ledger
                d_row["Cash code"] = "OSF005"
                d_row["Description"] = f"Zoho Merchant Fee BC000571 Page Fit Inc. DBA Intoxx Fitness_ZOHO PAYMENTS DES:{boa_desc}" [cite: 39]
                d_row["Debit"] = fee_amt
                d_row["Offset account"] = offset_acc
                output_rows.append(d_row)

            # ----------------------------------------------------
            # ROUTE 2: STRIPE PAYMENTS
            # ----------------------------------------------------
            elif "STRIPE" in boa_desc:
                gross_amt = boa_amt * 1.025
                fee_amt = gross_amt - boa_amt
                
                # Credit Line [cite: 74]
                c_row = create_base_row()
                c_row["Date"] = boa_date
                c_row["Account type"] = "Customer"
                c_row["Account name"] = "Elite Functional Wellness" [cite: 75]
                c_row["Account"] = "BC000327" [cite: 75]
                c_row["Posting Profile"] = "AutoPost"
                c_row["Cash code"] = "AR001" [cite: 71]
                c_row["Description"] = f"BC000327 Elite Functional Wellness_STRIPE DES:{boa_desc}" [cite: 75]
                c_row["Credit"] = gross_amt
                c_row["Offset account"] = offset_acc
                output_rows.append(c_row)
                
                # Debit Fee Line [cite: 77]
                d_row = create_base_row()
                d_row["Date"] = boa_date
                d_row["Account name"] = "Outside Service (Finance)"
                d_row["Account type"] = "Ledger"
                d_row["Account"] = default_debit_ledger
                d_row["Cash code"] = "OSF006"
                d_row["Description"] = f"Stripe Merchant Fee BC000327 Elite Functional Wellness_STRIPE DES:{boa_desc}" [cite: 78]
                d_row["Debit"] = fee_amt
                d_row["Offset account"] = offset_acc
                output_rows.append(d_row)

            # ----------------------------------------------------
            # ROUTE 3: BANKCARD PAYMENTS
            # ----------------------------------------------------
            elif "BANKCARD" in boa_desc:
                gross_amt = boa_amt * 1.035
                fee_amt = gross_amt - boa_amt
                
                # Credit Line [cite: 172]
                c_row = create_base_row()
                c_row["Date"] = boa_date
                c_row["Account type"] = "Customer"
                c_row["Account name"] = "Bankcard Customer Normalized"
                c_row["Account"] = "BC000422" [cite: 176]
                c_row["Posting Profile"] = "AutoPost"
                c_row["Cash code"] = "AR001" [cite: 169]
                c_row["Description"] = f"BC000422 Bankcard Customer_BANKCARD DES:{boa_desc}" [cite: 173]
                c_row["Credit"] = gross_amt
                c_row["Offset account"] = offset_acc
                output_rows.append(c_row)
                
                # Debit Fee Line [cite: 175]
                d_row = create_base_row()
                d_row["Date"] = boa_date
                d_row["Account name"] = "Outside Service (Finance)"
                d_row["Account type"] = "Ledger"
                d_row["Account"] = default_debit_ledger
                d_row["Cash code"] = "OSF007"
                d_row["Description"] = f"Authorization.net Merchant Fee BC000422 Bankcard Customer_BANKCARD DES:{boa_desc}" [cite: 176]
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
                    m_row["Cash code"] = "OSD002" [cite: 119]
                    m_row["Account type"] = "Vendor" [cite: 130]
                    m_row["Account"] = "BV000130" [cite: 133]
                    m_row["Description"] = f"Marketing Subscriptions_{boa_row.get(desc_col, '')}"
                elif "ADOBE INC" in boa_desc and abs(boa_amt) == 21.19:
                    m_row["Cash code"] = "OSD005" [cite: 121]
                    m_row["Account type"] = "Ledger" [cite: 130]
                    m_row["Account"] = "43170116-U26C06000-B735349" [cite: 133]
                    m_row["Description"] = f"Common Subscription_{boa_row.get(desc_col, '')}"
                elif "KIM LEE LLP" in boa_desc:
                    m_row["Cash code"] = "OSF008" [cite: 133]
                    m_row["Account type"] = "Ledger"
                    m_row["Account"] = "43170111-U26C05001-B735350-UOS003" [cite: 133]
                    m_row["Description"] = f"CPA Fee_Monthly retainer fee_{boa_row.get(desc_col, '')}" [cite: 137]
                else:
                    m_row["Cash code"] = "OSD001" [cite: 133]
                    m_row["Account type"] = "Ledger"
                    m_row["Account"] = "43170113-U26C00000-B000000-UIT001" [cite: 133]
                    m_row["Description"] = f"Outside Service(Due&Subs)_{boa_row.get(desc_col, '')}"
                    
                output_rows.append(m_row)

            # ----------------------------------------------------
            # ROUTE 5: FALLBACK TRACK (Manual Accounting Allocation)
            # ----------------------------------------------------
            else:
                f_row = create_base_row()
                f_row["Date"] = boa_date
                
                # Fields left completely empty to match Fallback Rules [cite: 96, 101]
                f_row["Account name"] = "" [cite: 96, 101]
                f_row["Account type"] = "" [cite: 96, 101]
                f_row["Account"] = "" [cite: 96, 101]
                f_row["Posting Profile"] = "" [cite: 101]
                f_row["Cash code"] = "" [cite: 101]
                
                # Exact literal string mapping [cite: 98, 101]
                f_row["Description"] = boa_row.get(desc_col, '') [cite: 98, 101]
                f_row["Debit"] = abs(boa_amt) [cite: 99, 101]
                f_row["Offset account"] = offset_acc
                output_rows.append(f_row)

    if output_rows:
        df_result = pd.DataFrame(output_rows, columns=D365_COLUMNS)
        str_lit.success("""Successfully processed and mapped transaction data entries!""")
        str_lit.dataframe(df_result)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='D365_Upload_Journal')
        
        str_lit.download_button(
            label="Download Generation Sheet (XLSX)",
            data=buffer.getvalue(),
            file_name="D365_Automated_General_Journal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        str_lit.warning("""Data extracted, but no matching non-empty transaction rows were found after processing.""")
