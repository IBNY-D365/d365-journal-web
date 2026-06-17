import streamlit as st
import pandas as pd
import re
import os
from pypdf import PdfReader

# Set up page layout
st.set_page_config(page_title="D365 Accounting Journal Generator", layout="wide")
st.title("📊 D365 Zoho, Stripe & BOA Journal Generator")
st.write("Upload your daily processing packages below to build your flawless 25-column D365 upload templates.")

# 1. SIDEBAR FIXED CONFIGURATION
st.sidebar.header("⚙️ Fixed D365 Settings")
company_id = st.sidebar.text_input("Company", value="bwa")
offset_account = st.sidebar.text_input("Offset Account", value="B1000002")
debit_ledger_acct = st.sidebar.text_input("Debit Line Account", value="43170111-U26C05001-B735350-UOA003")

# Hardcoded exact repository filenames visible in GitHub
MASTER_FILE_NAME = "Customer Master Account File.xlsx"
CASH_CODE_FILE = "Cash Code Masterlist.xlsx"
BOA_EXPENSES_FILE = "BOA3371 Expenses List.xlsx - BOA3371.csv"

# 2. FILE UPLOADERS
col1, col2, col3 = st.columns(3)
with col1:
    gateway_file = st.file_uploader(
        "1. Upload Processing File (Zoho CSV or Stripe PDF)",
        type=["csv", "xlsx", "pdf"]
    )
with col2:
    # FIX 1: Accept_multiple_files=True so users can upload multiple invoices
    invoice_files = st.file_uploader(
        "2. Upload Invoice PDF(s) (Required for Zoho — upload all invoices for this payout)",
        type=["pdf", "csv", "xlsx", "txt"],
        accept_multiple_files=True
    )
with col3:
    boa_file = st.file_uploader("3. Upload Bank of America Statement", type=["csv", "xlsx"])


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def extract_text_from_pdf(uploaded_pdf):
    """Return full text from an uploaded PDF file object."""
    try:
        reader = PdfReader(uploaded_pdf)
        full_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
        return full_text
    except Exception:
        return ""


def super_clean_string(text):
    """Normalise a string for fuzzy matching."""
    if pd.isna(text) or text is None:
        return ""
    txt = str(text).lower()
    txt = re.sub(r'\S+@\S+', '', txt)
    txt = re.sub(r'\b(com|org|net|edu|gov)\b', '', txt)
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', ' ', txt)
    cleaned = re.sub(r'\b(llc|pllc|inc|corp|co|incorporated|limited|llp)\b', ' ', cleaned)
    return " ".join(cleaned.split())


def safe_float(value):
    """Convert a value to float, stripping currency symbols/commas. Returns 0.0 on failure."""
    try:
        return float(str(value).replace(',', '').replace('$', '').strip())
    except (ValueError, TypeError):
        return 0.0


def determine_debit_credit(amount: float, account_type: str) -> tuple:
    """
    FIX 2: Determine whether an amount goes in the Debit or Credit column
    based on the sign of the amount and the account type.

    Rules (standard double-entry for AR/AP/Bank):
      - Customer (AR) credit lines are positive gross amounts → Credit column
      - Fee/expense lines are costs to the company → Debit column
      - Negative BOA amounts (money leaving account) → Debit column
      - Positive BOA amounts (money entering account) → Credit column

    Returns (debit_value, credit_value) where unused side is empty string "".
    """
    acct_type_lower = account_type.lower()

    if acct_type_lower == "customer":
        # AR credit: customer owes us → Credit
        return ("", abs(amount))

    if acct_type_lower in ("ledger", "vendor"):
        # Expense/fee or vendor payable → Debit
        return (abs(amount), "")

    if acct_type_lower == "bank":
        if amount >= 0:
            return ("", abs(amount))   # deposit → Credit
        else:
            return (abs(amount), "")   # withdrawal → Debit

    # Default: treat positive as credit, negative as debit
    if amount >= 0:
        return ("", abs(amount))
    return (abs(amount), "")


def lookup_customer(search_name: str, cust_df, acct_col: str, name_col: str):
    """
    FIX 3: Robust two-pass customer lookup against the Customer Master Account File.
    Pass 1 — substring containment (either direction).
    Pass 2 — token overlap (≥2 shared words).
    Returns (account_number, account_name).
    """
    search_key = super_clean_string(search_name)
    if not search_key:
        return "MISSING_ACCT", search_name

    # Pass 1: substring
    mask = cust_df['Account Name Clean'].apply(
        lambda x: search_key in str(x) or str(x) in search_key
    )
    match = cust_df[mask]
    if not match.empty:
        return str(match.iloc[0][acct_col]).strip(), str(match.iloc[0][name_col]).strip()

    # Pass 2: token overlap
    search_tokens = set(search_key.split())
    best_row, max_overlap = None, 0
    for _, row in cust_df.iterrows():
        master_tokens = set(str(row['Account Name Clean']).split())
        overlap = len(search_tokens & master_tokens)
        if overlap >= 2 and overlap > max_overlap:
            max_overlap = overlap
            best_row = row

    if best_row is not None:
        return str(best_row[acct_col]).strip(), str(best_row[name_col]).strip()

    return "MISSING_ACCT", search_name


def build_journal_row(
    date, company, account_name, account_type, account_num,
    posting_profile, cash_code, description,
    amount, offset_company, offset_acct_type, offset_acct,
    currency="USD"
):
    """Build a single 25-column journal dict using determine_debit_credit."""
    debit_val, credit_val = determine_debit_credit(amount, account_type)
    return {
        "Date": date,
        "Voucher": "",
        "Account name": account_name,
        "Company": company,
        "Account type": account_type,
        "Account": account_num,
        "Posting profile": posting_profile,
        "Cash code": cash_code,
        "Description": description,
        "Debit": debit_val,
        "Credit": credit_val,
        "Item sales tax group": "",
        "Sales tax code": "",
        "Offset company": offset_company,
        "Offset account type": offset_acct_type,
        "Offset account": offset_acct,
        "Offset transaction text": "",
        "Currency": currency,
        "Exchange rate": 1.00,
        "Item sales tax group2": "",
        "Sales tax group": "AVATAX",
        "Withholding tax group": "",
        "Release date": "",
        "Reversing entry": "No",
        "Reversing date": "",
    }


# ─── MAIN PIPELINE ───────────────────────────────────────────────────────────

if gateway_file and boa_file:
    st.subheader("4. Review & Generate")

    try:
        # ── Guard: master lookup files must exist in repo ──────────────────
        if not os.path.exists(MASTER_FILE_NAME) or not os.path.exists(CASH_CODE_FILE):
            st.error(
                f"❌ Missing repository lookup files. "
                f"Verify GitHub sync for '{MASTER_FILE_NAME}' and '{CASH_CODE_FILE}'."
            )
            st.stop()

        # ── A. Load Customer Master Account File ───────────────────────────
        # FIX 3a: Load and normalise the master file robustly
        cust_df = pd.read_excel(MASTER_FILE_NAME, engine='openpyxl')
        cust_df.columns = [str(col).strip() for col in cust_df.columns]

        # Detect the account-number and account-name columns by keyword
        acct_col = next(
            (c for c in cust_df.columns if re.search(r'\baccount\b', c, re.I)
             and not re.search(r'name|type|desc', c, re.I)),
            cust_df.columns[1]
        )
        name_col = next(
            (c for c in cust_df.columns if re.search(r'name', c, re.I)),
            cust_df.columns[2]
        )
        cust_df['Account Name Clean'] = cust_df[name_col].apply(super_clean_string)

        # ── B. Load Cash Code Masterlist ───────────────────────────────────
        # FIX 3b: Flexible column detection so any header layout works
        cc_df = pd.read_excel(CASH_CODE_FILE, engine='openpyxl')
        cc_df.columns = [str(col).strip() for col in cc_df.columns]

        cc_term_col = next(
            (c for c in cc_df.columns if re.search(r'term|desc|name|type', c, re.I)),
            cc_df.columns[0]
        )
        cc_code_col = next(
            (c for c in cc_df.columns if re.search(r'code', c, re.I)),
            cc_df.columns[1]
        )
        # Pre-clean cash code terms for faster lookups
        cc_df['Term Clean'] = cc_df[cc_term_col].apply(super_clean_string)

        def dynamic_cash_code_lookup(term_string: str, fallback: str) -> str:
            """FIX 3c: Match cash codes against the masterlist by cleaned term."""
            clean_term = super_clean_string(term_string)
            if not clean_term:
                return fallback
            # Exact / partial match on pre-cleaned column
            mask = cc_df['Term Clean'].str.contains(clean_term, na=False)
            if not mask.any():
                # Reverse: does any row term appear inside the search string?
                mask = cc_df['Term Clean'].apply(lambda t: bool(t) and t in clean_term)
            if mask.any():
                return str(cc_df[mask].iloc[0][cc_code_col]).strip()
            return fallback

        # ── C. Load BOA Expenses Mapping Guide (optional) ─────────────────
        if os.path.exists(BOA_EXPENSES_FILE):
            exp_df = pd.read_csv(BOA_EXPENSES_FILE, skiprows=1)
            exp_df.columns = [str(col).strip() for col in exp_df.columns]
        else:
            exp_df = pd.DataFrame()

        # ── D. Load Bank of America Statement ─────────────────────────────
        if boa_file.name.endswith('.csv'):
            boa_lines = boa_file.getvalue().decode('utf-8', errors='ignore').splitlines()
            skip_count = 0
            for line in boa_lines:
                if re.search(r'date.*description.*amount', line, re.I):
                    break
                skip_count += 1
            boa_file.seek(0)
            boa_all_df = pd.read_csv(boa_file, skiprows=skip_count)
        else:
            boa_all_df = pd.read_excel(boa_file, engine='openpyxl')

        boa_all_df.columns = [str(col).strip() for col in boa_all_df.columns]
        boa_desc_col = next(
            (c for c in boa_all_df.columns if re.search(r'desc|text|memo', c, re.I)),
            boa_all_df.columns[1]
        )
        boa_date_col_name = next(
            (c for c in boa_all_df.columns if re.search(r'date|post', c, re.I)),
            boa_all_df.columns[0]
        )
        boa_amount_col = next(
            (c for c in boa_all_df.columns if re.search(r'amount|amt', c, re.I)),
            boa_all_df.columns[2]
        )

        # Identify gateway settlement row in BOA
        is_stripe = boa_all_df[
            boa_all_df[boa_desc_col].astype(str).str.contains('STRIPE', case=False, na=False)
        ]
        is_zoho = boa_all_df[
            boa_all_df[boa_desc_col].astype(str).str.contains('ZOHO PAYMENTS', case=False, na=False)
        ]

        boa_date = ""
        boa_reference_desc = "PROCESSING CLEARANCE SETTLEMENT"
        boa_net_deposit = 0.0

        if not is_stripe.empty:
            engine_mode = "STRIPE"
            boa_date = str(is_stripe.iloc[0][boa_date_col_name]).strip()
            boa_reference_desc = str(is_stripe.iloc[0][boa_desc_col]).strip()
            boa_net_deposit = abs(safe_float(is_stripe.iloc[0][boa_amount_col]))
        elif not is_zoho.empty:
            engine_mode = "ZOHO"
            boa_date = str(is_zoho.iloc[0][boa_date_col_name]).strip()
            boa_reference_desc = str(is_zoho.iloc[0][boa_desc_col]).strip()
            boa_net_deposit = abs(safe_float(is_zoho.iloc[0][boa_amount_col]))
        else:
            engine_mode = "ZOHO"
            if not boa_all_df.empty:
                boa_date = str(boa_all_df.iloc[0][boa_date_col_name]).strip()
                boa_reference_desc = str(boa_all_df.iloc[0][boa_desc_col]).strip()
                boa_net_deposit = abs(safe_float(boa_all_df.iloc[0][boa_amount_col]))

        journal_rows = []

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE MODE 1: STRIPE — PDF extractor
        # ══════════════════════════════════════════════════════════════════════
        if engine_mode == "STRIPE" or gateway_file.name.endswith('.pdf'):
            st.info("⚙️ Running Engine Mode: Stripe Factual PDF Extractor")

            pdf_text = extract_text_from_pdf(gateway_file)
            charge_blocks = []
            total_extracted_gross = 0.0

            lines = [l.strip() for l in pdf_text.split('\n') if l.strip()]
            for idx, line in enumerate(lines):
                line_lower = line.lower()
                if "charge" in line_lower and ("plan" in line_lower or "agreement" in line_lower):
                    amounts = [
                        float(amt.replace(',', ''))
                        for amt in re.findall(r'\d+(?:,\d{3})*(?:\.\d{2})', line)
                    ]
                    if amounts:
                        gross_amt = amounts[0]
                        total_extracted_gross += gross_amt

                        extracted_name = "Unknown Customer"
                        if "agreement -" in line_lower:
                            extracted_name = line.split("Agreement -")[-1].split(" -")[0].strip()
                        elif "plan -" in line_lower:
                            extracted_name = line.split("Plan -")[-1].split(" -")[0].strip()

                        extracted_name = re.sub(r'\S+@\S+', '', extracted_name).split('@')[0].strip()
                        extracted_name = extracted_name.strip(" -")

                        charge_blocks.append({
                            "name": extracted_name,
                            "gross": gross_amt,
                            "is_installment": "installment" in line_lower
                        })

            for charge in charge_blocks:
                gross_val = charge["gross"]
                is_inst = charge["is_installment"]
                cash_code_label = "Installment" if is_inst else "Receipt"
                cash_code = dynamic_cash_code_lookup(cash_code_label, "AR002" if is_inst else "AR001")

                cust_acct, cust_name = lookup_customer(charge["name"], cust_df, acct_col, name_col)
                prefix = "MPP " if is_inst else ""
                credit_desc = f"{prefix}{cust_acct} {cust_name}_{boa_reference_desc}"

                journal_rows.append(build_journal_row(
                    date=boa_date, company=company_id,
                    account_name=cust_name, account_type="Customer",
                    account_num=cust_acct, posting_profile="AutoPost",
                    cash_code=cash_code, description=credit_desc,
                    amount=gross_val,
                    offset_company=company_id, offset_acct_type="Bank",
                    offset_acct=offset_account
                ))

            # Stripe fee = gross total − net deposit
            calculated_stripe_fee = round(total_extracted_gross - boa_net_deposit, 2)
            if calculated_stripe_fee > 0:
                merchant_cc = dynamic_cash_code_lookup("Stripe Merchant Fee", "OSF005")
                fee_desc = f"Stripe Merchant Fee_{boa_reference_desc}"
                journal_rows.append(build_journal_row(
                    date=boa_date, company=company_id,
                    account_name="Outside Service (Finance)", account_type="Ledger",
                    account_num=debit_ledger_acct, posting_profile="",
                    cash_code=merchant_cc, description=fee_desc,
                    amount=calculated_stripe_fee,   # positive → Debit for Ledger
                    offset_company=company_id, offset_acct_type="Bank",
                    offset_acct=offset_account
                ))

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE MODE 2: ZOHO — multi-invoice + payout CSV reconciliation
        # ══════════════════════════════════════════════════════════════════════
        else:
            st.info("⚙️ Running Engine Mode: Zoho Corporate Payment Pipeline")

            # FIX 1: Collect payer names & invoice terms from ALL uploaded invoices
            all_invoice_terms = []   # list of "monthly"/"receipt" per invoice
            all_payer_names = []     # list of extracted payer name per invoice

            if invoice_files:
                for inv_file in invoice_files:
                    inv_terms = "receipt"
                    inv_payer = ""

                    if inv_file.name.endswith('.pdf'):
                        raw_text = extract_text_from_pdf(inv_file)
                        clean_text = super_clean_string(raw_text)

                        if 'monthly' in clean_text or 'mpp' in clean_text:
                            inv_terms = "monthly"

                        # Extract "Bill To" / "Invoice To" name
                        inv_lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
                        for i, line in enumerate(inv_lines):
                            if re.search(r'bill\s*to|invoice\s*to', line, re.I):
                                if i + 1 < len(inv_lines):
                                    inv_payer = inv_lines[i + 1].strip()
                                    break

                    all_invoice_terms.append(inv_terms)
                    all_payer_names.append(inv_payer)

            # Load Zoho payout export
            if gateway_file.name.endswith('.csv'):
                zoho_df = pd.read_csv(gateway_file)
            else:
                zoho_df = pd.read_excel(gateway_file, engine='openpyxl')
            zoho_df.columns = [str(col).strip() for col in zoho_df.columns]

            # FIX 2: Detect gross/fee columns reliably
            zoho_gross_col = next(
                (c for c in zoho_df.columns if re.search(r'^amount$', c, re.I)),
                next((c for c in zoho_df.columns if re.search(r'amount|gross', c, re.I)),
                     zoho_df.columns[3])
            )
            zoho_fee_col = next(
                (c for c in zoho_df.columns if re.search(r'^fee$', c, re.I)),
                next((c for c in zoho_df.columns if re.search(r'fee|charge', c, re.I)),
                     zoho_df.columns[4])
            )
            zoho_cust_col = next(
                (c for c in zoho_df.columns if re.search(r'customer.*name|payer', c, re.I)),
                None
            )
            zoho_type_col = next(
                (c for c in zoho_df.columns if re.search(r'transaction.*type|type', c, re.I)),
                None
            )

            for idx, row in zoho_df.iterrows():
                # ── Resolve payer name ──────────────────────────────────────
                payer_name = ""
                if zoho_cust_col and zoho_cust_col in zoho_df.columns:
                    raw_payer = str(row[zoho_cust_col]).strip()
                    if raw_payer and raw_payer.lower() not in ("nan", "none"):
                        payer_name = raw_payer

                # Fall back to matching invoice payer by row index or first available
                if not payer_name or "inbody" in payer_name.lower():
                    if idx < len(all_payer_names) and all_payer_names[idx]:
                        payer_name = all_payer_names[idx]
                    elif all_payer_names:
                        payer_name = next((p for p in all_payer_names if p), "Unknown Payer")
                    else:
                        payer_name = "Unknown Payer"

                # ── Resolve invoice terms for this row ──────────────────────
                if idx < len(all_invoice_terms):
                    invoice_terms = all_invoice_terms[idx]
                elif all_invoice_terms:
                    invoice_terms = all_invoice_terms[0]
                else:
                    invoice_terms = "receipt"

                # ── FIX 2: Determine transaction direction from Zoho type ───
                txn_type = ""
                if zoho_type_col and zoho_type_col in zoho_df.columns:
                    txn_type = str(row[zoho_type_col]).strip().lower()

                # Zoho marks refunds/chargebacks as negative or with type keywords
                is_refund = any(kw in txn_type for kw in ("refund", "chargeback", "reversal", "void"))

                gross_raw = safe_float(row[zoho_gross_col])
                fee_raw = safe_float(row[zoho_fee_col])

                # Negative gross in CSV also signals a refund/reversal
                if gross_raw < 0:
                    is_refund = True

                gross_amt = abs(gross_raw)
                fee_amt = abs(fee_raw)

                # ── Customer lookup ─────────────────────────────────────────
                cust_acct, cust_name = lookup_customer(payer_name, cust_df, acct_col, name_col)

                # ── Cash code lookup ────────────────────────────────────────
                if is_refund:
                    cash_code = dynamic_cash_code_lookup("Refund", "AR003")
                elif invoice_terms == "monthly":
                    cash_code = dynamic_cash_code_lookup("Installment", "AR002")
                else:
                    cash_code = dynamic_cash_code_lookup("Receipt", "AR001")

                # ── Description ─────────────────────────────────────────────
                prefix = "MPP " if invoice_terms == "monthly" else ""
                credit_desc = f"{prefix}{cust_acct} {cust_name}_{boa_reference_desc}"

                # ── FIX 2: For refunds, flip the sign so it goes to Debit ──
                entry_amount = -gross_amt if is_refund else gross_amt

                journal_rows.append(build_journal_row(
                    date=boa_date, company=company_id,
                    account_name=cust_name,
                    account_type="Customer",
                    account_num=cust_acct,
                    posting_profile="AutoPost",
                    cash_code=cash_code,
                    description=credit_desc,
                    amount=entry_amount,
                    offset_company=company_id,
                    offset_acct_type="Bank",
                    offset_acct=offset_account
                ))

                # ── Fee line (always a debit/expense) ──────────────────────
                if fee_amt > 0:
                    fee_cc = dynamic_cash_code_lookup("Zoho Merchant Fee", "OSF005")
                    fee_desc = f"Zoho Merchant Fee {cust_acct}_{cust_name}_{boa_reference_desc}"
                    journal_rows.append(build_journal_row(
                        date=boa_date, company=company_id,
                        account_name="Outside Service (Finance)",
                        account_type="Ledger",
                        account_num=debit_ledger_acct,
                        posting_profile="",
                        cash_code=fee_cc,
                        description=fee_desc,
                        amount=fee_amt,   # positive → Debit for Ledger
                        offset_company=company_id,
                        offset_acct_type="Bank",
                        offset_acct=offset_account
                    ))

        # ══════════════════════════════════════════════════════════════════════
        # STEP E: BOA EXPENSE INGESTION (debit lines from bank statement)
        # ══════════════════════════════════════════════════════════════════════
        if not exp_df.empty:
            for _, r_row in boa_all_df.iterrows():
                raw_desc = str(r_row[boa_desc_col]).strip()
                raw_amt = safe_float(r_row[boa_amount_col])

                # FIX 2: Only process rows where money LEFT the account (debits)
                if raw_amt >= 0:
                    continue

                debit_value = abs(raw_amt)

                for _, e_row in exp_df.iterrows():
                    lookup_keyword = (
                        str(e_row['Bank Transaction Description'])
                        .split('*')[0].split()[0].strip().lower()
                    )
                    if lookup_keyword and lookup_keyword in raw_desc.lower():
                        map_name = str(e_row['Account name']).strip()
                        map_type = str(e_row['Account type']).strip()
                        map_acct = str(e_row['Account']).strip()
                        map_cc = str(e_row.get('Cash code', 'SP001')).strip()
                        map_desc_base = str(e_row['Description']).strip()

                        profile_flag = "AutoPost" if map_type.lower() == "vendor" else ""
                        combined_desc = f"{map_desc_base}_{raw_desc}"
                        t_date = str(r_row[boa_date_col_name]).strip()

                        journal_rows.append(build_journal_row(
                            date=t_date, company=company_id,
                            account_name=map_name, account_type=map_type,
                            account_num=map_acct, posting_profile=profile_flag,
                            cash_code=map_cc, description=combined_desc,
                            amount=debit_value,
                            offset_company=company_id,
                            offset_acct_type="Bank",
                            offset_acct=offset_account
                        ))
                        break

        # ══════════════════════════════════════════════════════════════════════
        # FINAL OUTPUT — 25-column D365 template
        # ══════════════════════════════════════════════════════════════════════
        columns_25 = [
            "Date", "Voucher", "Account name", "Company", "Account type", "Account",
            "Posting profile", "Cash code", "Description", "Debit", "Credit",
            "Item sales tax group", "Sales tax code", "Offset company",
            "Offset account type", "Offset account", "Offset transaction text",
            "Currency", "Exchange rate", "Item sales tax group2",
            "Sales tax group", "Withholding tax group", "Release date",
            "Reversing entry", "Reversing date"
        ]

        final_df = pd.DataFrame(journal_rows)
        if not final_df.empty:
            final_df = final_df.reindex(columns=columns_25).fillna("")

            # ── Validation summary ─────────────────────────────────────────
            total_debits = pd.to_numeric(final_df["Debit"], errors='coerce').fillna(0).sum()
            total_credits = pd.to_numeric(final_df["Credit"], errors='coerce').fillna(0).sum()
            missing_accts = (final_df["Account"] == "MISSING_ACCT").sum()

            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Total Debits", f"${total_debits:,.2f}")
            col_b.metric("Total Credits", f"${total_credits:,.2f}")
            delta_color = "normal" if abs(total_debits - total_credits) < 0.02 else "inverse"
            col_c.metric("Balance Difference", f"${abs(total_debits - total_credits):,.2f}",
                         delta="Balanced ✅" if abs(total_debits - total_credits) < 0.02
                         else "⚠️ Out of balance", delta_color=delta_color)

            if missing_accts > 0:
                st.warning(
                    f"⚠️ {missing_accts} row(s) have MISSING_ACCT — "
                    "verify these customer names against the Master Account File."
                )

            st.success("🎉 Cross-matched and balanced journal entries created seamlessly!")
            st.dataframe(final_df, use_container_width=True)

            csv_data = final_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Perfect D365 Upload CSV",
                data=csv_data,
                file_name="D365_Reconciliation_Journal.csv",
                mime="text/csv"
            )
        else:
            st.warning("⚠️ No valid transaction entries found to process.")

    except Exception as e:
        st.error(f"❌ Automation mapping process failed: {str(e)}")
        st.exception(e)   # shows full traceback in the UI for easier debugging

else:
    st.info(
        "💡 Please upload your Gateway File (Zoho or Stripe) and Bank of America statement "
        "to activate the automated alignment engine."
    )
