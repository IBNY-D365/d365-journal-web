import streamlit as st
import pandas as pd
import re
import os
from pypdf import PdfReader

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="D365 Accounting Journal Generator", layout="wide")
st.title("📊 D365 Zoho, Stripe & BOA Journal Generator")
st.write("Upload your daily processing packages below to build your D365 upload templates.")

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Fixed D365 Settings")
company_id = st.sidebar.text_input("Company", value="bwa")
offset_account = st.sidebar.text_input("Offset Account", value="B1000002")
debit_ledger_acct = st.sidebar.text_input(
    "Debit Line Account (Ledger)",
    value="43170111-U26C05001-B735350-UOA003",
)

# Repo-level lookup filenames (must be committed in the same GitHub directory as app.py)
MASTER_FILE_NAME = "Customer Master Account File.xlsx"
CASH_CODE_FILE = "Cash Code Masterlist.xlsx"

# ─── FILE UPLOADERS ──────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    gateway_file = st.file_uploader(
        "1. Upload Zoho Payout Export (CSV/XLSX) or Stripe PDF",
        type=["csv", "xlsx", "pdf"],
    )
with col2:
    invoice_files = st.file_uploader(
        "2. Upload Invoice PDF(s) — all invoices for this payout batch",
        type=["pdf", "csv", "xlsx", "txt"],
        accept_multiple_files=True,
    )
with col3:
    boa_file = st.file_uploader(
        "3. Upload Bank of America Statement (CSV/XLSX)",
        type=["csv", "xlsx"],
    )


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf(f) -> str:
    try:
        f.seek(0)
        return "\n".join(p.extract_text() or "" for p in PdfReader(f).pages)
    except Exception:
        return ""


def safe_float(v) -> float:
    try:
        return float(str(v).replace(",", "").replace("$", "").strip())
    except Exception:
        return 0.0


def clean_for_match(text) -> str:
    if pd.isna(text) or text is None:
        return ""
    t = str(text).lower()
    t = re.sub(r'\S+@\S+', '', t)
    t = re.sub(r'\b(llc|pllc|inc|corp|co|incorporated|limited|llp|dba)\b', ' ', t)
    t = re.sub(r'[^a-z0-9\s]', ' ', t)
    return " ".join(t.split())


def lookup_customer(name, cust_df, acct_col, name_col):
    key = clean_for_match(name)
    if not key:
        return "MISSING_ACCT", name

    # Exact / containment match first
    mask = cust_df["_name_clean"].apply(lambda x: key in x or x in key)
    hit = cust_df[mask]
    if not hit.empty:
        return str(hit.iloc[0][acct_col]).strip(), str(hit.iloc[0][name_col]).strip()

    # Token overlap fallback
    key_tok = set(key.split())
    best, best_s = None, 0
    for _, row in cust_df.iterrows():
        s = len(key_tok & set(str(row["_name_clean"]).split()))
        if s >= 2 and s > best_s:
            best_s, best = s, row
    if best is not None:
        return str(best[acct_col]).strip(), str(best[name_col]).strip()

    return "MISSING_ACCT", name


def cc_lookup(term, cc_df, term_col, code_col, fallback=""):
    """
    Match term against Cash Code Masterlist.
    Returns matched code, or fallback (default blank).
    """
    key = clean_for_match(term)
    if not key:
        return fallback

    mask = cc_df["_term_clean"].str.contains(key, na=False)
    if not mask.any():
        mask = cc_df["_term_clean"].apply(lambda t: bool(t) and t in key)

    return str(cc_df[mask].iloc[0][code_col]).strip() if mask.any() else fallback


def parse_invoice_term(text) -> str:
    lower = text.lower()

    if "due on receipt" in lower:
        return "receipt"

    if any(k in lower for k in ("monthly payment", "monthly plan", "mpp", "installment")):
        return "monthly"

    return "receipt"


def parse_invoice_customer(text) -> str:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for i, line in enumerate(lines):
        if re.search(r'bill\s*to|invoice\s*to', line, re.I) and i + 1 < len(lines):
            return lines[i + 1]
    return ""


def parse_invoice_number(text) -> str:
    patterns = [
        r'INV-\d+',
        r'Invoice\s*#\s*(INV-\d+)',
        r'Purchase\s*#\s*(INV-\d+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            if match.groups():
                return match.group(1).strip()
            return match.group(0).strip()
    return ""


def extract_invoice_name_map(invoice_files_list):
    """Build invoice number -> {customer, term, file_name}."""
    invoice_map = {}
    for inv in (invoice_files_list or []):
        if not inv.name.lower().endswith(".pdf"):
            continue

        text = extract_text_from_pdf(inv)
        invoice_no = parse_invoice_number(text)
        customer_name = parse_invoice_customer(text)
        term = parse_invoice_term(text)

        if invoice_no:
            invoice_map[invoice_no] = {
                "customer": customer_name,
                "term": term,
                "file_name": inv.name,
            }
    return invoice_map


def extract_invoice_number_from_text(text: str) -> str:
    if not text:
        return ""
    match = re.search(r'INV-\d+', text, re.I)
    return match.group(0).strip() if match else ""


def make_row(
    date, company, acct_name, acct_type, acct_num,
    posting_profile, cash_code, description,
    debit, credit,
    off_company, off_acct_type, off_acct
) -> dict:
    return {
        "Date": date, "Voucher": "", "Account name": acct_name,
        "Company": company, "Account type": acct_type, "Account": acct_num,
        "Posting profile": posting_profile, "Cash code": cash_code,
        "Description": description, "Debit": debit, "Credit": credit,
        "Item sales tax group": "", "Sales tax code": "",
        "Offset company": off_company, "Offset account type": off_acct_type,
        "Offset account": off_acct, "Offset transaction text": "",
        "Currency": "USD", "Exchange rate": 1.00,
        "Item sales tax group2": "", "Sales tax group": "AVATAX",
        "Withholding tax group": "", "Release date": "",
        "Reversing entry": "No", "Reversing date": "",
    }


COLUMNS_25 = [
    "Date", "Voucher", "Account name", "Company", "Account type", "Account",
    "Posting profile", "Cash code", "Description", "Debit", "Credit",
    "Item sales tax group", "Sales tax code", "Offset company",
    "Offset account type", "Offset account", "Offset transaction text",
    "Currency", "Exchange rate", "Item sales tax group2",
    "Sales tax group", "Withholding tax group", "Release date",
    "Reversing entry", "Reversing date",
]

# ─── Known BOA keyword → Cash Code map ───────────────────────────────────────
BOA_KEYWORD_MAP = {
    "amazon": "SP001",
    "amzn": "SP001",
    "adobe": "OSD005",
    "microsoft": "OSD005",
    "google": "OSD005",
    "zoom": "OSD005",
    "dropbox": "OSD005",
    "fedex": "SC004",
    "ups": "SC004",
    "usps": "SC004",
    "uber": "TD004",
    "lyft": "TD004",
    "delta": "TD002",
    "united": "TD002",
    "american air": "TD002",
    "hilton": "TD003",
    "marriott": "TD003",
    "guardian": "OSI002",
    "taskrabbit": "OSO007",
}


def categorise_boa_expense(desc: str, cc_df, term_col, code_col) -> str:
    """
    Return cash code for a BOA debit line.
    1. Check BOA_KEYWORD_MAP.
    2. Try fuzzy against Cash Code Masterlist.
    3. Return blank if no match.
    """
    desc_lower = desc.lower()

    for kw, code in BOA_KEYWORD_MAP.items():
        if kw in desc_lower:
            return code

    tokens = re.sub(r'[^a-z0-9\s]', ' ', desc_lower).split()[:3]
    for token in tokens:
        if len(token) < 4:
            continue
        result = cc_lookup(token, cc_df, term_col, code_col, fallback="")
        if result:
            return result

    return ""


# ═════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ═════════════════════════════════════════════════════════════════════════════

if gateway_file and boa_file:
    st.subheader("4. Review & Generate")

    try:
        # ── Guard: repo master files must exist ───────────────────────────────
        if not os.path.exists(MASTER_FILE_NAME) or not os.path.exists(CASH_CODE_FILE):
            st.error(
                f"❌ Missing lookup files. Ensure '{MASTER_FILE_NAME}' and '{CASH_CODE_FILE}' are committed to the repository."
            )
            st.stop()

        # ── Load Customer Master Account File ─────────────────────────────────
        cust_df = pd.read_excel(MASTER_FILE_NAME, engine="openpyxl")
        cust_df.columns = [str(c).strip() for c in cust_df.columns]
        acct_col = next(
            (c for c in cust_df.columns if re.search(r'\baccount\b', c, re.I) and not re.search(r'name|type|desc', c, re.I)),
            cust_df.columns[1],
        )
        name_col = next((c for c in cust_df.columns if re.search(r'name', c, re.I)), cust_df.columns[2])
        cust_df["_name_clean"] = cust_df[name_col].apply(clean_for_match)

        # ── Load Cash Code Masterlist ─────────────────────────────────────────
        cc_raw = pd.read_excel(CASH_CODE_FILE, engine="openpyxl", header=None)
        header_row = 0
        for i, row in cc_raw.iterrows():
            if any(str(v).strip().lower() == "cash code" for v in row):
                header_row = i
                break
        cc_df = pd.read_excel(CASH_CODE_FILE, engine="openpyxl", skiprows=header_row)
        cc_df.columns = [str(c).strip() for c in cc_df.columns]
        cc_code_col = next((c for c in cc_df.columns if re.search(r'cash\s*code', c, re.I)), cc_df.columns[0])
        cc_name_col = next((c for c in cc_df.columns if re.search(r'name|desc', c, re.I)), cc_df.columns[1])
        cc_df = cc_df.dropna(subset=[cc_code_col])
        cc_df = cc_df[cc_df[cc_code_col].astype(str).str.strip() != "Cash Code"]
        cc_df["_term_clean"] = cc_df[cc_name_col].apply(clean_for_match)

        # ── Load Bank of America Statement ────────────────────────────────────
        if boa_file.name.lower().endswith(".csv"):
            raw_lines = boa_file.getvalue().decode("utf-8", errors="ignore").splitlines()
            skip = 0
            for line in raw_lines:
                if re.search(r'date.*description.*amount', line, re.I):
                    break
                skip += 1
            boa_file.seek(0)
            boa_df = pd.read_csv(boa_file, skiprows=skip)
        else:
            boa_df = pd.read_excel(boa_file, engine="openpyxl")

        boa_df.columns = [str(c).strip() for c in boa_df.columns]
        boa_date_col = next((c for c in boa_df.columns if re.search(r'date|post', c, re.I)), boa_df.columns[0])
        boa_desc_col = next((c for c in boa_df.columns if re.search(r'desc|text|memo', c, re.I)), boa_df.columns[1])
        boa_amt_col = next((c for c in boa_df.columns if re.search(r'amount|amt', c, re.I)), boa_df.columns[2])

        boa_df = boa_df[boa_df[boa_amt_col].notna()]
        boa_df = boa_df[boa_df[boa_amt_col].astype(str).str.strip() != ""]
        boa_df["_amt_float"] = boa_df[boa_amt_col].apply(safe_float)
        boa_df = boa_df[boa_df["_amt_float"] != 0.0].copy()

        # ── Detect engine mode ────────────────────────────────────────────────
        stripe_rows = boa_df[boa_df[boa_desc_col].astype(str).str.contains("STRIPE", case=False, na=False)]
        zoho_rows = boa_df[boa_df[boa_desc_col].astype(str).str.contains("ZOHO PAYMENTS", case=False, na=False)]

        if not stripe_rows.empty or gateway_file.name.lower().endswith(".pdf"):
            engine_mode = "STRIPE"
            boa_settlement = stripe_rows.iloc[0] if not stripe_rows.empty else boa_df.iloc[0]
        else:
            engine_mode = "ZOHO"
            boa_settlement = zoho_rows.iloc[0] if not zoho_rows.empty else boa_df.iloc[0]

        boa_date = str(boa_settlement[boa_date_col]).strip()
        boa_ref_desc = str(boa_settlement[boa_desc_col]).strip()
        boa_net_deposit = abs(safe_float(boa_settlement[boa_amt_col]))

        journal_rows = []

        # ─────────────────────────────────────────────────────────────────────
        # Preload invoice mapping for Zoho/Stripe customer resolution
        # ─────────────────────────────────────────────────────────────────────
        invoice_map = extract_invoice_name_map(invoice_files)

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE A — STRIPE PDF
        # ══════════════════════════════════════════════════════════════════════
        if engine_mode == "STRIPE":
            st.info("⚙️ Engine: Stripe PDF Extractor")
            pdf_text = extract_text_from_pdf(gateway_file)
            charges = []
            total_gross = 0.0

            for line in [l.strip() for l in pdf_text.splitlines() if l.strip()]:
                ll = line.lower()
                if "charge" in ll and ("plan" in ll or "agreement" in ll):
                    amounts = [float(a.replace(",", "")) for a in re.findall(r'\d+(?:,\d{3})*\.\d{2}', line)]
                    if amounts:
                        gross = amounts[0]
                        total_gross += gross
                        name = "Unknown Customer"
                        if "agreement -" in ll:
                            name = line.split("Agreement -")[-1].split(" -")[0].strip()
                        elif "plan -" in ll:
                            name = line.split("Plan -")[-1].split(" -")[0].strip()
                        name = re.sub(r'\S+@\S+', '', name).strip(" -")
                        charges.append({"name": name, "gross": gross, "is_mpp": "installment" in ll})

            for ch in charges:
                # For Stripe, try to resolve via invoice/customer master if possible, otherwise keep raw extracted name.
                ca, cn = lookup_customer(ch["name"], cust_df, acct_col, name_col)
                code = cc_lookup(
                    "Installment" if ch["is_mpp"] else "Receipt",
                    cc_df,
                    cc_name_col,
                    cc_code_col,
                    "AR002" if ch["is_mpp"] else "AR001",
                )
                prefix = "MPP " if ch["is_mpp"] else ""
                journal_rows.append(make_row(
                    boa_date, company_id, cn, "Customer", ca,
                    "AutoPost", code, f"{prefix}{ca} {cn}_{boa_ref_desc}",
                    "", ch["gross"],
                    company_id, "Bank", offset_account,
                ))

            fee = round(total_gross - boa_net_deposit, 2)
            if fee > 0:
                journal_rows.append(make_row(
                    boa_date, company_id, "Outside Service (Finance)", "Ledger", debit_ledger_acct,
                    "", "OSF006", f"Stripe Merchant Fee_{boa_ref_desc}",
                    fee, "",
                    company_id, "Bank", offset_account,
                ))

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE B — ZOHO PAYOUT EXPORT
        # ══════════════════════════════════════════════════════════════════════
        else:
            st.info("⚙️ Engine: Zoho Corporate Payment Pipeline")

            # Load Zoho export
            zoho_df = pd.read_csv(gateway_file) if gateway_file.name.lower().endswith(".csv") else pd.read_excel(gateway_file, engine="openpyxl")
            zoho_df.columns = [str(c).strip() for c in zoho_df.columns]

            gross_col = next(
                (c for c in zoho_df.columns if c.strip().lower() == "amount"),
                next((c for c in zoho_df.columns if re.search(r'amount|gross', c, re.I)), zoho_df.columns[3]),
            )
            fee_col = next(
                (c for c in zoho_df.columns if c.strip().lower() == "fee"),
                next((c for c in zoho_df.columns if re.search(r'\bfee\b', c, re.I)), zoho_df.columns[4]),
            )
            cust_name_col = next((c for c in zoho_df.columns if re.search(r'customer.*name|customername', c, re.I)), None)
            txn_type_col = next((c for c in zoho_df.columns if re.search(r'transaction.*type|transactiontype', c, re.I)), None)
            desc_col = next((c for c in zoho_df.columns if re.search(r'desc|description|memo', c, re.I)), None)

            st.caption(
                f"Zoho columns → Gross: **{gross_col}** | Fee: **{fee_col}** | Customer: **{cust_name_col or 'not found'}**"
            )

            for _, zrow in zoho_df.iterrows():
                gross_raw = safe_float(zrow[gross_col])
                fee_raw = safe_float(zrow[fee_col])
                if gross_raw == 0.0:
                    continue

                txn_type_val = str(zrow.get(txn_type_col, "")).lower() if txn_type_col else ""
                is_refund = gross_raw < 0 or any(k in txn_type_val for k in ("refund", "chargeback", "reversal", "void"))
                gross_amt = abs(gross_raw)
                fee_amt = abs(fee_raw)

                zoho_cust = ""
                if cust_name_col:
                    raw = str(zrow.get(cust_name_col, "")).strip()
                    if raw.lower() not in ("nan", "none", ""):
                        zoho_cust = raw

                # If Zoho doesn't supply the customer, resolve customer from invoice number in the description.
                if not zoho_cust:
                    raw_desc = str(zrow.get(desc_col, "")) if desc_col else ""
                    invoice_no = extract_invoice_number_from_text(raw_desc)
                    if invoice_no and invoice_no in invoice_map:
                        zoho_cust = invoice_map[invoice_no].get("customer", "")

                # Term resolution from invoice text/file when possible.
                term = "receipt"
                raw_desc = str(zrow.get(desc_col, "")) if desc_col else ""
                invoice_no = extract_invoice_number_from_text(raw_desc)
                if invoice_no and invoice_no in invoice_map:
                    term = invoice_map[invoice_no].get("term", "receipt")

                ca, cn = lookup_customer(zoho_cust or "Unknown Payer", cust_df, acct_col, name_col)

                if is_refund:
                    code = cc_lookup("Refund", cc_df, cc_name_col, cc_code_col, "AR003")
                elif term == "monthly":
                    code = cc_lookup("Installment", cc_df, cc_name_col, cc_code_col, "AR002")
                else:
                    code = cc_lookup("Receipt", cc_df, cc_name_col, cc_code_col, "AR001")

                mpp = "MPP " if term == "monthly" else ""
                cr_desc = f"{mpp}{ca} {cn}_{boa_ref_desc}"
                fee_desc = f"Zoho Merchant Fee {ca}_{cn}_{boa_ref_desc}"

                if is_refund:
                    journal_rows.append(make_row(
                        boa_date, company_id, cn, "Customer", ca,
                        "AutoPost", code, f"REFUND {cr_desc}",
                        gross_amt, "",
                        company_id, "Bank", offset_account,
                    ))
                else:
                    journal_rows.append(make_row(
                        boa_date, company_id, cn, "Customer", ca,
                        "AutoPost", code, cr_desc,
                        "", gross_amt,
                        company_id, "Bank", offset_account,
                    ))

                if fee_amt > 0:
                    journal_rows.append(make_row(
                        boa_date, company_id, "Outside Service (Finance)", "Ledger", debit_ledger_acct,
                        "", "OSF005", fee_desc,
                        fee_amt, "",
                        company_id, "Bank", offset_account,
                    ))

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE C — ALL OTHER BOA TRANSACTIONS
        # Temporary expense workflow: keep only Date, Description, Debit.
        # ══════════════════════════════════════════════════════════════════════
        st.info("⚙️ Processing remaining BOA transactions…")

        gateway_desc_pattern = "ZOHO PAYMENTS" if engine_mode == "ZOHO" else "STRIPE"
        gateway_mask = boa_df[boa_desc_col].astype(str).str.contains(gateway_desc_pattern, case=False, na=False)

        for _, brow in boa_df[~gateway_mask].iterrows():
            raw_desc = str(brow[boa_desc_col]).strip()
            raw_date = str(brow[boa_date_col]).strip()
            raw_amt = brow["_amt_float"]

            if raw_desc in ("", "nan") or raw_amt == 0.0:
                continue
            if re.search(r'beginning balance|ending balance', raw_desc, re.I):
                continue

            if raw_amt < 0:
                # Temporary expense workflow: only populate date, description, debit.
                blank_row = {col: "" for col in COLUMNS_25}
                blank_row["Date"] = raw_date
                blank_row["Description"] = raw_desc
                blank_row["Debit"] = abs(raw_amt)
                journal_rows.append(blank_row)
            else:
                # Keep existing logic for non-expense credits.
                cash_code = categorise_boa_expense(raw_desc, cc_df, cc_name_col, cc_code_col)
                journal_rows.append(make_row(
                    raw_date, company_id,
                    "Outside Service (Finance)", "Ledger", debit_ledger_acct,
                    "", cash_code,
                    raw_desc,
                    "", raw_amt,
                    company_id, "Bank", offset_account,
                ))

        # ══════════════════════════════════════════════════════════════════════
        # OUTPUT
        # ══════════════════════════════════════════════════════════════════════
        final_df = pd.DataFrame(journal_rows)

        if not final_df.empty:
            final_df = final_df.reindex(columns=COLUMNS_25).fillna("")

            total_deb = pd.to_numeric(final_df["Debit"], errors="coerce").fillna(0).sum()
            total_cre = pd.to_numeric(final_df["Credit"], errors="coerce").fillna(0).sum()
            diff = abs(total_deb - total_cre)
            balanced = diff < 0.02
            missing = (final_df["Account"] == "MISSING_ACCT").sum()

            m1, m2, m3 = st.columns(3)
            m1.metric("Total Debits", f"${total_deb:,.2f}")
            m2.metric("Total Credits", f"${total_cre:,.2f}")
            m3.metric(
                "Balance Difference",
                f"${diff:,.2f}",
                delta="✅ Balanced" if balanced else "⚠️ Out of balance",
                delta_color="normal" if balanced else "inverse",
            )

            if missing:
                st.warning(
                    f"⚠️ {missing} row(s) show MISSING_ACCT — verify those customer names in the Customer Master Account File."
                )

            def hl(row):
                return ["background-color:#fff3cd" if row["Account"] == "MISSING_ACCT" else "" for _ in row]

            st.success("🎉 All transactions processed. Review below, then download.")
            st.dataframe(final_df.style.apply(hl, axis=1), use_container_width=True)

            st.download_button(
                "📥 Download D365 Upload CSV",
                data=final_df.to_csv(index=False).encode("utf-8"),
                file_name="D365_Reconciliation_Journal.csv",
                mime="text/csv",
            )
        else:
            st.warning("⚠️ No journal entries were produced. Check your input files.")

    except Exception as e:
        st.error(f"❌ Pipeline error: {e}")
        st.exception(e)

else:
    st.info("💡 Upload the payout export (Zoho/Stripe) and the Bank of America statement to begin.")
