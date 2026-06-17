import streamlit as st
import pandas as pd
import re
import os
from pypdf import PdfReader

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="D365 Accounting Journal Generator", layout="wide")
st.title("📊 D365 Zoho, Stripe & BOA Journal Generator")
st.write("Upload your daily processing packages below to build your flawless 25-column D365 upload templates.")

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Fixed D365 Settings")
company_id        = st.sidebar.text_input("Company",          value="bwa")
offset_account    = st.sidebar.text_input("Offset Account",   value="B1000002")
debit_ledger_acct = st.sidebar.text_input("Debit Line Account (Ledger)", value="43170111-U26C05001-B735350-UOA003")

# Repo-level lookup filenames (must be in same GitHub directory as app.py)
MASTER_FILE_NAME  = "Customer Master Account File.xlsx"
CASH_CODE_FILE    = "Cash Code Masterlist.xlsx"
BOA_EXPENSES_FILE = "BOA3371 Expenses List.xlsx - BOA3371.csv"

# ─── FILE UPLOADERS ───────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    gateway_file = st.file_uploader(
        "1. Upload Zoho Payout Export (CSV or XLSX) or Stripe PDF",
        type=["csv", "xlsx", "pdf"]
    )
with col2:
    invoice_files = st.file_uploader(
        "2. Upload Invoice PDF(s) — upload ALL invoices for this payout batch",
        type=["pdf", "csv", "xlsx", "txt"],
        accept_multiple_files=True
    )
with col3:
    boa_file = st.file_uploader("3. Upload Bank of America Statement (CSV or XLSX)", type=["csv", "xlsx"])


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf(uploaded_pdf) -> str:
    """Return concatenated text from every page of an uploaded PDF object."""
    try:
        uploaded_pdf.seek(0)
        reader = PdfReader(uploaded_pdf)
        return "\n".join(
            page.extract_text() or "" for page in reader.pages
        )
    except Exception:
        return ""


def safe_float(value) -> float:
    """Strip currency symbols / commas and cast to float. Returns 0.0 on failure."""
    try:
        return float(str(value).replace(",", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return 0.0


def clean_for_match(text) -> str:
    """
    Normalise a string for fuzzy name matching:
    • lowercase
    • strip e-mail addresses
    • remove legal suffixes (LLC, Inc, …)
    • remove punctuation
    • collapse whitespace
    """
    if pd.isna(text) or text is None:
        return ""
    t = str(text).lower()
    t = re.sub(r'\S+@\S+', '', t)                                      # emails
    t = re.sub(r'\b(llc|pllc|inc|corp|co|incorporated|limited|llp|dba)\b', ' ', t)
    t = re.sub(r'[^a-z0-9\s]', ' ', t)                                 # punctuation
    return " ".join(t.split())


def lookup_customer(name: str, cust_df: pd.DataFrame, acct_col: str, name_col: str):
    """
    Two-pass fuzzy lookup against the Customer Master Account File.
    Pass 1 — substring containment (either direction).
    Pass 2 — token-overlap fallback (≥2 shared words).
    Returns (account_number, official_account_name).
    """
    key = clean_for_match(name)
    if not key:
        return "MISSING_ACCT", name

    # Pass 1: substring
    mask = cust_df["_name_clean"].apply(lambda x: key in x or x in key)
    hit  = cust_df[mask]
    if not hit.empty:
        return str(hit.iloc[0][acct_col]).strip(), str(hit.iloc[0][name_col]).strip()

    # Pass 2: token overlap
    key_tokens = set(key.split())
    best, best_score = None, 0
    for _, row in cust_df.iterrows():
        score = len(key_tokens & set(str(row["_name_clean"]).split()))
        if score >= 2 and score > best_score:
            best_score, best = score, row
    if best is not None:
        return str(best[acct_col]).strip(), str(best[name_col]).strip()

    return "MISSING_ACCT", name


def cash_code_lookup(term: str, cc_df: pd.DataFrame,
                     term_col: str, code_col: str, fallback: str) -> str:
    """Match a search term against the Cash Code Masterlist. Returns the code or fallback."""
    key = clean_for_match(term)
    if not key:
        return fallback
    mask = cc_df["_term_clean"].str.contains(key, na=False)
    if not mask.any():
        mask = cc_df["_term_clean"].apply(lambda t: bool(t) and t in key)
    return str(cc_df[mask].iloc[0][code_col]).strip() if mask.any() else fallback


def parse_invoice_term(pdf_text: str) -> str:
    """
    Determine payment term from invoice text.
    Returns 'monthly' (AR002 / MPP) or 'receipt' (AR001).
    """
    lower = pdf_text.lower()
    if any(kw in lower for kw in ("monthly payment", "monthly plan", "mpp", "installment")):
        return "monthly"
    return "receipt"


def parse_invoice_customer(pdf_text: str) -> str:
    """
    Extract the customer / bill-to name from invoice text.
    Looks for lines immediately after 'Bill To' or 'Invoice To'.
    """
    lines = [l.strip() for l in pdf_text.splitlines() if l.strip()]
    for i, line in enumerate(lines):
        if re.search(r'bill\s*to|invoice\s*to', line, re.I):
            if i + 1 < len(lines):
                return lines[i + 1]
    return ""


def make_row(date, company, acct_name, acct_type, acct_num,
             posting_profile, cash_code, description,
             debit, credit,
             off_company, off_acct_type, off_acct) -> dict:
    """Assemble a complete 25-column D365 journal row dict."""
    return {
        "Date":                   date,
        "Voucher":                "",
        "Account name":           acct_name,
        "Company":                company,
        "Account type":           acct_type,
        "Account":                acct_num,
        "Posting profile":        posting_profile,
        "Cash code":              cash_code,
        "Description":            description,
        "Debit":                  debit,
        "Credit":                 credit,
        "Item sales tax group":   "",
        "Sales tax code":         "",
        "Offset company":         off_company,
        "Offset account type":    off_acct_type,
        "Offset account":         off_acct,
        "Offset transaction text":"",
        "Currency":               "USD",
        "Exchange rate":          1.00,
        "Item sales tax group2":  "",
        "Sales tax group":        "AVATAX",
        "Withholding tax group":  "",
        "Release date":           "",
        "Reversing entry":        "No",
        "Reversing date":         "",
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


# ═════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE — only runs when required files are uploaded
# ═════════════════════════════════════════════════════════════════════════════

if gateway_file and boa_file:
    st.subheader("4. Review & Generate")

    try:
        # ── Guard: repo master files must exist ───────────────────────────────
        if not os.path.exists(MASTER_FILE_NAME) or not os.path.exists(CASH_CODE_FILE):
            st.error(
                f"❌ Missing lookup files in repository. "
                f"Ensure '{MASTER_FILE_NAME}' and '{CASH_CODE_FILE}' are committed."
            )
            st.stop()

        # ── Load Customer Master Account File ─────────────────────────────────
        cust_df = pd.read_excel(MASTER_FILE_NAME, engine="openpyxl")
        cust_df.columns = [str(c).strip() for c in cust_df.columns]

        # Detect account-number column (contains "account", not "name/type/desc")
        acct_col = next(
            (c for c in cust_df.columns
             if re.search(r'\baccount\b', c, re.I)
             and not re.search(r'name|type|desc', c, re.I)),
            cust_df.columns[1]
        )
        # Detect account-name column
        name_col = next(
            (c for c in cust_df.columns if re.search(r'name', c, re.I)),
            cust_df.columns[2]
        )
        cust_df["_name_clean"] = cust_df[name_col].apply(clean_for_match)

        # ── Load Cash Code Masterlist ─────────────────────────────────────────
        cc_df = pd.read_excel(CASH_CODE_FILE, engine="openpyxl")
        cc_df.columns = [str(c).strip() for c in cc_df.columns]

        cc_term_col = next(
            (c for c in cc_df.columns if re.search(r'term|desc|name|type', c, re.I)),
            cc_df.columns[0]
        )
        cc_code_col = next(
            (c for c in cc_df.columns if re.search(r'code', c, re.I)),
            cc_df.columns[1]
        )
        cc_df["_term_clean"] = cc_df[cc_term_col].apply(clean_for_match)

        # ── Load BOA Expenses mapping (optional) ──────────────────────────────
        exp_df = pd.DataFrame()
        if os.path.exists(BOA_EXPENSES_FILE):
            exp_df = pd.read_csv(BOA_EXPENSES_FILE, skiprows=1)
            exp_df.columns = [str(c).strip() for c in exp_df.columns]

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

        boa_date_col = next(
            (c for c in boa_df.columns if re.search(r'date|post', c, re.I)),
            boa_df.columns[0]
        )
        boa_desc_col = next(
            (c for c in boa_df.columns if re.search(r'desc|text|memo', c, re.I)),
            boa_df.columns[1]
        )
        boa_amt_col = next(
            (c for c in boa_df.columns if re.search(r'amount|amt', c, re.I)),
            boa_df.columns[2]
        )

        # ── Detect engine mode from BOA ───────────────────────────────────────
        stripe_rows = boa_df[boa_df[boa_desc_col].astype(str).str.contains("STRIPE", case=False, na=False)]
        zoho_rows   = boa_df[boa_df[boa_desc_col].astype(str).str.contains("ZOHO PAYMENTS", case=False, na=False)]

        if not stripe_rows.empty or gateway_file.name.lower().endswith(".pdf"):
            engine_mode = "STRIPE"
            boa_settlement_row = stripe_rows.iloc[0] if not stripe_rows.empty else boa_df.iloc[0]
        else:
            engine_mode = "ZOHO"
            boa_settlement_row = zoho_rows.iloc[0] if not zoho_rows.empty else boa_df.iloc[0]

        # The BOA posting date and full BOA description — used in every D365 description
        boa_date      = str(boa_settlement_row[boa_date_col]).strip()
        boa_ref_desc  = str(boa_settlement_row[boa_desc_col]).strip()
        boa_net_deposit = abs(safe_float(boa_settlement_row[boa_amt_col]))

        journal_rows = []

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE A — STRIPE PDF
        # ══════════════════════════════════════════════════════════════════════
        if engine_mode == "STRIPE":
            st.info("⚙️ Engine: Stripe PDF Extractor")

            pdf_text = extract_text_from_pdf(gateway_file)
            charge_blocks = []
            total_gross = 0.0

            for line in [l.strip() for l in pdf_text.splitlines() if l.strip()]:
                ll = line.lower()
                if "charge" in ll and ("plan" in ll or "agreement" in ll):
                    amounts = [
                        float(a.replace(",", ""))
                        for a in re.findall(r'\d+(?:,\d{3})*\.\d{2}', line)
                    ]
                    if amounts:
                        gross = amounts[0]
                        total_gross += gross
                        name = "Unknown Customer"
                        if "agreement -" in ll:
                            name = line.split("Agreement -")[-1].split(" -")[0].strip()
                        elif "plan -" in ll:
                            name = line.split("Plan -")[-1].split(" -")[0].strip()
                        name = re.sub(r'\S+@\S+', '', name).strip(" -")
                        charge_blocks.append({
                            "name": name,
                            "gross": gross,
                            "is_installment": "installment" in ll,
                        })

            for charge in charge_blocks:
                cust_acct, cust_name = lookup_customer(charge["name"], cust_df, acct_col, name_col)
                is_mpp   = charge["is_installment"]
                cc       = cash_code_lookup("Installment" if is_mpp else "Receipt",
                                            cc_df, cc_term_col, cc_code_col,
                                            "AR002" if is_mpp else "AR001")
                prefix   = "MPP " if is_mpp else ""
                desc     = f"{prefix}{cust_acct} {cust_name}_{boa_ref_desc}"

                journal_rows.append(make_row(
                    date=boa_date, company=company_id,
                    acct_name=cust_name, acct_type="Customer", acct_num=cust_acct,
                    posting_profile="AutoPost", cash_code=cc, description=desc,
                    debit="", credit=charge["gross"],
                    off_company=company_id, off_acct_type="Bank", off_acct=offset_account,
                ))

            # Stripe fee = total gross − BOA net deposit
            stripe_fee = round(total_gross - boa_net_deposit, 2)
            if stripe_fee > 0:
                fee_cc   = cash_code_lookup("Stripe Merchant Fee", cc_df, cc_term_col, cc_code_col, "OSF005")
                fee_desc = f"Stripe Merchant Fee_{boa_ref_desc}"
                journal_rows.append(make_row(
                    date=boa_date, company=company_id,
                    acct_name="Outside Service (Finance)", acct_type="Ledger",
                    acct_num=debit_ledger_acct,
                    posting_profile="", cash_code=fee_cc, description=fee_desc,
                    debit=stripe_fee, credit="",
                    off_company=company_id, off_acct_type="Bank", off_acct=offset_account,
                ))

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE B — ZOHO PAYOUT EXPORT
        #
        # Rules (from Zoho Payment Workflow doc):
        #   • Each row in the Zoho export = one customer transaction
        #   • Credit line  → gross amount (Amount column), Account type = Customer
        #   • Debit line   → fee amount   (Fee column),   Account type = Ledger
        #   • BOA only shows the NET deposit — we NEVER use BOA amount per customer
        #   • Cash code    → AR001 (receipt) or AR002 (monthly/MPP)
        #   • Payment term → determined per customer from their uploaded invoice PDF
        #   • Credit desc  → "[MPP ]ACCT# Customer Name_BOA FULL DESCRIPTION"
        #   • Debit desc   → "Zoho Merchant Fee ACCT#_Customer Name_BOA FULL DESCRIPTION"
        # ══════════════════════════════════════════════════════════════════════
        else:
            st.info("⚙️ Engine: Zoho Corporate Payment Pipeline")

            # ── Step 1: Parse every uploaded invoice PDF ──────────────────────
            # Build a dict keyed by cleaned customer name → {"term": ..., "raw_name": ...}
            invoice_map: dict[str, dict] = {}

            if invoice_files:
                for inv in invoice_files:
                    if not inv.name.lower().endswith(".pdf"):
                        continue
                    text      = extract_text_from_pdf(inv)
                    term      = parse_invoice_term(text)
                    cust_name = parse_invoice_customer(text)
                    if cust_name:
                        invoice_map[clean_for_match(cust_name)] = {
                            "term":     term,
                            "raw_name": cust_name,
                        }

            # ── Step 2: Load Zoho payout export ──────────────────────────────
            if gateway_file.name.lower().endswith(".csv"):
                zoho_df = pd.read_csv(gateway_file)
            else:
                zoho_df = pd.read_excel(gateway_file, engine="openpyxl")
            zoho_df.columns = [str(c).strip() for c in zoho_df.columns]

            # Column detection — handle slight header variations
            # Gross amount column: "Amount" (exact first, then partial)
            gross_col = next(
                (c for c in zoho_df.columns if c.strip().lower() == "amount"),
                next((c for c in zoho_df.columns if re.search(r'amount|gross', c, re.I)),
                     zoho_df.columns[3]),
            )
            # Fee column: "Fee" (exact first, then partial)
            fee_col = next(
                (c for c in zoho_df.columns if c.strip().lower() == "fee"),
                next((c for c in zoho_df.columns if re.search(r'\bfee\b', c, re.I)),
                     zoho_df.columns[4]),
            )
            # Customer name column
            cust_name_col = next(
                (c for c in zoho_df.columns if re.search(r'customer.*name|customername', c, re.I)),
                None,
            )
            # Transaction type column (to detect refunds)
            txn_type_col = next(
                (c for c in zoho_df.columns if re.search(r'transaction.*type|transactiontype', c, re.I)),
                None,
            )

            st.caption(
                f"Zoho columns detected → Gross: **{gross_col}** | "
                f"Fee: **{fee_col}** | "
                f"Customer: **{cust_name_col or 'not found'}**"
            )

            # ── Step 3: Build journal rows — one credit + one debit per Zoho row ──
            for _, zrow in zoho_df.iterrows():

                # ── 3a. Gross and fee amounts ─────────────────────────────────
                gross_raw = safe_float(zrow[gross_col])
                fee_raw   = safe_float(zrow[fee_col])

                # Skip rows with no value (header echoes, subtotals, blanks)
                if gross_raw == 0.0:
                    continue

                # Detect refund/reversal by sign or TransactionType field
                txn_type_val = ""
                if txn_type_col:
                    txn_type_val = str(zrow.get(txn_type_col, "")).strip().lower()
                is_refund = gross_raw < 0 or any(
                    kw in txn_type_val for kw in ("refund", "chargeback", "reversal", "void")
                )

                gross_amt = abs(gross_raw)
                fee_amt   = abs(fee_raw)

                # ── 3b. Customer name from Zoho export ───────────────────────
                zoho_cust_name = ""
                if cust_name_col:
                    raw = str(zrow.get(cust_name_col, "")).strip()
                    if raw.lower() not in ("nan", "none", ""):
                        zoho_cust_name = raw

                # ── 3c. Match invoice to this customer to get payment term ────
                #  Priority: exact/fuzzy match on invoice_map keys
                #  Fallback: "receipt" (AR001)
                term = "receipt"
                if invoice_map:
                    zoho_key = clean_for_match(zoho_cust_name)
                    # Try direct key match first
                    if zoho_key in invoice_map:
                        term = invoice_map[zoho_key]["term"]
                    else:
                        # Fuzzy: find invoice whose cleaned name overlaps with Zoho name
                        zoho_tokens = set(zoho_key.split())
                        for inv_key, inv_data in invoice_map.items():
                            inv_tokens = set(inv_key.split())
                            if len(zoho_tokens & inv_tokens) >= 2:
                                term = inv_data["term"]
                                break

                # ── 3d. Customer account lookup ───────────────────────────────
                cust_acct, cust_name = lookup_customer(
                    zoho_cust_name or "Unknown Payer", cust_df, acct_col, name_col
                )

                # ── 3e. Cash code ─────────────────────────────────────────────
                if is_refund:
                    cc = cash_code_lookup("Refund", cc_df, cc_term_col, cc_code_col, "AR003")
                elif term == "monthly":
                    cc = cash_code_lookup("Installment", cc_df, cc_term_col, cc_code_col, "AR002")
                else:
                    cc = cash_code_lookup("Receipt", cc_df, cc_term_col, cc_code_col, "AR001")

                # ── 3f. Descriptions (per workflow doc) ───────────────────────
                # Credit:  "[MPP ]ACCT# Customer Name_BOA FULL REFERENCE"
                # Debit:   "Zoho Merchant Fee ACCT#_Customer Name_BOA FULL REFERENCE"
                mpp_prefix   = "MPP " if term == "monthly" else ""
                credit_desc  = f"{mpp_prefix}{cust_acct} {cust_name}_{boa_ref_desc}"
                fee_desc     = f"Zoho Merchant Fee {cust_acct}_{cust_name}_{boa_ref_desc}"

                # ── 3g. CREDIT LINE — gross amount, Customer account type ──────
                if is_refund:
                    # Refund: debit the customer (reversal of original credit)
                    journal_rows.append(make_row(
                        date=boa_date, company=company_id,
                        acct_name=cust_name, acct_type="Customer", acct_num=cust_acct,
                        posting_profile="AutoPost", cash_code=cc,
                        description=f"REFUND {credit_desc}",
                        debit=gross_amt, credit="",
                        off_company=company_id, off_acct_type="Bank", off_acct=offset_account,
                    ))
                else:
                    journal_rows.append(make_row(
                        date=boa_date, company=company_id,
                        acct_name=cust_name, acct_type="Customer", acct_num=cust_acct,
                        posting_profile="AutoPost", cash_code=cc,
                        description=credit_desc,
                        debit="", credit=gross_amt,
                        off_company=company_id, off_acct_type="Bank", off_acct=offset_account,
                    ))

                # ── 3h. DEBIT LINE — fee amount, Ledger account type ──────────
                # Per workflow: debit = processing fee shown in Zoho (NOT in BOA)
                if fee_amt > 0:
                    journal_rows.append(make_row(
                        date=boa_date, company=company_id,
                        acct_name="Outside Service (Finance)", acct_type="Ledger",
                        acct_num=debit_ledger_acct,
                        posting_profile="", cash_code="OSF005",
                        description=fee_desc,
                        debit=fee_amt, credit="",
                        off_company=company_id, off_acct_type="Bank", off_acct=offset_account,
                    ))

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE C — BOA EXPENSE LINES (debits from bank statement)
        # Only processes rows where money LEFT the account (negative amounts)
        # and description matches the BOA Expenses mapping file.
        # ══════════════════════════════════════════════════════════════════════
        if not exp_df.empty:
            for _, brow in boa_df.iterrows():
                raw_desc = str(brow[boa_desc_col]).strip()
                raw_amt  = safe_float(brow[boa_amt_col])

                if raw_amt >= 0:       # skip deposits / credits
                    continue

                debit_val = abs(raw_amt)

                for _, erow in exp_df.iterrows():
                    kw = str(erow.get("Bank Transaction Description", "")).split("*")[0].split()
                    if not kw:
                        continue
                    kw = kw[0].strip().lower()
                    if kw and kw in raw_desc.lower():
                        map_type    = str(erow["Account type"]).strip()
                        profile_flag = "AutoPost" if map_type.lower() == "vendor" else ""
                        journal_rows.append(make_row(
                            date=str(brow[boa_date_col]).strip(),
                            company=company_id,
                            acct_name=str(erow["Account name"]).strip(),
                            acct_type=map_type,
                            acct_num=str(erow["Account"]).strip(),
                            posting_profile=profile_flag,
                            cash_code=str(erow.get("Cash code", "SP001")).strip(),
                            description=f"{str(erow['Description']).strip()}_{raw_desc}",
                            debit=debit_val, credit="",
                            off_company=company_id,
                            off_acct_type="Bank",
                            off_acct=offset_account,
                        ))
                        break

        # ══════════════════════════════════════════════════════════════════════
        # OUTPUT — 25-column D365 upload template
        # ══════════════════════════════════════════════════════════════════════
        final_df = pd.DataFrame(journal_rows)

        if not final_df.empty:
            final_df = final_df.reindex(columns=COLUMNS_25).fillna("")

            # ── Validation metrics ────────────────────────────────────────────
            total_deb = pd.to_numeric(final_df["Debit"],  errors="coerce").fillna(0).sum()
            total_cre = pd.to_numeric(final_df["Credit"], errors="coerce").fillna(0).sum()
            diff      = abs(total_deb - total_cre)
            balanced  = diff < 0.02
            missing   = (final_df["Account"] == "MISSING_ACCT").sum()

            m1, m2, m3 = st.columns(3)
            m1.metric("Total Debits",  f"${total_deb:,.2f}")
            m2.metric("Total Credits", f"${total_cre:,.2f}")
            m3.metric(
                "Balance Difference", f"${diff:,.2f}",
                delta="✅ Balanced" if balanced else "⚠️ Out of balance",
                delta_color="normal" if balanced else "inverse",
            )

            if missing:
                st.warning(
                    f"⚠️ {missing} row(s) show MISSING_ACCT. "
                    "Check those customer names against the Customer Master Account File."
                )

            # ── Highlight MISSING_ACCT rows ───────────────────────────────────
            def highlight_missing(row):
                return ["background-color: #fff3cd" if row["Account"] == "MISSING_ACCT"
                        else "" for _ in row]

            st.success("🎉 Journal entries built! Review below, then download.")
            st.dataframe(
                final_df.style.apply(highlight_missing, axis=1),
                use_container_width=True,
            )

            csv_bytes = final_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Download D365 Upload CSV",
                data=csv_bytes,
                file_name="D365_Reconciliation_Journal.csv",
                mime="text/csv",
            )
        else:
            st.warning("⚠️ No journal entries were produced. Check your input files.")

    except Exception as e:
        st.error(f"❌ Pipeline error: {e}")
        st.exception(e)

else:
    st.info("💡 Upload the Zoho payout export and the Bank of America statement to begin.")
