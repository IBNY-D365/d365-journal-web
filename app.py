
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
TEMP_RECEIPT_ACCOUNT = "21040102"
TEMP_RECEIPT_ACCOUNT_NAME = "Temporary Receipt"
TEMP_RECEIPT_ACCOUNT_TYPE = "Ledger"

# ─── FILE UPLOADERS ──────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    gateway_file = st.file_uploader(
        "1. Upload Zoho Payout Export (CSV/XLSX/PDF) or Stripe PDF",
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

def normalize_text(text: str) -> str:
    if text is None:
        return ""
    text = str(text)
    for old, new in {
        "\u202f": " ",
        "\xa0": " ",
        "−": "-",
        "–": "-",
        "—": "-",
        "…": "...",
    }.items():
        text = text.replace(old, new)
    return text


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", normalize_text(text)).strip()


def extract_text_from_pdf(f) -> str:
    try:
        f.seek(0)
        reader = PdfReader(f)
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception:
        return ""


def safe_float(v) -> float:
    try:
        return float(str(v).replace(",", "").replace("$", "").replace(" ", "").strip())
    except Exception:
        return 0.0


def extract_amounts(text: str):
    text = normalize_text(text)
    pat = re.compile(r"-?\$?\s*\d[\d,\s]*\.\s*\d{2}")
    vals = []
    for raw in pat.findall(text):
        cleaned = raw.replace("$", "").replace(" ", "").replace(",", "")
        cleaned = cleaned.replace("−", "-").replace("–", "-").replace("—", "-")
        try:
            vals.append(float(cleaned))
        except Exception:
            pass
    return vals


def first_amount(text: str):
    vals = extract_amounts(text)
    return vals[0] if vals else None


def clean_for_match(text) -> str:
    if pd.isna(text) or text is None:
        return ""
    t = str(text).lower()
    t = re.sub(r"\S+@\S+", "", t)
    t = re.sub(r"\b(llc|pllc|inc|corp|co|incorporated|limited|llp|dba)\b", " ", t)
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    return " ".join(t.split())


def lookup_customer(name, cust_df, acct_col, name_col):
    key = clean_for_match(name)
    if not key:
        return "MISSING_ACCT", name

    mask = cust_df["_name_clean"].apply(lambda x: key in x or x in key)
    hit = cust_df[mask]
    if not hit.empty:
        return str(hit.iloc[0][acct_col]).strip(), str(hit.iloc[0][name_col]).strip()

    key_tok = set(key.split())
    best, best_s = None, 0
    for _, row in cust_df.iterrows():
        row_tokens = set(str(row["_name_clean"]).split())
        s = len(key_tok & row_tokens)
        if s >= 2 and s > best_s:
            best_s, best = s, row
    if best is not None:
        return str(best[acct_col]).strip(), str(best[name_col]).strip()

    return "MISSING_ACCT", name


def cc_lookup(term, cc_df, term_col, code_col, fallback=""):
    key = clean_for_match(term)
    if not key:
        return fallback

    mask = cc_df["_term_clean"].str.contains(key, na=False)
    if not mask.any():
        mask = cc_df["_term_clean"].apply(lambda t: bool(t) and t in key)

    return str(cc_df[mask].iloc[0][code_col]).strip() if mask.any() else fallback


def parse_invoice_term(text) -> str:
    lower = normalize_text(text).lower()
    if "due on receipt" in lower:
        return "receipt"
    if any(k in lower for k in ("monthly payment", "monthly plan", "mpp", "installment")):
        return "monthly"
    return "receipt"


def parse_invoice_customer(text) -> str:
    lines = [normalize_spaces(l) for l in normalize_text(text).splitlines() if normalize_spaces(l)]
    for i, line in enumerate(lines):
        if re.search(r"bill\s*to|invoice\s*to", line, re.I) and i + 1 < len(lines):
            return lines[i + 1].strip()
    return ""


def parse_invoice_number(text) -> str:
    patterns = [
        r"INV-\d+",
        r"Invoice\s*#\s*(INV-\d+)",
        r"Purchase\s*#\s*(INV-\d+)",
    ]
    normalized = normalize_text(text)
    for pattern in patterns:
        match = re.search(pattern, normalized, re.I)
        if match:
            return match.group(1).strip() if match.groups() else match.group(0).strip()
    return ""


def parse_invoice_total(text) -> float | None:
    lines = [normalize_spaces(l) for l in normalize_text(text).splitlines() if normalize_spaces(l)]
    priority_labels = ["Total", "Sub Total", "Amount"]

    for label in priority_labels:
        for line in lines:
            if re.search(rf"\b{re.escape(label)}\b", line, re.I):
                amt = first_amount(line)
                if amt is not None and amt > 0:
                    return round(amt, 2)

    # fallback: choose the largest positive amount found in the document
    amounts = [a for a in extract_amounts(text) if a > 0]
    if amounts:
        return round(max(amounts), 2)
    return None


def build_invoice_catalog(invoice_files_list):
    catalog = []
    for inv in (invoice_files_list or []):
        name = getattr(inv, "name", "")
        if not name.lower().endswith(".pdf"):
            continue
        text = extract_text_from_pdf(inv)
        total = parse_invoice_total(text)
        catalog.append(
            {
                "file_name": name,
                "customer": parse_invoice_customer(text),
                "term": parse_invoice_term(text),
                "invoice_no": parse_invoice_number(text),
                "total": total,
                "used": False,
                "raw_text": text,
            }
        )
    return catalog


def match_invoice_to_amount(amount: float, catalog):
    if not catalog:
        return None

    candidates = []
    for idx, inv in enumerate(catalog):
        if inv["used"]:
            continue
        if inv["total"] is None:
            continue
        if abs(float(inv["total"]) - float(amount)) <= 0.01:
            candidates.append((idx, abs(float(inv["total"]) - float(amount))))

    if not candidates:
        return None

    idx = sorted(candidates, key=lambda x: x[1])[0][0]
    catalog[idx]["used"] = True
    return catalog[idx]


def parse_zoho_payout_pdf(text: str) -> pd.DataFrame:
    """
    Best-effort parser for Zoho Payout Details PDFs.
    The PDF text shows each payment row with gross, fee and net amounts.
    """
    lines = [normalize_spaces(l) for l in normalize_text(text).splitlines() if normalize_spaces(l)]

    chunks = []
    current = []
    for line in lines:
        if re.match(r"^Payment\b", line):
            if current:
                chunks.append(" ".join(current))
            current = [line]
        elif current:
            if line.startswith(("Payout:", "Payout ID:", "OVERVIEW", "Payout Summary", "All Transactions", "Export")):
                chunks.append(" ".join(current))
                current = []
            else:
                current.append(line)
    if current:
        chunks.append(" ".join(current))

    rows = []
    for chunk in chunks:
        if not re.match(r"^Payment\b", chunk):
            continue
        amounts = extract_amounts(chunk)
        if len(amounts) < 2:
            continue
        gross = round(abs(amounts[0]), 2)
        fee = round(abs(amounts[1]), 2)
        txn_type = "Payment"
        rows.append(
            {
                "Amount": gross,
                "Fee": fee,
                "CustomerName": "",
                "Description": chunk,
                "TransactionType": txn_type,
            }
        )
    return pd.DataFrame(rows)


def is_stripe_pdf(text: str) -> bool:
    lower = normalize_text(text).lower()
    return "stripe" in lower and any(k in lower for k in ("agreement", "plan", "charge"))


def parse_zoho_csv_or_xlsx(uploaded):
    if uploaded.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded)
    return pd.read_excel(uploaded, engine="openpyxl")


def make_row(
    date, company, acct_name, acct_type, acct_num,
    posting_profile, cash_code, description,
    debit, credit,
    off_company, off_acct_type, off_acct
) -> dict:
    return {
        "Date": date,
        "Voucher": "",
        "Account name": acct_name,
        "Company": company,
        "Account type": acct_type,
        "Account": acct_num,
        "Posting profile": posting_profile,
        "Cash code": cash_code,
        "Description": description,
        "Debit": debit,
        "Credit": credit,
        "Item sales tax group": "",
        "Sales tax code": "",
        "Offset company": off_company,
        "Offset account type": off_acct_type,
        "Offset account": off_acct,
        "Offset transaction text": "",
        "Currency": "USD",
        "Exchange rate": 1.00,
        "Item sales tax group2": "",
        "Sales tax group": "AVATAX",
        "Withholding tax group": "",
        "Release date": "",
        "Reversing entry": "No",
        "Reversing date": "",
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
    desc_lower = normalize_text(desc).lower()

    for kw, code in BOA_KEYWORD_MAP.items():
        if kw in desc_lower:
            return code

    tokens = re.sub(r"[^a-z0-9\s]", " ", desc_lower).split()[:3]
    for token in tokens:
        if len(token) < 4:
            continue
        result = cc_lookup(token, cc_df, term_col, code_col, fallback="")
        if result:
            return result
    return ""


def resolve_customer_account(invoice_rec, cust_df, acct_col, name_col):
    """
    Returns: (acct_type, acct_num, acct_name, is_temp_receipt)
    """
    if invoice_rec is None:
        return TEMP_RECEIPT_ACCOUNT_TYPE, TEMP_RECEIPT_ACCOUNT, TEMP_RECEIPT_ACCOUNT_NAME, True

    cust_name = invoice_rec.get("customer", "") or ""
    acct_num, resolved_name = lookup_customer(cust_name, cust_df, acct_col, name_col)

    if acct_num == "MISSING_ACCT":
        return TEMP_RECEIPT_ACCOUNT_TYPE, TEMP_RECEIPT_ACCOUNT, TEMP_RECEIPT_ACCOUNT_NAME, True

    return "Customer", acct_num, resolved_name, False


def build_credit_description(cash_code, acct_num, acct_name, boa_ref_desc, monthly=False):
    prefix = f"{cash_code}: "
    if monthly:
        prefix += "MPP "
    return f"{prefix}{acct_num} {acct_name}_{boa_ref_desc}"


def build_fee_description(acct_num, acct_name, boa_ref_desc):
    return f"Zoho Merchant Fee {acct_num}_{acct_name}_{boa_ref_desc}"


# ═════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ═════════════════════════════════════════════════════════════════════════════

if boa_file:
    st.subheader("4. Review & Generate")

    try:
        if not os.path.exists(MASTER_FILE_NAME) or not os.path.exists(CASH_CODE_FILE):
            st.error(
                f"❌ Missing lookup files. Ensure '{MASTER_FILE_NAME}' and '{CASH_CODE_FILE}' are committed to the repository."
            )
            st.stop()

        # ── Load Customer Master Account File ────────────────────────────────
        cust_df = pd.read_excel(MASTER_FILE_NAME, engine="openpyxl")
        cust_df.columns = [str(c).strip() for c in cust_df.columns]

        acct_col = next(
            (
                c for c in cust_df.columns
                if re.search(r"\baccount\b", c, re.I) and not re.search(r"name|type|desc", c, re.I)
            ),
            cust_df.columns[1],
        )
        name_col = next((c for c in cust_df.columns if re.search(r"name", c, re.I)), cust_df.columns[2])
        cust_df["_name_clean"] = cust_df[name_col].apply(clean_for_match)

        # ── Load Cash Code Masterlist ────────────────────────────────────────
        cc_raw = pd.read_excel(CASH_CODE_FILE, engine="openpyxl", header=None)
        header_row = 0
        for i, row in cc_raw.iterrows():
            if any(str(v).strip().lower() == "cash code" for v in row):
                header_row = i
                break

        cc_df = pd.read_excel(CASH_CODE_FILE, engine="openpyxl", skiprows=header_row)
        cc_df.columns = [str(c).strip() for c in cc_df.columns]
        cc_code_col = next((c for c in cc_df.columns if re.search(r"cash\s*code", c, re.I)), cc_df.columns[0])
        cc_name_col = next((c for c in cc_df.columns if re.search(r"name|desc", c, re.I)), cc_df.columns[1])
        cc_df = cc_df.dropna(subset=[cc_code_col])
        cc_df = cc_df[cc_df[cc_code_col].astype(str).str.strip() != "Cash Code"]
        cc_df["_term_clean"] = cc_df[cc_name_col].apply(clean_for_match)

        # ── Load Bank of America Statement ──────────────────────────────────
        if boa_file.name.lower().endswith(".csv"):
            raw_lines = boa_file.getvalue().decode("utf-8", errors="ignore").splitlines()
            skip = 0
            for line in raw_lines:
                if re.search(r"date.*description.*amount", line, re.I):
                    break
                skip += 1
            boa_file.seek(0)
            boa_df = pd.read_csv(boa_file, skiprows=skip)
        else:
            boa_df = pd.read_excel(boa_file, engine="openpyxl")

        boa_df.columns = [str(c).strip() for c in boa_df.columns]
        boa_date_col = next((c for c in boa_df.columns if re.search(r"date|post", c, re.I)), boa_df.columns[0])
        boa_desc_col = next((c for c in boa_df.columns if re.search(r"desc|text|memo", c, re.I)), boa_df.columns[1])
        boa_amt_col = next((c for c in boa_df.columns if re.search(r"amount|amt", c, re.I)), boa_df.columns[2])

        boa_df = boa_df[boa_df[boa_amt_col].notna()]
        boa_df = boa_df[boa_df[boa_amt_col].astype(str).str.strip() != ""]
        boa_df["_amt_float"] = boa_df[boa_amt_col].apply(safe_float)
        boa_df = boa_df[boa_df["_amt_float"] != 0.0].copy()

        # ── Detect engine mode ────────────────────────────────────────────────
        engine_mode = "BOA_ONLY"
        gateway_text = ""
        if gateway_file:
            if gateway_file.name.lower().endswith(".pdf"):
                gateway_text = extract_text_from_pdf(gateway_file)
                gateway_file.seek(0)
                if is_stripe_pdf(gateway_text):
                    engine_mode = "STRIPE"
                else:
                    engine_mode = "ZOHO"
            else:
                # For non-PDF uploads, default to Zoho unless the file clearly contains Stripe payment labels.
                temp_df = parse_zoho_csv_or_xlsx(gateway_file)
                gateway_file.seek(0)
                flat = " ".join([str(c) for c in temp_df.columns] + temp_df.astype(str).head(5).fillna("").agg(" ".join, axis=1).tolist())
                if "stripe" in flat.lower():
                    engine_mode = "STRIPE"
                else:
                    engine_mode = "ZOHO"

        boa_date = ""
        boa_ref_desc = ""
        boa_net_deposit = 0.0
        if not boa_df.empty:
            # Prefer a settlement row for the header date/description if one exists.
            if engine_mode == "STRIPE":
                settlement_mask = boa_df[boa_desc_col].astype(str).str.contains("STRIPE", case=False, na=False)
            elif engine_mode == "ZOHO":
                settlement_mask = boa_df[boa_desc_col].astype(str).str.contains("ZOHO PAYMENTS", case=False, na=False)
            else:
                settlement_mask = pd.Series(False, index=boa_df.index)

            if settlement_mask.any():
                boa_settlement = boa_df[settlement_mask].iloc[0]
            else:
                boa_settlement = boa_df.iloc[0]

            boa_date = str(boa_settlement[boa_date_col]).strip()
            boa_ref_desc = str(boa_settlement[boa_desc_col]).strip()
            boa_net_deposit = abs(safe_float(boa_settlement[boa_amt_col]))

        journal_rows = []

        # ─────────────────────────────────────────────────────────────────────
        # Invoice catalog used for Zoho amount matching
        # ─────────────────────────────────────────────────────────────────────
        invoice_catalog = build_invoice_catalog(invoice_files)

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE A — STRIPE PDF
        # ══════════════════════════════════════════════════════════════════════
        if engine_mode == "STRIPE":
            st.info("⚙️ Engine: Stripe PDF Extractor")

            if gateway_file.name.lower().endswith(".pdf"):
                pdf_text = gateway_text
            else:
                # If a non-PDF gateway file is used for Stripe, treat it as text-like via CSV/XLSX load.
                pdf_text = ""

            total_gross = 0.0
            charges = []

            # Stripe PDFs can have different formats; keep the parsing conservative.
            for line in [normalize_spaces(l) for l in pdf_text.splitlines() if normalize_spaces(l)]:
                ll = line.lower()
                if "charge" in ll and ("plan" in ll or "agreement" in ll):
                    amounts = extract_amounts(line)
                    if amounts:
                        gross = round(abs(amounts[0]), 2)
                        total_gross += gross
                        name = "Unknown Customer"
                        if "agreement -" in ll:
                            name = line.split("Agreement -")[-1].split(" -")[0].strip()
                        elif "plan -" in ll:
                            name = line.split("Plan -")[-1].split(" -")[0].strip()
                        name = re.sub(r"\S+@\S+", "", name).strip(" -")
                        charges.append({"name": name, "gross": gross, "is_mpp": "installment" in ll})

            for ch in charges:
                acct_type, acct_num, acct_name, is_temp = resolve_customer_account(
                    {"customer": ch["name"]}, cust_df, acct_col, name_col
                )
                cash_code = cc_lookup(
                    "Installment" if ch["is_mpp"] else "Receipt",
                    cc_df,
                    cc_name_col,
                    cc_code_col,
                    "AR002" if ch["is_mpp"] else "AR001",
                )
                desc_name = TEMP_RECEIPT_ACCOUNT_NAME if is_temp else acct_name
                credit_desc = build_credit_description(
                    cash_code, acct_num, desc_name, boa_ref_desc, monthly=ch["is_mpp"]
                )

                journal_rows.append(
                    make_row(
                        boa_date, company_id, desc_name, acct_type, acct_num,
                        "AutoPost" if not is_temp else "", cash_code, credit_desc,
                        "", f"{round(ch['gross'], 2):.2f}",
                        company_id, "Bank", offset_account,
                    )
                )

            fee = round(total_gross - boa_net_deposit, 2)
            if fee > 0:
                journal_rows.append(
                    make_row(
                        boa_date, company_id, "Outside Service (Finance)", "Ledger", debit_ledger_acct,
                        "", "OSF006", f"Stripe Merchant Fee_{boa_ref_desc}",
                        f"{round(fee, 2):.2f}", "",
                        company_id, "Bank", offset_account,
                    )
                )

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE B — ZOHO PAYOUT EXPORT
        # ══════════════════════════════════════════════════════════════════════
        elif engine_mode == "ZOHO":
            st.info("⚙️ Engine: Zoho Corporate Payment Pipeline")

            if gateway_file.name.lower().endswith(".pdf"):
                zoho_df = parse_zoho_payout_pdf(gateway_text)
                gateway_file.seek(0)
            else:
                zoho_df = parse_zoho_csv_or_xlsx(gateway_file)

            zoho_df.columns = [str(c).strip() for c in zoho_df.columns]

            gross_col = next(
                (c for c in zoho_df.columns if c.strip().lower() in ("amount", "gross amount")),
                next((c for c in zoho_df.columns if re.search(r"amount|gross", c, re.I)), zoho_df.columns[0]),
            )
            fee_col = next(
                (c for c in zoho_df.columns if c.strip().lower() == "fee"),
                next((c for c in zoho_df.columns if re.search(r"\bfee\b", c, re.I)), None),
            )
            cust_name_col = next((c for c in zoho_df.columns if re.search(r"customer.*name|customername", c, re.I)), None)
            txn_type_col = next((c for c in zoho_df.columns if re.search(r"transaction.*type|transactiontype", c, re.I)), None)
            desc_col = next((c for c in zoho_df.columns if re.search(r"desc|description|memo", c, re.I)), None)

            st.caption(
                f"Zoho columns → Gross: **{gross_col}** | Fee: **{fee_col or 'not found'}** | Customer: **{cust_name_col or 'not found'}**"
            )

            for _, zrow in zoho_df.iterrows():
                gross_raw = safe_float(zrow[gross_col])
                if gross_raw == 0.0:
                    continue

                fee_raw = safe_float(zrow[fee_col]) if fee_col else 0.0
                txn_type_val = str(zrow.get(txn_type_col, "")).lower() if txn_type_col else ""
                is_refund = gross_raw < 0 or any(k in txn_type_val for k in ("refund", "chargeback", "reversal", "void"))
                gross_amt = round(abs(gross_raw), 2)
                fee_amt = round(abs(fee_raw), 2)

                raw_desc = str(zrow.get(desc_col, "")) if desc_col else ""
                zoho_cust = ""

                if cust_name_col:
                    raw = str(zrow.get(cust_name_col, "")).strip()
                    if raw.lower() not in ("nan", "none", ""):
                        zoho_cust = raw

                invoice_rec = match_invoice_to_amount(gross_amt, invoice_catalog)
                if not zoho_cust and invoice_rec is not None:
                    zoho_cust = invoice_rec.get("customer", "")

                term = invoice_rec.get("term", "receipt") if invoice_rec else "receipt"

                acct_type, acct_num, acct_name, is_temp = resolve_customer_account(
                    {"customer": zoho_cust}, cust_df, acct_col, name_col
                )
                desc_name = TEMP_RECEIPT_ACCOUNT_NAME if is_temp else acct_name
                cash_code = cc_lookup(
                    "Installment" if term == "monthly" else ("Refund" if is_refund else "Receipt"),
                    cc_df,
                    cc_name_col,
                    cc_code_col,
                    "AR002" if term == "monthly" else ("AR003" if is_refund else "AR001"),
                )

                monthly = term == "monthly"
                credit_desc = build_credit_description(
                    cash_code, acct_num, desc_name, boa_ref_desc, monthly=monthly
                )

                if is_refund:
                    journal_rows.append(
                        make_row(
                            boa_date, company_id, desc_name, acct_type, acct_num,
                            "AutoPost" if not is_temp else "", cash_code,
                            f"REFUND {credit_desc}",
                            f"{gross_amt:.2f}", "",
                            company_id, "Bank", offset_account,
                        )
                    )
                else:
                    journal_rows.append(
                        make_row(
                            boa_date, company_id, desc_name, acct_type, acct_num,
                            "AutoPost" if not is_temp else "", cash_code,
                            credit_desc,
                            "", f"{gross_amt:.2f}",
                            company_id, "Bank", offset_account,
                        )
                    )

                if fee_amt > 0:
                    journal_rows.append(
                        make_row(
                            boa_date, company_id, "Outside Service (Finance)", "Ledger", debit_ledger_acct,
                            "", "OSF005", build_fee_description(acct_num, desc_name, boa_ref_desc),
                            f"{fee_amt:.2f}", "",
                            company_id, "Bank", offset_account,
                        )
                    )

        # ══════════════════════════════════════════════════════════════════════
        # ENGINE C — ALL OTHER BOA TRANSACTIONS
        # Temporary expense workflow: keep only Date, Description, Debit for debits.
        # ══════════════════════════════════════════════════════════════════════
        st.info("⚙️ Processing remaining BOA transactions…")

        if engine_mode == "BOA_ONLY":
            gateway_mask = pd.Series(False, index=boa_df.index)
        else:
            gateway_desc_pattern = "ZOHO PAYMENTS" if engine_mode == "ZOHO" else "STRIPE"
            gateway_mask = boa_df[boa_desc_col].astype(str).str.contains(gateway_desc_pattern, case=False, na=False)

        for _, brow in boa_df[~gateway_mask].iterrows():
            raw_desc = str(brow[boa_desc_col]).strip()
            raw_date = str(brow[boa_date_col]).strip()
            raw_amt = brow["_amt_float"]

            if raw_desc in ("", "nan") or raw_amt == 0.0:
                continue
            if re.search(r"beginning balance|ending balance", raw_desc, re.I):
                continue

            if raw_amt < 0:
                blank_row = {col: "" for col in COLUMNS_25}
                blank_row["Date"] = raw_date
                blank_row["Description"] = raw_desc
                blank_row["Debit"] = f"{round(abs(raw_amt), 2):.2f}"
                blank_row["Offset account type"] = "Bank"
                journal_rows.append(blank_row)
            else:
                cash_code = categorise_boa_expense(raw_desc, cc_df, cc_name_col, cc_code_col)
                journal_rows.append(
                    make_row(
                        raw_date, company_id,
                        "Outside Service (Finance)", "Ledger", debit_ledger_acct,
                        "", cash_code,
                        raw_desc,
                        "", f"{round(raw_amt, 2):.2f}",
                        company_id, "Bank", offset_account,
                    )
                )

        # ══════════════════════════════════════════════════════════════════════
        # OUTPUT
        # ══════════════════════════════════════════════════════════════════════
        final_df = pd.DataFrame(journal_rows)

        if not final_df.empty:
            final_df = final_df.reindex(columns=COLUMNS_25).fillna("")

            # Force two decimal places for all monetary columns
            for col in ["Debit", "Credit"]:
                final_df[col] = pd.to_numeric(final_df[col], errors="coerce")
                final_df[col] = final_df[col].apply(lambda x: "" if pd.isna(x) else f"{x:.2f}")

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

            if engine_mode == "BOA_ONLY":
                st.info(
                    "BOA-only expense staging mode is active. Negative BOA transactions are exported with only Date, Description, Debit, and Offset account type = Bank."
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
