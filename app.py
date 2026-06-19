import os
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from pypdf import PdfReader

# ============================================================
# Page setup
# ============================================================
st.set_page_config(page_title="D365 Accounting Journal Generator", layout="wide")
st.title("D365 Transaction Journal Generator")
st.write("Upload your Bank of America statement plus any gateway/invoice files for the day.")

# ============================================================
# Config
# ============================================================
BASE_DIR = Path(".")
CUSTOMER_MASTER_FILE = BASE_DIR / "Customer Master Account File.xlsx"
CASH_CODE_FILE = BASE_DIR / "Cash Code Masterlist.xlsx"
MONTHLY_EXPENSE_FILE = BASE_DIR / "Monthly Expense Record.xlsx"
FORM_MASTER_FILE = BASE_DIR / "Form_Master_DB.xlsx"

SOURCE_TO_OFFSET = {
    "3371": "B1000002",
    "3924": "B1000003",
    "3384": "B1000001",
}

DEFAULT_OFFSET = "B1000002"
DEFAULT_COMPANY = "bwa"
DEFAULT_LEDGER_ACCOUNT = "43170111-U26C05001-B735350-UOA003"
TEMP_RECEIPT_ACCOUNT = "21040102"
TEMP_RECEIPT_NAME = "Temporary Receipt"

AR_CODES = {
    "receipt": "AR001",
    "monthly": "AR002",
    "financing": "AR003",
    "leasing": "AR004",
    "net_1day": "AR005",
    "net_10days": "AR006",
    "net_25days": "AR007",
    "net_30days": "AR008",
    "net_40days": "AR009",
    "net_45days": "AR010",
    "net_60days": "AR011",
    "other": "AR012",
}

PAYMENT_FEE_CODES = {
    "zoho": "OSF005",
    "stripe": "OSF006",
    "bankcard": "OSF007",
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

# ============================================================
# Uploaders
# ============================================================
col1, col2, col3 = st.columns(3)
with col1:
    gateway_file = st.file_uploader(
        "1. Upload Zoho / Stripe / Bankcard file (PDF, CSV, XLSX)",
        type=["pdf", "csv", "xlsx"],
    )
with col2:
    invoice_files = st.file_uploader(
        "2. Upload invoice files (PDF, CSV, XLSX, TXT)",
        type=["pdf", "csv", "xlsx", "txt"],
        accept_multiple_files=True,
    )
with col3:
    boa_file = st.file_uploader(
        "3. Upload Bank of America statement (CSV, XLSX)",
        type=["csv", "xlsx"],
    )

st.sidebar.header("D365 Defaults")
company_id = st.sidebar.text_input("Company", value=DEFAULT_COMPANY)
offset_account_default = st.sidebar.text_input("Default Offset Account", value=DEFAULT_OFFSET)
debit_ledger_acct = st.sidebar.text_input("Debit Line Account (Ledger)", value=DEFAULT_LEDGER_ACCOUNT)

# ============================================================
# Helpers
# ============================================================
def normalize_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    txt = str(value).lower()
    txt = txt.replace("\u2212", "-").replace("—", "-").replace("–", "-")
    txt = re.sub(r"\S+@\S+", " ", txt)
    txt = re.sub(r"\b(llc|inc|corp|co|pllc|llp|dba|limited|incorporated)\b", " ", txt)
    txt = re.sub(r"[^a-z0-9\s\-_.]", " ", txt)
    return " ".join(txt.split())

def clean_for_match(value) -> str:
    return normalize_text(value).replace("-", " ")

def safe_float(value) -> float:
    try:
        txt = str(value).replace("$", "").replace(",", "").replace("(", "-").replace(")", "").strip()
        txt = txt.replace("\u2212", "-").replace("—", "-").replace("–", "-")
        return float(txt)
    except Exception:
        return 0.0

def money2(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return f"{safe_float(value):.2f}"

def read_pdf_text(file_obj) -> str:
    try:
        file_obj.seek(0)
        return "\n".join(page.extract_text() or "" for page in PdfReader(file_obj).pages)
    except Exception:
        return ""

def resolve_offset_account(text: str, default_value: str) -> str:
    t = normalize_text(text)
    for src, acct in SOURCE_TO_OFFSET.items():
        if src in t:
            return acct
    return default_value

def detect_header_row(raw_df: pd.DataFrame, keywords: list[str], max_rows: int = 20) -> int:
    for idx in range(min(max_rows, len(raw_df))):
        row = [normalize_text(v) for v in raw_df.iloc[idx].tolist()]
        if sum(any(k in cell for k in keywords) for cell in row) >= 2:
            return idx
    return 0

def load_table(file_path_or_buffer, preferred_sheet: str | None = None) -> pd.DataFrame:
    if str(file_path_or_buffer).lower().endswith(".csv"):
        return pd.read_csv(file_path_or_buffer)
    if preferred_sheet:
        return pd.read_excel(file_path_or_buffer, sheet_name=preferred_sheet, engine="openpyxl")
    return pd.read_excel(file_path_or_buffer, engine="openpyxl")

def detect_col(df: pd.DataFrame, include: list[str], exclude: list[str] | None = None, fallback_index: int | None = None):
    exclude = exclude or []
    for col in df.columns:
        n = normalize_text(col)
        if any(i in n for i in include) and not any(e in n for e in exclude):
            return col
    if fallback_index is not None and len(df.columns) > fallback_index:
        return df.columns[fallback_index]
    return None

def load_customer_master():
    df = load_table(CUSTOMER_MASTER_FILE)
    df.columns = [str(c).strip() for c in df.columns]
    acct_col = detect_col(df, ["account"], ["type", "name", "desc"], 1)
    name_col = detect_col(df, ["name"], [], 2)
    if acct_col is None or name_col is None:
        raise ValueError("Unable to detect customer master columns.")
    df["_clean_name"] = df[name_col].map(clean_for_match)
    return df, acct_col, name_col

def lookup_customer(customer_name: str, cust_df: pd.DataFrame, acct_col: str, name_col: str):
    key = clean_for_match(customer_name)
    if not key:
        return None
    exact = cust_df[cust_df["_clean_name"].eq(key)]
    if not exact.empty:
        row = exact.iloc[0]
        return {
            "account": str(row[acct_col]).strip(),
            "name": str(row[name_col]).strip(),
            "found": True,
        }
    contains = cust_df[cust_df["_clean_name"].apply(lambda x: key in x or x in key)]
    if not contains.empty:
        row = contains.iloc[0]
        return {
            "account": str(row[acct_col]).strip(),
            "name": str(row[name_col]).strip(),
            "found": True,
        }
    key_tokens = set(key.split())
    best_row = None
    best_score = 0
    for _, row in cust_df.iterrows():
        score = len(key_tokens & set(str(row["_clean_name"]).split()))
        if score > best_score:
            best_score = score
            best_row = row
    if best_row is not None and best_score >= 2:
        return {
            "account": str(best_row[acct_col]).strip(),
            "name": str(best_row[name_col]).strip(),
            "found": True,
        }
    return None

def load_cash_codes():
    raw = pd.read_excel(CASH_CODE_FILE, engine="openpyxl", header=None)
    hdr = detect_header_row(raw, ["cash code", "code", "name"])
    df = pd.read_excel(CASH_CODE_FILE, engine="openpyxl", skiprows=hdr)
    df.columns = [str(c).strip() for c in df.columns]
    code_col = detect_col(df, ["cash code"], [], 0)
    name_col = detect_col(df, ["name", "desc"], [], 1)
    if code_col is None or name_col is None:
        raise ValueError("Unable to detect cash code columns.")
    df = df[df[code_col].notna()].copy()
    df["_clean_name"] = df[name_col].map(clean_for_match)
    return df, code_col, name_col

def cash_code_for_term(term: str) -> str:
    t = normalize_text(term)
    if any(k in t for k in ["due on receipt", "due upon receipt", "receipt"]):
        return AR_CODES["receipt"]
    if any(k in t for k in ["monthly payment", "monthly plan", "mpp", "installment"]):
        return AR_CODES["monthly"]
    if "financing" in t:
        return AR_CODES["financing"]
    if "leasing" in t:
        return AR_CODES["leasing"]
    if "1 day" in t or "1day" in t or "net 1" in t:
        return AR_CODES["net_1day"]
    if "10 day" in t or "10days" in t or "net 10" in t:
        return AR_CODES["net_10days"]
    if "25 day" in t or "25days" in t or "net 25" in t:
        return AR_CODES["net_25days"]
    if "30 day" in t or "30days" in t or "net 30" in t:
        return AR_CODES["net_30days"]
    if "40 day" in t or "40days" in t or "net 40" in t:
        return AR_CODES["net_40days"]
    if "45 day" in t or "45days" in t or "net 45" in t:
        return AR_CODES["net_45days"]
    if "60 day" in t or "60days" in t or "net 60" in t:
        return AR_CODES["net_60days"]
    return AR_CODES["other"]

def load_form_master():
    if not FORM_MASTER_FILE.exists():
        return None
    try:
        df = pd.read_excel(FORM_MASTER_FILE, sheet_name="Sales_PRF", engine="openpyxl")
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    return df

def load_monthly_rules():
    if not MONTHLY_EXPENSE_FILE.exists():
        return None
    try:
        raw = pd.read_excel(MONTHLY_EXPENSE_FILE, sheet_name=0, header=None, engine="openpyxl")
    except Exception:
        return None
    hdr = detect_header_row(raw, ["account", "cash", "description", "pattern", "amount"])
    df = pd.read_excel(MONTHLY_EXPENSE_FILE, sheet_name=0, skiprows=hdr, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df

def detect_rule_columns(df: pd.DataFrame):
    cols = {
        "pattern": detect_col(df, ["pattern", "description", "desc", "vendor", "merchant", "boa"], [], 0),
        "amount": detect_col(df, ["amount", "debit", "charge"], [], None),
        "account_name": detect_col(df, ["account name", "name"], ["description"], None),
        "account_type": detect_col(df, ["account type", "type"], None, None),
        "account": detect_col(df, ["account"], ["name", "type", "desc"], None),
        "cash_code": detect_col(df, ["cash code", "cashcode"], [], None),
        "desc_template": detect_col(df, ["description template", "description", "template", "memo"], [], None),
    }
    return cols

def monthly_rule_match(raw_desc: str, raw_amt: float, rules_df: pd.DataFrame):
    if rules_df is None or rules_df.empty:
        return None
    cols = detect_rule_columns(rules_df)
    desc_n = normalize_text(raw_desc)
    for _, row in rules_df.iterrows():
        pattern = str(row.get(cols["pattern"], "") if cols["pattern"] else "").strip()
        if not pattern or pattern.lower() in ["nan", "none"]:
            continue
        pat_n = normalize_text(pattern)
        if pat_n and pat_n not in desc_n:
            continue
        rule_amt = None
        if cols["amount"] and str(row.get(cols["amount"], "")).strip() not in ["", "nan", "None"]:
            rule_amt = safe_float(row.get(cols["amount"]))
            if rule_amt and abs(rule_amt - abs(raw_amt)) > 0.02:
                continue
        return {
            "account_name": str(row.get(cols["account_name"], "")).strip() if cols["account_name"] else "",
            "account_type": str(row.get(cols["account_type"], "")).strip() if cols["account_type"] else "",
            "account": str(row.get(cols["account"], "")).strip() if cols["account"] else "",
            "cash_code": str(row.get(cols["cash_code"], "")).strip() if cols["cash_code"] else "",
            "desc_template": str(row.get(cols["desc_template"], "")).strip() if cols["desc_template"] else "",
        }
    return None

def build_row(date, company, acct_name, acct_type, account, posting_profile, cash_code, description, debit, credit, offset_account_type="Bank", offset_account_value="", bank_desc_text=""):
    return {
        "Date": date,
        "Voucher": "",
        "Account name": acct_name,
        "Company": company,
        "Account type": acct_type,
        "Account": account,
        "Posting profile": posting_profile,
        "Cash code": cash_code,
        "Description": description,
        "Debit": debit,
        "Credit": credit,
        "Item sales tax group": "",
        "Sales tax code": "",
        "Offset company": company,
        "Offset account type": offset_account_type,
        "Offset account": offset_account_value,
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

def parse_invoice_file(file_obj):
    name = getattr(file_obj, "name", "invoice")
    lower = name.lower()
    text = ""
    if lower.endswith(".pdf"):
        text = read_pdf_text(file_obj)
    else:
        try:
            if lower.endswith(".csv"):
                df = pd.read_csv(file_obj)
                text = "\n".join(df.astype(str).fillna("").agg(" ".join, axis=1).tolist())
            else:
                df = pd.read_excel(file_obj, engine="openpyxl")
                text = "\n".join(df.astype(str).fillna("").agg(" ".join, axis=1).tolist())
        except Exception:
            try:
                text = file_obj.read().decode("utf-8", errors="ignore")
            except Exception:
                text = ""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    customer = ""
    for i, line in enumerate(lines):
        if re.search(r"bill\s*to|invoice\s*to", line, re.I) and i + 1 < len(lines):
            customer = lines[i + 1].strip()
            break
    total = None
    total_patterns = [
        r"\bTotal\b[^\d]{0,15}\$?([\d,]+\.\d{2})",
        r"\bBalance Due\b[^\d]{0,15}\$?([\d,]+\.\d{2})",
        r"\bPayment Made\b[^\d]{0,15}\(?-?\$?([\d,]+\.\d{2})\)?",
    ]
    for pat in total_patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            total = safe_float(m.group(1))
            break
    inv_no = ""
    m = re.search(r"INV-\d+", text, re.I)
    if m:
        inv_no = m.group(0).upper()
    term = ""
    t = normalize_text(text)
    if "due on receipt" in t or "due upon receipt" in t:
        term = "receipt"
    elif any(k in t for k in ["monthly payment", "monthly plan", "mpp", "installment"]):
        term = "monthly"
    elif "financing" in t:
        term = "financing"
    elif "leasing" in t:
        term = "leasing"
    elif "net 1 day" in t or "net 1" in t:
        term = "net_1day"
    elif "net 10" in t:
        term = "net_10days"
    elif "net 25" in t:
        term = "net_25days"
    elif "net 30" in t:
        term = "net_30days"
    elif "net 40" in t:
        term = "net_40days"
    elif "net 45" in t:
        term = "net_45days"
    elif "net 60" in t:
        term = "net_60days"
    else:
        term = "receipt"
    return {
        "file": name,
        "text": text,
        "customer": customer,
        "total": round(total, 2) if total is not None else None,
        "invoice_no": inv_no,
        "term": term,
        "used": False,
    }

def build_invoice_index(files):
    invs = []
    for f in files or []:
        invs.append(parse_invoice_file(f))
    return invs

def pick_invoice_by_amount(amount, invoices):
    amt = round(abs(float(amount)), 2)
    candidates = [i for i in invoices if i.get("total") is not None and round(i["total"], 2) == amt and not i["used"]]
    if candidates:
        candidates.sort(key=lambda x: (x.get("customer", "") == "", x.get("invoice_no", "") == ""))
        chosen = candidates[0]
        chosen["used"] = True
        return chosen
    return None


def resolve_payment_info(invoice_meta, form_df, customer_name, invoice_no, customer_account=""):
    """
    Resolve the Zoho cash code (AR001-AR012) and the normalized payment term.

    The form master is treated as the authoritative source when it contains a
    matching row and an explicit AR code or recognizable payment-term wording.
    """
    invoice_term = ""
    if invoice_meta and invoice_meta.get("term"):
        invoice_term = str(invoice_meta.get("term")).strip().lower()

    def term_to_code(term: str) -> str:
        return cash_code_for_term(term or "receipt")

    fallback_term = invoice_term or "receipt"
    fallback_code = term_to_code(fallback_term)

    if form_df is None or form_df.empty:
        return {"term": fallback_term, "cash_code": fallback_code}

    search_keys = [
        clean_for_match(customer_account),
        clean_for_match(customer_name),
        normalize_text(invoice_no),
    ]
    best_text = ""
    best_score = 0

    for _, row in form_df.iterrows():
        row_text = normalize_text(" ".join(str(v) for v in row.tolist()))
        score = sum(1 for k in search_keys if k and k in row_text)
        if score > best_score:
            best_score = score
            best_text = row_text

    if not best_text:
        return {"term": fallback_term, "cash_code": fallback_code}

    # Explicit AR code in the matched row wins.
    for term_key, code in AR_CODES.items():
        if normalize_text(code) in best_text:
            return {"term": term_key, "cash_code": code}

    # Infer the term from the matched row text.
    if any(k in best_text for k in ["monthly", "mpp", "installment"]):
        return {"term": "monthly", "cash_code": AR_CODES["monthly"]}
    if "financing" in best_text:
        return {"term": "financing", "cash_code": AR_CODES["financing"]}
    if "leasing" in best_text:
        return {"term": "leasing", "cash_code": AR_CODES["leasing"]}
    if "net 1 day" in best_text or "net 1" in best_text or "1 day" in best_text:
        return {"term": "net_1day", "cash_code": AR_CODES["net_1day"]}
    if "net 10 days" in best_text or "net 10" in best_text or "10 day" in best_text:
        return {"term": "net_10days", "cash_code": AR_CODES["net_10days"]}
    if "net 25 days" in best_text or "net 25" in best_text or "25 day" in best_text:
        return {"term": "net_25days", "cash_code": AR_CODES["net_25days"]}
    if "net 30 days" in best_text or "net 30" in best_text or "30 day" in best_text:
        return {"term": "net_30days", "cash_code": AR_CODES["net_30days"]}
    if "net 40 days" in best_text or "net 40" in best_text or "40 day" in best_text:
        return {"term": "net_40days", "cash_code": AR_CODES["net_40days"]}
    if "net 45 days" in best_text or "net 45" in best_text or "45 day" in best_text:
        return {"term": "net_45days", "cash_code": AR_CODES["net_45days"]}
    if "net 60 days" in best_text or "net 60" in best_text or "60 day" in best_text:
        return {"term": "net_60days", "cash_code": AR_CODES["net_60days"]}

    return {"term": fallback_term, "cash_code": fallback_code}



def detect_gateway_kind(file_obj):
    text = ""
    lower_name = file_obj.name.lower()

    if lower_name.endswith(".pdf"):
        text = read_pdf_text(file_obj)
    else:
        try:
            if lower_name.endswith(".csv"):
                text = pd.read_csv(file_obj).astype(str).fillna("").to_string(index=False)
            else:
                text = pd.read_excel(file_obj, engine="openpyxl").astype(str).fillna("").to_string(index=False)
        except Exception:
            text = ""

    t = normalize_text(text)

    # Explicit indicators first.
    if "bankcard" in t or "authorize" in t or "authorizenet" in t or "mtot disc" in t:
        return "bankcard", text

    if "stripe" in t or ("agreement" in t and "plan" in t):
        return "stripe", text

    # Zoho payouts often contain "Payout Summary" or "Payout Descriptor: ZOHO PAYMENTS".
    if "zoho payments" in t or "payout summary" in t or "payout descriptor" in t or "pay.zoho.com" in t:
        return "zoho", text

    # Default to Zoho for the generic gateway upload unless Stripe/Bankcard is detected.
    return "zoho", text



def parse_zoho_payout_pdf_text(text: str):
    """
    Extract every payment row from a Zoho payout PDF.

    The expected payout PDF contains one or more rows in the form:
      Payment Jun 16, 2026, 02:05 PM — ... $1,908.07 −$55.63 $1,852.44 USD

    The parser returns one dict per payment with gross and fee values.
    """
    txns = []
    if not text:
        return txns

    normalized = (
        text.replace("−", "-")
            .replace("—", " - ")
            .replace("–", " - ")
    )
    normalized = re.sub(r"\s+", " ", normalized).strip()

    # Primary pattern: the payment table row with gross, fee and total.
    row_pattern = re.compile(
        r"Payment\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4},\s+\d{2}:\d{2}\s+[AP]M\s+[-—]\s+"
        r"(?P<desc>.*?)\s+\$?(?P<gross>[\d,]+\.\d{2})\s+[-−]\s*\$?(?P<fee>[\d,]+\.\d{2})\s+\$?(?P<net>[\d,]+\.\d{2})\s+USD",
        re.I,
    )

    for m in row_pattern.finditer(normalized):
        txns.append(
            {
                "gross": round(abs(safe_float(m.group("gross"))), 2),
                "fee": round(abs(safe_float(m.group("fee"))), 2),
                "description": m.group("desc").strip(),
                "customer_name": "",
                "type": "payment",
                "date": "",
            }
        )

    if txns:
        return txns

    # Fallback: parse any line that starts with Payment and has at least 3 money fields.
    for line in [l.strip() for l in text.splitlines() if l.strip()]:
        line_norm = normalize_text(line)
        if not line_norm.startswith("payment"):
            continue
        amounts = re.findall(r"-?\$?[\d,]+\.\d{2}", line)
        if len(amounts) < 3:
            continue
        txns.append(
            {
                "gross": round(abs(safe_float(amounts[-3])), 2),
                "fee": round(abs(safe_float(amounts[-2])), 2),
                "description": line,
                "customer_name": "",
                "type": "payment",
                "date": "",
            }
        )
    return txns


def parse_gateway_pdf_text(text, kind):
    if kind == "zoho":
        return parse_zoho_payout_pdf_text(text)

    txns = []
    normalized = text.replace("−", "-").replace("—", "-").replace("–", "-")
    for line in [l.strip() for l in normalized.splitlines() if l.strip()]:
        if not re.search(r"\$?\d[\d,]*\.\d{2}", line):
            continue
        money = re.findall(r"-?\$?\d[\d,]*\.\d{2}", line)
        if len(money) < 1:
            continue
        gross = abs(safe_float(money[0]))
        fee = abs(safe_float(money[1])) if len(money) > 1 else 0.0
        if kind in ["stripe", "bankcard"] and not any(k in normalize_text(line) for k in ["charge", "payment", "agreement", "plan", "bankcard", "authorize"]):
            continue
        txns.append(
            {
                "gross": gross,
                "fee": fee,
                "description": line,
                "customer_name": "",
                "type": "payment",
                "date": "",
            }
        )
    return txns


def parse_gateway_file(file_obj):
    kind, text = detect_gateway_kind(file_obj)
    file_obj.seek(0)
    if file_obj.name.lower().endswith(".pdf"):
        txns = parse_gateway_pdf_text(text, kind)
    else:
        txns = parse_gateway_tabular(file_obj)
    return kind, txns, text


def build_payment_rows(kind, boarow, txns, invoices, cust_df, acct_col, name_col, form_df, offset_default):
    rows = []
    fee_total = 0.0
    unique_tags = []

    for txn in txns:
        gross = round(abs(float(txn["gross"])), 2)
        fee = round(abs(float(txn.get("fee", 0.0))), 2)
        fee_total += fee

        invoice = pick_invoice_by_amount(gross, invoices)
        customer_name = ""
        invoice_no = ""

        if invoice:
            customer_name = invoice.get("customer", "") or ""
            invoice_no = invoice.get("invoice_no", "") or ""
        if not customer_name:
            customer_name = txn.get("customer_name", "") or ""
        if not customer_name:
            customer_name = "Unknown"

        matched = lookup_customer(customer_name, cust_df, acct_col, name_col)

        if matched:
            account_name = matched["name"]
            account_no = matched["account"]
            account_type = "Customer"

            payment_info = resolve_payment_info(
                invoice_meta=invoice,
                form_df=form_df,
                customer_name=account_name,
                invoice_no=invoice_no,
                customer_account=account_no,
            )
            cash_code = payment_info["cash_code"]
            term = payment_info["term"]
            desc_prefix = "MPP " if term == "monthly" else ""
            desc = f"{desc_prefix}{account_no} {account_name}_{boarow['desc']}"
        else:
            account_name = TEMP_RECEIPT_NAME
            account_no = TEMP_RECEIPT_ACCOUNT
            account_type = "Ledger"
            cash_code = AR_CODES["receipt"]
            term = "receipt"
            desc = f"{TEMP_RECEIPT_NAME}_{boarow['desc']}"

        unique_tags.append(f"{account_no} {account_name}")

        rows.append(
            build_row(
                date=boarow["date"],
                company=company_id,
                acct_name=account_name,
                acct_type=account_type,
                account=account_no,
                posting_profile="AutoPost",
                cash_code=cash_code,
                description=desc,
                debit="",
                credit=gross,
                offset_account_type="Bank",
                offset_account_value=resolve_offset_account(boarow["desc"], offset_default),
            )
        )

    fee_desc_base = f"{kind.capitalize()} Merchant Fee "
    fee_desc = fee_desc_base + ", ".join(dict.fromkeys(unique_tags)) + f"_{boarow['desc']}"
    fee_code = PAYMENT_FEE_CODES.get(kind, "OSF005")

    if fee_total == 0 and txns:
        gross_sum = round(sum(abs(float(t["gross"])) for t in txns), 2)
        fee_total = round(max(0.0, gross_sum - abs(float(boarow["amount"]))), 2)

    if fee_total > 0:
        rows.append(
            build_row(
                date=boarow["date"],
                company=company_id,
                acct_name="Outside Service (Finance)",
                acct_type="Ledger",
                account=debit_ledger_acct if kind != "bankcard" else DEFAULT_LEDGER_ACCOUNT,
                posting_profile="",
                cash_code=fee_code,
                description=fee_desc,
                debit=fee_total,
                credit="",
                offset_account_type="Bank",
                offset_account_value=resolve_offset_account(boarow["desc"], offset_default),
            )
        )

    return rows


def build_monthly_and_other_rows(boa_df, skip_mask, monthly_rules, offset_default):
    rows = []
    for _, r in boa_df.loc[~skip_mask].iterrows():
        raw_amt = safe_float(r["_amt_float"])
        if raw_amt == 0:
            continue
        raw_desc = str(r["desc"]).strip()
        raw_date = str(r["date"]).strip()
        if raw_amt < 0:
            match = monthly_rule_match(raw_desc, abs(raw_amt), monthly_rules)
            if match:
                acct_name = match.get("account_name") or ""
                acct_type = match.get("account_type") or "Ledger"
                account_no = match.get("account") or ""
                cash_code = match.get("cash_code") or ""
                template = match.get("desc_template") or ""
                desc = f"{template}{raw_desc}" if template else raw_desc
                rows.append(build_row(
                    date=raw_date,
                    company=company_id,
                    acct_name=acct_name,
                    acct_type=acct_type,
                    account=account_no,
                    posting_profile="",
                    cash_code=cash_code,
                    description=desc,
                    debit=abs(raw_amt),
                    credit="",
                    offset_account_type="Bank",
                    offset_account_value=resolve_offset_account(raw_desc, offset_default),
                ))
            else:
                rows.append(build_row(
                    date=raw_date,
                    company=company_id,
                    acct_name="",
                    acct_type="",
                    account="",
                    posting_profile="",
                    cash_code="",
                    description=raw_desc,
                    debit=abs(raw_amt),
                    credit="",
                    offset_account_type="Bank",
                    offset_account_value=resolve_offset_account(raw_desc, offset_default),
                ))
        else:
            # positive BOA amounts are preserved as-is but with blank target fields
            rows.append(build_row(
                date=raw_date,
                company=company_id,
                acct_name="",
                acct_type="",
                account="",
                posting_profile="",
                cash_code="",
                description=raw_desc,
                debit="",
                credit=abs(raw_amt),
                offset_account_type="Bank",
                offset_account_value=resolve_offset_account(raw_desc, offset_default),
            ))
    return rows

def format_export_df(df):
    df = df.reindex(columns=COLUMNS_25).fillna("")
    for col in ["Debit", "Credit", "Exchange rate"]:
        if col in df.columns:
            def _fmt(v):
                if v == "" or pd.isna(v):
                    return ""
                try:
                    return f"{float(v):.2f}"
                except Exception:
                    return ""
            df[col] = df[col].map(_fmt)
    return df

# ============================================================
# Main
# ============================================================
if boa_file:
    try:
        if not CUSTOMER_MASTER_FILE.exists():
            st.error(f"Missing {CUSTOMER_MASTER_FILE.name}")
            st.stop()
        if not CASH_CODE_FILE.exists():
            st.error(f"Missing {CASH_CODE_FILE.name}")
            st.stop()

        cust_df, cust_acct_col, cust_name_col = load_customer_master()
        cash_df, cash_code_col, cash_name_col = load_cash_codes()
        monthly_rules = load_monthly_rules()
        form_df = load_form_master()

        # BOA load
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
        date_col = detect_col(boa_df, ["date", "post"], [], 0)
        desc_col = detect_col(boa_df, ["desc", "memo", "text", "description"], [], 1)
        amt_col = detect_col(boa_df, ["amount", "amt", "debit", "credit"], [], 2)
        if date_col is None or desc_col is None or amt_col is None:
            st.error("Could not detect BOA date/description/amount columns.")
            st.stop()

        boa_df = boa_df[boa_df[amt_col].notna()].copy()
        boa_df["_amt_float"] = boa_df[amt_col].map(safe_float)
        boa_df = boa_df[boa_df["_amt_float"] != 0].copy()
        boa_df = boa_df.rename(columns={date_col: "date", desc_col: "desc", amt_col: "amount"})

        gateway_kind = None
        gateway_txns = []
        gateway_text = ""
        if gateway_file:
            gateway_kind, gateway_txns, gateway_text = parse_gateway_file(gateway_file)

        settlement_pattern = {
            "zoho": "ZOHO PAYMENTS",
            "stripe": "STRIPE",
            "bankcard": "BANKCARD",
        }.get(gateway_kind or "", "")
        if settlement_pattern:
            gateway_mask = boa_df["desc"].astype(str).str.contains(settlement_pattern, case=False, na=False)
        else:
            gateway_mask = pd.Series(False, index=boa_df.index)

        journal_rows = []

        # Payment handler
        if gateway_kind in ["zoho", "stripe", "bankcard"] and gateway_txns:
            settlement_rows = boa_df.loc[gateway_mask].copy()
            if settlement_rows.empty:
                settlement_rows = boa_df.head(1).copy()

            # Use first settlement row for date/offset, but keep each BOA gateway row isolated if multiple found
            settle = settlement_rows.iloc[0]
            boarow = {
                "date": str(settle["date"]).strip(),
                "desc": str(settle["desc"]).strip(),
                "amount": safe_float(settle["_amt_float"]),
            }
            payment_rows = build_payment_rows(
                kind=gateway_kind,
                boarow=boarow,
                txns=gateway_txns,
                invoices=build_invoice_index(invoice_files),
                cust_df=cust_df,
                acct_col=cust_acct_col,
                name_col=cust_name_col,
                form_df=form_df,
                offset_default=offset_account_default,
            )
            journal_rows.extend(payment_rows)

        # Independent expense tracks
        journal_rows.extend(build_monthly_and_other_rows(boa_df, gateway_mask, monthly_rules, offset_account_default))

        final_df = pd.DataFrame(journal_rows)
        if final_df.empty:
            st.warning("No journal rows were produced.")
            st.stop()

        final_df = format_export_df(final_df)

        total_debit = pd.to_numeric(final_df["Debit"], errors="coerce").fillna(0).sum()
        total_credit = pd.to_numeric(final_df["Credit"], errors="coerce").fillna(0).sum()
        diff = abs(total_debit - total_credit)
        missing = (final_df["Account"] == "").sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Debits", f"${total_debit:,.2f}")
        c2.metric("Total Credits", f"${total_credit:,.2f}")
        c3.metric("Balance Difference", f"${diff:,.2f}", delta="✅ Balanced" if diff < 0.02 else "⚠️ Out of balance")

        if (final_df["Account name"] == "").any():
            st.info("Some fallback rows were intentionally left blank for downstream allocation.")

        st.success("All transactions processed.")
        st.dataframe(final_df, use_container_width=True)

        st.download_button(
            "Download D365 Upload CSV",
            data=final_df.to_csv(index=False).encode("utf-8"),
            file_name="D365_Reconciliation_Journal.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"Pipeline error: {e}")
        st.exception(e)
else:
    st.info("Upload the BOA statement to begin. Gateway and invoice files are optional depending on the day.")
