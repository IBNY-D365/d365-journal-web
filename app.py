"""
IBNY D365 Journal Entry Automation App
---------------------------------------
Automates cash closing by matching BOA + Zoho transactions,
resolving customer accounts, and generating a D365-ready Excel upload.

Source-of-truth files (bundled with app):
  - Cash_Code_Masterlist.xlsx
  - IBNY_Business_Customer_Account.xlsx
  - Posted_Journal_in_D365_Sample_Reference.xlsx  (format reference)
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import warnings
warnings.filterwarnings("ignore")

# ── All logic inlined — single file deployment ───────────────────────────

"""
parsers.py — File parsing for BOA, Zoho (PDF/CSV/Excel), invoices, and reference files.

Source of truth: Zoho_Payment_D365_Automation_Rules.docx

BOA CSV format:   5-row summary block, blank line, then real header + data
Zoho PDF format:  Screen-captured pay.zoho.com page — parsed from raw text
Invoice PDF:      InBody Purchase Statement — extract Bill To, total, terms, ticket#
"""

import re, io
import pandas as pd
import numpy as np

# ── BOA column aliases ────────────────────────────────────────────────────────
_BOA_COL_MAP = {
    "date":        ["date", "posting date", "post date", "transaction date"],
    "description": ["description", "payee", "details", "memo"],
    "amount":      ["amount", "net amount", "credit", "credit amount", "transaction amount"],
    "debit":       ["debit", "debit amount"],
    "balance":     ["balance", "running balance", "running bal", "running bal."],
}

# ── Zoho column aliases ───────────────────────────────────────────────────────
_ZOHO_COL_MAP = {
    "date":         ["date", "payment date", "transaction date", "created date", "date & time"],
    "customer":     ["customer name", "customer", "client name", "name", "bill to", "payer"],
    "gross_amount": ["gross amount", "gross", "amount", "total amount", "payment amount"],
    "fee":          ["processing fee", "merchant fee", "fee", "transaction fee", "fees"],
    "net_amount":   ["net amount", "net", "net payment", "settlement amount", "payout amount"],
    "invoice":      ["invoice", "invoice number", "invoice #", "reference", "description"],
    "payment_id":   ["payment id", "transaction id", "id"],
    "status":       ["status", "payment status", "type"],
    "email":        ["email", "email address"],
}


def _normalise_col(h):
    return re.sub(r"\s+", " ", str(h).strip().lower())


def _map_columns(df, col_map):
    norm_to_actual = {_normalise_col(c): c for c in df.columns}
    result = {}
    for canonical, aliases in col_map.items():
        for alias in aliases:
            if _normalise_col(alias) in norm_to_actual:
                result[canonical] = norm_to_actual[_normalise_col(alias)]
                break
    return result


def _to_numeric(series):
    return (
        series.astype(str)
        .str.replace(r"[\$,€£]", "", regex=True)
        .str.replace(r"[−–]", "-", regex=True)
        .str.replace(r"\((\d+\.?\d*)\)", r"-\1", regex=True)
        .str.strip()
        .replace("", np.nan)
        .replace("nan", np.nan)
        .pipe(pd.to_numeric, errors="coerce")
    )


# ─────────────────────────────────────────────────────────────────────────────
# BOA parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_boa(file_obj):
    """Parse Bank of America export (CSV or Excel).
    
    Filter rule per §3.1: only rows where description contains
    'ZOHO PAYMENTS' are processed as Zoho payment deposits.
    'ZOHO* ZOHO-ONE' and other Zoho software charges are excluded.
    """
    errors = []
    try:
        name = getattr(file_obj, "name", "")
        df = _read_boa_csv(file_obj) if name.lower().endswith(".csv") else _read_boa_excel(file_obj)
    except Exception as e:
        return None, [f"Failed to read BOA file: {e}"]

    if df is None or df.empty:
        return None, ["BOA file parsed to an empty DataFrame."]

    mapping = _map_columns(df, _BOA_COL_MAP)
    df = df.rename(columns={v: k for k, v in mapping.items()})

    for col in ("date", "description", "amount", "debit", "balance"):
        if col not in df.columns:
            df[col] = np.nan

    if "amount" in df.columns and df["amount"].notna().any():
        df["_boa_amount"] = _to_numeric(df["amount"])
    elif "debit" in df.columns:
        df["_boa_amount"] = _to_numeric(df["debit"]) * -1
    else:
        df["_boa_amount"] = np.nan

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # §3.1: filter to "ZOHO PAYMENTS" deposits only — excludes software subscriptions
    if "description" in df.columns:
        df["_is_zoho"] = df["description"].astype(str).str.upper().str.contains("ZOHO PAYMENTS")
    else:
        df["_is_zoho"] = False

    # Drop rows with no date (summary/balance rows)
    df = df[df["date"].notna()].reset_index(drop=True)
    return df, errors


def _read_boa_csv(file_obj):
    raw = file_obj.read()
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")
    lines = raw.splitlines()
    # Find header line: most keyword matches
    best_score, header_idx = 0, 0
    for i, line in enumerate(lines):
        score = sum(1 for kw in ["date", "description", "amount", "balance"] if kw in line.lower())
        if score > best_score:
            best_score, header_idx = score, i
    csv_content = "\n".join(lines[header_idx:])
    return pd.read_csv(io.StringIO(csv_content), dtype=str, on_bad_lines="skip", engine="python")


def _read_boa_excel(file_obj):
    data = file_obj.read()
    for header_row in range(10):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            cols = [str(c).lower() for c in df.columns]
            if sum(1 for kw in ["date", "description", "amount"] if any(kw in c for c in cols)) >= 2:
                return df
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(data), dtype=str)


# ─────────────────────────────────────────────────────────────────────────────
# Zoho parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_zoho(file_obj):
    """Parse Zoho Payments export (PDF, CSV, or Excel)."""
    errors = []
    name = getattr(file_obj, "name", "")
    try:
        if name.lower().endswith(".pdf"):
            df, pdf_errors = _parse_zoho_pdf(file_obj)
            errors += pdf_errors
        elif name.lower().endswith(".csv"):
            df = _read_zoho_csv(file_obj)
        else:
            df = _read_zoho_excel(file_obj)
    except Exception as e:
        return None, [f"Failed to read Zoho file: {e}"]

    if df is None or df.empty:
        return None, ["Zoho file parsed to an empty DataFrame."]

    mapping = _map_columns(df, _ZOHO_COL_MAP)
    df = df.rename(columns={v: k for k, v in mapping.items()})

    for col in _ZOHO_COL_MAP.keys():
        if col not in df.columns:
            df[col] = np.nan

    for col in ("gross_amount", "fee", "net_amount"):
        if col in df.columns:
            df[col] = _to_numeric(df[col])

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Drop only "Total" summary rows, NOT blank-customer rows
    # (Zoho PDF always has blank customer — name comes from invoice)
    if "customer" in df.columns:
        mask = ~df["customer"].astype(str).str.strip().str.lower().str.startswith("total")
        df = df[mask]

    return df.reset_index(drop=True), errors


def _parse_zoho_pdf(file_obj):
    """
    Parse Zoho Payout Details PDF (screen-captured from pay.zoho.com).
    
    Extracts:
    - Payout metadata: date, payout ID, bank ref, BOA account last 4
    - Summary: total count, AUTHORITATIVE gross and fee totals
    - Per-transaction gross amounts (for invoice matching)
    
    The summary fee is ALWAYS used as the debit row amount (never per-txn sum).
    """
    errors = []
    try:
        import pdfplumber
    except ImportError:
        return pd.DataFrame(), ["pdfplumber not installed."]

    raw_data = file_obj.read()
    pages_text = []
    try:
        with pdfplumber.open(io.BytesIO(raw_data)) as pdf:
            for page in pdf.pages:
                pages_text.append(page.extract_text() or "")
    except Exception as e:
        return pd.DataFrame(), [f"PDF read error: {e}"]

    full_text = "\n".join(pages_text)

    # ── Payout metadata ───────────────────────────────────────────────────────
    payout_date = None
    pd_m = re.search(r"Payout:\s+(\w+ \d{1,2}, \d{4})", full_text)
    if pd_m:
        payout_date = pd_m.group(1).strip()

    # Paid On date (split across 3 lines: "Jun\n18,\n2026")
    paid_on = None
    po_m = re.search(r"Paid\s*\nOn\s*\n(\w+)\s*\n(\d{1,2}),?\s*\n(\d{4})", full_text)
    if po_m:
        paid_on = f"{po_m.group(1)} {po_m.group(2)}, {po_m.group(3)}"
    if not paid_on:
        po_m2 = re.search(r"Paid\s+On\s+(\w+\s+\d{1,2},?\s*\d{4})", full_text)
        if po_m2:
            paid_on = po_m2.group(1).strip()

    posting_date = paid_on or payout_date or ""

    pid_m = re.search(r"Payout ID:\s*(\d+)", full_text)
    payout_id = pid_m.group(1) if pid_m else ""

    bref_m = re.search(r"Bank Reference(?:\s+ID)?:\s*(\d+)", full_text)
    bank_ref = bref_m.group(1) if bref_m else ""

    acct_m = re.search(r"\(\s*[•\*]*\s*(\d{4})\s*\)", full_text)
    boa_acct = acct_m.group(1) if acct_m else ""

    # ── Summary line (AUTHORITATIVE) ─────────────────────────────────────────
    summary_m = re.search(
        r"Payments\s+(\d+)\s+\$([\d,]+\.\d{2})\s+[−\-]\$?([\d,]+\.\d{2})\s+\$([\d,]+\.\d{2})",
        full_text
    )
    if not summary_m:
        return pd.DataFrame(), ["Could not find Payments summary line in Zoho PDF."]

    expected_count = int(summary_m.group(1))
    summary_gross  = float(summary_m.group(2).replace(",", ""))
    summary_fee    = float(summary_m.group(3).replace(",", ""))  # AUTHORITATIVE
    summary_net    = float(summary_m.group(4).replace(",", ""))

    # ── Per-transaction gross amounts (chunk-by-3 strategy) ──────────────────
    txn_section = full_text.split("All Transactions", 1)[1] if "All Transactions" in full_text else full_text
    for pat in [r"Refunds[^\n]*\n?", r"Adjustments[^\n]*\n?",
                r"Amount\s+\$[\d,]+\.\d{2}\s+USD\n?", r"https?://\S+", r"Page \d+ of \d+"]:
        txn_section = re.sub(pat, "", txn_section)

    all_dollars = [float(m.group(1).replace(",", ""))
                   for m in re.finditer(r"\$([\d,]+\.\d{2})", txn_section)
                   if float(m.group(1).replace(",", "")) > 0]

    # Strip summary values from front
    skip_vals, skip_idx, clean = [summary_gross, summary_fee, summary_net], 0, []
    for v in all_dollars:
        if skip_idx < len(skip_vals) and abs(v - skip_vals[skip_idx]) < 0.02:
            skip_idx += 1
            continue
        clean.append(v)

    transactions = []
    if len(clean) >= expected_count * 3:
        for i in range(expected_count):
            chunk = clean[i*3:i*3+3]
            transactions.append({
                "date":           posting_date,
                "customer":       "",
                "gross_amount":   str(chunk[0]),
                "fee":            str(chunk[2]),
                "net_amount":     str(chunk[1]),
                "payout_date":    posting_date,
                "payout_id":      payout_id,
                "bank_ref":       bank_ref,
                "boa_acct":       boa_acct,
                "_summary_fee":   str(summary_fee),
                "_summary_gross": str(summary_gross),
            })
    else:
        # Fallback: use available grosses, rely on summary fee for debit
        available = clean[:expected_count]
        for i in range(expected_count):
            gross = available[i] if i < len(available) else 0.0
            transactions.append({
                "date":           posting_date,
                "customer":       "",
                "gross_amount":   str(gross),
                "fee":            "0.0",
                "net_amount":     str(gross),
                "payout_date":    posting_date,
                "payout_id":      payout_id,
                "bank_ref":       bank_ref,
                "boa_acct":       boa_acct,
                "_summary_fee":   str(summary_fee),
                "_summary_gross": str(summary_gross),
            })
        parsed_sum = sum(float(t["gross_amount"]) for t in transactions)
        if abs(parsed_sum - summary_gross) > 1.0:
            errors.append(f"Per-transaction gross sum ${parsed_sum:,.2f} != summary ${summary_gross:,.2f}. "
                         f"Upload invoices for exact per-customer amounts.")

    return pd.DataFrame(transactions), errors


def _read_zoho_csv(file_obj):
    content = file_obj.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    lines = content.splitlines()
    header_idx = 0
    for i, line in enumerate(lines):
        if sum(1 for kw in ["amount", "fee", "date", "customer"] if kw in line.lower()) >= 2:
            header_idx = i
            break
    return pd.read_csv(io.StringIO("\n".join(lines[header_idx:])), dtype=str,
                       on_bad_lines="skip", engine="python")


def _read_zoho_excel(file_obj):
    data = file_obj.read()
    for header_row in range(8):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            cols = [str(c).lower() for c in df.columns]
            if sum(1 for kw in ["amount", "fee", "date", "customer"] if any(kw in c for c in cols)) >= 2:
                return df
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(data), dtype=str)


# ─────────────────────────────────────────────────────────────────────────────
# Invoice PDF parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_invoice_pdf(file_obj):
    """
    Extract from InBody Purchase Statement PDF:
      invoice_number, customer_name (Bill To), total, payment_terms, cs_ticket
    
    Per §3.2: if customer name absent from Zoho, invoice Bill To is used.
    Per §3.2: if CS/PS Ticket column matches Bill To name, use that Account Name.
    """
    try:
        import pdfplumber
    except ImportError:
        return {}

    raw = file_obj.read() if hasattr(file_obj, "read") else open(file_obj, "rb").read()
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
    except Exception:
        return {}

    result = {}

    inv_m = re.search(r"Purchase\s*#\s*(INV-[\w\d]+)", text, re.IGNORECASE)
    if inv_m:
        result["invoice_number"] = inv_m.group(1).strip()

    # Bill To: name on the line immediately after "Bill To"
    bill_m = re.search(r"Bill\s+To\s*\n([^\n]+)", text, re.IGNORECASE)
    if bill_m:
        result["customer_name"] = bill_m.group(1).strip()

    terms_m = re.search(r"Terms\s*:\s*([^\n]+)", text, re.IGNORECASE)
    if terms_m:
        result["payment_terms"] = terms_m.group(1).strip()

    date_m = re.search(r"Invoice\s+Date\s*:\s*([^\n]+)", text, re.IGNORECASE)
    if date_m:
        result["invoice_date"] = date_m.group(1).strip()

    # Total: last standalone "Total $X" (not SubTotal)
    total_matches = list(re.finditer(r"(?:^|\n)\s*Total\s+\$?([\d,]+\.\d{2})", text))
    if total_matches:
        result["total"] = float(total_matches[-1].group(1).replace(",", ""))
    else:
        t_m = re.search(r"Total\s+\$?([\d,]+\.\d{2})", text)
        if t_m:
            result["total"] = float(t_m.group(1).replace(",", ""))

    # CS/PS ticket number: "[## 653 ##]" pattern in item descriptions
    ticket_m = re.search(r"\[##\s*(\d+)\s*##\]", text)
    if ticket_m:
        result["cs_ticket"] = ticket_m.group(1).strip()

    return result


# ─────────────────────────────────────────────────────────────────────────────
# Reference file loaders
# ─────────────────────────────────────────────────────────────────────────────

def load_customer_master(path):
    """
    Load Account_Masterlist.xlsx.
    Columns: Account Type | Account | Account Name | CS/PS Ticket (optional)
    Returns only Customer rows (BC######).
    CS/PS Ticket column maps individual Bill To names to business accounts.
    """
    df = pd.read_excel(path, header=None, dtype=str)
    header_row = 0
    for i, row in df.iterrows():
        if "account" in [str(v).strip().lower() for v in row.values]:
            header_row = i
            break

    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [c.strip() for c in df.columns]
    cols_lower = [c.lower() for c in df.columns]

    # Filter to Customer rows
    if "account type" in cols_lower:
        type_col = df.columns[cols_lower.index("account type")]
        df = df[df[type_col].astype(str).str.strip().str.lower() == "customer"].copy()

    acct_col   = next((c for c in df.columns if c.lower() == "account"), None)
    name_col   = next((c for c in df.columns if c.lower() == "account name"), None)
    ticket_col = next((c for c in df.columns
                       if c.lower() in ("cs/ps ticket", "cs ticket", "ps ticket", "ticket")), None)

    if acct_col and name_col:
        keep = [acct_col, name_col] + ([ticket_col] if ticket_col else [])
        df = df[keep].copy()
        df.columns = ["Account", "Account Name"] + (["CS/PS Ticket"] if ticket_col else [])
    else:
        df = df.iloc[:, :2].copy()
        df.columns = ["Account", "Account Name"]

    if "CS/PS Ticket" not in df.columns:
        df["CS/PS Ticket"] = ""

    df = df.dropna(subset=["Account"]).reset_index(drop=True)
    df["Account"]      = df["Account"].astype(str).str.strip()
    df["Account Name"] = df["Account Name"].astype(str).str.strip()
    df["CS/PS Ticket"] = df["CS/PS Ticket"].fillna("").astype(str).str.strip()
    df = df[df["Account"].str.match(r"BC\d+")].reset_index(drop=True)
    return df


def load_cash_codes(path):
    df = pd.read_excel(path, header=None, dtype=str)
    header_row = 0
    for i, row in df.iterrows():
        if "cash code" in [str(v).strip().lower() for v in row.values]:
            header_row = i
            break
    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = ["Cash Code", "Cash Code Name"] + list(df.columns[2:])
    df = df.dropna(subset=["Cash Code"]).reset_index(drop=True)
    df["Cash Code"] = df["Cash Code"].str.strip()
    return df


"""
matcher.py — Match BOA ↔ Zoho transactions and resolve customer accounts.

Rules source: Zoho_Payment_D365_Automation_Rules.docx §3.2

Resolution pipeline per §3.2:
  1. Match Zoho transaction gross amount to invoice total
  2. Get customer name from invoice "Bill To"
  3. Look up name in Account_Masterlist:
     a. If found in Account Name column → use that account
     b. If found in CS/PS Ticket column → use the corresponding Account Name (business name)
     c. If not found → flag for review, leave Account blank, Account type = Customer
  4. Cash code from invoice payment terms
"""

import re
import numpy as np
import pandas as pd
from rapidfuzz import fuzz, process
from parsers import parse_invoice_pdf

_FUZZY_THRESHOLD = 80

# Payment terms → (cash_code, description_prefix)
# Source: Zoho_Payment_D365_Automation_Rules.docx §4
_TERMS_MAP = {
    "due on receipt":        ("AR001", ""),
    "due upon receipt":      ("AR001", ""),
    "monthly payment plan":  ("AR002", "MPP "),
    "monthly payment":       ("AR002", "MPP "),
    "mpp":                   ("AR002", "MPP "),
    "financing":             ("AR003", ""),
    "leasing":               ("AR004", ""),
    "net 1":                 ("AR005", ""),
    "net 10":                ("AR006", ""),
    "net 25":                ("AR007", ""),
    "net 30":                ("AR008", ""),
    "net 40":                ("AR009", ""),
    "net 45":                ("AR010", ""),
    "net 60":                ("AR011", ""),
    "due end of next month": ("AR011", ""),
    "net":                   ("AR008", ""),
}


def match_transactions(boa_df, zoho_df, customer_df, invoice_files=None):
    """
    Enrich each Zoho transaction row with D365 fields.
    
    Returns (enriched_df, log_entries).
    Each row gains: _account, _account_name, _cash_code, _cash_code_prefix,
                    _desc_prefix, _match_confidence, _needs_review, _review_reason,
                    _boa_date, _boa_description, _summary_fee
    """
    log = []

    # ── Step 1: Parse invoices → {amount: invoice_data} ──────────────────────
    invoice_by_amount = {}
    if invoice_files:
        for inv_file in invoice_files:
            try:
                inv = parse_invoice_pdf(inv_file)
                if inv and inv.get("total"):
                    key = round(float(inv["total"]), 2)
                    invoice_by_amount[key] = inv
                    log.append({"level": "OK",
                                "msg": f"Invoice: {inv.get('invoice_number')} | "
                                       f"Bill To: '{inv.get('customer_name')}' | "
                                       f"Total: ${float(inv['total']):,.2f} | "
                                       f"Terms: {inv.get('payment_terms')} | "
                                       f"Ticket: {inv.get('cs_ticket','')} "})
            except Exception as e:
                log.append({"level": "WARN", "msg": f"Could not parse invoice: {e}"})

    # ── Step 2: Build customer lookup (Account Name + CS/PS Ticket) ───────────
    lookup = _build_lookup(customer_df)

    # ── Step 3: Resolve each Zoho row ─────────────────────────────────────────
    results = []
    for idx, zrow in zoho_df.iterrows():
        record = _resolve_row(idx, zrow, lookup, invoice_by_amount, log)
        results.append(record)

    enriched = pd.concat(
        [zoho_df.reset_index(drop=True), pd.DataFrame(results).reset_index(drop=True)],
        axis=1
    )

    # ── Step 4: Attach BOA posting date and description ───────────────────────
    enriched = _attach_boa(enriched, boa_df, log)

    # ── Step 5: Balance check ─────────────────────────────────────────────────
    _check_balance(boa_df, enriched, log)

    return enriched, log


# ─────────────────────────────────────────────────────────────────────────────
# Customer lookup
# ─────────────────────────────────────────────────────────────────────────────

def _build_lookup(customer_df):
    """
    Build {normalised_name: {account, account_name}} dict.
    Two entries per row when CS/PS Ticket is present:
      - Account Name (business name)
      - CS/PS Ticket (individual Bill To name → same business account)
    """
    lookup = {}
    for _, row in customer_df.iterrows():
        acc    = str(row.get("Account", "")).strip()
        name   = str(row.get("Account Name", "")).strip()
        ticket = str(row.get("CS/PS Ticket", "")).strip()
        if not acc or not name:
            continue
        entry = {"account": acc, "account_name": name}
        lookup[_norm(name)] = entry
        if ticket and ticket.lower() not in ("", "nan", "none"):
            lookup[_norm(ticket)] = entry
    return lookup


def _norm(name):
    """Normalise name for fuzzy matching."""
    name = str(name).lower().strip()
    name = re.sub(r"[^\w\s]", " ", name)
    for sfx in [r"\bllc\b", r"\binc\b", r"\bltd\b", r"\bcorp\b",
                r"\bpllc\b", r"\bpc\b", r"\bdba\b", r"\bthe\b"]:
        name = re.sub(sfx, "", name)
    return re.sub(r"\s+", " ", name).strip()


def _fuzzy_lookup(raw_name, lookup):
    """Fuzzy-match raw_name against lookup keys."""
    if not lookup:
        return None, 0
    norm = _norm(raw_name)
    result  = process.extractOne(norm, list(lookup.keys()), scorer=fuzz.token_set_ratio)
    if not result:
        return None, 0
    best_key, score = result[0], result[1]
    if score >= _FUZZY_THRESHOLD:
        return lookup[best_key], score
    return None, score


# ─────────────────────────────────────────────────────────────────────────────
# Row resolution
# ─────────────────────────────────────────────────────────────────────────────

def _resolve_row(idx, zrow, lookup, invoice_by_amount, log):
    """
    Resolve one Zoho transaction row to D365 fields.
    
    Priority:
    1. Match gross amount to invoice → get Bill To name + terms
    2. Look up Bill To name in master (Account Name or CS/PS Ticket column)
    3. If not found → flag, leave account blank, still Customer type
    """
    gross = _safe_float(zrow.get("gross_amount"))

    # ── Match invoice by gross amount ─────────────────────────────────────────
    inv = None
    if gross:
        inv = invoice_by_amount.get(round(gross, 2))

    if inv:
        customer_name = inv.get("customer_name", "")
        payment_terms = inv.get("payment_terms", "")
        cs_ticket     = inv.get("cs_ticket", "")
        cash_code, pfx = _terms_to_cash_code(payment_terms)

        # Look up customer in master
        matched, score = _fuzzy_lookup(customer_name, lookup)

        if matched:
            is_cs = bool(cs_ticket)
            desc_prefix = f"{customer_name} CS Ticket #{cs_ticket}_" if is_cs else ""
            log.append({"level": "OK",
                        "msg": f"Row {idx}: '{customer_name}' → "
                               f"{matched['account']} {matched['account_name']} "
                               f"[score={score}] cash={cash_code}"})
            return {
                "_account":          matched["account"],
                "_account_name":     matched["account_name"],
                "_account_type":     "Customer",
                "_posting_profile":  "AutoPost",
                "_cash_code":        cash_code,
                "_cash_code_prefix": pfx,
                "_desc_prefix":      desc_prefix,
                "_match_confidence": "HIGH" if score >= 92 else "MEDIUM",
                "_needs_review":     score < 92,
                "_review_reason":    f"Fuzzy score {score}/100 — verify name" if score < 92 else "",
                "_invoice_number":   inv.get("invoice_number", ""),
                "_raw_name":         customer_name,
            }
        else:
            # Not in master — flag but keep Customer type per §5.1
            reason = (f"'{customer_name}' not found in Account_Masterlist. "
                      f"Add them with BC###### to resolve.")
            log.append({"level": "WARN", "msg": f"Row {idx}: {reason}"})
            is_cs = bool(cs_ticket)
            desc_prefix = f"{customer_name} CS Ticket #{cs_ticket}_" if is_cs else ""
            return {
                "_account":          "",
                "_account_name":     customer_name,
                "_account_type":     "Customer",
                "_posting_profile":  "AutoPost",
                "_cash_code":        cash_code,
                "_cash_code_prefix": pfx,
                "_desc_prefix":      desc_prefix,
                "_match_confidence": "LOW",
                "_needs_review":     True,
                "_review_reason":    reason,
                "_invoice_number":   inv.get("invoice_number", ""),
                "_raw_name":         customer_name,
            }

    # ── No invoice: try Zoho customer name directly ───────────────────────────
    raw_customer = str(zrow.get("customer", "")).strip()
    if raw_customer and raw_customer not in ("—", "-", "nan", ""):
        matched, score = _fuzzy_lookup(raw_customer, lookup)
        if matched:
            cash_code, pfx = "AR001", ""
            return {
                "_account":          matched["account"],
                "_account_name":     matched["account_name"],
                "_account_type":     "Customer",
                "_posting_profile":  "AutoPost",
                "_cash_code":        cash_code,
                "_cash_code_prefix": pfx,
                "_desc_prefix":      "",
                "_match_confidence": "HIGH" if score >= 92 else "MEDIUM",
                "_needs_review":     score < 92,
                "_review_reason":    "",
                "_invoice_number":   "",
                "_raw_name":         raw_customer,
            }

    # ── No match at all ───────────────────────────────────────────────────────
    reason = (f"No invoice uploaded matching ${gross:,.2f}. "
              f"Upload the invoice to resolve.") if gross else "No customer name and no invoice."
    log.append({"level": "WARN", "msg": f"Row {idx}: {reason}"})
    return {
        "_account":          "",
        "_account_name":     raw_customer or "",
        "_account_type":     "Customer",
        "_posting_profile":  "AutoPost",
        "_cash_code":        "",
        "_cash_code_prefix": "",
        "_desc_prefix":      "",
        "_match_confidence": "LOW",
        "_needs_review":     True,
        "_review_reason":    reason,
        "_invoice_number":   "",
        "_raw_name":         raw_customer or "",
    }


# ─────────────────────────────────────────────────────────────────────────────
# BOA attachment + balance check
# ─────────────────────────────────────────────────────────────────────────────

def _attach_boa(zoho_df, boa_df, log):
    """Attach BOA posting date and description string to each Zoho row."""
    zoho_df = zoho_df.copy()
    if boa_df is None or boa_df.empty:
        zoho_df["_boa_date"] = pd.NaT
        zoho_df["_boa_description"] = ""
        return zoho_df

    boa_zoho = boa_df[boa_df.get("_is_zoho", pd.Series(False, index=boa_df.index))].copy()
    if boa_zoho.empty:
        zoho_df["_boa_date"] = pd.NaT
        zoho_df["_boa_description"] = ""
        log.append({"level": "WARN",
                    "msg": "No 'ZOHO PAYMENTS' rows found in BOA file."})
        return zoho_df

    # Use the first (and usually only) ZOHO PAYMENTS row
    boa_row = boa_zoho.iloc[0]
    boa_date = pd.Timestamp(boa_row["date"]).normalize() if pd.notna(boa_row.get("date")) else pd.NaT
    boa_desc = str(boa_row.get("description", ""))

    zoho_df["_boa_date"]        = boa_date
    zoho_df["_boa_description"] = boa_desc
    return zoho_df


def _check_balance(boa_df, zoho_df, log):
    """Validate: sum(Zoho gross) - summary_fee == BOA net deposit."""
    if boa_df is None or boa_df.empty or "_boa_amount" not in boa_df.columns:
        return
    boa_zoho = boa_df[boa_df.get("_is_zoho", pd.Series(False, index=boa_df.index))]
    if boa_zoho.empty:
        return

    boa_net    = float(boa_zoho["_boa_amount"].sum())
    zoho_gross = float(zoho_df["gross_amount"].fillna(0).sum()) if "gross_amount" in zoho_df.columns else 0

    # Use authoritative summary fee
    if "_summary_fee" in zoho_df.columns:
        sf = pd.to_numeric(zoho_df["_summary_fee"].dropna(), errors="coerce").dropna()
        zoho_fee = float(sf.iloc[0]) if not sf.empty else 0
    else:
        zoho_fee = float(zoho_df["fee"].fillna(0).sum()) if "fee" in zoho_df.columns else 0

    zoho_net = zoho_gross - zoho_fee
    diff = abs(boa_net - zoho_net)
    tol  = max(abs(boa_net) * 0.005, 1.0)

    if diff <= tol:
        log.append({"level": "OK",
                    "msg": f"Balance PASSED: BOA ${boa_net:,.2f} = "
                           f"Zoho gross ${zoho_gross:,.2f} − fee ${zoho_fee:,.2f} = ${zoho_net:,.2f}"})
    else:
        log.append({"level": "WARN",
                    "msg": f"Balance MISMATCH: BOA ${boa_net:,.2f} vs "
                           f"Zoho net ${zoho_net:,.2f} (diff ${diff:,.2f})"})


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _terms_to_cash_code(terms_str):
    if not terms_str or str(terms_str).lower() in ("nan", ""):
        return "AR001", ""
    t = terms_str.lower().strip()
    for key, val in _TERMS_MAP.items():
        if key in t:
            return val
    return "AR001", ""


def _safe_float(val):
    try:
        f = float(str(val).replace(",", "").replace("$", "").strip())
        return None if np.isnan(f) else f
    except (ValueError, TypeError):
        return None


"""
builder.py — Build D365 journal entry rows from matched Zoho data.

Rules source: Zoho_Payment_D365_Automation_Rules.docx §5

Per payout batch:
  • 1 CREDIT line per customer (§5.1)
  • 1 DEBIT line for ALL grouped merchant fees (§5.2, §3.3)

Description format per §5.1 col 9 examples:
  AR001 (standard):  "{BC######} {Account Name}_{BOA description}"
  AR002 (MPP):       "MPP {BC######} {Account Name}_{BOA description}"
  CS ticket:         "{Bill To Name} CS Ticket #{ticket}_{BOA description}"
  No account yet:    "{Account Name}_{BOA description}"

Debit description per §5.2:
  Single:  "Zoho Merchant Fee {BC######} {Name}_{BOA description}"
  Batch:   "Zoho Merchant Fee {BC1} {Name1}, {BC2} {Name2}_{BOA description}"
"""

import re
import pandas as pd
import numpy as np

# Constants per §5.1 and §5.2
_COMPANY         = "bwa"
_POSTING_PROFILE = "AutoPost"
_DEBIT_ACCT_NAME = "Outside Service (Finance)"
_DEBIT_ACCT_TYPE = "Ledger"
_DEBIT_ACCT      = "43170111-U26C05001-B735350-UOA003"
_DEBIT_CASH_CODE = "OSF005"
_CURRENCY        = "USD"
_EXCHANGE_RATE   = 1.00
_SALES_TAX_GROUP = "AVATAX"
_REVERSING_ENTRY = "No"

D365_COLUMNS = [
    "Date", "Voucher", "Account name", "Company", "Account type",
    "Account", "Posting profile", "Cash code", "Description",
    "Debit", "Credit",
    "Item sales tax group", "Sales tax code",
    "Offset company", "Offset account type", "Offset account",
    "Offset transaction text",
    "Currency", "Exchange rate",
    "Item sales tax group2", "Sales tax group",
    "Withholding tax group", "Release date", "Reversing entry", "Reversing date",
    "_needs_review", "_review_reason", "_match_confidence",
]


def build_journal_entries(matched_df, customer_df, offset_account="B1000002"):
    log, rows = [], []

    matched_df = matched_df.copy()

    # Use BOA posting date as the D365 entry date per §5.1 col 1
    if "_boa_date" in matched_df.columns and matched_df["_boa_date"].notna().any():
        matched_df["_group_date"] = pd.to_datetime(matched_df["_boa_date"], errors="coerce")
    elif "date" in matched_df.columns:
        matched_df["_group_date"] = pd.to_datetime(matched_df["date"], errors="coerce")
    else:
        matched_df["_group_date"] = pd.NaT

    groups = matched_df.groupby("_group_date", dropna=False)

    for group_date, group in groups:
        group_rows   = list(group.iterrows())
        multi_batch  = len(group_rows) > 1
        credit_rows  = []
        total_fee    = 0.0
        summary_fee  = None
        fee_desc_parts = []

        date_str = ""
        if pd.notna(group_date):
            try:
                date_str = pd.Timestamp(group_date).strftime("%m/%d/%Y")
            except Exception:
                date_str = str(group_date)

        for _, zrow in group_rows:
            account       = str(zrow.get("_account", "")).strip()
            account_name  = str(zrow.get("_account_name", "")).strip()
            cash_code     = str(zrow.get("_cash_code", "AR001")).strip() or "AR001"
            cash_pfx      = str(zrow.get("_cash_code_prefix", "")).strip()
            desc_prefix   = str(zrow.get("_desc_prefix", "")).strip()
            needs_review  = bool(zrow.get("_needs_review", False))
            review_reason = str(zrow.get("_review_reason", ""))
            confidence    = str(zrow.get("_match_confidence", "LOW"))
            gross         = _safe_float(zrow.get("gross_amount"))
            fee           = _safe_float(zrow.get("fee"))
            boa_desc      = str(zrow.get("_boa_description", "")).strip()

            # Accumulate fee
            if fee and fee > 0:
                total_fee += fee
            sf = _safe_float(zrow.get("_summary_fee"))
            if sf and sf > 0:
                summary_fee = sf

            # Build credit description per §5.1 col 9
            credit_desc = _credit_description(
                cash_code, cash_pfx, desc_prefix, account, account_name, boa_desc
            )

            # Accumulate fee description parts for debit row
            part = " ".join(filter(None, [account, account_name])).strip()
            if part:
                fee_desc_parts.append(part)

            credit_row = {
                "Date":                  date_str,
                "Voucher":               "",
                "Account name":          account_name,
                "Company":               _COMPANY,
                "Account type":          "Customer",   # always Customer per §5.1 col 5
                "Account":               account,
                "Posting profile":       _POSTING_PROFILE if account else "",
                "Cash code":             cash_code,
                "Description":           credit_desc,
                "Debit":                 "",
                "Credit":                f"{gross:.2f}" if gross else "",
                "Item sales tax group":  "",
                "Sales tax code":        "",
                "Offset company":        _COMPANY,
                "Offset account type":   "Bank",
                "Offset account":        offset_account,
                "Offset transaction text": "",
                "Currency":              _CURRENCY,
                "Exchange rate":         _EXCHANGE_RATE,
                "Item sales tax group2": "",
                "Sales tax group":       _SALES_TAX_GROUP,
                "Withholding tax group": "",
                "Release date":          "",
                "Reversing entry":       _REVERSING_ENTRY,
                "Reversing date":        "",
                "_needs_review":         needs_review,
                "_review_reason":        review_reason,
                "_match_confidence":     confidence,
            }
            credit_rows.append(credit_row)
            rows.append(credit_row)

        # ── Debit row: use Zoho summary fee (authoritative) ───────────────────
        debit_fee = summary_fee if summary_fee else total_fee
        if debit_fee > 0:
            if multi_batch:
                fee_names = ", ".join(fee_desc_parts)
                debit_desc = (f"Zoho Merchant Fee {fee_names}_{boa_desc}"
                              if boa_desc else f"Zoho Merchant Fee {fee_names}")
            else:
                # Single: use same name/account part as credit, prefix with "Zoho Merchant Fee "
                single_part = fee_desc_parts[0] if fee_desc_parts else ""
                debit_desc = (f"Zoho Merchant Fee {single_part}_{boa_desc}"
                              if boa_desc else f"Zoho Merchant Fee {single_part}")

            rows.append({
                "Date":                  date_str,
                "Voucher":               "",
                "Account name":          _DEBIT_ACCT_NAME,
                "Company":               _COMPANY,
                "Account type":          _DEBIT_ACCT_TYPE,
                "Account":               _DEBIT_ACCT,
                "Posting profile":       "",
                "Cash code":             _DEBIT_CASH_CODE,
                "Description":           debit_desc,
                "Debit":                 f"{debit_fee:.2f}",
                "Credit":                "",
                "Item sales tax group":  "",
                "Sales tax code":        "",
                "Offset company":        _COMPANY,
                "Offset account type":   "Bank",
                "Offset account":        offset_account,
                "Offset transaction text": "",
                "Currency":              _CURRENCY,
                "Exchange rate":         _EXCHANGE_RATE,
                "Item sales tax group2": "",
                "Sales tax group":       _SALES_TAX_GROUP,
                "Withholding tax group": "",
                "Release date":          "",
                "Reversing entry":       _REVERSING_ENTRY,
                "Reversing date":        "",
                "_needs_review":         False,
                "_review_reason":        "",
                "_match_confidence":     "HIGH",
            })
            log.append({"level": "OK",
                        "msg": f"Debit row: ${debit_fee:.2f} "
                               f"({'summary' if summary_fee else 'per-txn sum'}) "
                               f"for {date_str}"})

    if not rows:
        log.append({"level": "WARN", "msg": "No journal entry rows generated."})
        return pd.DataFrame(columns=[c for c in D365_COLUMNS if not c.startswith("_")]), log

    journal_df = pd.DataFrame(rows)
    for col in D365_COLUMNS:
        if col not in journal_df.columns:
            journal_df[col] = ""
    return journal_df, log


def _credit_description(cash_code, cash_pfx, desc_prefix, account, account_name, boa_desc):
    """
    Build credit line description per §5.1 col 9.

    Examples from rules doc:
      AR001: "BC000571 Page Fit Inc. DBA Intoxx Fitness_ZOHO PAYMENTS DES:..."
      AR002: "MPP BC000327 Elite Functional Wellness_ZOHO PAYMENTS DES:..."
    
    Note: cash code does NOT appear in the description string.
    The description format is: [{MPP }]{BC######} {Name}_{BOA desc}
    
    CS ticket: "{Bill To} CS Ticket #{ticket}_{BOA desc}"
    No account: "{Account Name}_{BOA desc}"
    """
    name_part = " ".join(filter(None, [account, account_name])).strip()

    if desc_prefix:
        # CS ticket path: "Paul Fuss CS Ticket #654_ZOHO PAYMENTS DES:..."
        return f"{desc_prefix}{boa_desc}" if boa_desc else desc_prefix.rstrip("_")

    if cash_pfx:
        # MPP path: "MPP BC000327 Elite Functional Wellness_ZOHO PAYMENTS DES:..."
        return f"{cash_pfx}{name_part}_{boa_desc}" if boa_desc else f"{cash_pfx}{name_part}"

    if account:
        # Standard: "BC000571 Page Fit Inc._ZOHO PAYMENTS DES:..."
        return f"{name_part}_{boa_desc}" if boa_desc else name_part

    # No account resolved yet: "Legends Charter School_ZOHO PAYMENTS DES:..."
    return f"{account_name}_{boa_desc}" if boa_desc else account_name


def _safe_float(val):
    try:
        f = float(str(val).replace(",", "").replace("$", "").strip())
        return None if np.isnan(f) else f
    except (ValueError, TypeError):
        return None


"""
exporter.py — Export D365 journal entries to a formatted Excel file.

Output matches the D365_General_Journal_Template.xlsx column order exactly.
Rows needing review are highlighted in amber.
"""

import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ── D365 output column order (exact match to template) ───────────────────────
D365_OUTPUT_COLUMNS = [
    "Date", "Voucher", "Account name", "Company", "Account type",
    "Account", "Posting profile", "Cash code", "Description",
    "Debit", "Credit",
    "Item sales tax group", "Sales tax code",
    "Offset company", "Offset account type", "Offset account",
    "Offset transaction text",
    "Currency", "Exchange rate",
    "Item sales tax group2", "Sales tax group",
    "Withholding tax group", "Release date", "Reversing entry", "Reversing date",
]

# ── Styles ────────────────────────────────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", start_color="1A3C6E", end_color="1A3C6E")   # navy
HEADER_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
REVIEW_FILL   = PatternFill("solid", start_color="FFF3CD", end_color="FFF3CD")   # amber
REVIEW_FONT   = Font(name="Arial", color="856404", size=10)
NORMAL_FONT   = Font(name="Arial", size=10)
CREDIT_FILL   = PatternFill("solid", start_color="E8F5E9", end_color="E8F5E9")   # light green
DEBIT_FILL    = PatternFill("solid", start_color="E3F2FD", end_color="E3F2FD")   # light blue

THIN_BORDER = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)

COLUMN_WIDTHS = {
    "Date": 14,
    "Voucher": 18,
    "Account name": 32,
    "Company": 9,
    "Account type": 13,
    "Account": 20,
    "Posting profile": 16,
    "Cash code": 10,
    "Description": 72,
    "Debit": 14,
    "Credit": 14,
    "Item sales tax group": 12,
    "Sales tax code": 12,
    "Offset company": 9,
    "Offset account type": 14,
    "Offset account": 40,
    "Offset transaction text": 14,
    "Currency": 10,
    "Exchange rate": 12,
    "Item sales tax group2": 12,
    "Sales tax group": 13,
    "Withholding tax group": 13,
    "Release date": 12,
    "Reversing entry": 13,
    "Reversing date": 12,
}


def export_to_excel(journal_df: pd.DataFrame) -> bytes:
    """
    Write D365 journal entries to an openpyxl workbook and return raw bytes.

    Two sheets are produced:
      1. 'D365 Upload'  — clean export, no internal columns
      2. 'Review Items' — only flagged rows, with review reason column
    """
    # ── Separate clean data from internal review columns ──────────────────────
    internal_cols = [c for c in journal_df.columns if c.startswith("_")]
    needs_review_col = "_needs_review"  if "_needs_review"  in journal_df.columns else None
    review_reason_col= "_review_reason" if "_review_reason" in journal_df.columns else None

    # Build the upload-ready DataFrame (only D365 columns, in order)
    upload_df = _build_upload_df(journal_df)

    wb = Workbook()

    # ── Sheet 1: D365 Upload ──────────────────────────────────────────────────
    ws = wb.active
    ws.title = "D365 Upload"
    _write_sheet(ws, upload_df, journal_df, needs_review_col)

    # ── Sheet 2: Review Items ─────────────────────────────────────────────────
    if needs_review_col and journal_df[needs_review_col].any():
        ws2 = wb.create_sheet("Review Items")
        flagged_idx = journal_df[journal_df[needs_review_col] == True].index
        flagged_upload = upload_df.loc[flagged_idx].copy()
        if review_reason_col:
            flagged_upload.insert(
                len(flagged_upload.columns), "Review Reason",
                journal_df.loc[flagged_idx, review_reason_col].values
            )
        _write_sheet(ws2, flagged_upload, journal_df, needs_review_col)

    # ── Sheet 3: Legend ───────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Legend")
    _write_legend(ws3)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _build_upload_df(journal_df: pd.DataFrame) -> pd.DataFrame:
    """Extract D365 output columns in canonical order; fill missing with ''."""
    df = pd.DataFrame()
    for col in D365_OUTPUT_COLUMNS:
        if col in journal_df.columns:
            df[col] = journal_df[col].fillna("").astype(str).replace("nan", "").replace("None", "")
        else:
            df[col] = ""
    return df


def _write_sheet(ws, upload_df: pd.DataFrame, source_df: pd.DataFrame, needs_review_col):
    """Write headers + data rows with formatting to ws."""
    # ── Headers ───────────────────────────────────────────────────────────────
    for col_idx, col_name in enumerate(upload_df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = THIN_BORDER

    ws.row_dimensions[1].height = 28

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row_idx, (df_idx, row) in enumerate(upload_df.iterrows(), start=2):
        flagged = False
        if needs_review_col and needs_review_col in source_df.columns:
            flagged = bool(source_df.loc[df_idx, needs_review_col]) if df_idx in source_df.index else False

        # Determine row type: credit (has Credit amount) or debit (has Debit)
        has_credit = str(row.get("Credit", "")).strip() not in ("", "0", "0.0", "0.00")
        has_debit  = str(row.get("Debit",  "")).strip() not in ("", "0", "0.0", "0.00")
        row_fill = REVIEW_FILL if flagged else (CREDIT_FILL if has_credit else (DEBIT_FILL if has_debit else None))
        row_font = REVIEW_FONT if flagged else NORMAL_FONT

        for col_idx, (col_name, value) in enumerate(row.items(), start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font      = row_font
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(col_name == "Description"))
            if row_fill:
                cell.fill = row_fill

    # ── Column widths ─────────────────────────────────────────────────────────
    for col_idx, col_name in enumerate(upload_df.columns, start=1):
        width = COLUMN_WIDTHS.get(col_name, 14)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Freeze pane: keep headers visible ─────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Auto-filter ───────────────────────────────────────────────────────────
    if upload_df.shape[0] > 0:
        last_col = get_column_letter(len(upload_df.columns))
        ws.auto_filter.ref = f"A1:{last_col}1"


def _write_legend(ws):
    """Write a color legend and column field guide on the Legend sheet."""
    ws.title = "Legend"

    headers = ["Color", "Meaning"]
    ws.append(headers)
    ws["A1"].font = HEADER_FONT
    ws["A1"].fill = HEADER_FILL
    ws["B1"].font = HEADER_FONT
    ws["B1"].fill = HEADER_FILL

    legend_rows = [
        ("Green row",  "CREDIT line — customer payment (verified)"),
        ("Blue row",   "DEBIT line — Zoho merchant fee"),
        ("Amber row",  "⚠️ Flagged for manual review — check account, cash code, or amount"),
    ]

    fills = [CREDIT_FILL, DEBIT_FILL, REVIEW_FILL]
    fonts = [NORMAL_FONT, NORMAL_FONT, REVIEW_FONT]

    for i, (label, desc) in enumerate(legend_rows, start=2):
        ws.cell(row=i, column=1, value=label).fill = fills[i - 2]
        ws.cell(row=i, column=1).font = fonts[i - 2]
        ws.cell(row=i, column=2, value=desc).font = fonts[i - 2]

    ws.append([])
    ws.append(["Field", "Rule"])
    ws["A6"].font = HEADER_FONT
    ws["A6"].fill = HEADER_FILL
    ws["B6"].font = HEADER_FONT
    ws["B6"].fill = HEADER_FILL

    field_rules = [
        ("Date",             "BOA posting date (not journal entry date)"),
        ("Voucher",          "Leave blank — auto-generated by D365"),
        ("Account name",     "Canonical name from IBNY Business Customer Account file"),
        ("Company",          "Always: bwa"),
        ("Account type",     "Credit rows: Customer | Debit rows: Ledger"),
        ("Account",          "BC###### from customer master, or 43170111-… for fee rows"),
        ("Posting profile",  "AutoPost (credit rows only)"),
        ("Cash code",        "AR001=Due on receipt, AR002=MPP, OSF005=Merchant fee"),
        ("Description",      "Credit: {AR00X}: {BC######} {Customer Name}_{BOA desc}"),
        ("Debit",            "Merchant fee amount (debit rows only)"),
        ("Credit",           "Gross payment from Zoho (credit rows only, NOT BOA net)"),
        ("Offset account",   "B1000002 (acct 3371) | B1000003 (acct 3924) | B1000001 (acct 3384)"),
        ("Sales tax group",  "Always: AVATAX (credit rows)"),
        ("Reversing entry",  "Always: No"),
    ]

    for i, (field, rule) in enumerate(field_rules, start=7):
        ws.cell(row=i, column=1, value=field).font = NORMAL_FONT
        ws.cell(row=i, column=2, value=rule).font  = NORMAL_FONT

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 70


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IBNY D365 Journal Entry Automation",
    page_icon="📒",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {font-size:1.6rem; font-weight:700; color:#1a3c6e; margin-bottom:0.2rem;}
    .sub-header  {font-size:0.95rem; color:#555; margin-bottom:1.5rem;}
    .section-title {font-size:1.05rem; font-weight:600; color:#1a3c6e;
                    border-bottom:2px solid #d0dff5; padding-bottom:4px; margin-bottom:1rem;}
    .flag-box {background:#fff8e1; border-left:4px solid #ffc107;
               padding:0.6rem 1rem; border-radius:4px; margin:4px 0;}
    .ok-box   {background:#e8f5e9; border-left:4px solid #4caf50;
               padding:0.6rem 1rem; border-radius:4px; margin:4px 0;}
    .err-box  {background:#ffebee; border-left:4px solid #f44336;
               padding:0.6rem 1rem; border-radius:4px; margin:4px 0;}
    .metric-card {background:#f0f4ff; border-radius:8px; padding:1rem;
                  text-align:center;}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="main-header">📒 IBNY D365 Journal Entry Automation</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload BOA and Zoho files to generate a D365-ready journal entry Excel.</div>', unsafe_allow_html=True)
st.caption("v1.4 — fee fix: uses Zoho summary fee")

# ── Load reference data (auto-loaded from app directory) ─────────────────────
@st.cache_data(show_spinner=False)
def load_references():
    import os
    _base = os.path.dirname(os.path.abspath(__file__))
    customers  = load_customer_master(os.path.join(_base, "Account_Masterlist.xlsx"))
    cash_codes = load_cash_codes(os.path.join(_base, "Cash_Code_Masterlist.xlsx"))
    return customers, cash_codes

try:
    customer_df, cash_code_df = load_references()
    cs_count = int((customer_df.get("CS/PS Ticket", pd.Series()).fillna("").str.strip() != "").sum()) if "CS/PS Ticket" in customer_df.columns else 0
    st.markdown(f'<div class="ok-box">✅ Reference files loaded — {len(customer_df)} customer accounts · {len(cash_code_df)} cash codes · CS/PS entries: {cs_count}</div>', unsafe_allow_html=True)
except Exception as e:
    st.markdown(f'<div class="err-box">❌ Could not load reference files: {e}</div>', unsafe_allow_html=True)
    st.stop()

st.markdown("---")

# ── File Upload ───────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="section-title">🏦 Bank of America File</div>', unsafe_allow_html=True)
    boa_file = st.file_uploader(
        "Upload BOA transaction export",
        type=["xlsx", "csv", "xls"],
        key="boa",
        help="Excel or CSV export from Bank of America. Must contain date, description, and amount columns.",
    )

with col2:
    st.markdown('<div class="section-title">💳 Zoho Payments File</div>', unsafe_allow_html=True)
    zoho_file = st.file_uploader(
        "Upload Zoho payment export",
        type=["xlsx", "csv", "xls", "pdf"],
        key="zoho",
        help="Zoho Payments export showing gross amount and merchant fee per transaction.",
    )

# Optional invoice uploader
with st.expander("📄 Upload Customer Invoices (optional — used when name is missing from Zoho)"):
    invoice_files = st.file_uploader(
        "Upload one or more invoice PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="invoices",
    )

st.markdown("---")

# ── BOA Account routing (from automation rules §5.1 col 16) ──────────────────
st.markdown('<div class="section-title">⚙️ BOA Account Settings</div>', unsafe_allow_html=True)
acct_col1, acct_col2 = st.columns([1, 2])
with acct_col1:
    boa_account_last4 = st.selectbox(
        "BOA Source Account (last 4 digits)",
        options=["3371", "3924", "3384", "Unknown"],
        index=0,
        help="Determines the Offset Account (B1000002, B1000003, or B1000001).",
    )
with acct_col2:
    offset_map = {"3371": "B1000002", "3924": "B1000003", "3384": "B1000001", "Unknown": "B1000002"}
    selected_offset = offset_map[boa_account_last4]
    st.info(f"Offset Account → **{selected_offset}**")

st.markdown("---")

# ── Process ───────────────────────────────────────────────────────────────────
if st.button("🚀 Generate D365 Journal Entries", type="primary", use_container_width=True):
    if not boa_file or not zoho_file:
        st.error("Please upload both a BOA file and a Zoho file before proceeding.")
        st.stop()

    with st.spinner("Parsing files…"):
        boa_df, boa_errors = parse_boa(boa_file)
        zoho_df, zoho_errors = parse_zoho(zoho_file)

    # ── Parse feedback ──
    if boa_errors:
        for e in boa_errors:
            st.markdown(f'<div class="flag-box">⚠️ BOA: {e}</div>', unsafe_allow_html=True)
    if zoho_errors:
        for e in zoho_errors:
            st.markdown(f'<div class="flag-box">⚠️ Zoho: {e}</div>', unsafe_allow_html=True)

    if boa_df is None or boa_df.empty:
        st.error("Could not parse BOA file. Check the format and try again.")
        st.stop()
    if zoho_df is None or zoho_df.empty:
        st.error("Could not parse Zoho file. Check the format and try again.")
        st.stop()

    # ── Preview raw parses ──
    with st.expander("🔍 Raw BOA Rows Parsed"):
        st.dataframe(boa_df, use_container_width=True)
    with st.expander("🔍 Raw Zoho Rows Parsed"):
        st.dataframe(zoho_df, use_container_width=True)
        if "_summary_fee" in zoho_df.columns:
            st.success(f"✅ Zoho summary fee (authoritative): **${float(zoho_df['_summary_fee'].iloc[0]):,.2f}** — this will be used as the debit amount.")
        else:
            st.error("❌ _summary_fee column missing — old parsers.py is still running. Push the new file to GitHub.")

    with st.spinner("Matching transactions and resolving accounts…"):
        matched_df, match_log = match_transactions(
            boa_df, zoho_df, customer_df, invoice_files
        )

    with st.spinner("Building D365 journal entries…"):
        journal_df, build_log = build_journal_entries(
            matched_df, customer_df, selected_offset
        )

    # ── Match summary metrics ──
    st.markdown("---")
    st.markdown('<div class="section-title">📊 Processing Summary</div>', unsafe_allow_html=True)

    total     = len(matched_df)
    confident = int(matched_df["_match_confidence"].eq("HIGH").sum())  if "_match_confidence" in matched_df else 0
    flagged   = int(matched_df["_needs_review"].sum())                 if "_needs_review"     in matched_df else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Zoho Transactions", total)
    m2.metric("Matched (High Confidence)", confident)
    m3.metric("Flagged for Review", flagged)
    m4.metric("D365 Rows Generated", len(journal_df))

    # ── Match log details ──
    if match_log:
        with st.expander(f"📋 Match Log ({len(match_log)} entries)"):
            for entry in match_log:
                css_class = "flag-box" if entry.get("level") == "WARN" else "ok-box"
                st.markdown(f'<div class="{css_class}">{entry["msg"]}</div>', unsafe_allow_html=True)

    # ── Review table: flagged rows ──
    needs_review = journal_df[journal_df.get("_needs_review", False) == True] if "_needs_review" in journal_df else pd.DataFrame()
    if not needs_review.empty:
        st.markdown("---")
        st.markdown('<div class="section-title">🚩 Rows Requiring Manual Review</div>', unsafe_allow_html=True)
        st.warning(f"{len(needs_review)} row(s) could not be fully resolved. These appear highlighted in the export.")
        review_cols = [c for c in needs_review.columns if not c.startswith("_")]
        st.dataframe(needs_review[review_cols], use_container_width=True)

    # ── Full journal preview ──
    st.markdown("---")
    st.markdown('<div class="section-title">📄 D365 Journal Entry Preview</div>', unsafe_allow_html=True)
    display_cols = [c for c in journal_df.columns if not c.startswith("_")]
    st.dataframe(journal_df[display_cols], use_container_width=True, height=400)

    # ── Export ──
    with st.spinner("Building Excel export…"):
        excel_bytes = export_to_excel(journal_df)

    st.success("✅ D365 journal entry file ready for download.")
    st.download_button(
        label="⬇️ Download D365 Journal Entry Excel",
        data=excel_bytes,
        file_name="D365_Journal_Entry_Upload.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Sidebar: reference info ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📚 Reference Data")
    st.markdown(f"**Customer Accounts:** {len(customer_df)}")
    st.markdown(f"**Cash Codes:** {len(cash_code_df)}")

    st.markdown("---")
    st.markdown("### 🗂️ Cash Code Reference")
    st.dataframe(
        cash_code_df[cash_code_df["Cash Code"].str.startswith("AR")],
        use_container_width=True, height=320, hide_index=True
    )

    st.markdown("---")
    st.markdown("### 📘 How to Use")
    st.markdown("""
1. Upload **BOA Excel/CSV** export
2. Upload **Zoho Payments** export
3. Select the BOA account (last 4 digits)
4. Click **Generate Journal Entries**
5. Review flagged rows
6. Download the Excel file
""")
    st.markdown("---")
    st.caption("IBNY Cash Closing Automation v1.0")
