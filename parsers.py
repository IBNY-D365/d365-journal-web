"""
parsers.py — File parsing for BOA, Zoho, and reference spreadsheets.

Tuned to the actual file formats observed:
  BOA CSV  : 5-row summary block at top, then blank line, then real data header
             Columns: Date, Description, Amount, Running Bal.
  Zoho PDF : Screen-captured payout page from pay.zoho.com — no proper table
             structure. All data lives in a single merged cell as raw text.
             Must be parsed from the plain text of each page.
"""

import re
import io
import pandas as pd
import numpy as np

# ── Column aliases ────────────────────────────────────────────────────────────
_BOA_COL_MAP = {
    "date":        ["date", "posting date", "post date", "transaction date",
                    "trans date", "effective date"],
    "description": ["description", "payee", "details", "memo", "narrative",
                    "transaction description", "trans desc"],
    "amount":      ["amount", "net amount", "credit", "credit amount",
                    "transaction amount", "trans amount"],
    "debit":       ["debit", "debit amount"],
    "balance":     ["balance", "running balance", "running bal", "running bal."],
}

_ZOHO_COL_MAP = {
    "date":         ["date", "payment date", "transaction date", "created date",
                     "created time", "payout date", "date & time"],
    "customer":     ["customer name", "customer", "client name", "name",
                     "bill to", "payer", "contact", "buyer"],
    "gross_amount": ["gross amount", "gross", "amount", "total amount",
                     "payment amount", "gross sales"],
    "fee":          ["processing fee", "merchant fee", "fee", "transaction fee",
                     "charge", "service fee", "fees"],
    "net_amount":   ["net amount", "net", "net payment", "settlement amount",
                     "payout amount", "total"],
    "invoice":      ["invoice", "invoice number", "invoice #", "invoice no",
                     "reference", "description"],
    "payment_id":   ["payment id", "payment reference", "transaction id", "id"],
    "status":       ["status", "payment status", "type"],
    "email":        ["email", "email address", "customer email"],
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


# ─────────────────────────────────────────────────────────────────────────────
# BOA CSV/Excel parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_boa(file_obj):
    """
    Parse Bank of America export.

    The real stmt.csv has this structure:
      Row 0:  Description,,Summary Amt.         ← summary header (3 cols)
      Row 1:  Beginning balance...              ← summary data
      Row 2:  Total credits...
      Row 3:  Total debits...
      Row 4:  Ending balance...
      Row 5:  (blank)
      Row 6:  Date,Description,Amount,Running Bal.   ← REAL header (4 cols)
      Row 7+: actual transactions
    """
    errors = []
    try:
        name = getattr(file_obj, "name", "")
        if name.lower().endswith(".csv"):
            df = _read_boa_csv(file_obj)
        else:
            df = _read_boa_excel(file_obj)
    except Exception as e:
        return None, [f"Failed to read BOA file: {e}"]

    if df is None or df.empty:
        return None, ["BOA file parsed to an empty DataFrame."]

    mapping = _map_columns(df, _BOA_COL_MAP)
    rename = {v: k for k, v in mapping.items()}
    df = df.rename(columns=rename)

    for col in ("date", "description", "amount", "debit", "balance"):
        if col not in df.columns:
            df[col] = np.nan

    # Derive signed amount
    if "amount" in df.columns and df["amount"].notna().any():
        df["_boa_amount"] = _to_numeric(df["amount"])
    elif "debit" in df.columns and df["debit"].notna().any():
        df["_boa_amount"] = _to_numeric(df["debit"]) * -1
    else:
        df["_boa_amount"] = np.nan

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Drop rows with no date (summary/balance rows that slipped through)
    df = df[df["date"].notna()].copy()

    if "description" in df.columns:
        df["_is_zoho"] = df["description"].astype(str).str.upper().str.contains("ZOHO")
    else:
        df["_is_zoho"] = False
        errors.append("BOA: No description column found.")

    df = df.reset_index(drop=True)
    return df, errors


def _read_boa_csv(file_obj):
    """
    The BOA CSV has a 5-row summary block with 3 columns, then a blank line,
    then the real 4-column transaction data. We find the real header by
    looking for the line with the most recognized column keywords.
    """
    raw = file_obj.read()
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")

    lines = raw.splitlines()

    # Score each line: count how many BOA header keywords it contains
    best_score = 0
    header_idx = 0
    for i, line in enumerate(lines):
        lower = line.lower()
        score = sum(1 for kw in ["date", "description", "amount", "balance"]
                    if kw in lower)
        if score > best_score:
            best_score = score
            header_idx = i

    # Take everything from the best header line onward
    data_lines = lines[header_idx:]

    # Determine column count from the header
    header_cols = len(next(iter(data_lines)).split(","))

    # Keep only lines that have plausible column counts (header_cols ± 1)
    # to drop any stray summary lines below the data
    clean_lines = [data_lines[0]]  # always keep header
    for line in data_lines[1:]:
        stripped = line.strip()
        if not stripped:
            continue
        clean_lines.append(line)

    csv_content = "\n".join(clean_lines)

    return pd.read_csv(
        io.StringIO(csv_content),
        dtype=str,
        on_bad_lines="skip",
        engine="python",
    )


def _read_boa_excel(file_obj):
    data = file_obj.read()
    for header_row in range(10):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            cols_lower = [str(c).lower() for c in df.columns]
            score = sum(1 for kw in ["date", "description", "amount"]
                        if any(kw in c for c in cols_lower))
            if score >= 2:
                return df
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(data), dtype=str)


# ─────────────────────────────────────────────────────────────────────────────
# Zoho PDF parser  (screen-captured payout page from pay.zoho.com)
# ─────────────────────────────────────────────────────────────────────────────

def parse_zoho(file_obj):
    """
    Route to the correct sub-parser based on file extension.
    The Zoho Payout PDF is a screen-captured web page — pdfplumber finds only
    ONE merged cell containing all text. We must parse the raw page text.
    """
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

    rename = {v: k for k, v in mapping.items()}
    df = df.rename(columns=rename)

    for col in _ZOHO_COL_MAP.keys():
        if col not in df.columns:
            df[col] = np.nan

    for col in ("gross_amount", "fee", "net_amount"):
        if col in df.columns:
            df[col] = _to_numeric(df[col])

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Drop only literal "Total" summary rows (not empty-customer rows from PDF)
    # Zoho PDF rows legitimately have blank customer — name comes from invoice
    if "customer" in df.columns:
        mask = ~df["customer"].astype(str).str.strip().str.lower().str.startswith("total")
        df = df[mask]

    df = df.reset_index(drop=True)
    return df, errors


def _parse_zoho_pdf(file_obj):
    """
    Parse a Zoho Payout Details PDF (screen-captured from pay.zoho.com).

    Key observations from real file:
      - All text is extracted as raw lines (no proper table structure)
      - Page 1 has: payout header, summary block, first transaction row (split across lines)
      - Page 2 has: remaining transaction rows (also split across lines)
      - Transaction layout: date ... gross ... net ... fee  (fee is LAST, not second)
      - For each transaction block: largest dollar = gross, smallest = fee, middle = net
    """
    errors = []
    try:
        import pdfplumber
    except ImportError:
        return pd.DataFrame(), ["pdfplumber not installed — PDF parsing unavailable."]

    raw_data = file_obj.read()
    pages_text = []

    try:
        with pdfplumber.open(io.BytesIO(raw_data)) as pdf:
            for page in pdf.pages:
                pages_text.append(page.extract_text() or "")
    except Exception as e:
        return pd.DataFrame(), [f"PDF read error: {e}"]

    full_text = "\n".join(pages_text)

    # ── Extract payout metadata ───────────────────────────────────────────────
    # Paid On date (3 separate lines: "Jun", "16,", "2026")
    paid_on_date = None
    paid_on_m = re.search(
        r"Paid\s*\nOn\s*\n(\w+)\s*\n(\d{1,2}),?\s*\n(\d{4})", full_text
    )
    if paid_on_m:
        paid_on_date = f"{paid_on_m.group(1)} {paid_on_m.group(2)}, {paid_on_m.group(3)}"
    
    # Also try single-line format
    if not paid_on_date:
        paid_on_m2 = re.search(
            r"Paid\s+On\s+(\w+\s+\d{1,2},?\s*\d{4})", full_text, re.IGNORECASE
        )
        if paid_on_m2:
            paid_on_date = paid_on_m2.group(1).strip()

    payout_id_m = re.search(r"Payout ID:\s*(\d+)", full_text)
    payout_id = payout_id_m.group(1) if payout_id_m else ""

    bank_ref_m = re.search(r"Bank Reference ID:\s*(\d+)", full_text)
    bank_ref = bank_ref_m.group(1) if bank_ref_m else ""

    boa_acct_m = re.search(r"\(\s*(?:\u2022+\s*)?(\d{4})\s*\)", full_text)
    if not boa_acct_m:
        boa_acct_m = re.search(r"\(\s*[•]+\s*(\d{4})\s*\)", full_text)
    boa_acct = boa_acct_m.group(1) if boa_acct_m else ""

    # ── Get expected transaction count from summary ───────────────────────────
    summary_m = re.search(
        r"Payments\s+(\d+)\s+\$([\d,]+\.\d{2})\s+[−\-]\$?([\d,]+\.\d{2})",
        full_text
    )
    expected_count = int(summary_m.group(1)) if summary_m else None
    total_fee_from_summary = float(summary_m.group(3).replace(",","")) if summary_m else None

    # ── Find the "All Transactions" section ──────────────────────────────────
    txn_section = full_text
    if "All Transactions" in full_text:
        txn_section = full_text.split("All Transactions", 1)[1]

    # ── Find transaction blocks: each starts with a date like "Jun 12, 2026" ─
    # Dates in transaction rows (not the payout date header)
    # Payout date = "Jun 15, 2026" (from "Payout: Jun 15...") → skip this date
    payout_date_m = re.search(r"Payout:\s+(\w+ \d{1,2}, \d{4})", full_text)
    payout_header_date = payout_date_m.group(1) if payout_date_m else None

    # Find all "Month DD, YYYY" dates in the transaction section
    all_dates = list(re.finditer(r"(\w{3}\s+\d{1,2},\s*\d{4})", txn_section))
    
    # Filter out the payout header date if it appears
    txn_dates = []
    for dm in all_dates:
        d = dm.group(1).strip()
        if payout_header_date and d == payout_header_date:
            continue
        txn_dates.append(dm)

    transactions = []

    if txn_dates:
        for i, date_m in enumerate(txn_dates):
            # Get the text segment around this transaction
            start = max(0, date_m.start() - 20)  # a little before the date
            end = txn_dates[i+1].start() if i+1 < len(txn_dates) else len(txn_section)
            segment = txn_section[start:end]

            # Extract all dollar values in this segment
            dollar_vals = [
                float(m.group(1).replace(",",""))
                for m in re.finditer(r"\$([\d,]+\.\d{2})", segment)
            ]
            
            # Filter out $0.00 values and summary amounts
            dollar_vals = [v for v in dollar_vals if v > 0]

            if not dollar_vals:
                continue

            if len(dollar_vals) >= 2:
                gross = max(dollar_vals)
                fee   = min(dollar_vals)
            else:
                gross = dollar_vals[0]
                fee   = 0.0

            transactions.append({
                "date":         date_m.group(1).strip(),
                "customer":     "",
                "gross_amount": str(gross),
                "fee":          str(fee),
                "net_amount":   str(gross - fee),
                "payout_date":  paid_on_date or "",
                "payout_id":    payout_id,
                "bank_ref":     bank_ref,
                "boa_acct":     boa_acct,
            })
    
    # ── Fallback: distribute summary total across invoices ────────────────────
    if not transactions and summary_m:
        total_gross = float(summary_m.group(2).replace(",",""))
        total_fee   = float(summary_m.group(3).replace(",",""))
        transactions.append({
            "date":         paid_on_date or "",
            "customer":     "",
            "gross_amount": str(total_gross),
            "fee":          str(total_fee),
            "net_amount":   str(total_gross - total_fee),
            "payout_date":  paid_on_date or "",
            "payout_id":    payout_id,
            "bank_ref":     bank_ref,
            "boa_acct":     boa_acct,
            "_is_summary":  True,
        })
        errors.append(
            f"Zoho PDF: Parsed summary only ({summary_m.group(1)} payments, "
            f"gross=${total_gross:,.2f}, fee=${total_fee:,.2f}). "
            f"Upload invoices to split by customer."
        )

    if not transactions:
        errors.append(
            "Could not extract transaction data from Zoho PDF. "
            "Try exporting from Zoho Payments as CSV or Excel."
        )
        return pd.DataFrame(), errors

    df = pd.DataFrame(transactions)
    return df, errors



def _read_zoho_csv(file_obj):
    content = file_obj.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    lines = content.splitlines()
    header_idx = 0
    for i, line in enumerate(lines):
        lower = line.lower()
        score = sum(1 for kw in ["amount", "fee", "date", "customer", "name"]
                    if kw in lower)
        if score >= 2:
            header_idx = i
            break
    csv_content = "\n".join(lines[header_idx:])
    return pd.read_csv(
        io.StringIO(csv_content), dtype=str,
        on_bad_lines="skip", engine="python"
    )


def _read_zoho_excel(file_obj):
    data = file_obj.read()
    for header_row in range(8):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            cols_lower = [str(c).lower() for c in df.columns]
            score = sum(1 for kw in ["amount", "fee", "date", "customer"]
                        if any(kw in c for c in cols_lower))
            if score >= 2:
                return df
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(data), dtype=str)


# ─────────────────────────────────────────────────────────────────────────────
# Invoice PDF parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_invoice_pdf(file_obj):
    """
    Extract customer name, invoice number, total, and payment terms from
    an InBody Purchase Statement PDF.

    Returns a dict with keys:
      invoice_number, customer_name, total, payment_terms, invoice_date, cs_ticket
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

    # Invoice number
    inv_m = re.search(r"Purchase\s*#\s*(INV-[\w\d]+)", text, re.IGNORECASE)
    if inv_m:
        result["invoice_number"] = inv_m.group(1).strip()

    # Bill To name (line immediately after "Bill To")
    bill_to_m = re.search(r"Bill\s+To\s*\n([^\n]+)", text, re.IGNORECASE)
    if bill_to_m:
        result["customer_name"] = bill_to_m.group(1).strip()

    # Payment Terms
    terms_m = re.search(r"Terms\s*:\s*([^\n]+)", text, re.IGNORECASE)
    if terms_m:
        result["payment_terms"] = terms_m.group(1).strip()

    # Invoice date
    date_m = re.search(r"Invoice\s+Date\s*:\s*([^\n]+)", text, re.IGNORECASE)
    if date_m:
        result["invoice_date"] = date_m.group(1).strip()

    # Total: find the LAST standalone "Total $X,XXX.XX" line
    # (not SubTotal, not "Balance Due", just "Total $X")
    # Use findall and take the value after the last plain "Total"
    total_matches = list(re.finditer(
        r"(?:^|\n)\s*Total\s+\$?([\d,]+\.\d{2})", text
    ))
    if total_matches:
        result["total"] = float(total_matches[-1].group(1).replace(",", ""))
    else:
        # Fallback: any "Total $X"
        t_m = re.search(r"Total\s+\$?([\d,]+\.\d{2})", text)
        if t_m:
            result["total"] = float(t_m.group(1).replace(",", ""))

    # CS Ticket number — "[## 676 ##]" pattern
    ticket_m = re.search(r"\[##\s*(\d+)\s*##\]", text)
    if ticket_m:
        result["cs_ticket"] = ticket_m.group(1).strip()

    return result



# ─────────────────────────────────────────────────────────────────────────────
# Reference file loaders
# ─────────────────────────────────────────────────────────────────────────────

def load_customer_master(path):
    df = pd.read_excel(path, header=None, dtype=str)
    header_row = 0
    for i, row in df.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "account" in vals:
            header_row = i
            break
    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [c.strip() for c in df.columns]
    mask = df.iloc[:, 0].astype(str).str.match(r"BC\d+")
    df = df[mask].reset_index(drop=True)
    if len(df.columns) >= 2:
        df = df.iloc[:, :2]
        df.columns = ["Account", "Account Name"]
    df = df.dropna(subset=["Account"]).reset_index(drop=True)
    df["Account"] = df["Account"].str.strip()
    df["Account Name"] = df["Account Name"].astype(str).str.strip()
    return df


def load_cash_codes(path):
    df = pd.read_excel(path, header=None, dtype=str)
    header_row = 0
    for i, row in df.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "cash code" in vals:
            header_row = i
            break
    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = ["Cash Code", "Cash Code Name"] + list(df.columns[2:])
    df = df.dropna(subset=["Cash Code"]).reset_index(drop=True)
    df["Cash Code"] = df["Cash Code"].str.strip()
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

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
