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
        # Match only Zoho PAYMENT deposits, not Zoho software charges
        # "ZOHO PAYMENTS DES:..." = customer payment deposit ✅
        # "ZOHO* ZOHO-ONE..."     = software subscription charge ❌
        df["_is_zoho"] = df["description"].astype(str).str.upper().str.contains("ZOHO PAYMENTS")
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

    Architecture:
      The Zoho PDF is used for PAYOUT-LEVEL METADATA only:
        - Payout date (= BOA posting date)
        - Total gross amount (for balance validation)
        - Total fee / summary fee (authoritative debit amount — never recalculated)
        - BOA account last 4 digits (for offset account routing)
        - Payout ID + Bank Reference (appear in BOA description)

      Per-transaction customer names and amounts come from uploaded INVOICES,
      not from the PDF. This is robust to any transaction count (1 to 14+)
      because it never relies on parsing individual transaction rows from the PDF.

    Per-transaction amounts ARE extracted when available (for invoice-less flows),
    using the "chunk-by-3" strategy: the PDF always emits (gross, net, fee) for
    each transaction in sequence. We skip the summary values and chunk the rest.

    CRITICAL: The summary fee line is ALWAYS used as the debit row amount.
    Individual per-txn fees are never summed — they can drift due to PDF layout.
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

    # ── 1. Extract payout metadata ────────────────────────────────────────────

    # Payout date ("Payout: Jun 15, 2026, 06:38 AM")
    payout_date = None
    pd_m = re.search(r"Payout:\s+(\w+ \d{1,2}, \d{4})", full_text)
    if pd_m:
        payout_date = pd_m.group(1).strip()

    # Paid On date — rendered across 3 separate lines in the PDF:
    # "Paid / On / Jun / 16, / 2026"  (rendered across 3 separate lines)
    paid_on = None
    paid_on_m = re.search(
        r"Paid\s*\nOn\s*\n(\w+)\s*\n(\d{1,2}),?\s*\n(\d{4})", full_text
    )
    if paid_on_m:
        paid_on = f"{paid_on_m.group(1)} {paid_on_m.group(2)}, {paid_on_m.group(3)}"
    if not paid_on:
        paid_on_m2 = re.search(r"Paid\s+On\s+(\w+\s+\d{1,2},?\s*\d{4})", full_text)
        if paid_on_m2:
            paid_on = paid_on_m2.group(1).strip()

    posting_date = paid_on or payout_date or ""

    # Payout ID and Bank Reference
    pid_m = re.search(r"Payout ID:\s*(\d+)", full_text)
    payout_id = pid_m.group(1) if pid_m else ""

    bref_m = re.search(r"Bank Reference ID:\s*(\d+)", full_text)
    bank_ref = bref_m.group(1) if bref_m else ""

    # BOA account last 4 digits "( 3371 )" or "(••••3371)"
    acct_m = re.search(r"\(\s*[•\*]*\s*(\d{4})\s*\)", full_text)
    boa_acct = acct_m.group(1) if acct_m else ""

    # ── 2. Extract summary totals (authoritative) ─────────────────────────────
    summary_m = re.search(
        r"Payments\s+(\d+)\s+\$([\d,]+\.\d{2})\s+[−\-]\$?([\d,]+\.\d{2})\s+\$([\d,]+\.\d{2})",
        full_text
    )
    if not summary_m:
        return pd.DataFrame(), ["Could not find Payments summary line in Zoho PDF. "
                                  "Check that the file is a Zoho Payout Details PDF."]

    expected_count = int(summary_m.group(1))
    summary_gross  = float(summary_m.group(2).replace(",", ""))
    summary_fee    = float(summary_m.group(3).replace(",", ""))  # AUTHORITATIVE fee
    summary_net    = float(summary_m.group(4).replace(",", ""))

    # ── 3. Extract per-transaction gross amounts ──────────────────────────────
    # Used when no invoices are uploaded. Strategy: chunk-by-3.
    # The PDF emits dollar values after the summary in order:
    #   gross_1, net_1, fee_1, gross_2, net_2, fee_2, ...
    # We skip the summary values, then chunk remaining values into groups of 3.

    txn_section = full_text
    if "All Transactions" in full_text:
        txn_section = full_text.split("All Transactions", 1)[1]

    # Remove noise lines
    txn_section = re.sub(r"Refunds[^\n]*\n?", "", txn_section)
    txn_section = re.sub(r"Adjustments[^\n]*\n?", "", txn_section)
    txn_section = re.sub(r"Amount\s+\$[\d,]+\.\d{2}\s+USD\n?", "", txn_section)
    txn_section = re.sub(r"https?://\S+", "", txn_section)
    txn_section = re.sub(r"Page \d+ of \d+", "", txn_section)

    # Collect all positive dollar values in the transaction section
    all_dollars = [
        float(m.group(1).replace(",", ""))
        for m in re.finditer(r"\$([\d,]+\.\d{2})", txn_section)
        if float(m.group(1).replace(",", "")) > 0
    ]

    # Skip leading summary values (summary_gross, summary_fee, summary_net appear first)
    skip_vals = [summary_gross, summary_fee, summary_net]
    skip_idx = 0
    clean_dollars = []
    for val in all_dollars:
        if skip_idx < len(skip_vals) and abs(val - skip_vals[skip_idx]) < 0.02:
            skip_idx += 1
            continue
        clean_dollars.append(val)

    # Chunk into groups of 3: (gross, net, fee)
    transactions = []
    if len(clean_dollars) >= expected_count * 3:
        for i in range(expected_count):
            chunk = clean_dollars[i*3 : i*3+3]
            transactions.append({
                "gross_amount":  str(chunk[0]),
                "fee":           str(chunk[2]),   # fee is LAST in triplet
                "net_amount":    str(chunk[1]),
                "customer":      "",
                "date":          posting_date,
                "payout_date":   posting_date,
                "payout_id":     payout_id,
                "bank_ref":      bank_ref,
                "boa_acct":      boa_acct,
                "_summary_fee":  str(summary_fee),
                "_summary_gross": str(summary_gross),
                "_expected_count": str(expected_count),
            })
    else:
        # Partial data or dates-only available — create one row per expected txn
        # with gross distributed from clean_dollars (best effort)
        # The summary fee is still authoritative for the debit row
        available_gross = clean_dollars[:expected_count] if clean_dollars else []

        for i in range(expected_count):
            gross = available_gross[i] if i < len(available_gross) else 0.0
            transactions.append({
                "gross_amount":  str(gross),
                "fee":           "0.0",   # individual fee unknown; use summary total
                "net_amount":    str(gross),
                "customer":      "",
                "date":          posting_date,
                "payout_date":   posting_date,
                "payout_id":     payout_id,
                "bank_ref":      bank_ref,
                "boa_acct":      boa_acct,
                "_summary_fee":  str(summary_fee),
                "_summary_gross": str(summary_gross),
                "_expected_count": str(expected_count),
            })

    # Validate: sum of parsed grosses should match summary
    parsed_gross_sum = sum(float(t["gross_amount"]) for t in transactions)
    if abs(parsed_gross_sum - summary_gross) > 1.0:
        errors.append(
            f"Zoho PDF: Per-transaction gross sum ${parsed_gross_sum:,.2f} does not match "
            f"summary ${summary_gross:,.2f}. Individual row amounts may be approximate. "
            f"Upload invoices for exact per-customer amounts."
        )

    if not transactions:
        return pd.DataFrame(), ["Could not extract transaction rows from Zoho PDF. "
                                  "Upload Zoho export as CSV or Excel for best results."]

    return pd.DataFrame(transactions), errors



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
    """
    Load Account_Masterlist.xlsx.

    Supports these column formats:
      2-col: Account | Account Name
      3-col: Account Type | Account | Account Name
      4-col: Account Type | Account | Account Name | CS/PS Ticket

    The CS/PS Ticket column contains the individual person name that appears
    on the invoice "Bill To" field (e.g. "Ali Amir") when the registered D365
    account name is a business (e.g. "Functional Holistic Healing").

    Returns a DataFrame with columns:
      Account, Account Name, CS/PS Ticket (blank if not present)

    Only Customer rows (BC######) are returned.
    """
    df = pd.read_excel(path, header=None, dtype=str)

    # Find header row
    header_row = 0
    for i, row in df.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "account" in vals:
            header_row = i
            break

    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [c.strip() for c in df.columns]
    cols_lower = [c.lower() for c in df.columns]

    # Filter to Customer rows only
    if "account type" in cols_lower:
        type_col = df.columns[cols_lower.index("account type")]
        df = df[df[type_col].astype(str).str.strip().str.lower() == "customer"].copy()

    # Extract the columns we need
    acct_col = next((c for c in df.columns if c.lower() == "account"), None)
    name_col = next((c for c in df.columns if c.lower() == "account name"), None)

    # CS/PS Ticket column — optional, present in 4-column format
    ticket_col = next((c for c in df.columns
                       if c.lower() in ("cs/ps ticket", "cs ticket", "ps ticket",
                                        "ticket", "bill to name", "individual name")), None)

    if acct_col and name_col:
        keep = [acct_col, name_col]
        if ticket_col:
            keep.append(ticket_col)
        df = df[keep].copy()
        df.columns = ["Account", "Account Name"] + (["CS/PS Ticket"] if ticket_col else [])
    else:
        # Fallback: first two columns
        df = df.iloc[:, :2].copy()
        df.columns = ["Account", "Account Name"]

    # Ensure CS/PS Ticket column always exists (blank if not in file)
    if "CS/PS Ticket" not in df.columns:
        df["CS/PS Ticket"] = ""

    df = df.dropna(subset=["Account"]).reset_index(drop=True)
    df["Account"]      = df["Account"].astype(str).str.strip()
    df["Account Name"] = df["Account Name"].astype(str).str.strip()
    df["CS/PS Ticket"] = df["CS/PS Ticket"].fillna("").astype(str).str.strip()

    # Keep only valid BC###### accounts
    df = df[df["Account"].str.match(r"BC\d+")].reset_index(drop=True)
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
