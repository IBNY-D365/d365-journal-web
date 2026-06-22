"""
parsers.py — File parsing for BOA, Zoho, and reference spreadsheets.

All column detection is multi-name tolerant so it survives export format drift.
Returns (DataFrame, [error_strings]) tuples — never raises to the caller.
"""

import re
import io
import pandas as pd
import numpy as np

# ── Column name aliases ───────────────────────────────────────────────────────
_BOA_COL_MAP = {
    "date":        ["date", "posting date", "post date", "transaction date",
                    "trans date", "effective date"],
    "description": ["description", "payee", "details", "memo", "narrative",
                    "transaction description", "trans desc"],
    "amount":      ["amount", "net amount", "credit", "credit amount",
                    "transaction amount", "trans amount"],
    "debit":       ["debit", "debit amount"],
    "balance":     ["balance", "running balance"],
}

# Broad aliases — covers every Zoho PDF/CSV export variant we've seen
_ZOHO_COL_MAP = {
    "date": [
        "date", "payment date", "transaction date", "created date",
        "created time", "payout date", "settlement date", "paid on",
    ],
    "customer": [
        "customer name", "customer", "client name", "name", "bill to",
        "payer", "contact", "buyer", "account name",
    ],
    "gross_amount": [
        "gross amount", "gross", "gross payment", "amount", "total amount",
        "payment amount", "sale amount", "charged amount", "gross sales",
        "gross total", "subtotal", "revenue",
    ],
    "fee": [
        "processing fee", "merchant fee", "fee", "transaction fee",
        "charge", "service fee", "zoho fee", "platform fee",
        "fees & charges", "fees", "total fees", "fee amount",
        "stripe fee", "payment fee",
    ],
    "net_amount": [
        "net amount", "net", "net payment", "settlement amount",
        "payout amount", "net total", "net payout", "amount paid out",
    ],
    "invoice": [
        "invoice", "invoice number", "invoice #", "invoice no",
        "reference", "invoice id", "order id", "order number",
    ],
    "payment_id": [
        "payment id", "payment reference", "transaction id", "id",
        "zoho payment id", "txn id", "transaction #",
    ],
    "description": [
        "description", "memo", "notes", "payment description", "remarks",
    ],
    "status": ["status", "payment status", "transaction status"],
    "email":  ["email", "email address", "customer email"],
}


def _normalise_col(header: str) -> str:
    return re.sub(r"\s+", " ", str(header).strip().lower())


def _map_columns(df: pd.DataFrame, col_map: dict) -> dict:
    norm_to_actual = {_normalise_col(c): c for c in df.columns}
    result = {}
    for canonical, aliases in col_map.items():
        for alias in aliases:
            if _normalise_col(alias) in norm_to_actual:
                result[canonical] = norm_to_actual[_normalise_col(alias)]
                break
    return result


# ─────────────────────────────────────────────────────────────────────────────
# BOA parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_boa(file_obj):
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

    if "amount" in df.columns and df["amount"].notna().any():
        df["_boa_amount"] = _to_numeric(df["amount"])
    elif "debit" in df.columns and df["debit"].notna().any():
        df["_boa_amount"] = _to_numeric(df["debit"]) * -1
    else:
        df["_boa_amount"] = np.nan

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    if "description" in df.columns:
        df["_is_zoho"] = df["description"].astype(str).str.upper().str.contains("ZOHO")
    else:
        df["_is_zoho"] = False
        errors.append("BOA: No description column found; cannot filter Zoho rows.")

    df = df.reset_index(drop=True)
    return df, errors


def _read_boa_csv(file_obj):
    """
    BOA CSV exports have metadata rows at the top before the real header.
    Strategy:
      1. Read all lines raw
      2. Find the line whose comma-count best matches a data row (most consistent)
      3. Use that as the header; skip everything above it
      4. Drop rows after any blank/separator line at the bottom
    """
    raw = file_obj.read()
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")

    lines = [l for l in raw.splitlines()]

    # Find the header line: first line that contains date/description/amount keywords
    header_idx = None
    for i, line in enumerate(lines):
        lower = line.lower()
        if any(kw in lower for kw in ["date", "description", "amount", "balance"]):
            # Make sure it's actually a header (not a value row with a date)
            # Count how many of our expected keywords appear
            score = sum(1 for kw in ["date", "description", "amount", "balance"] if kw in lower)
            if score >= 2:
                header_idx = i
                break

    if header_idx is None:
        header_idx = 0

    # Count columns in the header
    header_cols = len(lines[header_idx].split(","))

    # Keep only lines from header_idx onward that have the same column count
    # (or close — allowing for quoted commas, use csv reader)
    csv_lines = [lines[header_idx]]
    for line in lines[header_idx + 1:]:
        stripped = line.strip()
        if not stripped:
            continue  # skip blank lines
        csv_lines.append(line)

    csv_content = "\n".join(csv_lines)

    try:
        df = pd.read_csv(
            io.StringIO(csv_content),
            dtype=str,
            on_bad_lines="skip",   # skip lines with wrong column count
            engine="python",
        )
        return df
    except Exception:
        # Last resort: read with error_bad_lines suppressed
        try:
            df = pd.read_csv(
                io.StringIO(csv_content),
                dtype=str,
                error_bad_lines=False,
                warn_bad_lines=False,
            )
            return df
        except Exception:
            return pd.read_csv(io.StringIO(csv_content), dtype=str, skiprows=0)


def _read_boa_excel(file_obj):
    data = file_obj.read()
    for header_row in range(8):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            cols_lower = [str(c).lower() for c in df.columns]
            if sum(1 for kw in ["date", "description", "amount"] if any(kw in c for c in cols_lower)) >= 2:
                return df
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(data), dtype=str)


# ─────────────────────────────────────────────────────────────────────────────
# Zoho parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_zoho(file_obj):
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

    # Show what columns were actually found (helps diagnose future mismatches)
    actual_cols = list(df.columns)

    mapping = _map_columns(df, _ZOHO_COL_MAP)

    # If gross_amount or fee still not found, try partial/fuzzy column matching
    if "gross_amount" not in mapping:
        mapping = _try_partial_match(df, mapping, "gross_amount",
                                     ["amount", "gross", "total", "sale", "revenue", "subtotal"])
    if "fee" not in mapping:
        mapping = _try_partial_match(df, mapping, "fee",
                                     ["fee", "charge", "deduct", "commission"])

    for must_have in ("gross_amount", "fee"):
        if must_have not in mapping:
            errors.append(
                f"Zoho: Required column '{must_have}' not found. "
                f"Columns in file: {actual_cols}"
            )

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

    # Drop summary/total rows (blank customer name or literal "Total" rows)
    if "customer" in df.columns:
        mask = (
            df["customer"].astype(str).str.strip().ne("") &
            df["customer"].astype(str).str.lower().ne("nan") &
            ~df["customer"].astype(str).str.lower().str.contains(r"^total")
        )
        df = df[mask]

    df = df.reset_index(drop=True)
    return df, errors


def _try_partial_match(df, mapping, canonical, keywords):
    """Try to match a canonical column via partial keyword match on actual column names."""
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if any(kw in col_lower for kw in keywords):
            # Don't re-use already mapped columns
            if col not in mapping.values():
                mapping[canonical] = col
                return mapping
    return mapping


def _read_zoho_csv(file_obj):
    content = file_obj.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    # Try to skip leading metadata rows
    lines = content.splitlines()
    header_idx = 0
    for i, line in enumerate(lines):
        lower = line.lower()
        score = sum(1 for kw in ["amount", "fee", "date", "customer", "name"] if kw in lower)
        if score >= 2:
            header_idx = i
            break
    csv_content = "\n".join(lines[header_idx:])
    return pd.read_csv(io.StringIO(csv_content), dtype=str, on_bad_lines="skip", engine="python")


def _read_zoho_excel(file_obj):
    data = file_obj.read()
    for header_row in range(8):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            cols_lower = [str(c).lower() for c in df.columns]
            if sum(1 for kw in ["amount", "fee", "date", "customer"] if any(kw in c for c in cols_lower)) >= 2:
                return df
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(data), dtype=str)


def _parse_zoho_pdf(file_obj):
    """
    Extract table data from a Zoho Payout PDF.

    Zoho PDFs typically have:
    - A summary section at the top (payout total, fees, etc.)
    - A transactions table listing individual payments

    Strategy:
      1. Extract all tables from every page via pdfplumber
      2. Try each table as a potential data table (look for amount/fee columns)
      3. If no clean table found, fall back to text extraction and parse line by line
    """
    errors = []
    try:
        import pdfplumber
    except ImportError:
        return pd.DataFrame(), ["pdfplumber not installed — PDF parsing unavailable."]

    raw_data = file_obj.read()
    all_tables = []
    full_text = ""

    try:
        with pdfplumber.open(io.BytesIO(raw_data)) as pdf:
            for page in pdf.pages:
                # Extract text for fallback parsing
                page_text = page.extract_text() or ""
                full_text += page_text + "\n"

                tables = page.extract_tables()
                for table in (tables or []):
                    if not table or len(table) < 2:
                        continue
                    # Clean header
                    header = [str(h).strip().replace("\n", " ") if h else f"col_{i}"
                              for i, h in enumerate(table[0])]
                    if not any(h.strip() for h in header):
                        continue
                    rows = []
                    for row in table[1:]:
                        if row and any(cell and str(cell).strip() for cell in row):
                            rows.append({
                                header[i]: str(cell).strip().replace("\n", " ") if cell else ""
                                for i, cell in enumerate(row)
                                if i < len(header)
                            })
                    if rows:
                        all_tables.append(pd.DataFrame(rows))
    except Exception as e:
        errors.append(f"PDF read error: {e}")
        return pd.DataFrame(), errors

    # ── Pick the best table: the one with the most amount/fee-like columns ────
    best_df = None
    best_score = -1
    amount_keywords = ["amount", "gross", "fee", "net", "total", "charge", "payout"]

    for tbl in all_tables:
        score = sum(
            1 for col in tbl.columns
            if any(kw in str(col).lower() for kw in amount_keywords)
        )
        if score > best_score and len(tbl) > 0:
            best_score = score
            best_df = tbl

    if best_df is not None and best_score > 0:
        return best_df, errors

    # ── Fallback: parse payout summary text into a single-row DataFrame ───────
    # Zoho payout PDFs often have a summary block like:
    #   Total Sales:  $X,XXX.XX
    #   Processing Fee:  $XX.XX
    #   Payout Amount:  $X,XXX.XX
    fallback = _parse_zoho_pdf_text(full_text, errors)
    if fallback is not None and not fallback.empty:
        return fallback, errors

    if all_tables:
        # Return biggest table even if no amount columns detected
        return max(all_tables, key=len), errors

    errors.append("Could not extract transaction data from Zoho PDF. "
                  "Please export from Zoho as CSV or Excel instead.")
    return pd.DataFrame(), errors


def _parse_zoho_pdf_text(text: str, errors: list):
    """
    Last-resort parser: scan raw PDF text for dollar amounts and labels.
    Handles Zoho payout summary PDFs that don't have proper table structure.
    """
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Patterns to extract
    # Examples from Zoho payout PDFs:
    #   "Gross Amount  $1,234.56"
    #   "Processing Fees  $12.34"
    #   "Net Amount  $1,222.22"
    #   "Date  Jun 10, 2025"
    #   "Customer  Acme Corp"

    money_re = re.compile(r"\$?\s*([\d,]+\.?\d{0,2})")
    date_re  = re.compile(r"\b(\w{3}\s+\d{1,2},?\s+\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b")

    record = {}
    for line in lines:
        lower = line.lower()
        # Try to extract label: value pairs
        # Look for lines that contain a dollar amount
        m = money_re.search(line)
        if m:
            val = m.group(1).replace(",", "")
            if any(kw in lower for kw in ["gross", "sale", "total amount", "charged"]):
                record["gross_amount"] = val
            elif any(kw in lower for kw in ["processing fee", "merchant fee", "fee", "charge"]):
                record["fee"] = val
            elif any(kw in lower for kw in ["net", "payout", "deposited", "settlement"]):
                record["net_amount"] = val

        dm = date_re.search(line)
        if dm and "date" not in record:
            record["date"] = dm.group(1)

        # Customer name lines (no dollar sign)
        if not m and any(kw in lower for kw in ["customer", "client", "payer", "name"]):
            # Try to get the value after a colon or tab
            parts = re.split(r"[:|\t]{1}", line, maxsplit=1)
            if len(parts) == 2 and parts[1].strip():
                record["customer"] = parts[1].strip()

    if not record:
        return None

    return pd.DataFrame([record])


# ─────────────────────────────────────────────────────────────────────────────
# Reference file loaders
# ─────────────────────────────────────────────────────────────────────────────

def load_customer_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, header=None, dtype=str)
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "account" in vals:
            header_row = i
            break
    if header_row is None:
        header_row = 0
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


def load_cash_codes(path: str) -> pd.DataFrame:
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

def _to_numeric(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.replace(r"[\$,€£]", "", regex=True)
        .str.replace(r"\((\d+\.?\d*)\)", r"-\1", regex=True)
        .str.strip()
        .replace("", np.nan)
        .replace("nan", np.nan)
        .pipe(pd.to_numeric, errors="coerce")
    )
