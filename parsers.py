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
# Each key = our canonical name; value = list of header strings to match
_BOA_COL_MAP = {
    "date":        ["date", "posting date", "post date", "transaction date", "trans date", "effective date"],
    "description": ["description", "payee", "details", "memo", "narrative", "transaction description", "trans desc"],
    "amount":      ["amount", "net amount", "credit", "credit amount", "transaction amount", "trans amount"],
    "debit":       ["debit", "debit amount"],
    "balance":     ["balance", "running balance"],
}

_ZOHO_COL_MAP = {
    "date":         ["date", "payment date", "transaction date", "created date", "created time"],
    "customer":     ["customer name", "customer", "client name", "name", "bill to", "payer", "contact"],
    "gross_amount": ["gross amount", "gross", "gross payment", "amount", "total amount", "payment amount"],
    "fee":          ["processing fee", "merchant fee", "fee", "transaction fee", "charge", "service fee"],
    "net_amount":   ["net amount", "net", "net payment", "settlement amount"],
    "invoice":      ["invoice", "invoice number", "invoice #", "invoice no", "reference"],
    "payment_id":   ["payment id", "payment reference", "transaction id", "id", "zoho payment id"],
    "description":  ["description", "memo", "notes", "payment description"],
    "status":       ["status", "payment status"],
    "email":        ["email", "email address", "customer email"],
}


def _normalise_col(header: str) -> str:
    """Lower, strip, collapse whitespace."""
    return re.sub(r"\s+", " ", str(header).strip().lower())


def _map_columns(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Return {canonical_name: actual_df_column_name} for every alias that
    finds a match.  Unmatched canonicals are absent from the result.
    """
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

def parse_boa(file_obj) -> tuple[pd.DataFrame | None, list]:
    """Parse a Bank of America transaction export (Excel or CSV)."""
    errors = []
    try:
        name = getattr(file_obj, "name", "")
        if name.endswith(".csv"):
            df = _read_boa_csv(file_obj)
        else:
            df = _read_boa_excel(file_obj)
    except Exception as e:
        return None, [f"Failed to read BOA file: {e}"]

    if df is None or df.empty:
        return None, ["BOA file parsed to an empty DataFrame."]

    mapping = _map_columns(df, _BOA_COL_MAP)
    errors += [f"BOA: column '{k}' not found — expected one of {v}"
               for k, v in _BOA_COL_MAP.items()
               if k in ("date", "description") and k not in mapping]

    # Rename to canonical names
    rename = {v: k for k, v in mapping.items()}
    df = df.rename(columns=rename)

    # Ensure canonical columns exist
    for col in ("date", "description", "amount", "debit", "balance"):
        if col not in df.columns:
            df[col] = np.nan

    # Derive a single signed amount column
    # BOA CSV: amount column is the credit; debit column is separate
    if "amount" in df.columns and df["amount"].notna().any():
        df["_boa_amount"] = _to_numeric(df["amount"])
    elif "debit" in df.columns and df["debit"].notna().any():
        df["_boa_amount"] = _to_numeric(df["debit"]) * -1
    else:
        df["_boa_amount"] = np.nan

    # Parse dates
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Filter: keep only Zoho-tagged rows (description contains ZOHO)
    if "description" in df.columns:
        df["_is_zoho"] = df["description"].astype(str).str.upper().str.contains("ZOHO")
    else:
        df["_is_zoho"] = False
        errors.append("BOA: No description column found; cannot filter Zoho rows.")

    df = df.reset_index(drop=True)
    return df, errors


def _read_boa_csv(file_obj) -> pd.DataFrame:
    """Handle BOA CSV exports which often have header rows or metadata lines."""
    content = file_obj.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")

    lines = content.splitlines()
    # Find the first line that looks like a proper header
    header_idx = 0
    for i, line in enumerate(lines):
        if any(kw in line.lower() for kw in ["date", "description", "amount"]):
            header_idx = i
            break

    csv_content = "\n".join(lines[header_idx:])
    return pd.read_csv(io.StringIO(csv_content), dtype=str)


def _read_boa_excel(file_obj) -> pd.DataFrame:
    """Read BOA Excel; tries multiple header rows if row 0 looks like metadata."""
    data = file_obj.read()
    for header_row in range(6):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            if any(any(kw in str(c).lower() for kw in ["date", "description", "amount"])
                   for c in df.columns):
                return df
        except Exception:
            continue
    # Fallback: just read row 0
    return pd.read_excel(io.BytesIO(data), dtype=str)


# ─────────────────────────────────────────────────────────────────────────────
# Zoho parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_zoho(file_obj) -> tuple[pd.DataFrame | None, list]:
    """Parse a Zoho Payments export (Excel, CSV, or PDF)."""
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

    for must_have in ("gross_amount", "fee"):
        if must_have not in mapping:
            errors.append(f"Zoho: Required column '{must_have}' not found.")

    rename = {v: k for k, v in mapping.items()}
    df = df.rename(columns=rename)

    # Ensure all canonical columns exist
    for col in _ZOHO_COL_MAP.keys():
        if col not in df.columns:
            df[col] = np.nan

    # Parse numerics
    for col in ("gross_amount", "fee", "net_amount"):
        if col in df.columns:
            df[col] = _to_numeric(df[col])

    # Parse dates
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Drop summary/total rows
    if "customer" in df.columns:
        df = df[df["customer"].astype(str).str.strip() != ""]

    df = df.reset_index(drop=True)
    return df, errors


def _read_zoho_csv(file_obj) -> pd.DataFrame:
    content = file_obj.read()
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    return pd.read_csv(io.StringIO(content), dtype=str)


def _read_zoho_excel(file_obj) -> pd.DataFrame:
    data = file_obj.read()
    for header_row in range(5):
        try:
            df = pd.read_excel(io.BytesIO(data), header=header_row, dtype=str)
            if any(any(kw in str(c).lower() for kw in ["amount", "customer", "fee"])
                   for c in df.columns):
                return df
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(data), dtype=str)


def _parse_zoho_pdf(file_obj) -> tuple[pd.DataFrame, list]:
    """Extract table data from a Zoho PDF export using pdfplumber."""
    errors = []
    rows = []
    try:
        import pdfplumber
        data = file_obj.read()
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    header = [str(h).strip() if h else "" for h in table[0]]
                    for row in table[1:]:
                        if row:
                            rows.append(dict(zip(header, [str(c).strip() if c else "" for c in row])))
    except ImportError:
        errors.append("pdfplumber not installed — PDF parsing unavailable.")
        return pd.DataFrame(), errors
    except Exception as e:
        errors.append(f"PDF parse error: {e}")
        return pd.DataFrame(), errors

    if not rows:
        errors.append("No table data found in Zoho PDF.")
        return pd.DataFrame(), errors

    return pd.DataFrame(rows), errors


# ─────────────────────────────────────────────────────────────────────────────
# Reference file loaders
# ─────────────────────────────────────────────────────────────────────────────

def load_customer_master(path: str) -> pd.DataFrame:
    """Load IBNY_Business_Customer_Account.xlsx → {Account, Account Name}."""
    df = pd.read_excel(path, header=None, dtype=str)

    # Find the row that contains 'Account' and 'Account Name' headers
    header_row = None
    for i, row in df.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "account" in vals:
            header_row = i
            break

    if header_row is None:
        # Fallback: assume first row
        header_row = 0

    df.columns = df.iloc[header_row].astype(str).str.strip()
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    # Normalise column names
    df.columns = [c.strip() for c in df.columns]
    # Keep only rows where Account looks like BC######
    mask = df.iloc[:, 0].astype(str).str.match(r"BC\d+")
    df = df[mask].reset_index(drop=True)

    # Standardise to two columns: Account, Account Name
    if len(df.columns) >= 2:
        df = df.iloc[:, :2]
        df.columns = ["Account", "Account Name"]
    df = df.dropna(subset=["Account"]).reset_index(drop=True)
    df["Account"] = df["Account"].str.strip()
    df["Account Name"] = df["Account Name"].astype(str).str.strip()
    return df


def load_cash_codes(path: str) -> pd.DataFrame:
    """Load Cash_Code_Masterlist.xlsx → {Cash Code, Cash Code Name}."""
    df = pd.read_excel(path, header=None, dtype=str)

    # Find header row
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
    """Strip currency symbols/commas and coerce to float."""
    return (
        series.astype(str)
        .str.replace(r"[\$,€£]", "", regex=True)
        .str.replace(r"\((\d+\.?\d*)\)", r"-\1", regex=True)  # (123) → -123
        .str.strip()
        .replace("", np.nan)
        .replace("nan", np.nan)
        .pipe(pd.to_numeric, errors="coerce")
    )
