"""D365 journal entry builder.

Purpose
-------
Read source files from:
- Bank of America CSV exports
- Zoho payment summary files (PDF, CSV, XLSX)
- Customer invoice PDFs

Then normalize the data and write a D365 General Journal workbook using a
configuration file that defines the exact column mapping.

Important
---------
This script is intentionally configuration-driven. It does NOT assume field
names from your source files. You must supply the exact mappings from your
Zoho Payment_D365_Automation_Rules document in a YAML file.

Suggested repo layout
---------------------
.
├── d365_journal_builder.py
├── mappings.yaml
└── requirements.txt

Example mappings.yaml
---------------------
source_files:
  boa_csv:
    transaction_date: Transaction Date
    description: Description
    amount: Amount
    reference: Reference
  zoho_payment_summary:
    payment_date: Payment Date
    invoice_number: Invoice Number
    customer_name: Customer Name
    payment_amount: Amount
  invoice_pdf:
    invoice_number_regex: 'Invoice\s*#\s*([A-Z0-9-]+)'

d365_template:
  header_row: 1
  columns:
    journal_batch_number: Journal batch number
    account_type: Account type
    account: Account
    debit: Debit
    credit: Credit
    text: Text
    offset_account_type: Offset account type
    offset_account: Offset account
    transaction_date: Transaction date

journal_rules:
  journal_batch_number: AUTO-001
  debit_line:
    account_type: Ledger
    account: 111100
  credit_line:
    account_type: Customer
    offset_account_type: Bank
  description_template: 'Payment {invoice_number} - {customer_name}'
"""

from __future__ import annotations

import argparse
import json
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

try:
    import pdfplumber
except ImportError:  # pragma: no cover
    pdfplumber = None

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger(__name__)


@dataclass
class NormalizedTransaction:
    transaction_date: Optional[str]
    source_type: str
    reference: Optional[str]
    invoice_number: Optional[str]
    customer_name: Optional[str]
    amount: float
    description: str
    raw: Dict[str, Any]


class MappingError(RuntimeError):
    pass


def load_yaml(path: Path) -> Dict[str, Any]:
    if yaml is None:
        raise RuntimeError("PyYAML is required. Install with pip install pyyaml")
    with path.open("r", encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def normalize_amount(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    text = text.replace(",", "")
    # Parentheses imply negative values in many bank exports.
    is_negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()$")
    text = text.replace("$", "")
    amount = float(text)
    return -amount if is_negative else amount


def first_nonempty(*values: Any) -> Optional[str]:
    for value in values:
        if value is None:
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        text = str(value).strip()
        if text:
            return text
    return None


def read_csv_file(path: Path) -> pd.DataFrame:
    return pd.read_csv(path)


def read_xlsx_file(path: Path) -> pd.DataFrame:
    return pd.read_excel(path)


def extract_pdf_text(path: Path) -> str:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is required. Install with pip install pdfplumber")
    text_chunks: List[str] = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            text_chunks.append(page.extract_text() or "")
    return "\n".join(text_chunks)


def parse_pdf_table_like_text(text: str) -> pd.DataFrame:
    """Fallback parser for text-based PDFs.

    This does not assume a single PDF format. It attempts to split lines into
    columns after removing obvious header noise. For production use, map the
    exact PDF layout in your config and add a dedicated parser.
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    rows: List[List[str]] = []
    for line in lines:
        # Split on 2+ spaces to preserve fields that contain single spaces.
        parts = re.split(r"\s{2,}", line)
        if len(parts) >= 2:
            rows.append(parts)
    if not rows:
        return pd.DataFrame()
    max_len = max(len(r) for r in rows)
    padded = [r + [None] * (max_len - len(r)) for r in rows]
    return pd.DataFrame(padded)


def extract_invoice_number(text: str, regex_pattern: str) -> Optional[str]:
    match = re.search(regex_pattern, text, flags=re.IGNORECASE)
    return match.group(1).strip() if match else None


def load_source_file(path: Path, kind: str) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return read_csv_file(path)
    if suffix in {".xlsx", ".xls"}:
        return read_xlsx_file(path)
    if suffix == ".pdf":
        text = extract_pdf_text(path)
        return parse_pdf_table_like_text(text)
    raise MappingError(f"Unsupported file type for {kind}: {path.name}")


def map_boa_transactions(df: pd.DataFrame, cfg: Dict[str, Any]) -> List[NormalizedTransaction]:
    m = cfg["source_files"]["boa_csv"]
    tx_date_col = m["transaction_date"]
    desc_col = m["description"]
    amount_col = m["amount"]
    ref_col = m.get("reference")

    out: List[NormalizedTransaction] = []
    for _, row in df.iterrows():
        amount = normalize_amount(row.get(amount_col))
        out.append(
            NormalizedTransaction(
                transaction_date=first_nonempty(row.get(tx_date_col)),
                source_type="boa_csv",
                reference=first_nonempty(row.get(ref_col)) if ref_col else None,
                invoice_number=None,
                customer_name=None,
                amount=amount,
                description=first_nonempty(row.get(desc_col)) or "",
                raw=row.to_dict(),
            )
        )
    return out


def map_zoho_summary(df: pd.DataFrame, cfg: Dict[str, Any]) -> List[NormalizedTransaction]:
    m = cfg["source_files"]["zoho_payment_summary"]
    payment_date_col = m["payment_date"]
    invoice_col = m["invoice_number"]
    customer_col = m["customer_name"]
    amount_col = m["payment_amount"]

    out: List[NormalizedTransaction] = []
    for _, row in df.iterrows():
        amount = normalize_amount(row.get(amount_col))
        out.append(
            NormalizedTransaction(
                transaction_date=first_nonempty(row.get(payment_date_col)),
                source_type="zoho_payment_summary",
                reference=first_nonempty(row.get(invoice_col)),
                invoice_number=first_nonempty(row.get(invoice_col)),
                customer_name=first_nonempty(row.get(customer_col)),
                amount=amount,
                description="",
                raw=row.to_dict(),
            )
        )
    return out


def map_invoice_pdf(path: Path, cfg: Dict[str, Any]) -> Dict[str, Any]:
    m = cfg["source_files"]["invoice_pdf"]
    regex_pattern = m["invoice_number_regex"]
    text = extract_pdf_text(path)
    return {
        "file": path.name,
        "invoice_number": extract_invoice_number(text, regex_pattern),
        "raw_text": text,
    }


def build_d365_rows(transactions: Iterable[NormalizedTransaction], cfg: Dict[str, Any]) -> pd.DataFrame:
    rules = cfg["journal_rules"]
    cols = cfg["d365_template"]["columns"]

    rows: List[Dict[str, Any]] = []
    batch = rules["journal_batch_number"]
    debit_rule = rules["debit_line"]
    credit_rule = rules["credit_line"]
    description_template = rules.get("description_template", "")

    for tx in transactions:
        context = {
            "invoice_number": tx.invoice_number or tx.reference or "",
            "customer_name": tx.customer_name or "",
            "amount": tx.amount,
            "description": tx.description,
            "transaction_date": tx.transaction_date or "",
        }
        text = description_template.format(**context) if description_template else tx.description
        amount = abs(tx.amount)

        # Debit line
        rows.append({
            cols["journal_batch_number"]: batch,
            cols["account_type"]: debit_rule["account_type"],
            cols["account"]: debit_rule["account"],
            cols["debit"]: amount if amount > 0 else 0,
            cols["credit"]: 0,
            cols["text"]: text,
            cols["offset_account_type"]: credit_rule["offset_account_type"],
            cols["offset_account"]: credit_rule.get("offset_account", tx.reference or ""),
            cols["transaction_date"]: tx.transaction_date,
            "Source Type": tx.source_type,
            "Source Reference": tx.reference,
            "Invoice Number": tx.invoice_number,
            "Customer Name": tx.customer_name,
            "Raw Amount": tx.amount,
        })

        # Credit line
        rows.append({
            cols["journal_batch_number"]: batch,
            cols["account_type"]: credit_rule["account_type"],
            cols["account"]: credit_rule["account"],
            cols["debit"]: 0,
            cols["credit"]: amount if amount > 0 else 0,
            cols["text"]: text,
            cols["offset_account_type"]: debit_rule["account_type"],
            cols["offset_account"]: debit_rule["account"],
            cols["transaction_date"]: tx.transaction_date,
            "Source Type": tx.source_type,
            "Source Reference": tx.reference,
            "Invoice Number": tx.invoice_number,
            "Customer Name": tx.customer_name,
            "Raw Amount": tx.amount,
        })

    return pd.DataFrame(rows)


def reorder_columns(df: pd.DataFrame, desired_order: List[str]) -> pd.DataFrame:
    existing = [c for c in desired_order if c in df.columns]
    extras = [c for c in df.columns if c not in existing]
    return df[existing + extras]


def process_inputs(input_dir: Path, cfg: Dict[str, Any]) -> pd.DataFrame:
    all_transactions: List[NormalizedTransaction] = []

    for path in sorted(input_di
