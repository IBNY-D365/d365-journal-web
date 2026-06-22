"""
builder.py — Build D365 journal entry rows from matched Zoho data.

Rules source: Zoho_Payment_D365_Automation_Rules.docx §4 & §5

Per transaction creates exactly:
  • 1 CREDIT line  (customer payment)
  • 1 DEBIT line   (Zoho merchant fee)

Multi-payment batch rule (§3.3):
  When multiple Zoho rows share the same BOA date, merchant fees are
  GROUPED into one debit line; customer payment credits remain individual.

Cash code assignment (§4):
  AR001 — due on receipt (default)
  AR002 — Monthly Payment Plan  → description prefixed "MPP "
  AR003–AR012 — other payment terms (require invoice lookup)
  Unknown → blank, flagged for review
"""

import re
import pandas as pd
import numpy as np

# ── Constants from automation rules ──────────────────────────────────────────
_COMPANY          = "bwa"
_ACCT_TYPE_CREDIT = "Customer"
_ACCT_TYPE_DEBIT  = "Ledger"
_POSTING_PROFILE  = "AutoPost"
_DEBIT_ACCT_NAME  = "Outside Service (Finance)"
_DEBIT_ACCT       = "43170111-U26C05001-B735350-UOA003"
_DEBIT_CASH_CODE  = "OSF005"
_CURRENCY         = "USD"
_EXCHANGE_RATE    = 1.00
_SALES_TAX_GROUP  = "AVATAX"
_REVERSING_ENTRY  = "No"

# The 25 D365 column names in exact required order
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
    # Internal review columns (stripped before export)
    "_needs_review", "_review_reason", "_match_confidence",
]


def build_journal_entries(
    matched_df: pd.DataFrame,
    customer_df: pd.DataFrame,
    offset_account: str = "B1000002",
) -> tuple[pd.DataFrame, list]:
    """
    Convert matched Zoho rows → D365 journal entry rows.
    Returns (journal_df, build_log).
    """
    log  = []
    rows = []

    # ── Group by BOA date to implement batch merchant fee rule ────────────────
    date_col = "_boa_date" if "_boa_date" in matched_df.columns else "date"
    matched_df = matched_df.copy()
    matched_df["_group_date"] = pd.to_datetime(
        matched_df[date_col] if date_col in matched_df.columns else matched_df.get("date"),
        errors="coerce"
    )

    groups = matched_df.groupby("_group_date", dropna=False)

    for group_date, group in groups:
        group_rows  = list(group.iterrows())
        multi_batch = len(group_rows) > 1

        credit_rows  = []
        total_fee    = 0.0
        fee_desc_parts = []

        for _, zrow in group_rows:
            # ── Resolve fields ────────────────────────────────────────────────
            account      = str(zrow.get("_account", "")).strip()
            account_name = str(zrow.get("_account_name", "")).strip()
            needs_review = bool(zrow.get("_needs_review", False))
            review_reason= str(zrow.get("_review_reason", ""))
            confidence   = str(zrow.get("_match_confidence", "LOW"))
            gross        = _safe_float(zrow.get("gross_amount"))
            fee          = _safe_float(zrow.get("fee"))
            boa_desc     = str(zrow.get("_boa_description", "")).strip()
            cash_code    = str(zrow.get("_cash_code", "")).strip()

            # ── Determine cash code from Zoho description / invoice ───────────
            if not cash_code:
                cash_code, prefix = _infer_cash_code(zrow, boa_desc)
            else:
                prefix = "MPP " if cash_code == "AR002" else ""

            if not cash_code:
                needs_review  = True
                review_reason = (review_reason + "; Cash code unknown — check payment term").strip("; ")

            # ── Build credit description ──────────────────────────────────────
            credit_desc = _build_credit_description(
                prefix, account, account_name, boa_desc, cash_code
            )

            # ── Use date: prefer BOA posting date, else Zoho date ────────────
            entry_date = group_date
            if pd.isna(entry_date):
                entry_date = pd.to_datetime(zrow.get("date"), errors="coerce")

            date_str = entry_date.strftime("%m/%d/%Y") if pd.notna(entry_date) else ""

            # ── Credit line ───────────────────────────────────────────────────
            credit_row = {
                "Date":                  date_str,
                "Voucher":               "",
                "Account name":          account_name or "",
                "Company":               _COMPANY,
                "Account type":          _ACCT_TYPE_CREDIT if account else "",
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

            if fee and fee > 0:
                total_fee += fee

            # Build fee description parts for batch debit
            if account or account_name:
                fee_desc_parts.append(f"{account} {account_name}".strip())

        # ── Debit (merchant fee) line: one per date group ─────────────────────
        if total_fee > 0:
            if multi_batch:
                # Multi-payment: aggregate description (§5.2)
                fee_desc_str = ", ".join(fee_desc_parts)
                debit_desc = _build_debit_description_multi(fee_desc_str, boa_desc)
            else:
                # Single payment: use the single credit description
                debit_desc = _build_debit_description_single(
                    credit_rows[0]["Description"] if credit_rows else ""
                )

            # Use date from first credit row in this group
            date_str = credit_rows[0]["Date"] if credit_rows else ""

            debit_row = {
                "Date":                  date_str,
                "Voucher":               "",
                "Account name":          _DEBIT_ACCT_NAME,
                "Company":               _COMPANY,
                "Account type":          _ACCT_TYPE_DEBIT,
                "Account":               _DEBIT_ACCT,
                "Posting profile":       "",
                "Cash code":             _DEBIT_CASH_CODE,
                "Description":           debit_desc,
                "Debit":                 f"{total_fee:.2f}",
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
            }
            rows.append(debit_row)
            log.append({"level": "OK",
                        "msg": f"Built {'batch' if multi_batch else 'single'} debit row for "
                               f"{date_str or 'undated'} — fee ${total_fee:.2f}"})

    if not rows:
        log.append({"level": "WARN", "msg": "No journal entry rows were generated."})
        return pd.DataFrame(columns=[c for c in D365_COLUMNS if not c.startswith("_")]), log

    journal_df = pd.DataFrame(rows)

    # Ensure all D365 columns exist
    for col in D365_COLUMNS:
        if col not in journal_df.columns:
            journal_df[col] = ""

    return journal_df, log


# ─────────────────────────────────────────────────────────────────────────────
# Cash code inference
# ─────────────────────────────────────────────────────────────────────────────

_MPP_PATTERNS = [
    re.compile(r"\bMPP\b", re.IGNORECASE),
    re.compile(r"monthly\s+payment\s+plan", re.IGNORECASE),
    re.compile(r"monthly\s+payment", re.IGNORECASE),
]

_PAYMENT_TERM_MAP = {
    re.compile(r"\bnet\s*1\b",  re.IGNORECASE): "AR005",
    re.compile(r"\bnet\s*10\b", re.IGNORECASE): "AR006",
    re.compile(r"\bnet\s*25\b", re.IGNORECASE): "AR007",
    re.compile(r"\bnet\s*30\b", re.IGNORECASE): "AR008",
    re.compile(r"\bnet\s*40\b", re.IGNORECASE): "AR009",
    re.compile(r"\bnet\s*45\b", re.IGNORECASE): "AR010",
    re.compile(r"\bnet\s*60\b", re.IGNORECASE): "AR011",
    re.compile(r"\bfinancing\b",re.IGNORECASE): "AR003",
    re.compile(r"\bleasing\b",  re.IGNORECASE): "AR004",
}


def _infer_cash_code(zrow: pd.Series, boa_desc: str) -> tuple[str, str]:
    """
    Infer cash code from available text signals.
    Returns (cash_code, description_prefix).
    """
    # Combine all text fields for inspection
    search_text = " ".join([
        str(zrow.get("description", "")),
        str(zrow.get("invoice", "")),
        boa_desc,
    ])

    # Check MPP signals first
    for pat in _MPP_PATTERNS:
        if pat.search(search_text):
            return "AR002", "MPP "

    # Check Net/Financing/Leasing patterns
    for pat, code in _PAYMENT_TERM_MAP.items():
        if pat.search(search_text):
            return code, ""

    # Default: AR001 (due on receipt)
    return "AR001", ""


# ─────────────────────────────────────────────────────────────────────────────
# Description builders — exact format per automation rules §5.1 / §5.2
# ─────────────────────────────────────────────────────────────────────────────

def _build_credit_description(
    prefix: str,          # "MPP " or ""
    account: str,         # BC000422
    account_name: str,    # Vizzhy Inc
    boa_desc: str,        # ZOHO PAYMENTS DES:... CCD
    cash_code: str,
) -> str:
    """
    Format per SOP (Zoho_Payment_D365_Automation_Rules §5.1):
      AR001 (no prefix):  "AR001: BC000571 Page Fit Inc. DBA Intoxx Fitness_ZOHO PAYMENTS..."
      AR002 (MPP prefix): "MPP BC000327 Elite Functional Wellness_BANKCARD-1869 DES:..."
        → MPP replaces the "AR002:" code prefix; cash code does NOT appear in description.
      Other AR codes:     "{cash_code}: {account} {name}_{boa_desc}"
    """
    name_part = " ".join(filter(None, [account, account_name])).strip()

    if prefix:
        # MPP (and any future term prefix): prefix replaces the code token
        # "MPP BC000327 Elite Functional Wellness_BOA_DESC"
        if boa_desc:
            return f"{prefix}{name_part}_{boa_desc}"
        return f"{prefix}{name_part}"
    else:
        # Standard: lead with cash code
        code_prefix = f"{cash_code}: " if cash_code else ""
        if boa_desc:
            return f"{code_prefix}{name_part}_{boa_desc}"
        return f"{code_prefix}{name_part}"


def _build_debit_description_single(credit_desc: str) -> str:
    """Single payment: prefix credit description with 'Zoho Merchant Fee_'."""
    if not credit_desc:
        return "Zoho Merchant Fee_"
    # Remove the cash code prefix (e.g. "AR001: ") for the debit description
    clean = re.sub(r"^(AR\d{3}|MPP\s+AR\d{3}):\s*", "", credit_desc)
    # Per SOP examples: "Zoho Merchant Fee BC000422 Vizzhy Inc_ZOHO PAYMENTS..."
    # (no underscore between "Fee" and the account)
    return f"Zoho Merchant Fee {clean}"


def _build_debit_description_multi(fee_desc_str: str, boa_desc: str) -> str:
    """
    Multi-payment batch: 'Zoho Merchant Fee {acct1} name1, {acct2} name2_BOA_DESC'
    Per automation rules §5.2 multi-payment example.
    """
    if boa_desc:
        return f"Zoho Merchant Fee {fee_desc_str}_{boa_desc}"
    return f"Zoho Merchant Fee {fee_desc_str}"


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _safe_float(val) -> float | None:
    try:
        f = float(str(val).replace(",", "").replace("$", "").strip())
        return f if not np.isnan(f) else None
    except (ValueError, TypeError):
        return None
