"""
builder.py — Build D365 journal entry rows from matched Zoho data.

Rules source: Zoho_Payment_D365_Automation_Rules.docx §4 & §5

Per transaction (grouped by BOA posting date):
  • 1 CREDIT line per customer payment
  • 1 DEBIT line for ALL merchant fees (grouped per payout batch)

Special account handling (from real D365 reference data):
  • Normal customers: Account type=Customer, Account=BC######, Posting Profile=AutoPost
  • CS/repair tickets: Account type=Ledger, Account=21040102-B1000002,
    Account Name=Temporary Receipt, Posting Profile=AutoPost
    Description: "Nicole Holovach CS Ticket #676_ZOHO PAYMENTS DES:..."
"""

import re
import pandas as pd
import numpy as np

_COMPANY          = "bwa"
_POSTING_PROFILE  = "AutoPost"
_DEBIT_ACCT_NAME  = "Outside Service (Finance)"
_DEBIT_ACCT_TYPE  = "Ledger"
_DEBIT_ACCT       = "43170111-U26C05001-B735350-UOA003"
_DEBIT_CASH_CODE  = "OSF005"
_CURRENCY         = "USD"
_EXCHANGE_RATE    = 1.00
_SALES_TAX_GROUP  = "AVATAX"
_REVERSING_ENTRY  = "No"

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
    log  = []
    rows = []

    # Determine grouping date: use BOA posting date (_boa_date) if available,
    # fall back to Zoho transaction date
    matched_df = matched_df.copy()
    if "_boa_date" in matched_df.columns and matched_df["_boa_date"].notna().any():
        matched_df["_group_date"] = pd.to_datetime(
            matched_df["_boa_date"], errors="coerce"
        )
    elif "date" in matched_df.columns:
        matched_df["_group_date"] = pd.to_datetime(
            matched_df["date"], errors="coerce"
        )
    else:
        matched_df["_group_date"] = pd.NaT

    groups = matched_df.groupby("_group_date", dropna=False)

    for group_date, group in groups:
        group_rows   = list(group.iterrows())
        multi_batch  = len(group_rows) > 1
        credit_rows      = []
        total_fee        = 0.0   # fallback: sum of per-txn fees
        summary_fee      = None  # authoritative: from Zoho summary line
        fee_desc_parts   = []

        for _, zrow in group_rows:
            account       = str(zrow.get("_account", "")).strip()
            account_name  = str(zrow.get("_account_name", "")).strip()
            account_type  = str(zrow.get("_account_type",  "Customer")).strip()
            posting_prof  = str(zrow.get("_posting_profile", _POSTING_PROFILE)).strip()
            needs_review  = bool(zrow.get("_needs_review", False))
            review_reason = str(zrow.get("_review_reason", ""))
            confidence    = str(zrow.get("_match_confidence", "LOW"))
            gross         = _safe_float(zrow.get("gross_amount"))
            fee           = _safe_float(zrow.get("fee"))
            boa_desc      = str(zrow.get("_boa_description", "")).strip()
            cash_code     = str(zrow.get("_cash_code", "")).strip()
            cash_pfx      = str(zrow.get("_cash_code_prefix", "")).strip()
            desc_prefix   = str(zrow.get("_desc_prefix", "")).strip()

            # Default cash code if not set
            if not cash_code:
                cash_code = "AR001"

            # Format entry date from BOA posting date
            date_str = ""
            if pd.notna(group_date):
                try:
                    date_str = pd.Timestamp(group_date).strftime("%m/%d/%Y")
                except Exception:
                    date_str = str(group_date)

            # Build credit description
            credit_desc = _build_credit_description(
                cash_code, cash_pfx, desc_prefix,
                account, account_name, boa_desc
            )

            credit_row = {
                "Date":                  date_str,
                "Voucher":               "",
                "Account name":          account_name,
                "Company":               _COMPANY,
                "Account type":          account_type,
                "Account":               account,
                "Posting profile":       posting_prof if account else "",
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

            # Prefer the Zoho summary-level fee over per-txn sum
            sf = _safe_float(zrow.get("_summary_fee"))
            if sf and sf > 0:
                summary_fee = sf

            # Accumulate fee description parts
            part = " ".join(filter(None, [account, account_name])).strip()
            if part:
                fee_desc_parts.append(part)

        # ── Use Zoho summary fee (authoritative) over per-txn sum ───────────
        # The summary line "Payments N $X −$Y $Z" is the source of truth per SOP.
        # Never assume or recalculate — use exactly what Zoho reports.
        debit_fee = summary_fee if summary_fee else total_fee

        # ── Single grouped debit row for all fees in this batch ───────────────
        if debit_fee > 0:
            date_str = credit_rows[0]["Date"] if credit_rows else ""

            if multi_batch:
                fee_name_str = ", ".join(fee_desc_parts)
                debit_desc = f"Zoho Merchant Fee {fee_name_str}_{boa_desc}" if boa_desc else f"Zoho Merchant Fee {fee_name_str}"
            else:
                # Single payment: "Zoho Merchant Fee_ {credit_desc without code prefix}"
                base = credit_rows[0]["Description"] if credit_rows else ""
                # Remove cash code prefix (e.g. "AR001: ") or MPP prefix from start
                base_clean = re.sub(r"^(AR\d{3}:\s*|MPP\s+)", "", base)
                debit_desc = f"Zoho Merchant Fee_ {base_clean}".strip()

            debit_row = {
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
            }
            rows.append(debit_row)
            log.append({
                "level": "OK",
                "msg": (f"Built {'batch' if multi_batch else 'single'} debit row "
                        f"for {date_str or 'undated'} — fee ${debit_fee:.2f} "
                        f"(source: {'Zoho summary' if summary_fee else 'per-txn sum'})")
            })

    if not rows:
        log.append({"level": "WARN", "msg": "No journal entry rows were generated."})
        return pd.DataFrame(columns=[c for c in D365_COLUMNS if not c.startswith("_")]), log

    journal_df = pd.DataFrame(rows)
    for col in D365_COLUMNS:
        if col not in journal_df.columns:
            journal_df[col] = ""

    return journal_df, log


# ─────────────────────────────────────────────────────────────────────────────
# Description builders
# ─────────────────────────────────────────────────────────────────────────────

def _build_credit_description(cash_code, cash_pfx, desc_prefix,
                               account, account_name, boa_desc):
    """
    Build the D365 credit line description.

    Normal (AR001):  "AR001: BC000649 Equity Now Inc_ZOHO PAYMENTS DES:..."
    MPP    (AR002):  "MPP BC000327 Elite Functional Wellness_ZOHO PAYMENTS DES:..."
    CS Ticket:       "Nicole Holovach CS Ticket #676_ZOHO PAYMENTS DES:..."

    Rules:
      - If desc_prefix is set (CS ticket), it replaces everything before the BOA desc
      - If cash_pfx is "MPP ", no cash code token in description
      - Otherwise, lead with "{cash_code}: "
    """
    name_part = " ".join(filter(None, [account, account_name])).strip()

    if desc_prefix:
        # CS ticket: "Nicole Holovach CS Ticket #676_ZOHO PAYMENTS DES:..."
        if boa_desc:
            return f"{desc_prefix}{boa_desc}"
        return desc_prefix.rstrip("_")

    if cash_pfx:
        # MPP: "MPP BC000327 Name_ZOHO PAYMENTS DES:..."
        if boa_desc:
            return f"{cash_pfx}{name_part}_{boa_desc}"
        return f"{cash_pfx}{name_part}"

    # Standard: "AR001: BC000649 Equity Now Inc_ZOHO PAYMENTS DES:..."
    code_token = f"{cash_code}: " if cash_code else ""
    if boa_desc:
        return f"{code_token}{name_part}_{boa_desc}"
    return f"{code_token}{name_part}"


def _safe_float(val):
    try:
        f = float(str(val).replace(",", "").replace("$", "").strip())
        return None if np.isnan(f) else f
    except (ValueError, TypeError):
        return None
