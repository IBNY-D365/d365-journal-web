"""
matcher.py — Match BOA ↔ Zoho transactions and resolve customer accounts.

Key insight from real data analysis:
  - The Zoho Payout PDF has NO customer names (field shows "—")
  - Customer names must come from uploaded invoice PDFs
  - Invoices are matched to Zoho transactions by amount (Total on invoice == gross_amount in Zoho)
  - BOA posting date (payout date) is used as the D365 entry date, not the Zoho transaction date
  - BOA description string (ZOHO PAYMENTS DES:...) is used in the D365 description field

Special account types (from real D365 reference data):
  - Normal business customers: Customer type, BC###### account
  - CS/repair tickets (individuals): Ledger type, account 21040102-B1000002,
    name "Temporary Receipt", description prefixed with "Nicole Holovach CS Ticket #676_"
"""

import re
import numpy as np
import pandas as pd
from rapidfuzz import fuzz, process
from parsers import parse_invoice_pdf

_FUZZY_THRESHOLD = 80

# Special handling: individual/CS customers use Temporary Receipt account
# These are identified when the customer is an individual (not a business)
# and typically come from CS ticket invoices
_TEMP_RECEIPT_ACCOUNT = "21040102-B1000002"
_TEMP_RECEIPT_NAME    = "Temporary Receipt"

# Payment terms → cash code mapping (from automation rules §4)
_TERMS_TO_CASH_CODE = {
    "due on receipt":         ("AR001", ""),
    "due upon receipt":       ("AR001", ""),
    "monthly payment":        ("AR002", "MPP "),
    "monthly payment plan":   ("AR002", "MPP "),
    "mpp":                    ("AR002", "MPP "),
    "financing":              ("AR003", ""),
    "leasing":                ("AR004", ""),
    "net 1":                  ("AR005", ""),
    "net 10":                 ("AR006", ""),
    "net 25":                 ("AR007", ""),
    "net 30":                 ("AR008", ""),
    "net 40":                 ("AR009", ""),
    "net 45":                 ("AR010", ""),
    "net 60":                 ("AR011", ""),
    "due end of next month":  ("AR010", ""),   # ~45-day term
    "net":                    ("AR008", ""),   # generic net → AR008
}


def match_transactions(boa_df, zoho_df, customer_df, invoice_files=None):
    """
    Main entry point. Enriches zoho_df rows with D365 fields.

    Steps:
      1. Parse all uploaded invoice PDFs → {amount: invoice_data}
      2. For each Zoho row: match by amount to invoice → get customer name + terms
      3. Look up customer name in master → get BC###### account
      4. Attach BOA posting date and description
      5. Validate balance invariant
    """
    log = []

    # ── Step 1: Parse invoices ────────────────────────────────────────────────
    invoice_by_amount = {}   # {rounded_amount: invoice_dict}
    invoice_by_name   = {}   # {normalised_name: invoice_dict}

    if invoice_files:
        for inv_file in invoice_files:
            try:
                inv_data = parse_invoice_pdf(inv_file)
                if inv_data:
                    amt = inv_data.get("total")
                    if amt:
                        key = round(float(amt), 2)
                        invoice_by_amount[key] = inv_data
                    name = inv_data.get("customer_name", "")
                    if name:
                        invoice_by_name[_normalise_name(name)] = inv_data
                    log.append({
                        "level": "OK",
                        "msg": (f"Invoice parsed: {inv_data.get('invoice_number','?')} | "
                                f"Customer: {inv_data.get('customer_name','?')} | "
                                f"Total: ${amt:,.2f}" if amt else "Invoice parsed (no total)")
                    })
            except Exception as e:
                log.append({"level": "WARN", "msg": f"Could not parse invoice: {e}"})

    # ── Step 2: Build customer lookup ─────────────────────────────────────────
    customer_lookup = _build_customer_lookup(customer_df)

    # ── Step 3: Resolve each Zoho row ─────────────────────────────────────────
    results = []
    for idx, zrow in zoho_df.iterrows():
        record = _resolve_row(
            zrow, customer_lookup, invoice_by_amount, invoice_by_name, log
        )
        results.append(record)

    resolved = pd.DataFrame(results)
    enriched = pd.concat(
        [zoho_df.reset_index(drop=True), resolved.reset_index(drop=True)], axis=1
    )

    # ── Step 4: Attach BOA date and description ───────────────────────────────
    enriched = _attach_boa_data(enriched, boa_df, log)

    # ── Step 5: Balance check ─────────────────────────────────────────────────
    _validate_balance(boa_df, enriched, log)

    return enriched, log


# ─────────────────────────────────────────────────────────────────────────────
# Row resolution
# ─────────────────────────────────────────────────────────────────────────────

def _resolve_row(zrow, customer_lookup, invoice_by_amount, invoice_by_name, log):
    """Resolve one Zoho transaction row to its D365 fields."""
    idx          = zrow.name
    raw_customer = str(zrow.get("customer", "")).strip()
    gross        = _safe_float(zrow.get("gross_amount"))
    invoice_ref  = str(zrow.get("invoice", "")).strip()

    # ── Try invoice match by amount first (most reliable for this PDF) ────────
    inv = None
    if gross:
        key = round(gross, 2)
        inv = invoice_by_amount.get(key)

    # If not found by amount, try by invoice reference string
    if not inv and invoice_ref and invoice_ref not in ("nan", ""):
        for norm_name, inv_data in invoice_by_name.items():
            if inv_data.get("invoice_number", "") in invoice_ref:
                inv = inv_data
                break

    # ── If we have an invoice, use it ─────────────────────────────────────────
    if inv:
        customer_name  = inv.get("customer_name", "")
        payment_terms  = inv.get("payment_terms", "")
        cs_ticket      = inv.get("cs_ticket", "")
        cash_code, pfx = _terms_to_cash_code(payment_terms)

        log.append({
            "level": "OK",
            "msg": (f"Row {idx}: matched invoice {inv.get('invoice_number','?')} "
                    f"→ '{customer_name}' | terms='{payment_terms}' | code={cash_code}")
        })

        # Check if this is a CS/repair ticket invoice (has [## XXX ##] in items)
        is_cs_ticket = bool(cs_ticket)

        # Normal business customer lookup — works for both business names AND
        # individual names that appear in the CS/PS Ticket column of the master
        master_result = _fuzzy_lookup(customer_name, customer_lookup)
        found_in_master = master_result["_match_confidence"] in ("HIGH", "MEDIUM")

        if is_cs_ticket and not found_in_master:
            # Individual not found in master at all — use Temporary Receipt
            desc_prefix = f"{customer_name} CS Ticket #{cs_ticket}_"
            return {
                "_account":          _TEMP_RECEIPT_ACCOUNT,
                "_account_name":     _TEMP_RECEIPT_NAME,
                "_account_type":     "Ledger",
                "_posting_profile":  "AutoPost",
                "_match_method":     "INVOICE_CS_TICKET_UNREGISTERED",
                "_match_confidence": "MEDIUM",
                "_needs_review":     True,
                "_review_reason":    (f"'{customer_name}' not in Account Masterlist. "
                                      f"Add them with their BC###### to resolve automatically."),
                "_cash_code":        cash_code,
                "_cash_code_prefix": "",
                "_desc_prefix":      desc_prefix,
                "_raw_name":         customer_name,
                "_invoice_number":   inv.get("invoice_number", ""),
            }

        if is_cs_ticket and found_in_master:
            # Individual found via CS/PS Ticket column — use their business account
            # but keep the CS ticket description prefix for the D365 description
            log.append({
                "level": "OK",
                "msg": (f"Row {idx}: CS ticket '{customer_name}' matched to "
                        f"'{master_result['_account_name']}' via CS/PS Ticket column.")
            })
            desc_prefix = f"{customer_name} CS Ticket #{cs_ticket}_"
            master_result["_desc_prefix"]      = desc_prefix
            master_result["_cash_code"]        = cash_code
            master_result["_cash_code_prefix"] = pfx
            master_result["_invoice_number"]   = inv.get("invoice_number", "")
            master_result["_raw_name"]         = customer_name
            master_result.setdefault("_account_type",    "Customer")
            master_result.setdefault("_posting_profile", "AutoPost")
            return master_result
        master_result["_cash_code"]        = cash_code
        master_result["_cash_code_prefix"] = pfx
        master_result["_desc_prefix"]      = ""
        master_result["_invoice_number"]   = inv.get("invoice_number", "")
        master_result["_raw_name"]         = customer_name
        master_result.setdefault("_account_type",    "Customer")
        master_result.setdefault("_posting_profile", "AutoPost")

        if master_result["_match_confidence"] == "LOW":
            master_result["_needs_review"]  = True
            master_result["_review_reason"] = (
                f"Customer '{customer_name}' from invoice not found in master "
                f"(best match score {master_result.get('_fuzzy_score',0)}/100). "
                f"Add this account to IBNY_Business_Customer_Account.xlsx."
            )
            log.append({
                "level": "WARN",
                "msg":   f"Row {idx}: '{customer_name}' not in master — flagged for review."
            })

        return master_result

    # ── No invoice: try customer name from Zoho directly ─────────────────────
    if raw_customer and raw_customer not in ("—", "-", "nan", ""):
        result = _fuzzy_lookup(raw_customer, customer_lookup)
        result["_cash_code"]        = "AR001"  # default
        result["_cash_code_prefix"] = ""
        result["_desc_prefix"]      = ""
        result["_invoice_number"]   = invoice_ref
        result["_raw_name"]         = raw_customer
        result.setdefault("_account_type",    "Customer")
        result.setdefault("_posting_profile", "AutoPost")
        return result

    # ── No name, no invoice match ─────────────────────────────────────────────
    log.append({
        "level": "WARN",
        "msg":   (f"Row {idx}: No customer name and no matching invoice for "
                  f"gross=${gross}. Upload the invoice to resolve.")
    })
    return {
        "_account":          "",
        "_account_name":     raw_customer or "",
        "_account_type":     "Customer",
        "_posting_profile":  "AutoPost",
        "_match_method":     "UNRESOLVED",
        "_match_confidence": "LOW",
        "_needs_review":     True,
        "_review_reason":    f"No customer name and no invoice matched amount ${gross}",
        "_cash_code":        "",
        "_cash_code_prefix": "",
        "_desc_prefix":      "",
        "_raw_name":         raw_customer or "",
        "_invoice_number":   "",
    }


# ─────────────────────────────────────────────────────────────────────────────
# BOA data attachment
# ─────────────────────────────────────────────────────────────────────────────

def _attach_boa_data(zoho_df, boa_df, log):
    """
    Attach the BOA posting date and description to each Zoho row.

    The BOA CSV has one row per payout date. The payout date in Zoho is
    stored in 'payout_date' (from the PDF header). If that matches a BOA
    row date, use it; otherwise fall back to ±3-day window.
    """
    zoho_df = zoho_df.copy()

    if boa_df is None or boa_df.empty:
        zoho_df["_boa_date"]        = pd.NaT
        zoho_df["_boa_description"] = ""
        return zoho_df

    boa_zoho = boa_df[boa_df.get("_is_zoho", pd.Series(False, index=boa_df.index))].copy()

    if boa_zoho.empty:
        zoho_df["_boa_date"]        = pd.NaT
        zoho_df["_boa_description"] = ""
        log.append({
            "level": "WARN",
            "msg":   "No ZOHO rows found in BOA file. "
                     "The BOA description must contain 'ZOHO' to be recognised."
        })
        return zoho_df

    # Build date → {description, amount} map from BOA
    boa_map = {}
    for _, brow in boa_zoho.iterrows():
        if pd.notna(brow.get("date")):
            d = pd.Timestamp(brow["date"]).normalize()
            boa_map[d] = {
                "description": str(brow.get("description", "")),
                "amount":      brow.get("_boa_amount", np.nan),
            }

    def get_boa_for_row(row):
        # Priority 1: use payout_date from Zoho PDF if available
        for date_field in ["payout_date", "date"]:
            val = row.get(date_field)
            if val and str(val) not in ("nan", "NaT", ""):
                try:
                    d = pd.Timestamp(val).normalize()
                    for delta in [0, 1, -1, 2, -2, 3, -3]:
                        target = d + pd.Timedelta(days=delta)
                        if target in boa_map:
                            return boa_map[target]
                except Exception:
                    pass
        # Use the first BOA row if nothing matched
        if boa_map:
            return next(iter(boa_map.values()))
        return {"description": "", "amount": np.nan}

    boa_dates = []
    boa_descs = []
    for _, row in zoho_df.iterrows():
        boa_entry = get_boa_for_row(row)
        boa_dates.append(
            pd.Timestamp(list(boa_map.keys())[0])
            if boa_map else pd.NaT
        )
        boa_descs.append(boa_entry["description"])

    # Use the actual BOA date (from boa_map keys) not the Zoho transaction date
    # The BOA posting date is what goes into D365 (per SOP)
    if boa_map:
        boa_posting_date = list(boa_map.keys())[0]  # first (and usually only) ZOHO row
        zoho_df["_boa_date"] = boa_posting_date
    else:
        zoho_df["_boa_date"] = pd.NaT

    zoho_df["_boa_description"] = boa_descs
    return zoho_df


# ─────────────────────────────────────────────────────────────────────────────
# Balance validation
# ─────────────────────────────────────────────────────────────────────────────

def _validate_balance(boa_df, zoho_df, log):
    if boa_df is None or boa_df.empty:
        return
    if "_boa_amount" not in boa_df.columns:
        return

    boa_zoho = boa_df[boa_df.get("_is_zoho", pd.Series(False, index=boa_df.index))]
    if boa_zoho.empty:
        return

    boa_net    = boa_zoho["_boa_amount"].sum()
    zoho_gross = zoho_df["gross_amount"].fillna(0).sum() if "gross_amount" in zoho_df.columns else 0

    # Use Zoho summary fee (authoritative) if present; fall back to per-txn sum
    if "_summary_fee" in zoho_df.columns:
        summary_fees = zoho_df["_summary_fee"].dropna()
        summary_fees = pd.to_numeric(summary_fees, errors="coerce").dropna()
        zoho_fee = summary_fees.iloc[0] if not summary_fees.empty else zoho_df["fee"].fillna(0).sum()
    else:
        zoho_fee = zoho_df["fee"].fillna(0).sum() if "fee" in zoho_df.columns else 0

    zoho_net   = zoho_gross - zoho_fee

    diff = abs(float(boa_net) - float(zoho_net))
    tol  = max(abs(float(boa_net)) * 0.005, 1.0)

    if diff <= tol:
        log.append({
            "level": "OK",
            "msg":   (f"Balance check PASSED: BOA net=${boa_net:,.2f} "
                      f"≈ Zoho gross${zoho_gross:,.2f} − fee${zoho_fee:,.2f} = ${zoho_net:,.2f}")
        })
    else:
        log.append({
            "level": "WARN",
            "msg":   (f"Balance mismatch: BOA net=${boa_net:,.2f}, "
                      f"Zoho net=${zoho_net:,.2f}, diff=${diff:,.2f} "
                      f"(tolerance=${tol:.2f})")
        })


# ─────────────────────────────────────────────────────────────────────────────
# Customer lookup helpers
# ─────────────────────────────────────────────────────────────────────────────

def _build_customer_lookup(customer_df):
    """
    Build lookup dict: normalised_name → {account, account_name}

    Two entries are created per row when CS/PS Ticket is present:
      1. Normalised Account Name  → used when Zoho/invoice has the business name
      2. Normalised CS/PS Ticket  → used when invoice "Bill To" has an individual name
         (e.g. "Ali Amir" → maps to BC000654 Functional Holistic Healing)

    This means even if the invoice says "Bill To: Ali Amir", the app will
    record the correct D365 account name "Functional Holistic Healing".
    """
    lookup = {}
    for _, row in customer_df.iterrows():
        acc    = str(row.get("Account", "")).strip()
        name   = str(row.get("Account Name", "")).strip()
        ticket = str(row.get("CS/PS Ticket", "")).strip()

        if not acc or not name:
            continue

        entry = {"account": acc, "account_name": name}

        # Primary key: business/account name
        lookup[_normalise_name(name)] = entry

        # Secondary key: individual name from CS/PS Ticket column
        if ticket and ticket.lower() not in ("", "nan", "none"):
            lookup[_normalise_name(ticket)] = entry

    return lookup


def _normalise_name(name):
    name = str(name).lower().strip()
    name = re.sub(r"[^\w\s]", " ", name)
    for suffix in [r"\bllc\b", r"\binc\b", r"\bltd\b", r"\bcorp\b",
                   r"\bpllc\b", r"\bpc\b", r"\bdba\b", r"\bthe\b"]:
        name = re.sub(suffix, "", name)
    return re.sub(r"\s+", " ", name).strip()


def _fuzzy_lookup(raw_name, customer_lookup):
    norm = _normalise_name(raw_name)
    keys = list(customer_lookup.keys())

    if not keys:
        return _unresolved(raw_name, "Customer master is empty")

    result   = process.extractOne(norm, keys, scorer=fuzz.token_set_ratio)
    best_key = result[0]
    score    = result[1]
    matched  = customer_lookup[best_key]

    if score >= _FUZZY_THRESHOLD:
        confidence = "HIGH" if score >= 92 else "MEDIUM"
        return {
            "_account":          matched["account"],
            "_account_name":     matched["account_name"],
            "_account_type":     "Customer",
            "_posting_profile":  "AutoPost",
            "_match_method":     f"FUZZY({score})",
            "_match_confidence": confidence,
            "_needs_review":     confidence == "MEDIUM",
            "_review_reason":    (f"Fuzzy match '{matched['account_name']}' "
                                  f"score {score}/100 — verify") if confidence == "MEDIUM" else "",
            "_fuzzy_score":      score,
        }
    return _unresolved(raw_name,
                       f"Best match '{matched['account_name']}' scored {score}/100 < {_FUZZY_THRESHOLD}")


def _unresolved(raw_name, reason):
    return {
        "_account":          "",
        "_account_name":     raw_name,
        "_account_type":     "Customer",
        "_posting_profile":  "AutoPost",
        "_match_method":     "UNRESOLVED",
        "_match_confidence": "LOW",
        "_needs_review":     True,
        "_review_reason":    reason,
        "_fuzzy_score":      0,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Payment term → cash code
# ─────────────────────────────────────────────────────────────────────────────

def _terms_to_cash_code(terms_str):
    """Return (cash_code, description_prefix) from a payment terms string."""
    if not terms_str or str(terms_str).lower() in ("nan", ""):
        return "AR001", ""
    terms_lower = terms_str.lower().strip()
    for key, val in _TERMS_TO_CASH_CODE.items():
        if key in terms_lower:
            return val
    return "AR001", ""   # default: due on receipt


def _safe_float(val):
    try:
        f = float(str(val).replace(",", "").replace("$", "").strip())
        return None if np.isnan(f) else f
    except (ValueError, TypeError):
        return None
