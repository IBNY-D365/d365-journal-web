"""
matcher.py — Match BOA transactions to Zoho records and resolve customer accounts.

Matching strategy (in priority order):
  1. Exact customer name match against customer master
  2. Fuzzy name match (token_set_ratio ≥ 85)
  3. Invoice number cross-reference
  4. Email-to-name lookup (if customer has email in Zoho but no name)
  5. Flag for manual review if confidence < threshold

For the BOA ↔ Zoho financial reconciliation:
  - BOA Net = sum(Zoho Gross) - sum(Zoho Fees)  [Mathematical Balance Invariant]
  - Tolerance: max(0.5% of amount, $1.00)
"""

import re
import numpy as np
import pandas as pd
from fuzzywuzzy import fuzz, process

# Fuzzy match threshold (0–100); below this → NEEDS_REVIEW
_FUZZY_THRESHOLD = 82


def match_transactions(
    boa_df: pd.DataFrame,
    zoho_df: pd.DataFrame,
    customer_df: pd.DataFrame,
    invoice_files=None,
) -> tuple[pd.DataFrame, list]:
    """
    Main entry point.  Returns (enriched_zoho_df, log_entries).

    Adds columns to zoho_df:
      _account        — resolved BC###### account number
      _account_name   — canonical account name from customer master
      _match_method   — how the name was resolved
      _match_confidence — HIGH / MEDIUM / LOW
      _needs_review   — bool
      _review_reason  — human-readable reason if flagged
      _boa_date       — posting date from BOA record
      _boa_description — full BOA description string
      _cash_code      — AR001 / AR002 etc. (may be blank if unknown)
    """
    log = []
    customer_lookup = _build_customer_lookup(customer_df)

    # ── Step 1: resolve each Zoho row to a customer account ──────────────────
    results = []
    for idx, zrow in zoho_df.iterrows():
        record = _resolve_zoho_row(zrow, customer_lookup, invoice_files, log)
        results.append(record)

    resolved = pd.DataFrame(results)
    enriched = pd.concat([zoho_df.reset_index(drop=True), resolved.reset_index(drop=True)], axis=1)

    # ── Step 2: link BOA date & description to Zoho rows ─────────────────────
    zoho_rows = enriched  # alias

    # Filter BOA to Zoho-only rows
    boa_zoho = boa_df[boa_df.get("_is_zoho", pd.Series(False, index=boa_df.index))].copy()

    if not boa_zoho.empty and "date" in boa_zoho.columns and "date" in zoho_rows.columns:
        # Group Zoho rows by date; attach matching BOA description by date
        boa_date_map = {}
        for _, brow in boa_zoho.iterrows():
            if pd.notna(brow.get("date")):
                d = pd.Timestamp(brow["date"]).normalize()
                # Store as list to handle multiple BOA rows on same date
                boa_date_map.setdefault(d, []).append({
                    "description": str(brow.get("description", "")),
                    "amount":      brow.get("_boa_amount", np.nan),
                })

        def attach_boa(row):
            if pd.isna(row.get("date")):
                return pd.Series({"_boa_date": pd.NaT, "_boa_description": ""})
            d = pd.Timestamp(row["date"]).normalize()
            # Try exact date, then ±3 days
            for delta in [0, 1, -1, 2, -2, 3, -3]:
                target = d + pd.Timedelta(days=delta)
                if target in boa_date_map:
                    entries = boa_date_map[target]
                    return pd.Series({
                        "_boa_date":        target,
                        "_boa_description": entries[0]["description"],  # closest match
                    })
            return pd.Series({"_boa_date": pd.NaT, "_boa_description": ""})

        attached = zoho_rows.apply(attach_boa, axis=1)
        zoho_rows = pd.concat([zoho_rows, attached], axis=1)

        # Validate mathematical balance invariant per date group
        _validate_balance(boa_zoho, zoho_rows, log)
    else:
        zoho_rows["_boa_date"]        = pd.NaT
        zoho_rows["_boa_description"] = ""
        if boa_zoho.empty:
            log.append({"level": "WARN", "msg": "No Zoho-tagged rows found in BOA file (description must contain 'ZOHO')."})

    return zoho_rows, log


# ─────────────────────────────────────────────────────────────────────────────
# Customer resolution
# ─────────────────────────────────────────────────────────────────────────────

def _build_customer_lookup(customer_df: pd.DataFrame) -> dict:
    """Build a normalised name → {Account, Account Name} lookup."""
    lookup = {}
    for _, row in customer_df.iterrows():
        acc  = str(row.get("Account", "")).strip()
        name = str(row.get("Account Name", "")).strip()
        if acc and name:
            key = _normalise_name(name)
            lookup[key] = {"account": acc, "account_name": name}
    return lookup


def _normalise_name(name: str) -> str:
    """Lower-case, strip punctuation and common legal suffixes for fuzzy matching."""
    name = str(name).lower().strip()
    name = re.sub(r"[^\w\s]", " ", name)
    # Remove common suffixes that vary between sources
    for suffix in [r"\bllc\b", r"\binc\b", r"\bltd\b", r"\bcorp\b", r"\bpllc\b",
                   r"\bpc\b", r"\bdba\b", r"\bthe\b"]:
        name = re.sub(suffix, "", name)
    return re.sub(r"\s+", " ", name).strip()


def _resolve_zoho_row(zrow: pd.Series, customer_lookup: dict, invoice_files, log: list) -> dict:
    """Attempt to resolve a single Zoho row to a known customer account."""
    raw_name    = str(zrow.get("customer", "")).strip()
    email       = str(zrow.get("email", "")).strip()
    invoice_ref = str(zrow.get("invoice", "")).strip()
    row_idx     = zrow.name

    # ── Pattern 1 & 2: Name present ──────────────────────────────────────────
    if raw_name and raw_name.lower() not in ("nan", "none", ""):
        result = _fuzzy_lookup(raw_name, customer_lookup)
        if result["_match_confidence"] in ("HIGH", "MEDIUM"):
            result["_raw_name"] = raw_name
            return result
        # Low confidence but name present — still flag
        log.append({"level": "WARN",
                    "msg":   f"Row {row_idx}: '{raw_name}' fuzzy score {result['_fuzzy_score']} < threshold — REVIEW."})
        result["_raw_name"] = raw_name
        return result

    # ── Pattern 3: Invoice number present, no name ───────────────────────────
    if invoice_ref and invoice_ref.lower() not in ("nan", "none", ""):
        # Try to extract account from invoice number pattern (e.g. INV-BC000327-...)
        acct_match = re.search(r"BC\d{6}", invoice_ref, re.IGNORECASE)
        if acct_match:
            acct = acct_match.group(0).upper()
            rev_lookup = {v["account"]: v for v in customer_lookup.values()}
            if acct in rev_lookup:
                log.append({"level": "OK",
                             "msg": f"Row {row_idx}: resolved via invoice ref '{invoice_ref}' → {acct}."})
                return {
                    "_account":           acct,
                    "_account_name":      rev_lookup[acct]["account_name"],
                    "_match_method":      "INVOICE_REF",
                    "_match_confidence":  "HIGH",
                    "_needs_review":      False,
                    "_review_reason":     "",
                    "_cash_code":         "",
                    "_fuzzy_score":       100,
                    "_raw_name":          "",
                }

    # ── Pattern 4: Only email present ────────────────────────────────────────
    # We cannot resolve email → name without an external source; flag for review
    reason = "No customer name"
    if email and email.lower() not in ("nan", "none", ""):
        reason = f"Only email available ({email}) — upload invoice to resolve"

    log.append({"level": "WARN", "msg": f"Row {row_idx}: {reason} — flagged for manual review."})
    return {
        "_account":          "",
        "_account_name":     raw_name or email,
        "_match_method":     "UNRESOLVED",
        "_match_confidence": "LOW",
        "_needs_review":     True,
        "_review_reason":    reason,
        "_cash_code":        "",
        "_fuzzy_score":      0,
        "_raw_name":         raw_name,
    }


def _fuzzy_lookup(raw_name: str, customer_lookup: dict) -> dict:
    """Fuzzy-match raw_name against customer master; return resolution dict."""
    norm = _normalise_name(raw_name)
    keys = list(customer_lookup.keys())

    if not keys:
        return _unresolved(raw_name, "Customer master is empty")

    # token_set_ratio handles re-ordered words and DBA names well
    best_key, score = process.extractOne(norm, keys, scorer=fuzz.token_set_ratio)
    matched = customer_lookup[best_key]

    if score >= _FUZZY_THRESHOLD:
        confidence = "HIGH" if score >= 95 else "MEDIUM"
        return {
            "_account":          matched["account"],
            "_account_name":     matched["account_name"],
            "_match_method":     f"FUZZY({score})",
            "_match_confidence": confidence,
            "_needs_review":     confidence == "MEDIUM",
            "_review_reason":    f"Fuzzy match score {score}/100 — verify name" if confidence == "MEDIUM" else "",
            "_cash_code":        "",
            "_fuzzy_score":      score,
            "_raw_name":         raw_name,
        }
    else:
        return _unresolved(raw_name, f"Best fuzzy match '{matched['account_name']}' scored only {score}/100")


def _unresolved(raw_name: str, reason: str) -> dict:
    return {
        "_account":          "",
        "_account_name":     raw_name,
        "_match_method":     "UNRESOLVED",
        "_match_confidence": "LOW",
        "_needs_review":     True,
        "_review_reason":    reason,
        "_cash_code":        "",
        "_fuzzy_score":      0,
        "_raw_name":         raw_name,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Mathematical balance validation
# ─────────────────────────────────────────────────────────────────────────────

def _validate_balance(boa_df: pd.DataFrame, zoho_df: pd.DataFrame, log: list):
    """
    Per the Automation Rules §3.3 Balance Invariant:
      sum(Zoho Gross) - sum(Zoho Fees) == BOA Net

    Groups both sides by date and checks within tolerance.
    Tolerance: max(0.5% of amount, $1.00)
    """
    if "date" not in boa_df.columns or "date" not in zoho_df.columns:
        return
    if "_boa_amount" not in boa_df.columns:
        return

    boa_by_date  = boa_df.groupby(boa_df["date"].dt.normalize())["_boa_amount"].sum()
    zoho_net_by_date = zoho_df.groupby(zoho_df["date"].dt.normalize()).apply(
        lambda g: g["gross_amount"].fillna(0).sum() - g["fee"].fillna(0).sum()
    )

    for date, boa_net in boa_by_date.items():
        if date not in zoho_net_by_date.index:
            continue
        zoho_net = zoho_net_by_date[date]
        diff     = abs(float(boa_net) - float(zoho_net))
        tol      = max(abs(float(boa_net)) * 0.005, 1.0)
        if diff > tol:
            log.append({
                "level": "WARN",
                "msg":   (f"Balance mismatch on {date.date()}: "
                          f"BOA net={boa_net:,.2f}, Zoho net={zoho_net:,.2f}, "
                          f"diff={diff:,.2f} (tolerance ${tol:.2f})"),
            })
        else:
            log.append({
                "level": "OK",
                "msg":   (f"Balance check PASSED for {date.date()}: "
                          f"BOA net={boa_net:,.2f} ≈ Zoho net={zoho_net:,.2f}"),
            })
