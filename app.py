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

from parsers import parse_boa, parse_zoho, load_customer_master, load_cash_codes
from matcher import match_transactions
from builder import build_journal_entries
from exporter import export_to_excel

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

# ── Load reference data (auto-loaded from app directory) ─────────────────────
@st.cache_data(show_spinner=False)
def load_references():
    customers = load_customer_master("IBNY_Business_Customer_Account.xlsx")
    cash_codes = load_cash_codes("Cash_Code_Masterlist.xlsx")
    return customers, cash_codes

try:
    customer_df, cash_code_df = load_references()
    st.markdown(f'<div class="ok-box">✅ Reference files loaded — {len(customer_df)} customer accounts · {len(cash_code_df)} cash codes</div>', unsafe_allow_html=True)
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
