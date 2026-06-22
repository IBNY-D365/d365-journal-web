import io
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

TEMPLATE_HEADERS = [
    "Date", "Voucher", "Account name", "Company", "Account type", "Account",
    "Posting Profile", "Cash code", "Description", "Debit", "Credit",
    "Item sales tax group", "Sales tax code", "Offset company",
    "Bank Account Type", "Offset account", "Offset transaction text",
    "Currency", "Exchange rate", "Item sales tax group2", "Sales tax group",
    "Withholding tax group", "Release date", "Reversing entry", "Reversing date",
]

AR_CODE_MAP = {
    "DUE": ("AR001", "AR Collection_AP"),
    "DUE ON RECEIPT": ("AR001", "AR Collection_AP"),
    "MONTHLY": ("AR002", "AR Collection_MPP"),
    "MPP": ("AR002", "AR Collection_MPP"),
    "FINANCING": ("AR003", "AR Collection_Financing"),
    "LEASING": ("AR004", "AR Collection_Leasing"),
    "NET 1 DAY": ("AR005", "AR Collection_Net_1Day"),
    "NET 10 DAYS": ("AR006", "AR Collection_Net_10Days"),
    "NET 25 DAYS": ("AR007", "AR Collection_Net_25Days"),
    "NET 30 DAYS": ("AR008", "AR Collection_Net_30Days"),
    "NET 40 DAYS": ("AR009", "AR Collection_Net_40Days"),
    "NET 45 DAYS": ("AR010", "AR Collection_Net_45Days"),
    "NET 60 DAYS": ("AR011", "AR Collection_Net_60Days"),
}

OFFSET_ACCOUNT_MAP = {
    "3371": "B1000002",
    "3924": "B1000003",
    "3384": "B1000001",
}

OUTPUT_TEMPLATE_VALUES = {
    "Company": "bwa",
    "Posting Profile": "AutoPost",
    "Offset company": "bwa",
    "Bank Account Type": "Bank",
    "Currency": "USD",
    "Exchange rate": 1.00,
    "Sales tax group": "AVATAX",
    "Reversing entry": "No",
    "Account type_credit": "Customer",
    "Account type_debit": "Ledger",
    "Account name_debit": "Outside Service (Finance)",
    "Account_debit": "43170111-U26C05001-B735350-UOA003",
    "Cash code_debit": "OSF005",
}

TEXT_NORMALIZATION = re.compile(r"[^A-Z0-9 ]+")
WHITESPACE = re.compile(r"\s+")


def normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).upper().strip()
    text = TEXT_NORMALIZATION.sub(" ", text)
    text = WHITESPACE.sub(" ", text)
    return text.strip()


def guess_column(columns: List[str], aliases: List[str]) -> Optional[str]:
    normalized = {normalize_text(c): c for c in columns}
    for alias in aliases:
        alias_norm = normalize_text(alias)
        if alias_norm in normalized:
            return normalized[alias_norm]
    for col in columns:
        col_norm = normalize_text(col)
        if any(alias_norm in col_norm for alias_norm in map(normalize_text, aliases)):
            return col
    return None


def read_uploaded_table(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(file)
    raise ValueError(f"Unsupported file type: {file.name}")


def resolve_cash_code(payment_term: str) -> Tuple[str, str]:
    term = normalize_text(payment_term)
    for key, value in AR_CODE_MAP.items():
        if key in term:
            return value
    return "AR012", "AR Collection_Other"


def resolve_offset_account(source_account: object) -> str:
    if source_account is None:
        return ""
    digits = re.sub(r"\D", "", str(source_account))
    for key, value in OFFSET_ACCOUNT_MAP.items():
        if key in digits:
            return value
    return ""


def to_number(value: object) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    text = str(value).replace(",", "").strip()
    if not text:
        return 0.0
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    return float(text)


def to_date(value: object):
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return None
    return dt.date()


def build_credit_description(account: str, customer_name: str, boa_desc: str, monthly: bool) -> str:
    base = f"{account} {customer_name} {boa_desc}".strip()
    if monthly:
        return f"MPP {base}".strip()
    return base


def build_debit_description(account_lines: List[Tuple[str, str]], boa_desc: str) -> str:
    if not account_lines:
        return f"Zoho Merchant Fee_{boa_desc}".strip()
    unique = []
    for acct, name in account_lines:
        token = f"{acct} {name}".strip()
        if token not in unique:
            unique.append(token)
    prefix = "Zoho Merchant Fee " + ", ".join(unique)
    return f"{prefix}_{boa_desc}".strip()


def safe_str(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def pick_field(row: pd.Series, column_name: str) -> object:
    if column_name in row.index:
        return row.get(column_name)
    for suffix in ("_boa", "_zoho"):
        candidate = f"{column_name}{suffix}"
        if candidate in row.index:
            return row.get(candidate)
    return None


@dataclass
class PaymentRow:
    boa_date: object
    boa_description: str
    source_account: str
    customer_name: str
    account_number: str
    payment_term: str
    gross_amount: float
    merchant_fee: float
    batch_id: str
    monthly: bool


def build_payment_rows(merged_df: pd.DataFrame, mapping: Dict[str, str]) -> List[PaymentRow]:
    rows: List[PaymentRow] = []
    for _, raw in merged_df.iterrows():
        customer_name = safe_str(pick_field(raw, mapping["customer_name"]))
        account_number = safe_str(pick_field(raw, mapping["account_number"]))
        boa_description = safe_str(pick_field(raw, mapping["description"]))
        boa_date = to_date(pick_field(raw, mapping["boa_date"]))
        source_account = safe_str(pick_field(raw, mapping["source_account"]))
        payment_term = safe_str(pick_field(raw, mapping["payment_term"]))
        gross_amount = to_number(pick_field(raw, mapping["gross_amount"]))
        merchant_fee = to_number(pick_field(raw, mapping["merchant_fee"]))
        batch_id = safe_str(pick_field(raw, mapping.get("batch_id", ""))) if mapping.get("batch_id") else ""
        monthly = "MONTH" in normalize_text(payment_term) or "MPP" in normalize_text(payment_term)
        rows.append(
            PaymentRow(
                boa_date=boa_date,
                boa_description=boa_description,
                source_account=source_account,
                customer_name=customer_name,
                account_number=account_number,
                payment_term=payment_term,
                gross_amount=gross_amount,
                merchant_fee=merchant_fee,
                batch_id=batch_id,
                monthly=monthly,
            )
        )
    return rows


def create_output_rows(payments: List[PaymentRow]) -> List[List[object]]:
    output_rows: List[List[object]] = []
    grouped: Dict[str, List[PaymentRow]] = {}

    for payment in payments:
        key = payment.batch_id or f"__SINGLE__{id(payment)}"
        grouped.setdefault(key, []).append(payment)

    for batch_key, items in grouped.items():
        fee_total = round(sum(p.merchant_fee for p in items), 2)
        credit_items = []
        for p in items:
            credit_items.append((p.account_number, p.customer_name))
            cash_code, _cash_name = resolve_cash_code(p.payment_term)
            offset_account = resolve_offset_account(p.source_account)
            credit_desc = build_credit_description(
                account=p.account_number,
                customer_name=p.customer_name,
                boa_desc=p.boa_description,
                monthly=(cash_code == "AR002" or p.monthly),
            )
            credit_row = [
                p.boa_date,
                None,
                p.customer_name,
                OUTPUT_TEMPLATE_VALUES["Company"],
                OUTPUT_TEMPLATE_VALUES["Account type_credit"],
                p.account_number,
                OUTPUT_TEMPLATE_VALUES["Posting Profile"],
                cash_code,
                credit_desc,
                None,
                round(p.gross_amount, 2),
                None,
                None,
                OUTPUT_TEMPLATE_VALUES["Offset company"],
                OUTPUT_TEMPLATE_VALUES["Bank Account Type"],
                offset_account,
                None,
                OUTPUT_TEMPLATE_VALUES["Currency"],
                OUTPUT_TEMPLATE_VALUES["Exchange rate"],
                None,
                OUTPUT_TEMPLATE_VALUES["Sales tax group"],
                None,
                None,
                OUTPUT_TEMPLATE_VALUES["Reversing entry"],
                None,
            ]
            output_rows.append(credit_row)

        debit_desc = build_debit_description(credit_items, items[0].boa_description)
        debit_row = [
            items[0].boa_date,
            None,
            OUTPUT_TEMPLATE_VALUES["Account name_debit"],
            OUTPUT_TEMPLATE_VALUES["Company"],
            OUTPUT_TEMPLATE_VALUES["Account type_debit"],
            OUTPUT_TEMPLATE_VALUES["Account_debit"],
            None,
            OUTPUT_TEMPLATE_VALUES["Cash code_debit"],
            debit_desc,
            round(fee_total, 2),
            None,
            None,
            None,
            OUTPUT_TEMPLATE_VALUES["Offset company"],
            OUTPUT_TEMPLATE_VALUES["Bank Account Type"],
            resolve_offset_account(items[0].source_account),
            None,
            OUTPUT_TEMPLATE_VALUES["Currency"],
            OUTPUT_TEMPLATE_VALUES["Exchange rate"],
            None,
            OUTPUT_TEMPLATE_VALUES["Sales tax group"],
            None,
            None,
            OUTPUT_TEMPLATE_VALUES["Reversing entry"],
            None,
        ]
        output_rows.append(debit_row)

    return output_rows


def write_to_template(template_bytes: bytes, output_rows: List[List[object]]) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Preserve the template header and clear old data below it.
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for idx, row in enumerate(output_rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=idx, column=col_idx, value=value)

    # Basic polish.
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.sheet_view.zoomScale = 90

    widths = {
        1: 12, 3: 28, 4: 10, 5: 14, 6: 18, 7: 14, 8: 12, 9: 90, 10: 14, 11: 14,
        14: 12, 15: 14, 16: 22, 18: 10, 19: 12, 21: 14, 24: 14,
    }
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    date_cols = [1, 24, 25]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for col_idx in date_cols:
            cell = row[col_idx - 1]
            if col_idx == 1 and cell.value is not None:
                cell.number_format = "yyyy-mm-dd"
            if col_idx in (24, 25) and cell.value is not None:
                cell.number_format = "@"
        for col_idx in (10, 11, 19):
            cell = row[col_idx - 1]
            if cell.value is not None and col_idx in (10, 11):
                cell.number_format = '0.00'
            elif cell.value is not None and col_idx == 19:
                cell.number_format = '0.00'

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def detect_mapping(df: pd.DataFrame) -> Dict[str, str]:
    cols = list(df.columns)
    mapping = {
        "boa_date": guess_column(cols, ["boa date", "posting date", "date", "transaction date"]),
        "description": guess_column(cols, ["description", "transaction description", "reference", "memo", "details"]),
        "source_account": guess_column(cols, ["source account", "account number", "account", "source acct", "bank account"]),
        "customer_name": guess_column(cols, ["customer name", "account name", "bill to", "business name", "client", "merchant"]),
        "account_number": guess_column(cols, ["account #", "account number", "customer account", "d365 account", "account"]),
        "payment_term": guess_column(cols, ["payment term", "terms", "invoice sent", "term"]),
        "gross_amount": guess_column(cols, ["gross amount", "payment amount", "amount", "gross", "credit amount"]),
        "merchant_fee": guess_column(cols, ["merchant fee", "processing fee", "fee", "charge", "debit amount"]),
        "batch_id": guess_column(cols, ["batch id", "batch", "payment id", "transaction id", "group id"]),
    }
    return {k: v for k, v in mapping.items() if v is not None}


def merge_boa_and_zoho(
    boa_df: pd.DataFrame,
    zoho_df: pd.DataFrame,
    boa_join_col: Optional[str],
    zoho_join_col: Optional[str],
) -> pd.DataFrame:
    if boa_join_col and zoho_join_col:
        boa_tmp = boa_df.copy()
        zoho_tmp = zoho_df.copy()
        boa_tmp["__JOIN_KEY__"] = boa_tmp[boa_join_col].astype(str).map(normalize_text)
        zoho_tmp["__JOIN_KEY__"] = zoho_tmp[zoho_join_col].astype(str).map(normalize_text)
        merged = zoho_tmp.merge(
            boa_tmp,
            on="__JOIN_KEY__",
            how="left",
            suffixes=("_zoho", "_boa"),
        )
        return merged

    # Best-effort fallback: repeat the first BOA row across Zoho rows.
    if len(boa_df) == 0:
        return zoho_df.copy()
    repeat = pd.concat([boa_df.iloc[[0]]] * max(len(zoho_df), 1), ignore_index=True)
    repeat = repeat.iloc[: len(zoho_df)].reset_index(drop=True)
    merged = pd.concat([zoho_df.reset_index(drop=True), repeat.add_suffix("_boa")], axis=1)
    return merged


def require_mapping(mapping: Dict[str, str], key: str) -> bool:
    return key in mapping and mapping[key]


def main():
    import streamlit as st
    st.set_page_config(page_title="D365 Journal Builder", layout="wide")
    st.title("D365 General Journal Builder for Zoho Payments")
    st.caption("Upload BOA and Zoho extracts, then download a D365-ready Excel journal file.")

    with st.sidebar:
        st.header("Files")
        template_file = st.file_uploader("D365 template (.xlsx)", type=["xlsx"], help="Use your D365_General Journal_Template.xlsx")
        boa_file = st.file_uploader("BOA report (.xlsx or .csv)", type=["xlsx", "csv"])
        zoho_file = st.file_uploader("Zoho record (.xlsx or .csv)", type=["xlsx", "csv"])

    if template_file is None:
        st.info("Upload your D365 template to begin.")
        return
    if boa_file is None or zoho_file is None:
        st.info("Upload both the BOA report and the Zoho record.")
        return

    boa_df = read_uploaded_table(boa_file)
    zoho_df = read_uploaded_table(zoho_file)

    st.subheader("Detected columns")
    col1, col2 = st.columns(2)
    with col1:
        st.write("BOA columns:")
        st.code("\n".join(map(str, boa_df.columns)))
    with col2:
        st.write("Zoho columns:")
        st.code("\n".join(map(str, zoho_df.columns)))

    boa_detected = detect_mapping(boa_df)
    zoho_detected = detect_mapping(zoho_df)

    st.subheader("Column mapping")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        boa_date_col = st.selectbox("BOA date column", boa_df.columns, index=list(boa_df.columns).index(boa_detected.get("boa_date", boa_df.columns[0])) if boa_detected.get("boa_date") in boa_df.columns else 0)
        boa_desc_col = st.selectbox("BOA description column", boa_df.columns, index=list(boa_df.columns).index(boa_detected.get("description", boa_df.columns[0])) if boa_detected.get("description") in boa_df.columns else 0)
        boa_source_account_col = st.selectbox("BOA source account column", boa_df.columns, index=list(boa_df.columns).index(boa_detected.get("source_account", boa_df.columns[0])) if boa_detected.get("source_account") in boa_df.columns else 0)
        boa_join_col = st.selectbox("BOA join key column (optional)", ["<none>"] + list(boa_df.columns), index=0)
    with m2:
        zoho_customer_col = st.selectbox("Zoho customer name column", zoho_df.columns, index=list(zoho_df.columns).index(zoho_detected.get("customer_name", zoho_df.columns[0])) if zoho_detected.get("customer_name") in zoho_df.columns else 0)
        zoho_account_col = st.selectbox("Zoho account number column", zoho_df.columns, index=list(zoho_df.columns).index(zoho_detected.get("account_number", zoho_df.columns[0])) if zoho_detected.get("account_number") in zoho_df.columns else 0)
        zoho_term_col = st.selectbox("Zoho payment term column", zoho_df.columns, index=list(zoho_df.columns).index(zoho_detected.get("payment_term", zoho_df.columns[0])) if zoho_detected.get("payment_term") in zoho_df.columns else 0)
        zoho_join_col = st.selectbox("Zoho join key column (optional)", ["<none>"] + list(zoho_df.columns), index=0)
    with m3:
        zoho_gross_col = st.selectbox("Zoho gross amount column", zoho_df.columns, index=list(zoho_df.columns).index(zoho_detected.get("gross_amount", zoho_df.columns[0])) if zoho_detected.get("gross_amount") in zoho_df.columns else 0)
        zoho_fee_col = st.selectbox("Zoho merchant fee column", zoho_df.columns, index=list(zoho_df.columns).index(zoho_detected.get("merchant_fee", zoho_df.columns[0])) if zoho_detected.get("merchant_fee") in zoho_df.columns else 0)
        zoho_batch_col = st.selectbox("Zoho batch/group column (optional)", ["<none>"] + list(zoho_df.columns), index=0)

    mapping = {
        "boa_date": boa_date_col,
        "description": boa_desc_col,
        "source_account": boa_source_account_col,
        "customer_name": zoho_customer_col,
        "account_number": zoho_account_col,
        "payment_term": zoho_term_col,
        "gross_amount": zoho_gross_col,
        "merchant_fee": zoho_fee_col,
    }
    if zoho_batch_col != "<none>":
        mapping["batch_id"] = zoho_batch_col

    st.subheader("Filtered preview")
    boa_zoho_mask = boa_df[boa_desc_col].astype(str).str.contains("ZOHO", case=False, na=False)
    st.dataframe(boa_df.loc[boa_zoho_mask].head(20), use_container_width=True)

    if st.button("Build D365 Excel", type="primary"):
        template_bytes = template_file.getvalue()

        if boa_join_col != "<none>" and zoho_join_col != "<none>":
            merged_df = merge_boa_and_zoho(boa_df, zoho_df, boa_join_col, zoho_join_col)
        else:
            merged_df = merge_boa_and_zoho(boa_df, zoho_df, None, None)

        # Use Zoho as the detail source and BOA as the journal context source.
        # The merge keeps the selected BOA fields available for the output rows.
        merged_for_rows = merged_df.copy()
        if boa_join_col == "<none>" or zoho_join_col == "<none>":
            # Fallback path prefixes BOA columns with _boa; normalize the selected names back into place.
            for selected in [boa_date_col, boa_desc_col, boa_source_account_col]:
                candidate = f"{selected}_boa"
                if candidate in merged_for_rows.columns and selected not in merged_for_rows.columns:
                    merged_for_rows[selected] = merged_for_rows[candidate]

        payments = build_payment_rows(merged_for_rows, mapping)
        if not payments:
            st.error("No payment rows could be created from the merged BOA/Zoho data.")
            return

        output_rows = create_output_rows(payments)
        excel_bytes = write_to_template(template_bytes, output_rows)

        st.success(f"Built {len(output_rows)} D365 rows from {len(payments)} Zoho payment rows.")
        st.download_button(
            label="Download D365 journal workbook",
            data=excel_bytes,
            file_name="D365_General_Journal_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        preview_df = pd.DataFrame(output_rows, columns=TEMPLATE_HEADERS)
        st.subheader("Output preview")
        st.dataframe(preview_df, use_container_width=True)

        st.info(
            "This app mirrors the source rules: the credit line uses the Zoho gross amount, the fee is posted as a separate debit line, and monthly terms get the MPP prefix."
        )


if __name__ == "__main__":
    main()
