# IBNY D365 Journal Entry Automation

Streamlit app that automates the InBody New York cash closing workflow:
- Parses Bank of America (BOA) and Zoho Payments exports
- Resolves customer names to canonical accounts using the master file (555+ accounts)
- Assigns cash codes (AR001–AR012, OSF005) per the SOP
- Generates a D365-ready Excel upload file with two rows per transaction (credit + debit)
- Flags ambiguous rows for manual review

---

## Project Structure

```
ibny_app/
├── app.py                              # Streamlit UI
├── parsers.py                          # BOA, Zoho, reference file parsers
├── matcher.py                          # Transaction matching & customer resolution
├── builder.py                          # D365 journal entry row builder
├── exporter.py                         # Excel export with formatting
├── Cash_Code_Masterlist.xlsx           # Reference — auto-loaded
├── Account Masterlist.xlsx # Reference — auto-loaded
├── Posted_Journal_in_D365_Sample_Reference.xlsx  # Format reference
└── README.md
```

---

## Setup

### Requirements
- Python 3.10+
- pip

### Install dependencies

```bash
pip install streamlit pandas openpyxl fuzzywuzzy python-levenshtein pdfplumber xlsxwriter
```

### Run the app

```bash
cd ibny_app
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`.

---

## How to Use

### Step 1 — Upload source files
| File | Format | What it provides |
|------|--------|-----------------|
| Bank of America export | `.xlsx` or `.csv` | Posting date, description (with ZOHO tag), net amount |
| Zoho Payments export | `.xlsx`, `.csv`, or `.pdf` | Customer name, gross amount, merchant fee |
| Customer invoices *(optional)* | `.pdf` | Customer name when absent from Zoho |

### Step 2 — Select BOA account
Choose the last 4 digits of the source BOA account. This determines the Offset Account:
- `3371` → B1000002
- `3924` → B1000003
- `3384` → B1000001

### Step 3 — Click Generate
The app will:
1. Parse and normalise both files
2. Match Zoho rows to customer master accounts (fuzzy name matching)
3. Infer cash codes from payment term signals
4. Build credit + debit journal entry rows
5. Validate that BOA net ≈ Zoho gross − fees (balance invariant check)

### Step 4 — Review flagged rows
Any row the app cannot confidently resolve will be:
- Highlighted in **amber** in the preview table
- Listed in the "Review Items" sheet in the Excel export
- Explained in the "Review Reason" column

Common flags:
- Customer name not found in master (fuzzy score too low)
- Only email in Zoho — upload the invoice to resolve
- Cash code ambiguous (no payment term signal)
- BOA ↔ Zoho balance mismatch

### Step 5 — Download Excel
The export contains three sheets:
1. **D365 Upload** — all rows, ready to import
2. **Review Items** — flagged rows only, with reason column
3. **Legend** — color guide and field rules

---

## D365 Column Rules (from Automation Rules document)

### Credit Line (Customer Payment)
| Field | Rule |
|-------|------|
| Date | BOA posting date |
| Account name | Canonical name from customer master |
| Company | bwa |
| Account type | Customer |
| Account | BC###### from customer master |
| Posting profile | AutoPost |
| Cash code | AR001 (due on receipt) or AR002 (MPP) etc. |
| Description | `{prefix}{cash_code}: {BC######} {Name}_{BOA description}` |
| Credit | Gross amount from Zoho (NOT BOA net) |
| Sales tax group | AVATAX |
| Reversing entry | No |

### Debit Line (Merchant Fee)
| Field | Rule |
|-------|------|
| Account name | Outside Service (Finance) |
| Account type | Ledger |
| Account | 43170111-U26C05001-B735350-UOA003 |
| Cash code | OSF005 |
| Description | `Zoho Merchant Fee {credit description}` |
| Debit | Processing fee from Zoho |

### Batch / Multi-Payment Rule
When multiple Zoho payments land on the same BOA date:
- Each customer gets its **own credit row**
- All merchant fees are **grouped into a single debit row**
- The debit description lists all account numbers and names

---

## Cash Code Reference
| Code | Name | Condition |
|------|------|-----------|
| AR001 | AR Collection_AP | Default / Due on receipt |
| AR002 | AR Collection_MPP | Monthly Payment Plan — prefix description with "MPP " |
| AR003 | AR Collection_Financing | Payment term = Financing |
| AR004 | AR Collection_Leasing | Payment term = Leasing |
| AR005–AR011 | Net 1/10/25/30/40/45/60 Days | Per invoice payment term |
| AR012 | AR Collection_Other | Fallback for unclassified terms |
| OSF005 | Zoho Payment_Merchant fee | Always used on debit (fee) rows |

---

## Input File Format Notes

### BOA Excel/CSV
The app auto-detects:
- `Date` / `Posting Date` / `Transaction Date`
- `Description` / `Payee` / `Details` / `Memo`
- `Amount` / `Net Amount` / `Credit Amount`

Rows are filtered: only those with "ZOHO" in the description are processed as Zoho payments.

### Zoho Payments Export
The app auto-detects:
- `Customer Name` / `Customer` / `Name` / `Payer`
- `Gross Amount` / `Amount` / `Total Amount`
- `Processing Fee` / `Merchant Fee` / `Fee`
- `Date` / `Payment Date`
- `Invoice` / `Invoice Number`
- `Email`

---

## Known Limitations / Future Work
- Cash code for AR003–AR012 requires invoice lookup; currently inferred from text signals only
- PDF invoice parsing extracts names from "Bill To" field but requires consistent formatting
- No direct D365 API push yet; export file must be manually imported
- `Form_Master_DB` / `Sales_PRF` lookup (referenced in automation rules §4) not yet integrated
