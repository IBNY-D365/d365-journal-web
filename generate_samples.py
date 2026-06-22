"""
generate_samples.py — Create sample BOA and Zoho test files matching real formats.

Run: python generate_samples.py
Output: sample_boa.xlsx, sample_zoho.xlsx in the current directory.
"""

import pandas as pd
from datetime import date

# ── Sample BOA (mimics Bank of America Excel export) ─────────────────────────
boa_rows = [
    {
        "Date":        "06/10/2025",
        "Description": "ZOHO PAYMENTS DES:ZOHO PAYME ID:ST-T3C7P5P8R7B1 INDN:LIZA CO ID:800948598 CCD",
        "Amount":      "37435.15",  # Net (gross - fees)
        "Balance":     "112500.00",
    },
    {
        "Date":        "06/12/2025",
        "Description": "ZOHO PAYMENTS DES:ZOHO PAYME ID:ST-T3C7P5P8R7B2 INDN:LIZA CO ID:800948598 CCD",
        "Amount":      "4852.58",
        "Balance":     "117352.58",
    },
    {
        "Date":        "06/12/2025",
        "Description": "ACH CREDIT SOME OTHER VENDOR",  # non-Zoho row — should be ignored
        "Amount":      "500.00",
        "Balance":     "117852.58",
    },
]
boa_df = pd.DataFrame(boa_rows)
boa_df.to_excel("sample_boa.xlsx", index=False)
print("Created sample_boa.xlsx")

# ── Sample Zoho (mimics Zoho Payments CSV/Excel export) ───────────────────────
zoho_rows = [
    # June 10 — single customer payment
    {
        "Payment Date":    "06/10/2025",
        "Customer Name":   "LegacyMD LLC",            # maps to BC000605
        "Gross Amount":    "32582.07",
        "Processing Fee":  "945.18",
        "Net Amount":      "31636.89",
        "Invoice":         "INV-BC000605-0045",
        "Payment ID":      "ST-T3C7P5P8R7B1",
        "Description":     "ZOHO PAYMENTS DES:ZOHO PAYME ID:ST-T3C7P5P8R7B1 INDN:LIZA CO ID:800948598 CCD",
        "Status":          "Completed",
    },
    {
        "Payment Date":    "06/10/2025",
        "Customer Name":   "Vizzhy Inc",               # maps to BC000422 — MPP
        "Gross Amount":    "4998.33",
        "Processing Fee":  "145.25",
        "Net Amount":      "4853.08",
        "Invoice":         "MPP BC000422 INV-2025-06",
        "Payment ID":      "ST-T3C7P5P8R7B1",
        "Description":     "ZOHO PAYMENTS DES:ZOHO PAYME ID:ST-T3C7P5P8R7B1 INDN:LIZA CO ID:800948598 CCD",
        "Status":          "Completed",
    },
    # June 12 — single customer (slightly different name to test fuzzy)
    {
        "Payment Date":    "06/12/2025",
        "Customer Name":   "Page Fit DBA Intoxx Fitness",  # fuzzy → BC000571 Page Fit Inc. DBA Intoxx Fitness
        "Gross Amount":    "4998.33",
        "Processing Fee":  "145.75",
        "Net Amount":      "4852.58",
        "Invoice":         "INV-2025-0612",
        "Payment ID":      "ST-T3C7P5P8R7B2",
        "Description":     "ZOHO PAYMENTS DES:ZOHO PAYME ID:ST-T3C7P5P8R7B2 INDN:LIZA CO ID:800948598 CCD",
        "Status":          "Completed",
    },
]
zoho_df = pd.DataFrame(zoho_rows)
zoho_df.to_excel("sample_zoho.xlsx", index=False)
print("Created sample_zoho.xlsx")
print("\nExpected D365 output (6 rows):")
print("  Row 1: CREDIT  LegacyMD LLC       BC000605  AR001  $32,582.07")
print("  Row 2: CREDIT  Vizzhy Inc         BC000422  AR002  $4,998.33")
print("  Row 3: DEBIT   Outside Svc(Fin)             OSF005 $1,090.43  ← June 10 batch fee")
print("  Row 4: CREDIT  Page Fit Inc. DBA  BC000571  AR001  $4,998.33")
print("  Row 5: DEBIT   Outside Svc(Fin)             OSF005 $145.75    ← June 12 single fee")
