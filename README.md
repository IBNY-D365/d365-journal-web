# D365 General Journal Builder

## What it does

This Streamlit app reads:

- a Bank of America export
- a Zoho payment export
- your D365 template workbook

It then generates a D365-ready journal workbook using the Zoho payment rules and the template column layout.

## Files

- `app.py` - Streamlit app
- `requirements.txt` - Python dependencies
- `README.md` - setup notes

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes

- Use your D365 template workbook as the upload template.
- If the BOA and Zoho files share a common key, select the join columns in the app.
- Monthly payment terms get the `MPP` prefix.
- Credit lines use the Zoho gross amount; debit lines use the merchant fee.
