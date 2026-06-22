[README.md](https://github.com/user-attachments/files/29210005/README.md)
# D365 Journal Builder

## What changed
- The D365 template is bundled with the app, so you do **not** need to upload it every day.
- BOA and Zoho files remain the daily inputs.
- Zoho record upload now accepts **PDF** in addition to XLSX and CSV.

## Files that must live in the repo
- `app.py`
- `requirements.txt`
- `D365_General_Journal_Template.xlsx`

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes
- Template upload is optional and only used as an override.
- PDF support is best-effort. If the Zoho PDF is text-based with tables, extraction usually works well. If the PDF is image-only or heavily formatted, CSV/XLSX is more reliable.
