# GSTR-1 PDF → Excel Analyzer

## What it does
- Upload one ZIP containing filed GSTR-1 PDFs.
- Reads the Financial Year and Tax Period from each filed copy.
- Extracts table-wise summary values from the filed GSTR-1.
- Fills the supplied `Sample_format.xlsx` structure.
- Calculates **Invoice Value = Taxable Value + IGST + CGST + SGST + Cess**.
- Produces one row per month plus a Total row.

## Run locally
```bash
python -m venv .venv
.venv\\Scripts\\activate       # Windows
pip install -r requirements.txt
streamlit run app.py
```

## Important design choice
The filed GSTR-1 PDF supplied as reference is a summary return, not an invoice-level JSON. Therefore this version extracts the summary totals actually visible in the filed PDF and does not invent invoice-level fields.

## Production upgrades recommended
1. Add OCR fallback for scanned PDFs.
2. Add a confidence score and a review queue for ambiguous extraction.
3. Validate each extracted month against GSTR-1 section totals and the final liability.
4. Support multiple GSTINs and prevent duplicate tax periods.
5. Add an audit log showing PDF → section → extracted numbers → Excel cell.
6. Add password/encryption handling if confidential PDFs are protected.
