# PDF → Master Excel

A production-ready Streamlit application that consolidates multiple **ALL Summary Report** PDFs (Shriram Finance loan reports) into a single downloadable Excel file.

---

## Features

| Feature | Detail |
|---|---|
| Multi-PDF upload | Upload 5-20 PDFs at once |
| Automatic extraction | pdfplumber — no Java, no OCR |
| Smart parsing | Detects header rows, skips ZTotal row, handles multi-page |
| Data cleaning | Removes empty rows, repeated headers, normalises column names |
| Metadata columns | Source File + optional Processed At timestamp |
| Duplicate detection | Optional flag for duplicate Loan No across PDFs |
| Formatted Excel | Navy header, auto column widths, frozen row |
| Error resilience | Failed PDFs are logged; rest continue processing |
| Deployment-ready | Streamlit Cloud compatible |

---

## Project Structure

```
PDF_to_EXCEL/
├── app.py                  # Streamlit UI — entry point
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── .streamlit/
│   └── config.toml         # Theme + server settings
└── src/
    ├── __init__.py
    ├── extractor.py        # PDF table extraction logic
    ├── cleaner.py          # Data cleaning pipeline
    └── utils.py            # Excel export + result dataclasses
```

---

## Local Setup

```bash
# 1. Clone / copy the project
cd "PDF_to_EXCEL"

# 2. Create and activate virtual environment
python -m venv p2e
p2e\Scripts\activate          # Windows
# source p2e/bin/activate     # macOS / Linux

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```




---

## How It Works

```
User uploads PDFs
       │
       ▼
extractor.py  ←  pdfplumber opens each PDF
  - Detect header row (by keyword matching)
  - Skip title rows, repeated headers, ZTotal rows
  - Return raw DataFrame per PDF
       │
       ▼
cleaner.py
  - Standardise column names
  - Remove empty / repeated rows
  - Add Source File + Processed At columns
  - Optional: flag duplicate Loan Nos
       │
       ▼
pd.concat  →  master DataFrame
       │
       ▼
utils.py  →  to_excel_bytes()
  - Formatted Excel (navy header, auto-width)
       │
       ▼
Streamlit download button
```

---

## Supported PDF Format

- **Report title**: `ALL Summary Report as on DD/MM/YYYY`
- **Header keywords used for detection**: `Loan No`, `Cust Name`, `Veh ID`, `SNo`, `Loan Amount`, `CHANNEL`, `Loan Status`
- **Rows skipped**: title row, ZTotal summary row, empty rows, repeated header rows
- **Pages**: single or multi-page PDFs supported
- **Encoding**: machine-readable (digitally generated) PDFs only — not scanned images

---

## Configuration

Edit `.streamlit/config.toml` to change the theme, max upload size, etc.

```toml
[server]
maxUploadSize = 200    # MB
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `streamlit` | Web UI framework |
| `pdfplumber` | PDF table extraction |
| `pandas` | DataFrame operations |
| `openpyxl` | Excel file writing |
