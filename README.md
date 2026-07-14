# LCC: PDF to Master Excel

A Streamlit tool that reconstructs tabular data out of position-only PDF reports and merges
multiple files into one type-correct Excel workbook, replacing a manual, error-prone,
several-times-a-month process with a single click.

**Live App:** [lcc-pdf2excel.streamlit.app](https://lcc-pdf2excel.streamlit.app/)

---

## Impact

| Before | After |
|---|---|
| Each Branch Manager downloaded one LCC PDF per Relationship Executive (5 to 10 REs), and manually copy-pasted every column into a master sheet, repeated every month | Upload all REs' PDFs at once, click convert, download one consolidated Master LCC Excel |
| Most online PDF→Excel converters are blocked on office machines, forcing staff to use personal phones to convert files | Fully self-hosted, works entirely within the browser, no external conversion service |
| Manual copy-paste across 10+ files × 91 columns is inherently error-prone (wrong row, wrong column, skipped cell) | Deterministic, positional extraction: same PDF always produces the same output, columns never drift |
| Every time the source system's report template changed (new fields added), someone had to notice and update the manual process | New trailing fields are detected and carried through automatically, with **zero code changes required** |
| No consolidated multi-RE view without manual reconciliation | One master sheet, correctly typed (real Excel numbers and dates, not text) and ready for pivot tables / formulas immediately |

This isn't a wrapper around an existing PDF-table library: none of the standard ones (Camelot,
Tabula, pdfplumber's table mode) handle this report's actual failure modes correctly out of
the box (dynamic column widths, columns that are empty for an entire file, header text that
doesn't align with data below it). The extraction logic here was built and iterated against
real production bugs, documented and fixed one at a time.

---

## The Problem

The internal system this integrates with only exports the **LCC (Live Contract Collection)**
report as a PDF, one file per Relationship Executive (RE). A Branch Manager needs all REs'
data consolidated into a single Master LCC Excel for review and follow-up, every month.

PDFs don't store tables; they store text painted at (x, y) coordinates on a page. The report
itself compounds the difficulty:

- **Column widths are dynamic.** A long customer name pushes every value on that row right;
  there are no fixed pixel boundaries to hardcode.
- **Some columns are always empty** for a given file, which makes them invisible to a naive
  extractor and shifts every column after them by one position.
- **Multi-page PDFs** repeat the header on every page and reset Y-coordinates to zero per
  page, so naive row-detection collides data from different pages.
- **The schema itself drifts.** Report templates have added new trailing fields over time
  (`NPA_Date`, `Paymethod`, and others) without warning: a hardcoded column count silently
  drops the new data.

## Technical Approach

Rather than a traditional PDF-table extractor, the app reads each page through two different
PyMuPDF text-extraction modes and reconciles them:

- **Paragraph-level HTML rendering** (`get_text("html")`) is used to *learn* column
  boundaries: PyMuPDF's own clustering groups nearby glyphs into blocks, giving a free
  first-pass estimate of where each column starts.
- **Word-level extraction** (`get_text("words")`) is used to *read* actual cell values:
  whitespace-delimited, so two adjacent columns can never fuse into one string the way
  paragraph mode can when a long name runs into the next column.

Column boundaries are then **learned from real data rows, not the header** (header labels are
often narrower/centered relative to the data below them), with header positions injected only
as a fallback for columns that have zero data anywhere in the file; otherwise those columns
would vanish and shift everything after them.

The 91-column schema is treated as a *known prefix*, not a hard cap: any column detected
beyond it is a genuinely new field, named directly from that PDF's own header text and
appended after the fixed columns. This means new report fields are absorbed automatically
instead of being silently dropped: the actual failure mode of the original implementation,
fixed once and permanently closed off by removing the hardcoded column-count cap.

When merging multiple files in one batch, the app computes the union of every file's columns
before concatenating, so a batch that mixes older- and newer-template PDFs still produces one
consistent, correctly-ordered master sheet; files missing a given column just get blank cells
for it rather than being dropped or misaligned.

## Output

| What | Detail |
|---|---|
| Columns | A fixed 91-column core schema, in guaranteed order, plus any new trailing fields the source PDF introduces |
| Numeric columns | Auto-detected and stored as real numbers (usable in Excel formulas/pivots), not text |
| Date columns | Parsed per-format (mixed `DD/MM/YYYY` and ISO `YYYY-MM-DD` sources) and written as real Excel dates, formatted `DD-MM-YYYY` |
| Total row | Automatically excluded |
| Multiple PDFs, mixed schemas | All rows merged into one sheet; missing columns per file are blank-filled, never dropped |

## Screenshots

**Landing page**

![Landing page](screenshots/before_uploading.png)

**PDFs selected, ready to convert**

![Files selected](screenshots/after_uploading.png)

**After extraction: 7 PDFs, 832 rows, 91 core columns**

![After extraction](screenshots/after_extraction.png)

## Tech Stack

| | |
|---|---|
| **PyMuPDF (`fitz`)** | Dual-mode PDF text extraction: HTML rendering for structure discovery, word-level for fusion-proof content reading |
| **pandas** | DataFrame assembly, cross-file schema reconciliation, and type coercion |
| **openpyxl** | Excel generation with per-column custom number formatting |
| **lxml** | Parsing PyMuPDF's HTML output to recover precise per-element (x, y) positions |
| **Streamlit** | Web UI, deployed on Streamlit Cloud |

## Engineering Notes

- No third-party PDF-table library is used: the extraction algorithm is purpose-built for
  this report's specific layout quirks, because general-purpose table extractors (Camelot,
  Tabula) assume ruled/bordered tables or fixed column grids that this report doesn't have.
- Every non-obvious constant in the code (tolerance values, row-length cutoffs, clustering
  thresholds) is derived from the actual observed range of real values, not picked
  arbitrarily.
- The schema-drift handling (auto-detecting and naming new columns) was added specifically
  because a hardcoded column cap had already caused a silent data-loss bug in production:
  the fix targets the class of bug, not just that one instance.
