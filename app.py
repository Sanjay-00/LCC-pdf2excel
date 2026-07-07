import io
import re
from collections import defaultdict

import fitz
import pandas as pd
import streamlit as st
from lxml import html as lxml_html

from report_issue import render_report_issue

st.set_page_config(page_title="LCC PDF to Excel", page_icon="🗐", layout="wide")

st.markdown("""
<style>
.hero {
    background: linear-gradient(135deg, #1a3c5e 0%, #2e75b6 100%);
    padding: 2rem 2.5rem;
    border-radius: 14px;
    margin-bottom: 2rem;
}
.hero h1 { color: #ffffff; font-size: 2.2rem; font-weight: 700; margin: 0; }
.hero p  { color: #cce0f5; font-size: 1rem; margin: 0.4rem 0 0 0; }
.stat-box {
    background: #f0f6ff;
    border: 1px solid #c5d8f0;
    border-radius: 10px;
    padding: 1rem 1.5rem;
    text-align: center;
}
.stat-box h3 { color: #1a3c5e; font-size: 1.8rem; margin: 0; }
.stat-box p  { color: #4a6fa5; font-size: 0.8rem; margin: 0; font-weight: 600; text-transform: uppercase; }
div[data-testid="stDownloadButton"] button {
    background: #1a3c5e; color: white; border-radius: 8px;
    font-weight: 600; width: 100%; font-size: 1rem; padding: 0.6rem;
}
div[data-testid="stDownloadButton"] button:hover { background: #2e75b6; color: white; }
</style>

<div class="hero">
    <h1>📑 LCC: PDF to Master Excel</h1>
    <p>Upload ALL RE's PDFs · Get one consolidated Excel file as Master LCC</p>
</div>
""", unsafe_allow_html=True)

# ── Fixed schema — 85 columns in exact order ──────────────────────────────────
EXPECTED_COLS = [
    "SNo","Loan No","CHANNEL","BU","StateName","Zone","RegionName","Unit","Ag_Date",
    "SRC Code","SRC Name","MNT CODE","MNT NAME","Due Dt","Tenure","Loan Status",
    "Loan Amount","Veh ID","Cust Name","Guar Name","Cust Mob No","Guar Mob No",
    "Segment","SegmentName","Make","Vehicle Description","Year Of Manufacture",
    "Arrear Opening","ARREARS OPEN AGAINST INST","ARREARS OPEN AGAINST EXP",
    "ARREARS OPEN AGAINST BC","ARREARS OPEN AGAINST PC","OPENING RESERVE COLLECTION",
    "Month Due-Inst","Month Due-Exp","MONTH DUE (BC)","MONTH DUE PC",
    "Month Receipt Amount","MONTHCOLL INST","MONTHCOLL EXP","MONTHCOLL BC","MONTHCOLL PC",
    "Month Collection (Excluding Reserve Collection)","Closing Arrears",
    "UN-CLEARED CHEQUE FOR THE MONTH/Amount Not remitted by RE",
    "Cum Due-Inst","Cum Due-Exp","CUM DUE (BC)","Cum Due PC",
    "Cum Coll (Inst+Exp)","Total Cum Collection",
    "ARREARS AGAINST INST","ARREARS AGAINST EXP","ARREARS AGAINST BC","ARREARS AGAINST PC",
    "CLOSING RESERVE COLLECTION","Arrears against Inst+Exp",
    "Uncleared Cheque/Amount Not remitted by RE",
    "LCC%","Arrears / EMI","DelinquencyDays","VehEMI Accrued","ClosingPC","POS","scheme",
    "Non Starter","Strike","RCEndors(>90Days)","RTO / INSURANCE",
    "NET Collection Demand Inst+Exp","Net Collection Demand Inst+Exp+BC",
    "NET COLLECTION","NET COLLECTION EXCLUDING RESERVE COLL",
    "Last Receipt Date","Last Receipt Amount","ParentLDueDate",
    "No Coll 3 Months and >6 EMI","NACHStatus","SaleType","CoLending_Loans",
    "CUSTOMER_STATUS","LGL_FLAG","LGL_DESCRIPTION","TyreFlag","FUEL_TYPE",
]

DATE_COLS     = {"Ag_Date", "Last Receipt Date", "ParentLDueDate", "NPA_Date"}


def _parse_date_col(series: pd.Series) -> pd.Series:
    """Parse a DATE_COLS column, honoring whichever format each value is actually in.

    Ag_Date renders as DD/MM/YYYY but Last Receipt Date and ParentLDueDate render
    as ISO YYYY-MM-DD. A single dayfirst=True pd.to_datetime call mangles the ISO
    values instead: dateutil treats the middle component as the day, silently
    swapping month/day when both are <=12 and blanking the cell to NaT otherwise.
    Parsing each recognized format explicitly avoids that misinterpretation.
    """
    s = series.astype(str).str.strip().replace("", pd.NA)
    result = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    iso_mask = s.str.match(r"^\d{4}-\d{2}-\d{2}$", na=False)
    slash_mask = s.str.match(r"^\d{1,2}/\d{1,2}/\d{4}$", na=False)
    other_mask = s.notna() & ~iso_mask & ~slash_mask

    if iso_mask.any():
        result.loc[iso_mask] = pd.to_datetime(s[iso_mask], format="%Y-%m-%d", errors="coerce")
    if slash_mask.any():
        result.loc[slash_mask] = pd.to_datetime(s[slash_mask], format="%d/%m/%Y", errors="coerce")
    if other_mask.any():
        result.loc[other_mask] = pd.to_datetime(s[other_mask], dayfirst=True, errors="coerce")

    return result
HEADER_SIGNAL = {"SNo", "CHANNEL", "Tenure"}
Y_TOL         = 3.0
N_COLS        = len(EXPECTED_COLS)   # 85


def _parse_html_paragraphs(page) -> list[tuple[float, float, str]]:
    """Return (top, left, text) for every non-empty <p> on the page."""
    out = []
    tree = lxml_html.fromstring(page.get_text("html"))
    for p in tree.findall(".//p"):
        style  = p.get("style", "")
        top_m  = re.search(r"top:([\d.]+)pt",  style)
        left_m = re.search(r"left:([\d.]+)pt", style)
        if not top_m or not left_m:
            continue
        text = "".join(p.itertext()).strip()
        if text:
            out.append((float(top_m.group(1)), float(left_m.group(1)), text))
    return out


def _parse_words(page) -> list[tuple[float, float, str]]:
    """Return (top, left, text) for every word on the page.

    Word-level bboxes split on whitespace, unlike get_text("html") which can
    fuse two adjacent cells into a single <p> when their text sits close
    together (e.g. a long MNT NAME running right up against the Due Dt
    value). Using words for the actual cell data avoids that fusion; column
    positions are still learned from paragraphs, where multi-word cells
    collapsing into one cluster is desirable.
    """
    return [(w[1], w[0], w[4]) for w in page.get_text("words")]


def _group_by_row(paragraphs: list) -> dict[float, list]:
    row_map: dict[float, list] = {}
    for top, left, text in paragraphs:
        key = next((k for k in row_map if abs(k - top) <= Y_TOL), top)
        row_map.setdefault(key, []).append((left, text))
    return row_map


def _data_clusters(row_map: dict, header_y: float,
                   x_tol: float = 5.0, max_rows: int = 20) -> list[float]:
    """
    Collect column X positions from data rows by clustering left-edge values.
    Skips title rows and ZTotal row. Returns sorted cluster representatives.
    Columns that are entirely empty in the PDF won't appear here.
    """
    all_xs: list[float] = []
    rows_checked = 0

    for y in sorted(row_map):
        if abs(y - header_y) <= Y_TOL:
            continue
        items = sorted(row_map[y], key=lambda w: w[0])
        if not items:
            continue
        if items[0][1] == "0" and any("ZTotal" in t for _, t in items):
            continue
        if len(items) < 5:          # skip title / separator rows
            continue
        all_xs.extend(left for left, _ in items)
        rows_checked += 1
        if rows_checked >= max_rows:
            break

    if not all_xs:
        return []

    all_xs.sort()
    clusters: list[float] = [all_xs[0]]
    for x in all_xs[1:]:
        if x - clusters[-1] <= x_tol:
            pass
        else:
            clusters.append(x)
    return clusters


def _build_col_xs(header_items: list, row_map: dict, header_y: float) -> list[float]:
    """
    Build the list of column X positions.

    Strategy:
    1. Learn positions from data rows (accurate, handles dynamic column widths,
       but misses columns that are entirely empty in this PDF).
    2. Supplement with header X positions for any column that had no data —
       those columns will remain empty in the output, but their X slot must
       exist so subsequent columns aren't shifted by one position.
    3. Sort.

    Not capped to a fixed column count: report templates occasionally add new
    trailing fields, and truncating here would silently drop them.
    """
    header_xs = [x for x, _ in header_items]
    data_xs   = _data_clusters(row_map, header_y)

    if not data_xs:
        return sorted(header_xs)

    # Inject header positions that have no matching data cluster.
    # These represent columns that are genuinely empty in this PDF.
    combined = list(data_xs)
    for hx in header_xs:
        if not any(abs(hx - dx) <= 20 for dx in combined):
            combined.append(hx)

    return sorted(combined)


def _resolve_col_names(page, col_xs: list[float], header_y: float) -> list[str]:
    """Name each detected column.

    The first len(EXPECTED_COLS) positions are assumed to be the known fixed
    schema, in order — that order has held across every report variant seen
    so far; only new fields get appended after it when a template changes.
    Any columns beyond that are new fields: their names are read straight
    from the header row's words, grouped into column buckets the same way
    data rows are (paragraph extraction fuses adjacent header cells, so we
    re-read the header at word level instead of reusing header_items).
    """
    word_row_map = _group_by_row(_parse_words(page))
    header_words: list[tuple[float, str]] = []
    for y, items in word_row_map.items():
        if abs(y - header_y) <= Y_TOL:
            header_words = sorted(items, key=lambda x: x[0])
            break

    detected = _assign_boundary(header_words, col_xs)

    names: list[str] = []
    seen: set[str] = set()
    for i, text in enumerate(detected):
        if i < len(EXPECTED_COLS):
            name = EXPECTED_COLS[i]
        else:
            name = text or f"Extra_{i - len(EXPECTED_COLS) + 1}"
            base, suffix = name, 2
            while name in seen:
                name = f"{base}_{suffix}"
                suffix += 1
        seen.add(name)
        names.append(name)
    return names


def _assign_boundary(items: list[tuple], col_xs: list[float]) -> list[str]:
    """
    Assign each (left, text) span to a column using boundary-based logic:
      column i owns X values in  [col_xs[i], col_xs[i+1]).
    This captures right-aligned or center-aligned values that sit inside the
    column's range but don't exactly match the column's left edge.
    """
    # col_xs is learned from the HTML "left:XXXpt" text (rounded to a fixed
    # precision), while item positions come from word-level extraction whose
    # raw coordinates differ by a fraction of a point. Without slack, a word
    # landing just short of its own column's boundary falls back into the
    # previous column, and the error cascades across the whole row.
    EPS = 1.0

    # Build right-side boundaries for each column
    boundaries = [col_xs[i + 1] for i in range(len(col_xs) - 1)]
    boundaries.append(float("inf"))

    cells: dict[int, str] = {}
    for left, text in items:
        # Assign to rightmost column whose left edge <= this value's X
        idx = 0
        for i, cx in enumerate(col_xs):
            if cx - EPS <= left:
                idx = i
            else:
                break
        # Verify the value is actually before the next column boundary
        if left < boundaries[idx] + EPS:
            cells[idx] = (cells.get(idx, "") + " " + text).strip()

    return [cells.get(i, "") for i in range(len(col_xs))]


def pdf_to_dataframe(file_bytes: bytes) -> pd.DataFrame:
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    col_xs: list[float] | None = None
    col_names: list[str] | None = None
    data: list[list[str]] = []

    for page in doc:
        paragraphs = _parse_html_paragraphs(page)
        row_map    = _group_by_row(paragraphs)

        # Find header row on this page (Y coords reset to 0 on each page)
        page_header_items = None
        page_header_y: float | None = None
        for y in sorted(row_map):
            items = sorted(row_map[y], key=lambda x: x[0])
            if HEADER_SIGNAL & {t for _, t in items}:
                page_header_items = items
                page_header_y     = y
                break

        # Learn column positions once from the first page that has a header
        if page_header_items is not None and col_xs is None:
            col_xs = _build_col_xs(page_header_items, row_map, page_header_y)
            col_names = _resolve_col_names(page, col_xs, page_header_y)

        if col_xs is None:
            continue

        # Extract data rows from word-level positions, not paragraphs — the
        # HTML-paragraph view can fuse two adjacent cells into one <p> when
        # their text sits close together (e.g. a long name running into the
        # next column's value), silently dropping the boundary between them.
        word_row_map = _group_by_row(_parse_words(page))
        for y in sorted(word_row_map):
            if page_header_y is not None and abs(y - page_header_y) <= Y_TOL:
                continue
            items = sorted(word_row_map[y], key=lambda x: x[0])
            # Word-level rows have far more items than paragraph-level ones
            # (each multi-word cell splits into several entries), so a real
            # data row has 80+ words while a title/separator line has a
            # handful — 30 cleanly separates the two.
            if not items or len(items) < 30:
                continue
            if any("ZTotal" in t for _, t in items):
                continue
            row = _assign_boundary(items, col_xs)
            if any(row):
                data.append(row)

    doc.close()

    if col_xs is None or col_names is None:
        return pd.DataFrame(columns=EXPECTED_COLS)

    n = len(col_xs)
    df = pd.DataFrame(
        [r[:n] + [""] * max(0, n - len(r)) for r in data],
        columns=col_names,
    )
    extra_cols = [c for c in col_names if c not in EXPECTED_COLS]
    return df.reindex(columns=EXPECTED_COLS + extra_cols, fill_value="")


# ── Streamlit UI ──────────────────────────────────────────────────────────────
col_up, col_info = st.columns([2, 1])

with col_up:
    files = st.file_uploader(
        "Drop PDF files here or click Upload",
        type="pdf",
        accept_multiple_files=True,
        label_visibility="visible",
    )

with col_info:
    st.markdown("""
    <div style="background:#f8fafd;border:1px solid #d0e4f7;border-radius:10px;padding:0.8rem 1.2rem;margin-top:0.2rem">
    <b>📋 How it works</b><br><br>
    1. Upload one or more <b>LCC in PDF format</b><br>
    2. Click <b>Convert to Excel</b><br>
    3. Download the master Excel file<br><br>
    <span style="color:#4a6fa5;font-size:0.85rem"> 91 fixed columns &nbsp;|&nbsp;  All REs data merged</span>
    </div>
    """, unsafe_allow_html=True)

if files:
    st.info(f"**{len(files)} file(s) selected:** " + ", ".join(f.name for f in files))

    if st.button("▶  Convert to Excel", type="primary", use_container_width=True):
        dfs:    list[pd.DataFrame] = []
        failed: list[str]          = []
        bar = st.progress(0)

        for i, f in enumerate(files):
            bar.progress(i / len(files), text=f"Processing {f.name}…")
            try:
                df = pdf_to_dataframe(f.read())
                if df.empty or df.shape[1] == 0:
                    failed.append(f.name)
                    st.warning(f"⚠ No table found in {f.name}")
                else:
                    dfs.append(df)
                    st.success(f"✔ {f.name} -  {len(df):,} rows & {len(df.columns)} columns extracted")
            except Exception as e:
                failed.append(f.name)
                st.error(f"✗ {f.name}: {e}")

        bar.progress(1.0, text="Done!")

        if dfs:
            # Union columns across all files — a batch can mix older-template
            # PDFs (fewer columns) with newer ones that add trailing fields.
            # Known columns keep their fixed order; new ones are appended in
            # the order they were first encountered. Missing columns for a
            # given file are blank-filled rather than dropping the file's data.
            extra_cols: list[str] = []
            for df in dfs:
                for c in df.columns:
                    if c not in EXPECTED_COLS and c not in extra_cols:
                        extra_cols.append(c)
            master_cols = EXPECTED_COLS + extra_cols
            dfs = [df.reindex(columns=master_cols, fill_value="") for df in dfs]

            master = pd.concat(dfs, ignore_index=True)

            st.divider()

            # Summary metrics
            m1, m2, m3 = st.columns(3)
            m1.markdown(f'<div class="stat-box"><h3>{len(files)}</h3><p>PDFs Processed</p></div>', unsafe_allow_html=True)
            m2.markdown(f'<div class="stat-box"><h3>{len(master):,}</h3><p>Total Rows</p></div>', unsafe_allow_html=True)
            m3.markdown(f'<div class="stat-box"><h3>{len(master.columns)}</h3><p>Columns</p></div>', unsafe_allow_html=True)

            master = master.fillna("")
            for col in master.columns:
                if col in DATE_COLS:
                    master[col] = _parse_date_col(master[col])
                    continue
                non_empty = master[col] != ""
                if not non_empty.any():
                    continue
                numeric = pd.to_numeric(master[col].replace("", pd.NA), errors="coerce")
                if numeric[non_empty].notna().all():
                    master[col] = numeric
                else:
                    master[col] = master[col].astype(str)

            st.markdown(f"#### Preview: first 20 rows of {len(master):,} total")
            preview = master.head(20).copy()
            for col in DATE_COLS:
                if col in preview.columns:
                    preview[col] = preview[col].dt.strftime("%d-%m-%Y").fillna("")
            st.dataframe(preview, use_container_width=True)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                master.to_excel(w, index=False, sheet_name="Master")
                ws = w.sheets["Master"]
                for col_idx, col_name in enumerate(master.columns, start=1):
                    if col_name in DATE_COLS:
                        col_letter = ws.cell(1, col_idx).column_letter
                        for cell in ws[col_letter][1:]:
                            cell.number_format = "DD-MM-YYYY"
            buf.seek(0)

            st.divider()
            st.download_button(
                "⬇  Download Master LCC Excel",
                data=buf.getvalue(),
                file_name="Master_LCC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if failed:
            st.error(f"Failed PDFs: {', '.join(failed)}")

render_report_issue()
