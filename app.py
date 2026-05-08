import io
import re
from collections import defaultdict

import fitz
import pandas as pd
import streamlit as st
from lxml import html as lxml_html

st.set_page_config(page_title="LCC PDF to Excel", page_icon="📊", layout="wide")

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
    Build the final list of N_COLS column X positions.

    Strategy:
    1. Learn positions from data rows (accurate, handles dynamic column widths,
       but misses columns that are entirely empty in this PDF).
    2. Supplement with header X positions for any column that had no data —
       those columns will remain empty in the output, but their X slot must
       exist so subsequent columns aren't shifted by one position.
    3. Sort and cap at N_COLS.
    """
    header_xs = [x for x, _ in header_items]
    data_xs   = _data_clusters(row_map, header_y)

    if not data_xs:
        return sorted(header_xs)[:N_COLS]

    # Inject header positions that have no matching data cluster.
    # These represent columns that are genuinely empty in this PDF.
    combined = list(data_xs)
    for hx in header_xs:
        if not any(abs(hx - dx) <= 20 for dx in combined):
            combined.append(hx)

    return sorted(combined)[:N_COLS]


def _assign_boundary(items: list[tuple], col_xs: list[float]) -> list[str]:
    """
    Assign each (left, text) span to a column using boundary-based logic:
      column i owns X values in  [col_xs[i], col_xs[i+1]).
    This captures right-aligned or center-aligned values that sit inside the
    column's range but don't exactly match the column's left edge.
    """
    # Build right-side boundaries for each column
    boundaries = [col_xs[i + 1] for i in range(len(col_xs) - 1)]
    boundaries.append(float("inf"))

    cells: dict[int, str] = {}
    for left, text in items:
        # Assign to rightmost column whose left edge <= this value's X
        idx = 0
        for i, cx in enumerate(col_xs):
            if cx <= left:
                idx = i
            else:
                break
        # Verify the value is actually before the next column boundary
        if left < boundaries[idx]:
            cells[idx] = (cells.get(idx, "") + " " + text).strip()

    return [cells.get(i, "") for i in range(len(col_xs))]


def pdf_to_dataframe(file_bytes: bytes) -> pd.DataFrame:
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    all_para: list[tuple[float, float, str]] = []
    for page in doc:
        all_para.extend(_parse_html_paragraphs(page))
    doc.close()

    row_map = _group_by_row(all_para)

    # Find header row
    header_items: list | None = None
    header_y: float | None    = None
    for y in sorted(row_map):
        items = sorted(row_map[y], key=lambda x: x[0])
        if HEADER_SIGNAL & {t for _, t in items}:
            header_items = items
            header_y     = y
            break

    if header_y is None or header_items is None:
        return pd.DataFrame(columns=EXPECTED_COLS)

    # Build column X positions: data rows for accuracy + header for empty columns
    col_xs = _build_col_xs(header_items, row_map, header_y)

    if not col_xs:
        return pd.DataFrame(columns=EXPECTED_COLS)

    # Extract data rows
    data: list[list[str]] = []
    for y in sorted(row_map):
        if abs(y - header_y) <= Y_TOL:
            continue
        items = sorted(row_map[y], key=lambda x: x[0])
        if not items or len(items) < 5:          # skip title/separator rows
            continue
        if items[0][1] == "0" and any("ZTotal" in t for _, t in items):
            continue
        row = _assign_boundary(items, col_xs)
        if any(row):
            data.append(row)

    n = len(col_xs)
    df = pd.DataFrame(
        [r[:n] + [""] * max(0, n - len(r)) for r in data],
        columns=EXPECTED_COLS[:n],
    )
    return df.reindex(columns=EXPECTED_COLS, fill_value="")


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
    <span style="color:#4a6fa5;font-size:0.85rem"> 85 fixed columns &nbsp;|&nbsp;  All RE's data merged</span>
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
                    st.success(f"✔ {f.name} — {len(df):,} rows extracted")
            except Exception as e:
                failed.append(f.name)
                st.error(f"✗ {f.name}: {e}")

        bar.progress(1.0, text="Done!")

        if dfs:
            master = pd.concat(dfs, ignore_index=True)

            st.divider()

            # Summary metrics
            m1, m2, m3 = st.columns(3)
            m1.markdown(f'<div class="stat-box"><h3>{len(files)}</h3><p>PDFs Processed</p></div>', unsafe_allow_html=True)
            m2.markdown(f'<div class="stat-box"><h3>{len(master):,}</h3><p>Total Rows</p></div>', unsafe_allow_html=True)
            m3.markdown(f'<div class="stat-box"><h3>{len(master.columns)}</h3><p>Columns</p></div>', unsafe_allow_html=True)

            st.markdown(f"#### Preview — first 20 rows of {len(master):,} total")
            st.dataframe(master.head(20), use_container_width=True)

            master = master.fillna("").astype(str).replace("nan", "")

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                master.to_excel(w, index=False, sheet_name="Master")
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
