import csv
import html
import io
import re
import zipfile
from pathlib import Path

import polars as pl
import streamlit as st
from xlsxwriter import Workbook

st.set_page_config(
    page_title="Excel ⇄ CSV Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# Styling (works in both light and dark themes)
# ----------------------------------------------------------------------------
CSS = """
<style>
.block-container { padding-top: 2rem; max-width: 1200px; }
#MainMenu, footer { visibility: hidden; }

/* Hero banner */
.hero {
    background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 50%, #db2777 100%);
    border-radius: 20px;
    padding: 2.2rem 2.5rem;
    color: #fff;
    margin-bottom: 1.6rem;
    box-shadow: 0 10px 30px rgba(124, 58, 237, 0.30);
}
.hero h1 { margin: 0; font-size: 2.2rem; font-weight: 800; color: #fff; padding: 0; }
.hero p  { margin: .4rem 0 0 0; font-size: 1.05rem; opacity: .92; color: #fff; }

/* Generic cards */
.card {
    background: rgba(128, 128, 128, 0.07);
    border: 1px solid rgba(128, 128, 128, 0.20);
    border-radius: 16px;
    padding: 1.2rem 1.3rem;
    height: 100%;
}
.card .icon  { font-size: 1.8rem; }
.card .title { font-weight: 700; font-size: 1.05rem; margin-top: .3rem; }
.card .desc  { opacity: .75; font-size: .92rem; margin-top: .15rem; }

/* Sheet selector (radio styled as chips) */
div[role="radiogroup"] { gap: .6rem; flex-wrap: wrap; }
div[role="radiogroup"] > label {
    background: rgba(128, 128, 128, 0.07);
    border: 1px solid rgba(128, 128, 128, 0.30);
    border-radius: 12px;
    padding: .55rem 1rem;
    cursor: pointer;
    transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}
div[role="radiogroup"] > label:hover {
    transform: translateY(-2px);
    border-color: #7c3aed;
    box-shadow: 0 6px 16px rgba(0,0,0,.12);
}
div[role="radiogroup"] > label:has(input:checked) {
    background: rgba(124, 58, 237, 0.16);
    border: 2px solid #7c3aed;
    font-weight: 700;
}

/* Metrics */
[data-testid="stMetric"] {
    background: linear-gradient(135deg, rgba(79,70,229,.12), rgba(219,39,119,.10));
    border: 1px solid rgba(128, 128, 128, 0.20);
    border-radius: 16px;
    padding: 1rem 1.2rem;
}
[data-testid="stMetricLabel"] { opacity: .8; }
[data-testid="stMetricValue"] { font-weight: 800; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { gap: .4rem; }
.stTabs [data-baseweb="tab"] {
    border-radius: 10px 10px 0 0;
    padding: .6rem 1.2rem;
    font-weight: 600;
}

/* Buttons */
.stDownloadButton button, .stButton button {
    border-radius: 12px;
    padding: .65rem 1.4rem;
    font-weight: 700;
    border: none;
    color: #fff;
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    box-shadow: 0 4px 14px rgba(124, 58, 237, .35);
    transition: transform .15s ease, box-shadow .15s ease;
}
.stDownloadButton button:hover, .stButton button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(124, 58, 237, .45);
    color: #fff;
}

/* Uploader */
[data-testid="stFileUploader"] section {
    border: 2px dashed rgba(124, 58, 237, .55);
    border-radius: 16px;
    background: rgba(124, 58, 237, .05);
}

/* Dataframes */
[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid rgba(128,128,128,.2); }

/* Filename pill */
.file-pill {
    display: inline-block;
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    background: rgba(124, 58, 237, .14);
    border: 1px solid rgba(124, 58, 237, .45);
    border-radius: 999px;
    padding: .35rem 1rem;
    font-weight: 600;
}

/* ---------- Shadow effects ---------- */
.hero {
    box-shadow: 0 20px 40px -12px rgba(124, 58, 237, .45),
                0 8px 16px -8px rgba(0, 0, 0, .25);
}
.card {
    box-shadow: 0 4px 6px rgba(0,0,0,.06), 0 10px 24px rgba(0,0,0,.10);
    transition: transform .2s ease, box-shadow .2s ease;
}
.card:hover {
    transform: translateY(-4px);
    box-shadow: 0 14px 30px rgba(124, 58, 237, .25), 0 6px 12px rgba(0,0,0,.12);
}
[data-testid="stMetric"] {
    box-shadow: 0 4px 6px rgba(0,0,0,.06), 0 8px 20px rgba(79, 70, 229, .15);
    transition: transform .2s ease, box-shadow .2s ease;
}
[data-testid="stMetric"]:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 26px rgba(79, 70, 229, .28);
}
[data-testid="stDataFrame"] {
    box-shadow: 0 6px 20px rgba(0,0,0,.14);
}
[data-testid="stFileUploader"] section {
    box-shadow: 0 6px 18px rgba(124, 58, 237, .18);
    transition: box-shadow .2s ease, transform .2s ease;
}
[data-testid="stFileUploader"] section:hover {
    box-shadow: 0 10px 28px rgba(124, 58, 237, .32);
    transform: translateY(-2px);
}
div[role="radiogroup"] > label {
    box-shadow: 0 2px 8px rgba(0,0,0,.10);
}
div[role="radiogroup"] > label:has(input:checked) {
    box-shadow: 0 6px 18px rgba(124, 58, 237, .40);
}
[data-testid="stExpander"] details {
    border-radius: 14px;
    box-shadow: 0 4px 14px rgba(0,0,0,.10);
}
[data-baseweb="select"] > div, [data-baseweb="input"] {
    box-shadow: 0 2px 8px rgba(0,0,0,.10);
    border-radius: 10px;
}
[data-testid="stSidebar"] {
    box-shadow: 4px 0 24px rgba(0,0,0,.14);
}
.file-pill {
    box-shadow: 0 4px 12px rgba(124, 58, 237, .25);
}
.stTabs [aria-selected="true"] {
    text-shadow: 0 1px 8px rgba(124, 58, 237, .35);
}

/* ---------- Sidebar content ---------- */
.sb-brand {
    font-size: 1.35rem; font-weight: 800; line-height: 1.2;
    background: linear-gradient(135deg, #4f46e5, #db2777);
    -webkit-background-clip: text; background-clip: text; color: transparent;
    margin-bottom: .1rem;
}
.sb-sub { opacity: .7; font-size: .85rem; margin-bottom: .6rem; }
.sb-heading { font-weight: 700; font-size: 1rem; margin: .9rem 0 .35rem 0; }
.sb-group { font-weight: 600; font-size: .85rem; opacity: .85; margin: .7rem 0 .2rem 0; }
.sb-step { display: flex; gap: .65rem; align-items: flex-start; margin: .42rem 0; }
.sb-num {
    flex: 0 0 24px; height: 24px; border-radius: 50%;
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    color: #fff; font-weight: 700; font-size: .78rem;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 3px 8px rgba(124, 58, 237, .40);
}
.sb-txt { font-size: .88rem; line-height: 1.35; padding-top: 2px; }
.sb-tip {
    background: rgba(124, 58, 237, .10);
    border: 1px solid rgba(124, 58, 237, .30);
    border-radius: 12px; padding: .7rem .85rem; font-size: .84rem; line-height: 1.4;
    box-shadow: 0 4px 12px rgba(124, 58, 237, .12);
    margin: .3rem 0;
}
.sb-foot { opacity: .6; font-size: .78rem; text-align: center; margin-top: 1rem; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


PREVIEW_ROWS = 10  # preview is limited to this many rows


# ----------------------------------------------------------------------------
# Shared helpers
# ----------------------------------------------------------------------------
def show_df(data, **kwargs):
    """st.dataframe that works on both old and new Streamlit versions."""
    try:
        st.dataframe(data, width="stretch", **kwargs)
    except Exception:
        st.dataframe(data, use_container_width=True, **kwargs)


def safe_name(text: str) -> str:
    return re.sub(r'[\\/*?:"<>|]+', "_", text).strip()


def human_size(n: int) -> str:
    size = float(n)
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{size:.1f} TB"


def empty_state(cards):
    cols = st.columns(len(cards))
    for col, (icon, title, desc) in zip(cols, cards):
        col.markdown(
            f'<div class="card"><div class="icon">{icon}</div>'
            f'<div class="title">{title}</div><div class="desc">{desc}</div></div>',
            unsafe_allow_html=True,
        )


# ----------------------------------------------------------------------------
# Excel -> CSV helpers
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Reading Excel file...")
def load_all_sheets(file_bytes: bytes) -> dict[str, pl.DataFrame]:
    """Load every sheet in the workbook as {sheet_name: DataFrame}."""
    return pl.read_excel(
        io.BytesIO(file_bytes),
        sheet_id=0,  # 0 = read ALL sheets
        infer_schema_length=None,  # scan all rows to avoid mixed-type errors
    )


def to_csv_bytes(df: pl.DataFrame, sep: str, header: bool) -> bytes:
    return df.write_csv(separator=sep, include_header=header).encode("utf-8")


# ----------------------------------------------------------------------------
# CSV -> Excel helpers
# ----------------------------------------------------------------------------
DELIMITERS = {
    "Auto-detect": None,
    "Comma ( , )": ",",
    "Semicolon ( ; )": ";",
    "Tab": "\t",
    "Pipe ( | )": "|",
}


def detect_delimiter(file_bytes: bytes) -> str:
    sample = file_bytes[:8192].decode("utf-8", errors="ignore")
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except Exception:
        return ","


@st.cache_data(show_spinner="Reading CSV...")
def read_csv_bytes(file_bytes: bytes, sep: str, infer_types: bool) -> pl.DataFrame:
    """Read a CSV.

    infer_types=False (default): every column is read as text, so IDs, leading
    zeros and long numbers are preserved exactly and mixed columns never fail.
    infer_types=True: detect numbers/dates; falls back to text if that fails.
    """
    common = dict(
        separator=sep,
        encoding="utf8-lossy",  # tolerate non-UTF-8 characters
        truncate_ragged_lines=True,
    )
    if infer_types:
        try:
            return pl.read_csv(io.BytesIO(file_bytes), infer_schema_length=None, **common)
        except Exception:
            pass  # fall through to the safe text-only read
    return pl.read_csv(io.BytesIO(file_bytes), infer_schema_length=0, **common)


def make_sheet_name(raw: str, used: set[str]) -> str:
    """Excel sheet names: max 31 chars, no []:*?/\\ and must be unique (case-insensitive)."""
    name = re.sub(r"[\[\]:*?/\\]+", "_", raw).strip().strip("'") or "Sheet"
    name = name[:31]
    candidate, i = name, 2
    while candidate.lower() in used:
        suffix = f"_{i}"
        candidate = name[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate.lower())
    return candidate


def build_workbook(frames: list[tuple[str, pl.DataFrame]]) -> bytes:
    buf = io.BytesIO()
    with Workbook(buf, {"constant_memory": False}) as wb:
        for sheet_name, df in frames:
            df.write_excel(
                workbook=wb,
                worksheet=sheet_name,
                autofit=True,
                header_format={"bold": True, "bg_color": "#EDE9FE", "border": 1},
            )
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Header + sidebar
# ----------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
        <h1>📊 Excel ⇄ CSV Converter</h1>
        <p>Convert Excel sheets to CSV, or combine CSV files into one Excel workbook.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

def steps_html(steps: list[str]) -> str:
    return "".join(
        f'<div class="sb-step"><div class="sb-num">{i}</div><div class="sb-txt">{t}</div></div>'
        for i, t in enumerate(steps, 1)
    )


with st.sidebar:
    st.markdown(
        '<div class="sb-brand">📊 Converter Toolkit</div>'
        '<div class="sb-sub">Excel ⇄ CSV, fast and private-friendly.</div>',
        unsafe_allow_html=True,
    )

    # ---- How it works ----
    st.markdown('<div class="sb-heading">📖 How it works</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sb-group">📊 Excel → CSV</div>'
        + steps_html(
            [
                "Upload an <b>.xlsx / .xlsm / .xls</b> file.",
                "Pick the <b>sheet</b> you want from the chips.",
                "Check the <b>preview</b> (first 10 rows), columns and stats.",
                "Download <b>filename_sheetname.csv</b>, or every sheet as a ZIP.",
            ]
        )
        + '<div class="sb-group">🔄 CSV → Excel</div>'
        + steps_html(
            [
                "Upload <b>one or more</b> CSV files.",
                "Choose the <b>delimiter</b> or leave it on auto-detect.",
                "Review the <b>preview</b> of each file.",
                "Name the workbook and download the <b>.xlsx</b> (one sheet per CSV).",
            ]
        ),
        unsafe_allow_html=True,
    )

    st.divider()

    # ---- Settings ----
    st.markdown('<div class="sb-heading">⚙️ Excel → CSV settings</div>', unsafe_allow_html=True)
    delimiter_label = st.selectbox(
        "CSV delimiter",
        ["Comma ( , )", "Semicolon ( ; )", "Tab", "Pipe ( | )"],
    )
    out_delimiter = {
        "Comma ( , )": ",",
        "Semicolon ( ; )": ";",
        "Tab": "\t",
        "Pipe ( | )": "|",
    }[delimiter_label]
    include_header = st.toggle("Include header row", value=True)

    st.divider()

    # ---- Tips ----
    st.markdown('<div class="sb-heading">💡 Good to know</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sb-tip">🔢 CSV → Excel keeps every column as <b>text</b> by default, '
        "so IDs and long numbers are never altered. Tick <i>Detect numbers &amp; dates</i> "
        "if you need real numbers.</div>"
        '<div class="sb-tip">📑 Excel sheet names are limited to <b>31 characters</b>; '
        "longer CSV names are shortened automatically.</div>"
        '<div class="sb-tip">👀 Previews show only the first <b>10 rows</b>, '
        "but downloads always contain <b>all rows</b>.</div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="sb-foot">Built with Streamlit &amp; Polars</div>',
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------------
# Page 1: Excel -> CSV
# ----------------------------------------------------------------------------
def excel_to_csv_page():
    uploaded = st.file_uploader(
        "Drop your Excel file here",
        type=["xlsx", "xlsm", "xls"],
        label_visibility="collapsed",
        key="excel_uploader",
    )

    if uploaded is None:
        empty_state(
            [
                ("📂", "Load every sheet", "All sheets in the workbook are read automatically."),
                ("👀", "Preview instantly", "Check rows, column types and statistics before exporting."),
                ("💾", "Export your way", "Named as filename_sheetname.csv, with your delimiter."),
            ]
        )
        return

    try:
        sheets = load_all_sheets(uploaded.getvalue())
    except Exception as e:
        st.error(f"Could not read the file: {e}")
        return

    if not sheets:
        st.warning("No sheets found in this workbook.")
        return

    sheet_names = list(sheets.keys())
    base = safe_name(Path(uploaded.name).stem)

    m1, m2, m3 = st.columns(3)
    m1.metric("📁 File", uploaded.name if len(uploaded.name) < 24 else uploaded.name[:21] + "...")
    m2.metric("🗂️ Sheets", len(sheet_names))
    m3.metric("📦 Size", human_size(uploaded.size))

    st.markdown("### 🗂️ Choose a sheet")

    def sheet_label(s: str) -> str:
        return f"{s}  ({sheets[s].height:,} rows × {sheets[s].width} cols)"

    widget_key = f"sheet_choice_{uploaded.name}_{uploaded.size}"
    if len(sheet_names) <= 8:
        selected = st.radio(
            "Select the sheet to preview and export",
            sheet_names,
            format_func=sheet_label,
            horizontal=True,
            key=widget_key,
        )
    else:
        selected = st.selectbox(
            "Select the sheet to preview and export",
            sheet_names,
            format_func=sheet_label,
            key=widget_key,
        )

    df = sheets[selected]

    tab_preview, tab_cols, tab_stats, tab_export = st.tabs(
        ["👀 Preview", "🧬 Columns", "📈 Statistics", "💾 Export"]
    )

    with tab_preview:
        k1, k2, k3 = st.columns(3)
        k1.metric("Rows", f"{df.height:,}")
        k2.metric("Columns", f"{df.width:,}")
        k3.metric("Null cells", f"{sum(df.null_count().row(0)):,}")

        if df.height == 0:
            st.info("This sheet is empty.")
        else:
            show_df(df.head(PREVIEW_ROWS))
            st.caption(f"Showing the first {min(PREVIEW_ROWS, df.height)} of {df.height:,} rows.")

    with tab_cols:
        nulls = df.null_count().row(0)
        show_df(
            pl.DataFrame(
                {
                    "Column": df.columns,
                    "Type": [str(t) for t in df.dtypes],
                    "Nulls": list(nulls),
                    "Unique": [df[c].n_unique() for c in df.columns],
                }
            ),
            hide_index=True,
        )

    with tab_stats:
        if df.height == 0:
            st.info("This sheet is empty.")
        else:
            try:
                show_df(df.describe())
            except Exception as e:
                st.warning(f"Could not compute statistics for this sheet: {e}")

    with tab_export:
        csv_name = f"{base}_{safe_name(selected)}.csv"
        csv_bytes = to_csv_bytes(df, out_delimiter, include_header)

        st.markdown("#### Selected sheet")
        st.markdown(f'<span class="file-pill">{html.escape(csv_name)}</span>', unsafe_allow_html=True)
        st.caption(f"{df.height:,} rows · {df.width} columns · about {human_size(len(csv_bytes))}")
        st.download_button(
            "⬇️ Download CSV",
            data=csv_bytes,
            file_name=csv_name,
            mime="text/csv",
            type="primary",
        )

        if len(sheet_names) > 1:
            st.divider()
            st.markdown("#### Every sheet at once")
            st.caption("All sheets as separate CSV files in one ZIP, each named filename_sheetname.csv.")
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for s in sheet_names:
                    zf.writestr(
                        f"{base}_{safe_name(s)}.csv",
                        to_csv_bytes(sheets[s], out_delimiter, include_header),
                    )
            st.download_button(
                "📦 Download all sheets (ZIP)",
                data=buf.getvalue(),
                file_name=f"{base}_all_sheets.zip",
                mime="application/zip",
            )


# ----------------------------------------------------------------------------
# Page 2: CSV -> Excel
# ----------------------------------------------------------------------------
def csv_to_excel_page():
    files = st.file_uploader(
        "Drop one or more CSV files here",
        type=["csv", "tsv", "txt"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key="csv_uploader",
    )

    if not files:
        empty_state(
            [
                ("📄", "One or many CSVs", "Upload several files and each becomes its own sheet."),
                ("🔍", "Smart delimiter", "Comma, semicolon, tab or pipe is detected automatically."),
                ("📗", "Clean Excel output", "Bold, shaded headers and auto-fitted column widths."),
            ]
        )
        return

    o1, o2 = st.columns(2)
    with o1:
        delim_label = st.selectbox(
            "Delimiter used in your CSV file(s)",
            list(DELIMITERS.keys()),
            key="csv_in_delimiter",
        )
    with o2:
        st.write("")  # vertical alignment
        infer_types = st.checkbox(
            "Detect numbers & dates",
            value=False,
            help="Off (recommended): all columns stay as text, so IDs and long numbers are kept exactly. "
            "On: columns that look numeric or date-like are converted.",
            key="csv_infer_types",
        )

    frames: list[tuple[str, pl.DataFrame]] = []
    used_names: set[str] = set()
    problems = []

    for f in files:
        raw = f.getvalue()
        sep = DELIMITERS[delim_label] or detect_delimiter(raw)
        try:
            df = read_csv_bytes(raw, sep, infer_types)
        except Exception as e:
            problems.append((f.name, str(e)))
            continue
        frames.append((make_sheet_name(Path(f.name).stem, used_names), df))

    for name, err in problems:
        st.error(f"Could not read **{name}**: {err}")

    if not frames:
        return

    total_rows = sum(df.height for _, df in frames)
    m1, m2, m3 = st.columns(3)
    m1.metric("📄 CSV files", len(frames))
    m2.metric("📑 Total rows", f"{total_rows:,}")
    m3.metric("📦 Uploaded", human_size(sum(f.size for f in files)))

    st.markdown("### 👀 Preview")
    for sheet_name, df in frames:
        with st.expander(f"{sheet_name}  ({df.height:,} rows × {df.width} cols)", expanded=len(frames) == 1):
            if df.height == 0:
                st.info("This file has no data rows.")
            else:
                show_df(df.head(PREVIEW_ROWS))
                st.caption(f"Showing the first {min(PREVIEW_ROWS, df.height)} of {df.height:,} rows.")

    st.markdown("### 💾 Export")
    default_name = frames[0][0] if len(frames) == 1 else "combined"
    wb_name = st.text_input("Workbook name (without .xlsx)", value=default_name, key="wb_name")
    wb_name = safe_name(wb_name) or "converted"
    xlsx_name = f"{wb_name}.xlsx"

    try:
        xlsx_bytes = build_workbook(frames)
    except Exception as e:
        st.error(f"Could not build the Excel file: {e}")
        return

    st.markdown(f'<span class="file-pill">{html.escape(xlsx_name)}</span>', unsafe_allow_html=True)
    sheet_list = ", ".join(n for n, _ in frames)
    st.caption(f"{len(frames)} sheet(s): {sheet_list} · about {human_size(len(xlsx_bytes))}")
    st.download_button(
        "⬇️ Download Excel",
        data=xlsx_bytes,
        file_name=xlsx_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# ----------------------------------------------------------------------------
# Main layout
# ----------------------------------------------------------------------------
tab_x2c, tab_c2x = st.tabs(["📊 Excel → CSV", "🔄 CSV → Excel"])
with tab_x2c:
    excel_to_csv_page()
with tab_c2x:
    csv_to_excel_page()
