"""
PDF to Excel / txt
------------------
Pull tables and text out of a PDF — supplier invoices, delivery notes,
price lists, statements — and download them as Excel or plain text.

Runs as a second page of the same Streamlit app, so it shares the
deployment and the URL. Nothing here touches the negative stock data.
"""

import io
import re

import pandas as pd
import streamlit as st

try:
    import pdfplumber
    PDF_OK, PDF_ERR = True, None
except Exception as e:  # keep the page usable, say why it is not
    PDF_OK, PDF_ERR = False, f"{type(e).__name__}: {e}"

st.set_page_config(page_title="PDF to Excel", page_icon="📄", layout="wide")
st.title("📄 PDF to Excel / txt")

if not PDF_OK:
    st.error(f"pdfplumber is not installed on the server. {PDF_ERR}")
    st.code("pdfplumber>=0.11", language=None)
    st.caption("Add that line to requirements.txt and redeploy.")
    st.stop()

XL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

up = st.file_uploader("PDF file", type=["pdf"])
if up is None:
    st.info("Upload a PDF to start. Works on invoices, delivery notes, "
            "price lists and statements that have a real text layer.")
    st.caption("A scanned or photographed PDF has no text layer — nothing "
               "can be extracted from it without OCR, which this page does "
               "not do.")
    st.stop()


@st.cache_data(show_spinner=False)
def open_pdf(raw: bytes):
    """Returns per-page text and tables, plus a flag for a missing text layer."""
    pages = []
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            tables = page.extract_tables() or []
            pages.append({"n": i, "text": text, "tables": tables,
                          "chars": len(text)})
    return pages


with st.spinner("Reading the PDF…"):
    try:
        pages = open_pdf(up.getvalue())
    except Exception as e:
        st.error(f"Could not read that file: {type(e).__name__}: {e}")
        st.stop()

total_chars = sum(p["chars"] for p in pages)
total_tables = sum(len(p["tables"]) for p in pages)

c1, c2, c3 = st.columns(3)
c1.metric("Pages", len(pages))
c2.metric("Tables found", total_tables)
c3.metric("Characters of text", f"{total_chars:,}")

if total_chars == 0:
    st.error("No text layer in this PDF — it is a scan or a photo. "
             "Extraction cannot work on it. You would need an OCR tool, "
             "or the original file from the supplier.")
    st.stop()

tab_tables, tab_text = st.tabs(["Tables → Excel", "Text → txt"])

# ---------------------------------------------------------------- tables
with tab_tables:
    if not total_tables:
        st.warning("pdfplumber found no ruled tables. If the data is laid out "
                   "in columns without lines, use the Text tab and split it "
                   "there instead.")
    else:
        first_row_header = st.checkbox("Use the first row of each table as "
                                       "the column headings", True)
        drop_empty = st.checkbox("Drop completely empty rows and columns", True)

        frames = []
        for p in pages:
            for j, t in enumerate(p["tables"], start=1):
                if not t:
                    continue
                df = pd.DataFrame(t)
                if drop_empty:
                    df = df.replace({None: ""}).astype(str)
                    df = df.loc[~(df == "").all(axis=1), ~(df == "").all(axis=0)]
                if first_row_header and len(df) > 1:
                    df.columns = [str(c).strip() or f"col{k}"
                                  for k, c in enumerate(df.iloc[0], start=1)]
                    df = df.iloc[1:].reset_index(drop=True)
                if len(df):
                    frames.append((f"p{p['n']}_t{j}", df))

        if not frames:
            st.warning("Tables were detected but came out empty after cleaning.")
        else:
            st.caption(f"{len(frames)} table(s). Pick one to preview.")
            which = st.selectbox("Table", [n for n, _ in frames])
            prev = dict(frames)[which]
            st.dataframe(prev, use_container_width=True, height=360)

            buf_one = io.BytesIO()
            with pd.ExcelWriter(buf_one, engine="openpyxl") as xw:
                prev.to_excel(xw, sheet_name=which[:31], index=False)

            buf_all = io.BytesIO()
            with pd.ExcelWriter(buf_all, engine="openpyxl") as xw:
                for name, df in frames:
                    df.to_excel(xw, sheet_name=name[:31], index=False)

            stacked = pd.concat(
                [df.assign(_source=name) for name, df in frames
                 if list(df.columns) == list(frames[0][1].columns)],
                ignore_index=True) if len(frames) > 1 else frames[0][1]
            buf_stack = io.BytesIO()
            with pd.ExcelWriter(buf_stack, engine="openpyxl") as xw:
                stacked.to_excel(xw, sheet_name="ALL", index=False)

            d1, d2 = st.columns(2)
            d1.download_button(f"⬇ Excel — this table ({which})",
                               buf_one.getvalue(), f"{which}.xlsx", XL,
                               use_container_width=True, type="primary")
            d2.download_button(f"⬇ Csv — this table ({which})",
                               prev.to_csv(index=False).encode(),
                               f"{which}.csv", "text/csv",
                               use_container_width=True)

            d3, d4 = st.columns(2)
            d3.download_button(f"⬇ Excel — all {len(frames)} tables, one tab each",
                               buf_all.getvalue(), "tables_by_page.xlsx", XL,
                               use_container_width=True)
            d4.download_button("⬇ Excel — all tables stacked into one sheet",
                               buf_stack.getvalue(), "tables_stacked.xlsx", XL,
                               use_container_width=True,
                               help="Only stacks tables whose columns match "
                                    "the first one.")

# ---------------------------------------------------------------- text
with tab_text:
    pick = st.multiselect("Pages", [p["n"] for p in pages],
                          default=[p["n"] for p in pages])
    keep = [p for p in pages if p["n"] in pick]
    joiner = st.radio("Between pages", ["Blank line", "Page marker", "Nothing"],
                      horizontal=True)

    parts = []
    for p in keep:
        if joiner == "Page marker":
            parts.append(f"--- page {p['n']} ---\n{p['text']}")
        else:
            parts.append(p["text"])
    text = ("\n\n" if joiner == "Blank line" else "\n").join(parts)

    strip_blanks = st.checkbox("Collapse repeated blank lines", True)
    if strip_blanks:
        text = re.sub(r"\n{3,}", "\n\n", text)

    st.text_area("Preview", text[:6000], height=320)
    if len(text) > 6000:
        st.caption(f"Showing the first 6,000 of {len(text):,} characters.")

    t1, t2 = st.columns(2)
    t1.download_button("⬇ Txt — selected pages", text.encode("utf-8"),
                       f"{up.name.rsplit('.', 1)[0]}.txt", "text/plain",
                       use_container_width=True, type="primary")

    rows = [{"page": p["n"], "line": i + 1, "text": ln}
            for p in keep for i, ln in enumerate(p["text"].splitlines()) if ln.strip()]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name="LINES", index=False)
    t2.download_button("⬇ Excel — one row per line", buf.getvalue(),
                       f"{up.name.rsplit('.', 1)[0]}_lines.xlsx", XL,
                       use_container_width=True,
                       help="Useful when the layout has no ruled table — "
                            "split the column in Excel afterwards.")

    with st.expander("Find barcodes and numbers in the text"):
        pat = st.text_input("Pattern (regex)", r"\b\d{8,14}\b",
                            help="Default finds 8 to 14 digit runs, which "
                                 "covers most barcodes.")
        try:
            found = re.findall(pat, text)
        except re.error as e:
            st.error(f"Bad pattern: {e}")
            found = []
        if found:
            uniq = list(dict.fromkeys(found))
            st.caption(f"{len(found)} matches, {len(uniq)} unique")
            st.code("\n".join(uniq), language=None)
        else:
            st.caption("No matches.")
