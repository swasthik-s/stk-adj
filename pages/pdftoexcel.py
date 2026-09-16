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
def _clean(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()


def header_score(row):
    """How much a row looks like column headings: several short, distinct,
    non-numeric cells."""
    cells = [_clean(c) for c in row]
    filled = [c for c in cells if c]
    if len(filled) < 3:
        return 0
    short = sum(1 for c in filled if len(c) <= 24)
    wordy = sum(1 for c in filled if re.search(r"[A-Za-z]", c))
    numeric = sum(1 for c in filled if re.fullmatch(r"[\d,.\-]+", c))
    distinct = len(set(filled))
    return (short + wordy + distinct - numeric * 2) if numeric < len(filled) / 2 \
        else 0


def tidy_table(raw_rows, min_rows=2):
    """Find the heading row inside a raw table, drop everything above it,
    and return a frame. None when it is not really a data table."""
    if not raw_rows or len(raw_rows) < 2:
        return None, None
    best_i, best = None, 0
    for i, r in enumerate(raw_rows[:8]):          # headings are near the top
        sc = header_score(r)
        if sc > best:
            best_i, best = i, sc
    if best_i is None:
        return None, None

    hdr = [_clean(c) for c in raw_rows[best_i]]
    hdr = [h if h else f"col{k}" for k, h in enumerate(hdr, start=1)]
    body = []
    for r in raw_rows[best_i + 1:]:
        cells = [_clean(c) for c in r]
        if not any(cells):
            continue
        # a repeated heading row further down (page breaks) — skip it
        if [c for c in cells if c] == [c for c in hdr if not c.startswith("col")]:
            continue
        if header_score(r) >= best and len([c for c in cells if c]) >= 3:
            continue
        body.append(cells[:len(hdr)] + [""] * max(0, len(hdr) - len(cells)))

    if len(body) < min_rows:
        return None, None
    df = pd.DataFrame(body, columns=hdr)
    df = df.loc[:, ~(df == "").all(axis=0)]       # drop empty columns
    df = df.loc[~(df == "").all(axis=1)]          # drop empty rows
    if df.shape[1] < 3 or len(df) < min_rows:
        return None, None
    return df, tuple(df.columns)


with tab_tables:
    only_data = st.checkbox(
        "Only real data tables", True,
        help="Skips address blocks, invoice headers and totals boxes — "
             "anything without proper column headings and at least two rows.")
    merge_same = st.checkbox(
        "Join tables that share the same columns", True,
        help="A line-item table split across pages becomes one table.")

    raw_tables = [(p["n"], j, t) for p in pages
                  for j, t in enumerate(p["tables"], start=1)]
    kept, skipped = [], []
    for pno, j, t in raw_tables:
        df, sig = tidy_table(t, min_rows=2 if only_data else 1)
        if df is None:
            skipped.append((f"p{pno}_t{j}", len(t)))
            if only_data:
                continue
            df = pd.DataFrame(t).replace({None: ""}).astype(str)
            sig = None
        kept.append((f"p{pno}_t{j}", df, sig))

    if merge_same and kept:
        groups, order = {}, []
        for name, df, sig in kept:
            key = sig or name
            if key not in groups:
                groups[key] = [name, df]
                order.append(key)
            else:
                groups[key][0] += f"+p{name.split('_')[0][1:]}"
                groups[key][1] = pd.concat([groups[key][1], df],
                                           ignore_index=True)
        kept = [(groups[k][0], groups[k][1], k) for k in order]

    c1, c2 = st.columns(2)
    c1.metric("Data tables kept", len(kept))
    c2.metric("Blocks skipped", len(skipped))
    if skipped:
        with st.expander("What was skipped"):
            st.caption("Address blocks, invoice header boxes, totals panels — "
                       "no column headings or fewer than two rows.")
            st.dataframe(pd.DataFrame(skipped, columns=["Block", "Raw rows"]),
                         use_container_width=True, hide_index=True)

    if not kept:
        st.warning("Nothing that looks like a data table. Untick the box above "
                   "to see every block, or use the Text tab.")
    else:
        which = st.selectbox("Table",
                             [f"{n}  ({len(d)} rows × {d.shape[1]} cols)"
                              for n, d, _ in kept])
        prev = kept[[f"{n}  ({len(d)} rows × {d.shape[1]} cols)"
                     for n, d, _ in kept].index(which)][1]
        st.dataframe(prev, use_container_width=True, height=380,
                     hide_index=True)

        buf_one = io.BytesIO()
        with pd.ExcelWriter(buf_one, engine="openpyxl") as xw:
            prev.to_excel(xw, sheet_name="TABLE", index=False)

        buf_all = io.BytesIO()
        with pd.ExcelWriter(buf_all, engine="openpyxl") as xw:
            for name, df, _ in kept:
                df.to_excel(xw, sheet_name=re.sub(r"[^A-Za-z0-9_]", "",
                                                  name)[:31] or "T",
                            index=False)

        d1, d2 = st.columns(2)
        d1.download_button("⬇ Excel — this table", buf_one.getvalue(),
                           "table.xlsx", XL, use_container_width=True,
                           type="primary")
        d2.download_button("⬇ Csv — this table",
                           prev.to_csv(index=False).encode(), "table.csv",
                           "text/csv", use_container_width=True)
        st.download_button(f"⬇ Excel — all {len(kept)} tables, one tab each",
                           buf_all.getvalue(), "tables.xlsx", XL,
                           use_container_width=True)

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
