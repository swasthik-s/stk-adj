"""
PDF → Excel / txt
=================
One page for every PDF job:

* supplier invoices — line items pulled out, arithmetic checked against the
  invoice's own totals, and the iTrade import file written for you
* anything else — tables cleaned up, address blocks and totals panels
  dropped, tables split across pages joined back together
* raw text when there is no ruled table at all
"""

import io
import re

import pandas as pd
import streamlit as st

try:
    import pdfplumber
    PDF_OK, PDF_ERR = True, None
except Exception as e:
    PDF_OK, PDF_ERR = False, f"{type(e).__name__}: {e}"

# OCR for scanned PDFs. RapidOCR is preferred: its models ship inside the pip
# package, so there is nothing to download at runtime and no system binary to
# install. Tesseract is kept as a fallback but reads decimals badly on dense
# tables — on this supplier's invoice it scored 2/5 on prices where RapidOCR
# scored 5/5.
OCR_ENGINE, OCR_ERR = None, None
_errs = []
# 1) rapidocr — the maintained package. Works on Python 3.8 to 3.14, models
#    bundled in the wheel, nothing downloaded at runtime.
try:
    import numpy as np
    import pypdfium2 as pdfium
    from rapidocr import RapidOCR
    OCR_ENGINE = "rapidocr"
except Exception as e:
    _errs.append(f"rapidocr: {type(e).__name__}: {e}")
# 2) rapidocr_onnxruntime — the older package. Same accuracy, but it refuses
#    to install on Python 3.13 and later, so it is only a local fallback.
if OCR_ENGINE is None:
    try:
        import numpy as np
        import pypdfium2 as pdfium
        from rapidocr_onnxruntime import RapidOCR
        OCR_ENGINE = "rapidocr_legacy"
    except Exception as e:
        _errs.append(f"rapidocr_onnxruntime: {type(e).__name__}: {e}")
# 3) tesseract — last resort; misreads decimals on dense tables.
if OCR_ENGINE is None:
    try:
        import pypdfium2 as pdfium
        import pytesseract
        pytesseract.get_tesseract_version()
        OCR_ENGINE = "tesseract"
    except Exception as e:
        _errs.append(f"tesseract: {type(e).__name__}: {e}")
OCR_ERR = " | ".join(_errs) or None
OCR_OK = OCR_ENGINE is not None


@st.cache_resource(show_spinner=False)
def _rapid():
    return RapidOCR()

st.title(":material/picture_as_pdf: PDF → Excel / txt")

if not PDF_OK:
    st.error(f"pdfplumber is not installed on the server. {PDF_ERR}")
    st.code("pdfplumber>=0.11", language=None)
    st.stop()

XL = ("application/vnd.openxmlformats-officedocument"
      ".spreadsheetml.sheet")

ALIAS = {
    "BARCODE": ["BARCODE", "ITEM CODE", "ITEMCODE", "CODE", "EAN"],
    "DESCRIPTION": ["DESCRIPTION", "ITEM NAME", "PARTICULARS", "ITEM"],
    "QTY": ["QTY", "QUANTITY", "RECD QTY", "REC QTY"],
    "PRICE": ["UNIT PRICE", "RATE", "PRICE", "COST", "UNIT COST"],
    "AMOUNT": ["AMOUNT", "VALUE", "NET AMOUNT"],
    "UNIT": ["UNIT", "UOM"],
    "SL": ["SL.NO", "SL NO", "S.NO", "SR.NO", "SR"],
    "VAT": ["VAT AMT", "VAT AMOUNT", "TAX AMT"],
    "AFTER": ["AFTER VAT AMT", "AFTER VAT", "GROSS"],
}


def _clean(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()


def _num(x):
    try:
        return float(str(x).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _find(hdr, keys):
    for k in keys:
        if k in hdr:
            return hdr.index(k)
    for i, h in enumerate(hdr):
        for k in keys:
            if h and k in h:
                return i
    return None


def header_score(row):
    """How much a row looks like column headings."""
    cells = [_clean(c) for c in row]
    filled = [c for c in cells if c]
    if len(filled) < 3:
        return 0
    numeric = sum(1 for c in filled if re.fullmatch(r"[\d,.\-]+", c))
    if numeric >= len(filled) / 2:
        return 0
    short = sum(1 for c in filled if len(c) <= 24)
    wordy = sum(1 for c in filled if re.search(r"[A-Za-z]", c))
    return short + wordy + len(set(filled))


def tidy_table(rows, min_rows=2):
    """Strip the preamble above the headings and return a clean frame."""
    if not rows or len(rows) < 2:
        return None, None
    best_i, best = None, 0
    for i, r in enumerate(rows[:8]):
        sc = header_score(r)
        if sc > best:
            best_i, best = i, sc
    if best_i is None:
        return None, None

    hdr = [_clean(c) for c in rows[best_i]]
    hdr = [h or f"col{k}" for k, h in enumerate(hdr, start=1)]
    body = []
    for r in rows[best_i + 1:]:
        cells = [_clean(c) for c in r]
        if not any(cells):
            continue
        if header_score(r) >= best:          # heading repeated at a page break
            continue
        body.append(cells[:len(hdr)] + [""] * max(0, len(hdr) - len(cells)))

    if len(body) < min_rows:
        return None, None
    df = pd.DataFrame(body, columns=hdr)
    df = df.loc[:, ~(df == "").all(axis=0)]
    df = df.loc[~(df == "").all(axis=1)]
    if df.shape[1] < 3 or len(df) < min_rows:
        return None, None
    return df, tuple(df.columns)


@st.cache_data(show_spinner=False)
def read_pdf(raw: bytes):
    pages, tables = [], []
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            pages.append({"n": i, "text": text})
            for j, t in enumerate(page.extract_tables() or [], start=1):
                tables.append((i, j, t))
    return pages, tables


def invoice_lines(tables):
    """Line items, if any table carries a BARCODE and a QTY column."""
    out = []
    for pno, _, tb in tables:
        hrow = hdr = None
        for i, r in enumerate(tb):
            h = [_clean(c).upper() for c in r]
            if (_find(h, ALIAS["BARCODE"]) is not None
                    and _find(h, ALIAS["QTY"]) is not None):
                hrow, hdr = i, h
                break
        if hrow is None:
            continue
        ix = {k: _find(hdr, v) for k, v in ALIAS.items()}
        for r in tb[hrow + 1:]:
            code = (_clean(r[ix["BARCODE"]]) if ix["BARCODE"] is not None
                    and ix["BARCODE"] < len(r) else "")
            if not re.fullmatch(r"\d{6,14}", code):
                continue

            def cell(key, as_num=False):
                i = ix.get(key)
                if i is None or i >= len(r):
                    return None if as_num else ""
                return _num(r[i]) if as_num else _clean(r[i])

            out.append({"SL": cell("SL", True), "Barcode": code,
                        "Description": cell("DESCRIPTION"),
                        "Unit": cell("UNIT"), "Qty": cell("QTY", True),
                        "UnitPrice": cell("PRICE", True),
                        "Amount": cell("AMOUNT", True),
                        "VatAmt": cell("VAT", True),
                        "AfterVat": cell("AFTER", True)})
    df = pd.DataFrame(out)
    if len(df):
        df = df.drop_duplicates(subset=["SL", "Barcode"]).reset_index(drop=True)
    return df



NUMTOK = re.compile(r"^[\d,]+(?:\.\d+)?$")


def lines_from_text(text, trailing=6):
    """Rebuild invoice rows from plain text, for pages that were OCR'd or
    have no ruled table.

    Read from the RIGHT: the money columns are always the last few numeric
    tokens (qty, price, amount, vat %, vat amount, after vat). Anchoring on
    the left breaks whenever OCR fumbles a narrow column — the CF column's
    "1" came back as "L" and ">" on two rows of the test invoice."""
    out = []
    for ln in text.splitlines():
        toks = ln.split()
        i = next((k for k, t in enumerate(toks)
                  if re.fullmatch(r"\d{6,14}", t)), None)
        if i is None:
            continue
        tail = []
        for t in reversed(toks):
            if NUMTOK.match(t):
                tail.append(t)
            else:
                break
        if len(tail) < trailing:
            continue
        tail = list(reversed(tail))[-trailing:]
        qty, price, amount, _vatpct, vatamt, after = [_num(x) for x in tail]
        sl = _num(toks[0]) if i > 0 and toks[0].isdigit() else None
        desc = " ".join(toks[i + 1:len(toks) - len(tail)])
        out.append({"SL": sl, "Barcode": toks[i], "Description": desc,
                    "Unit": "", "Qty": qty, "UnitPrice": price,
                    "Amount": amount, "VatAmt": vatamt, "AfterVat": after})
    df = pd.DataFrame(out)
    return (df.drop_duplicates("Barcode").reset_index(drop=True)
            if len(df) else df)


def invoice_head(full):
    def grab(pat):
        m = re.search(pat, full, re.I)
        return m.group(1).strip() if m else None
    return {"Invoice no": grab(r"INV NO\s*:?\s*([A-Z0-9\-/]+)"),
            "Date": grab(r"Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})"),
            "Supplier TRN": grab(r"TRN\s*:?\s*(\d{10,20})"),
            "Total": _num(grab(r"\bTotal\s+([\d,]+\.\d\d)")),
            "VAT": _num(grab(r"\bVAT\s+([\d,]+\.\d\d)")),
            "Net Total": _num(grab(r"Net Total\s+([\d,]+\.\d\d)"))}


def fmt(x, dp=7):
    if x is None:
        return "0"
    return f"{float(x):.{dp}f}".rstrip("0").rstrip(".") or "0"


def to_excel(frames: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in frames.items():
            sheet = re.sub(r"[^A-Za-z0-9_ ]", "", str(name))[:31] or "SHEET"
            df.to_excel(xw, sheet_name=sheet, index=False)
            ws = xw.sheets[sheet]
            for col in ws.iter_cols(min_row=1, max_row=1):
                letter = col[0].column_letter
                width = max(len(str(col[0].value or "")) + 2, 12)
                body = [len(str(ws.cell(row=r, column=col[0].column).value or ""))
                        for r in range(2, min(ws.max_row, 60) + 1)]
                ws.column_dimensions[letter].width = min(
                    max([width] + body) + 2, 48)
    return buf.getvalue()


def download_row(items, ns="dl"):
    """items: list of (label, data, filename, mime, primary).

    ns keeps widget keys unique — two tabs can legitimately offer files with
    the same name, and a key collision takes the whole page down."""
    cols = st.columns(len(items))
    for i, (col, (label, data, fname, mime, primary)) in enumerate(
            zip(cols, items)):
        col.download_button(label, data, fname, mime,
                            width="stretch",
                            type="primary" if primary else "secondary",
                            key=f"{ns}_{i}_{fname}")


# ==================================================================== input
ups = st.file_uploader("PDF file(s)", type=["pdf"],
                       accept_multiple_files=True)
if not ups:
    st.info("Upload one or more PDFs. Scanned pages are read automatically — "
            "nothing to switch on.")
    st.stop()

up = ups[0] if len(ups) == 1 else None
if up is None:
    names = [f.name for f in ups]
    up = ups[names.index(st.selectbox("File", names))]

with st.spinner("Reading…"):
    try:
        pages, raw_tables = read_pdf(up.getvalue())
    except Exception as e:
        st.error(f"Could not read that file: {type(e).__name__}: {e}")
        st.stop()

def _ocr_boxes(img):
    """(box, text) pairs from whichever RapidOCR is installed. The new
    package returns an object with .boxes/.txts; the old one a list."""
    out = _rapid()(np.array(img))
    if OCR_ENGINE == "rapidocr":
        if out is None or getattr(out, "boxes", None) is None:
            return []
        return list(zip(out.boxes, out.txts))
    res = out[0] if isinstance(out, tuple) else out
    return [(r[0], r[1]) for r in (res or [])]


def _rows_from_boxes(res, tol=14):
    """Rebuild visual rows from OCR boxes using their y positions, then order
    each row left to right. Without this the columns interleave."""
    items = sorted((float(sum(p[1] for p in box)) / 4,
                    float(sum(p[0] for p in box)) / 4, txt)
                   for box, txt in (res or []))
    rows, cur, last = [], [], None
    for y, x, txt in items:
        if last is None or abs(y - last) <= tol:
            cur.append((x, txt))
        else:
            rows.append(cur)
            cur = [(x, txt)]
        last = y
    rows.append(cur)
    return [" ".join(t for _, t in sorted(r)) for r in rows if r]


@st.cache_data(show_spinner=False)
def ocr_pdf(raw: bytes, dpi: int = 300):
    """Rasterise each page and read it. Returns page texts."""
    doc = pdfium.PdfDocument(io.BytesIO(raw))
    out = []
    for i in range(len(doc)):
        img = doc[i].render(scale=dpi / 72).to_pil().convert("RGB")
        if OCR_ENGINE in ("rapidocr", "rapidocr_legacy"):
            res = _ocr_boxes(img)
            text = "\n".join(_rows_from_boxes(res))
        else:
            text = pytesseract.image_to_string(img, config="--psm 6")
        out.append({"n": i + 1, "text": text})
    return out


# --------------------------------------------------------------- auto OCR
# A PDF can be mixed: some pages generated digitally, others scanned in.
# Decide page by page rather than treating the file as all one or the other.
THIN = 40          # characters — below this a page has no usable text layer


def ocr_page(raw: bytes, index: int, dpi: int):
    doc = pdfium.PdfDocument(io.BytesIO(raw))
    img = doc[index].render(scale=dpi / 72).to_pil().convert("RGB")
    if OCR_ENGINE in ("rapidocr", "rapidocr_legacy"):
        res = _ocr_boxes(img)
        return "\n".join(_rows_from_boxes(res))
    return pytesseract.image_to_string(img, config="--psm 6")


@st.cache_data(show_spinner=False)
def ocr_pages(raw: bytes, which: tuple, dpi: int = 300):
    return {i: ocr_page(raw, i, dpi) for i in which}


need = [p["n"] for p in pages if len(p["text"].strip()) < THIN]
ocr_used = []

if need and OCR_OK:
    with st.spinner(f"Reading {len(need)} scanned page(s)…"):
        try:
            got = ocr_pages(up.getvalue(), tuple(n - 1 for n in need), 300)
            for p in pages:
                if p["n"] in need and got.get(p["n"] - 1, "").strip():
                    p["text"] = got[p["n"] - 1]
                    ocr_used.append(p["n"])
        except Exception as e:
            st.error(f"OCR failed: {type(e).__name__}: {e}")
elif need and not OCR_OK:
    st.warning(f"Page(s) {', '.join(map(str, need))} are scanned and OCR is "
               f"not installed on this server.")
    st.caption("Add `pypdfium2` and `rapidocr-onnxruntime` to "
               "requirements.txt.")

full = "\n".join(p["text"] for p in pages)
if not full.strip():
    st.error("Nothing could be read from this file at all.")
    st.stop()

# Take rows from the ruled tables where they exist, and from the text for
# anything they missed — a mixed PDF has some of each, so merging beats
# choosing. Table rows win on a clash: they come from real cell boundaries.
inv = invoice_lines(raw_tables)
from_text = lines_from_text(full)
if len(from_text):
    if len(inv):
        have = set(inv["Barcode"].astype(str))
        extra = from_text[~from_text["Barcode"].astype(str).isin(have)]
        if len(extra):
            inv = pd.concat([inv, extra], ignore_index=True)
            if inv["SL"].notna().all():
                inv = inv.sort_values("SL").reset_index(drop=True)
    else:
        inv = from_text
head = invoice_head(full)
stem = re.sub(r"[^A-Za-z0-9]+", "_",
              (head.get("Invoice no") or up.name.rsplit(".", 1)[0]))[:40]

m1, m2, m3, m4 = st.columns(4)
m1.metric("Pages", len(pages),
          f"{len(ocr_used)} read by OCR" if ocr_used else None,
          delta_color="off")
m2.metric("Blocks found", len(raw_tables))
m3.metric("Invoice lines", len(inv) if len(inv) else "—")
m4.metric("Invoice no", head.get("Invoice no") or "—")

tabs = st.tabs((["Invoice"] if len(inv) else []) + ["Tables", "Text"])
ti = 0

# ================================================================== invoice
if len(inv):
    with tabs[0]:
        ti = 1
        c1, c2, c3 = st.columns(3)
        s_amt = float(inv["Amount"].sum()) if inv["Amount"].notna().any() else None
        if s_amt is not None and head.get("Total"):
            d = round(s_amt - head["Total"], 2)
            c1.metric("Lines vs invoice Total", f"{s_amt:,.2f}", f"{d:+.2f}",
                      delta_color="off")
            (c1.success if abs(d) <= 0.05 else c1.error)(
                "Matches" if abs(d) <= 0.05 else "Lines may be missing")
        s_gross = (float(inv["AfterVat"].sum())
                   if inv["AfterVat"].notna().any() else None)
        if s_gross is not None and head.get("Net Total"):
            d2 = round(s_gross - head["Net Total"], 2)
            c2.metric("After VAT vs Net Total", f"{s_gross:,.2f}", f"{d2:+.2f}",
                      delta_color="off")
            (c2.success if abs(d2) <= 0.05 else c2.error)(
                "Matches" if abs(d2) <= 0.05 else "Does not match")
        calc = (inv["Qty"] * inv["UnitPrice"]).round(2)
        off = inv[(calc - inv["Amount"].round(2)).abs() > 0.02]
        c3.metric("qty × price ≠ amount", len(off))
        if len(off):
            with c3.expander("Which rows"):
                st.dataframe(off[["SL", "Description", "Qty", "UnitPrice",
                                  "Amount"]], width="stretch",
                             hide_index=True)

        st.dataframe(inv, width="stretch", height=400, hide_index=True,
                     column_config={
                         "Barcode": st.column_config.TextColumn(width="medium"),
                         "Description": st.column_config.TextColumn(
                             width="large"),
                         "Qty": st.column_config.NumberColumn(format="%.3f"),
                         "UnitPrice": st.column_config.NumberColumn(
                             format="%.4f"),
                         "Amount": st.column_config.NumberColumn(format="%.2f")})

        o1, o2 = st.columns([1, 2])
        prefix = o1.text_input("Import prefix", "SML")
        o2.caption("Import file layout: prefix, barcode, unit price, quantity "
                   "— no header, no quotes.")
        txt = ("\n".join(f"{prefix},{r.Barcode},{fmt(r.UnitPrice)},"
                         f"{fmt(r.Qty, 3)}" for r in inv.itertuples()
                         if r.UnitPrice is not None) + "\n").encode("ascii",
                                                                    "ignore")
        xl = to_excel({"LINES": inv.drop(columns=["SL"]),
                       "INVOICE": pd.DataFrame(
                           [(k, v) for k, v in head.items()],
                           columns=["Field", "Value"])})
        reconciled = (s_amt is not None and head.get("Total")
                      and abs(round(s_amt - head["Total"], 2)) <= 0.05)
        from_ocr = bool(ocr_used)

        if from_ocr and not reconciled:
            st.error(
                "**Import file withheld.** These figures came from OCR and "
                "the line total does not match the invoice total, so at "
                "least one number was misread. Type the lines in by hand, or "
                "get the original PDF from the supplier."
            )
            download_row([
                ("Excel — check every figure", xl, f"{stem}_OCR_UNVERIFIED.xlsx",
                 XL, False),
                ("CSV", inv.to_csv(index=False).encode(),
                 f"{stem}_OCR_UNVERIFIED.csv", "text/csv", False),
            ], ns="inv")
        else:
            if from_ocr:
                st.warning("Figures came from OCR but the totals reconcile, "
                           "so the arithmetic holds. Still spot-check a few "
                           "prices against the paper before importing.")
            download_row([
                ("Excel", xl, f"{stem}.xlsx", XL, True),
                ("iTrade import", txt, f"{stem}.txt", "text/plain", False),
                ("CSV", inv.to_csv(index=False).encode(), f"{stem}.csv",
                 "text/csv", False),
            ], ns="inv")
        with st.expander("Preview the import file"):
            st.code(txt.decode(), language=None)
        with st.expander("Copy barcodes"):
            st.code("\n".join(inv["Barcode"]), language=None)

# =================================================================== tables
with tabs[ti]:
    f1, f2 = st.columns(2)
    only_data = f1.checkbox("Data tables only", True,
                            help="Drops address blocks, invoice header boxes "
                                 "and totals panels.")
    merge_same = f2.checkbox("Join tables with matching columns", True,
                             help="A table split across pages becomes one.")

    kept, skipped = [], []
    for pno, j, t in raw_tables:
        df, sig = tidy_table(t, 2 if only_data else 1)
        if df is None:
            skipped.append({"Block": f"page {pno}, block {j}",
                            "Raw rows": len(t)})
            if only_data:
                continue
            df, sig = pd.DataFrame(t).fillna("").astype(str), None
        kept.append([f"page {pno}", df, sig])

    if merge_same and kept:
        groups, order = {}, []
        for name, df, sig in kept:
            key = sig or name
            if key in groups:
                groups[key][1] = pd.concat([groups[key][1], df],
                                           ignore_index=True)
                groups[key][0] += f", {name.split()[-1]}"
            else:
                groups[key] = [name, df]
                order.append(key)
        kept = [[groups[k][0], groups[k][1], k] for k in order]

    k1, k2 = st.columns(2)
    k1.metric("Tables kept", len(kept))
    k2.metric("Blocks skipped", len(skipped))
    if skipped:
        with st.expander("What was skipped"):
            st.dataframe(pd.DataFrame(skipped), width="stretch",
                         hide_index=True)

    if not kept:
        st.warning("Nothing that looks like a data table. Untick the box "
                   "above to see everything, or use the Text tab.")
    else:
        labels = [f"{n} · {len(d)} rows × {d.shape[1]} cols"
                  for n, d, _ in kept]
        pick = labels[0] if len(labels) == 1 else st.selectbox("Table", labels)
        cur = kept[labels.index(pick)][1]
        st.dataframe(cur, width="stretch", height=380, hide_index=True)

        tsv = cur.to_csv(index=False, sep="\t").encode()
        download_row([
            ("Excel", to_excel({"TABLE": cur}), f"{stem}_table.xlsx", XL, True),
            ("CSV", cur.to_csv(index=False).encode(), f"{stem}_table.csv",
             "text/csv", False),
            ("Text", tsv, f"{stem}_table.txt", "text/plain", False),
        ], ns="tbl")
        if len(kept) > 1:
            st.download_button(
                f"Excel — all {len(kept)} tables, one tab each",
                to_excel({f"T{i+1}_{n}": d
                          for i, (n, d, _) in enumerate(kept)}),
                f"{stem}_all_tables.xlsx", XL, width="stretch",
                key="dl_all_tables")

# ===================================================================== text
with tabs[ti + 1]:
    pick_pages = st.multiselect("Pages", [p["n"] for p in pages],
                                default=[p["n"] for p in pages])
    keep = [p for p in pages if p["n"] in pick_pages]
    marker = st.checkbox("Mark page breaks", False)
    parts = [(f"--- page {p['n']} ---\n{p['text']}" if marker else p["text"])
             for p in keep]
    text = re.sub(r"\n{3,}", "\n\n", "\n\n".join(parts))

    st.text_area("Preview", text[:6000], height=300)
    if len(text) > 6000:
        st.caption(f"Showing the first 6,000 of {len(text):,} characters.")

    rows = [{"page": p["n"], "line": i + 1, "text": ln}
            for p in keep
            for i, ln in enumerate(p["text"].splitlines()) if ln.strip()]
    download_row([
        ("Text", text.encode("utf-8"), f"{stem}.txt", "text/plain", True),
        ("Excel — one row per line", to_excel({"LINES": pd.DataFrame(rows)}),
         f"{stem}_lines.xlsx", XL, False),
    ], ns="txt")

    with st.expander("Find barcodes or any pattern"):
        pat = st.text_input("Regex", r"\b\d{8,14}\b")
        try:
            found = list(dict.fromkeys(re.findall(pat, text)))
        except re.error as e:
            st.error(f"Bad pattern: {e}")
            found = []
        if found:
            st.caption(f"{len(found)} unique matches")
            st.code("\n".join(found), language=None)
        else:
            st.caption("No matches.")
