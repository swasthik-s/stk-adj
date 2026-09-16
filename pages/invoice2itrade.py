"""
Supplier Invoice → Excel / iTrade txt
-------------------------------------
Reads a supplier tax invoice PDF, pulls out the line items across every
page, checks the arithmetic against the invoice's own totals, and gives
you a clean Excel plus the flat import file:

    SML,BARCODE,UNIT PRICE,QTY

Built around the Al Madina Logistic Centre layout but driven by column
names, so other suppliers with a BARCODE / QTY / UNIT PRICE table work too.
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

st.title("🧾 Supplier invoice → Excel / iTrade txt")

if not PDF_OK:
    st.error(f"pdfplumber is not installed on the server. {PDF_ERR}")
    st.code("pdfplumber>=0.11", language=None)
    st.stop()



XL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Column names are matched loosely so other suppliers' layouts still work.
ALIAS = {
    "BARCODE": ["BARCODE", "ITEM CODE", "ITEMCODE", "CODE", "EAN"],
    "DESCRIPTION": ["DESCRIPTION", "ITEM NAME", "PARTICULARS", "ITEM"],
    "QTY": ["QTY", "QUANTITY", "RECD QTY", "REC QTY"],
    "PRICE": ["UNIT PRICE", "RATE", "PRICE", "COST", "UNIT COST"],
    "AMOUNT": ["AMOUNT", "VALUE", "NET AMOUNT", "TOTAL"],
    "UNIT": ["UNIT", "UOM"],
    "SL": ["SL.NO", "SL NO", "S.NO", "SR.NO", "SR", "#"],
    "VAT": ["VAT AMT", "VAT AMOUNT", "TAX AMT"],
    "AFTER": ["AFTER VAT AMT", "AFTER VAT", "GROSS"],
}


def _norm(h):
    return re.sub(r"\s+", " ", str(h or "")).strip().upper()


def _num(x):
    try:
        return float(str(x).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _find(hdr, keys):
    for k in keys:                       # exact first
        if k in hdr:
            return hdr.index(k)
    for i, h in enumerate(hdr):          # then contains
        for k in keys:
            if h and k in h:
                return i
    return None


@st.cache_data(show_spinner=False)
def parse_invoice(raw: bytes):
    """Returns (line items, invoice header info, full text)."""
    rows = []
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        full = "\n".join((p.extract_text() or "") for p in pdf.pages)
        for pno, page in enumerate(pdf.pages, start=1):
            for tb in page.extract_tables() or []:
                if not tb:
                    continue
                # the header is not always the first row — some pages merge
                # the party block into the same table
                hrow, hdr = None, None
                for i, r in enumerate(tb):
                    h = [_norm(c) for c in r]
                    if (_find(h, ALIAS["BARCODE"]) is not None
                            and _find(h, ALIAS["QTY"]) is not None):
                        hrow, hdr = i, h
                        break
                if hrow is None:
                    continue
                ix = {k: _find(hdr, v) for k, v in ALIAS.items()}
                for r in tb[hrow + 1:]:
                    code = (str(r[ix["BARCODE"]] or "").strip()
                            if ix["BARCODE"] is not None else "")
                    if not re.fullmatch(r"\d{6,14}", code):
                        continue

                    def cell(key, as_num=False):
                        i = ix.get(key)
                        if i is None or i >= len(r):
                            return None if as_num else ""
                        return _num(r[i]) if as_num else str(r[i] or "").strip()

                    rows.append({
                        "Page": pno, "SL": cell("SL", True), "Barcode": code,
                        "Description": cell("DESCRIPTION"),
                        "Unit": cell("UNIT"), "Qty": cell("QTY", True),
                        "UnitPrice": cell("PRICE", True),
                        "Amount": cell("AMOUNT", True),
                        "VatAmt": cell("VAT", True),
                        "AfterVat": cell("AFTER", True),
                    })

    df = pd.DataFrame(rows)
    if len(df):
        df = df.drop_duplicates(subset=["SL", "Barcode"]).reset_index(drop=True)

    def grab(pat):
        m = re.search(pat, full, re.I)
        return m.group(1).strip() if m else None

    head = {
        "Invoice no": grab(r"INV NO\s*:?\s*([A-Z0-9\-/]+)"),
        "Date": grab(r"Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})"),
        "Supplier TRN": grab(r"TRN\s*:?\s*(\d{10,20})"),
        "Total": _num(grab(r"\bTotal\s+([\d,]+\.\d\d)")),
        "VAT": _num(grab(r"\bVAT\s+([\d,]+\.\d\d)")),
        "Net Total": _num(grab(r"Net Total\s+([\d,]+\.\d\d)")),
    }
    return df, head, full


up = st.file_uploader("Supplier invoice (PDF)", type=["pdf"])
if up is None:
    st.info("Upload a supplier tax invoice. It reads every page, keeps only "
            "the line items, and checks the total against the invoice's own "
            "figure before you import anything.")
    st.stop()

with st.spinner("Reading the invoice…"):
    try:
        df, head, full = parse_invoice(up.getvalue())
    except Exception as e:
        st.error(f"Could not read that file: {type(e).__name__}: {e}")
        st.stop()

if not len(df):
    st.error("No line items found. Either the PDF has no text layer (a scan), "
             "or the table headings differ from the ones this page looks for "
             "— BARCODE and QTY at minimum.")
    with st.expander("What the file does contain"):
        st.text(full[:3000] or "(no text at all — it is a scan)")
    st.stop()

# ---------------------------------------------------------------- header
h1, h2, h3, h4 = st.columns(4)
h1.metric("Invoice", head.get("Invoice no") or "—")
h2.metric("Date", head.get("Date") or "—")
h3.metric("Line items", len(df))
h4.metric("Pages", int(df["Page"].max()))

# ---------------------------------------------------------------- checks
st.subheader("Checks")
c1, c2, c3 = st.columns(3)

sum_amt = float(df["Amount"].sum()) if df["Amount"].notna().any() else None
inv_tot = head.get("Total")
if sum_amt is not None and inv_tot:
    diff = round(sum_amt - inv_tot, 2)
    c1.metric("Sum of lines vs invoice Total", f"{sum_amt:,.2f}",
              f"{diff:+.2f}", delta_color="off")
    if abs(diff) <= 0.05:
        c1.success("Matches")
    else:
        c1.error("Does not match — some lines may be missing")
else:
    c1.info("No invoice Total found to compare against")

sum_gross = float(df["AfterVat"].sum()) if df["AfterVat"].notna().any() else None
if sum_gross is not None and head.get("Net Total"):
    d2 = round(sum_gross - head["Net Total"], 2)
    c2.metric("Sum after VAT vs Net Total", f"{sum_gross:,.2f}",
              f"{d2:+.2f}", delta_color="off")
    (c2.success if abs(d2) <= 0.05 else c2.error)(
        "Matches" if abs(d2) <= 0.05 else "Does not match")

calc = (df["Qty"] * df["UnitPrice"]).round(2)
off = df[(calc - df["Amount"].round(2)).abs() > 0.02]
c3.metric("Rows where qty × price ≠ amount", len(off))
if len(off):
    with c3.expander("Which rows"):
        st.dataframe(off[["SL", "Description", "Qty", "UnitPrice", "Amount"]],
                     use_container_width=True, hide_index=True)
    c3.caption("Usually the supplier rounding the line, not a parsing error. "
               "Check before importing.")

dupes = df[df.duplicated("Barcode", keep=False)]
if len(dupes):
    with st.expander(f"{dupes['Barcode'].nunique()} barcode(s) appear more "
                     f"than once — {len(dupes)} lines"):
        st.dataframe(dupes.sort_values("Barcode"), use_container_width=True,
                     hide_index=True)
        st.caption("Import them as-is and the quantities add up, which is "
                   "usually right for a split delivery.")

# ---------------------------------------------------------------- lines
st.subheader("Line items")
st.dataframe(df.drop(columns=["Page"]), use_container_width=True, height=420,
             hide_index=True,
             column_config={
                 "Barcode": st.column_config.TextColumn(width="medium"),
                 "Description": st.column_config.TextColumn(width="large"),
                 "Qty": st.column_config.NumberColumn(format="%.3f"),
                 "UnitPrice": st.column_config.NumberColumn(format="%.4f"),
                 "Amount": st.column_config.NumberColumn(format="%.2f"),
                 "AfterVat": st.column_config.NumberColumn(format="%.2f")})

# ---------------------------------------------------------------- output
st.subheader("Download")
o1, o2 = st.columns(2)
prefix = o1.text_input("Import file prefix", "SML")
price_from = o2.selectbox(
    "Price to use in the txt", ["Unit price (before VAT)",
                                "Unit price including VAT"],
    help="iTrade normally takes cost before VAT. Use the second option only "
         "if your import expects the gross figure.")


def fmt(x, dp=7):
    return f"{float(x):.{dp}f}".rstrip("0").rstrip(".") or "0"


lines = []
for r in df.itertuples():
    price = r.UnitPrice
    if price is None:
        continue
    if price_from.endswith("VAT") and r.Amount and r.AfterVat and r.Amount:
        price = price * (r.AfterVat / r.Amount)
    lines.append(f"{prefix},{r.Barcode},{fmt(price)},{fmt(r.Qty, 3)}")
txt = ("\n".join(lines) + "\n").encode("ascii", "ignore")

clean = df[["Barcode", "Description", "Unit", "Qty", "UnitPrice", "Amount",
            "VatAmt", "AfterVat"]].copy()
xbuf = io.BytesIO()
with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
    clean.to_excel(xw, sheet_name="LINES", index=False)
    pd.DataFrame([head]).T.reset_index().rename(
        columns={"index": "Field", 0: "Value"}).to_excel(
        xw, sheet_name="INVOICE", index=False)
    sh = xw.sheets["LINES"]
    for cell in sh["A"]:
        cell.number_format = "@"
    sh.column_dimensions["A"].width = 18
    sh.column_dimensions["B"].width = 44

stem = re.sub(r"[^A-Za-z0-9]+", "_",
              (head.get("Invoice no") or up.name.rsplit(".", 1)[0]))[:40]

d1, d2 = st.columns(2)
d1.download_button(f"⬇ Excel — {len(df)} lines", xbuf.getvalue(),
                   f"{stem}.xlsx", XL, use_container_width=True,
                   type="primary")
d2.download_button(f"⬇ Txt — {len(lines)} import lines", txt, f"{stem}.txt",
                   "text/plain", use_container_width=True)

with st.expander("Preview the import file"):
    st.code(txt.decode(), language=None)

with st.expander("Copy barcodes"):
    st.code("\n".join(df["Barcode"].tolist()), language=None)
