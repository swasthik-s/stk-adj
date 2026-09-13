"""
Negative Stock Adjustment Tool - Shams Al Madina
Streamlit app.  Run locally:  streamlit run app.py
"""

import io
import math
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Negative Stock Tool", page_icon="📦", layout="wide")

# ==================== house settings ====================
COMPANY = "AL MADINA HYPERMARKET"
BRANCH = "SHAMS AL MADINA HYPERMARKET LLC"
COST_DP = 7
PAIRS_PER_FILE = 7
BLANK_PAD = True          # pad each sheet out to PAIRS_PER_FILE ruled rows

THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GREY = PatternFill("solid", fgColor="BFBFBF")
DARK = PatternFill("solid", fgColor="808080")
WHITE_BOLD = Font(bold=True, color="FFFFFF")
CENTER = Alignment(horizontal="center", vertical="center")
HEADERS = ["SL NO", "OUTER BARCODE", "SINGLE BARCODE", "DESCRIPTION",
           "UNIT", "QTY", "COST", "VALUE"]
WIDTHS = [8, 20, 20, 52, 9, 11, 12, 13]
RED = PatternFill("solid", fgColor="C00000")
RED_FONT = Font(bold=True, color="C00000")
TEXT_FMT = "@"

NOISE = re.compile(r"[^A-Z0-9 ]+")
MULT = re.compile(
    r"\b(?:X\s?(\d{1,3})|(\d{1,3})\s?X|(\d{1,3})\s?PCS|TWIN|COMBO|SET|PK|PACK"
    r"|OFFER|OFR|SPCL|OTR|CTN|CARTON|BOX|BUNDLE|DZN|DOZEN)\b",
    re.I,
)
SIZE = re.compile(r"\b\d+(?:\.\d+)?\s?(?:GM|G|KG|ML|LTR|L)\b", re.I)


# ==================== parsing ====================
@st.cache_data(show_spinner=False)
def load_master(raw: bytes) -> pd.DataFrame:
    want = ["Item Barcode", "Item No", "Item Name", "Stock", "Cost",
            "Net MRP", "Is Active", "Category", "Group", "Brand"]
    head = pd.read_csv(io.BytesIO(raw), dtype=str, encoding="latin-1", nrows=0)
    use = [c for c in want if c in head.columns]
    m = pd.read_csv(io.BytesIO(raw), dtype=str, encoding="latin-1", usecols=use)
    m["Item Barcode"] = m["Item Barcode"].astype(str).str.strip()
    for src, dst in (("Stock", "stock"), ("Cost", "cost"), ("Net MRP", "mrp")):
        m[dst] = (
            pd.to_numeric(m[src].astype(str).str.replace(",", "", regex=False),
                          errors="coerce").fillna(0)
            if src in m.columns else 0.0
        )
    if "Is Active" not in m.columns:
        m["Is Active"] = "Checked"
    return m.drop_duplicates("Item Barcode").reset_index(drop=True)


@st.cache_data(show_spinner=False)
def inspect_negatives(raw: bytes):
    """Return every sheet name and a preview grid, so the header row can be
    found automatically or picked by hand."""
    xls = pd.ExcelFile(io.BytesIO(raw))
    return xls.sheet_names


CODE_KEYS = ["itemcode", "item code", "barcode", "code", "single barcode",
             "itembarcode", "item barcode", "sku"]
NAME_KEYS = ["item name", "itemname", "description", "item description",
             "particulars", "product", "name"]
QTY_KEYS = ["quantity", "qty", "stock qty", "stockqty", "balance qty",
            "closing qty", "stock", "balance"]
VAL_KEYS = ["stock value", "value", "stockvalue", "amount", "net value",
            "total value"]
SP_KEYS = ["selling price", "sellingprice", "sale price", "mrp", "retail",
           "price"]
CST_KEYS = ["cost", "cost price", "costprice", "unit cost", "avg cost"]


def _norm(x):
    return re.sub(r"[^a-z0-9 ]", "", str(x).strip().lower())


def _match(cols, keys):
    n = [_norm(c) for c in cols]
    for k in keys:                       # exact first
        if k in n:
            return cols[n.index(k)]
    for k in keys:                       # then contains
        for i, c in enumerate(n):
            if k in c:
                return cols[i]
    return None


@st.cache_data(show_spinner=False)
def detect_header(raw: bytes, sheet):
    """Scan the first 40 rows for the row that looks like a header."""
    grid = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=None,
                         dtype=str, nrows=40)
    best, best_score = None, 0
    for i in range(len(grid)):
        cells = [c for c in grid.iloc[i].tolist() if str(c) != "nan"]
        if len(cells) < 3:
            continue
        score = sum(bool(_match(cells, k)) for k in (CODE_KEYS, NAME_KEYS, QTY_KEYS))
        if score > best_score:
            best, best_score = i, score
    return best, best_score, grid


@st.cache_data(show_spinner=False)
def load_negatives(raw: bytes, sheet=0, header_row=None, mapping=None) -> pd.DataFrame:
    if header_row is None:
        header_row, score, _ = detect_header(raw, sheet)
        if header_row is None or score < 3:
            raise ValueError("HEADER_NOT_FOUND")
    t = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=header_row,
                      dtype=str).dropna(axis=1, how="all")
    cols = list(t.columns)
    mapping = mapping or {}
    c_code = mapping.get("code") or _match(cols, CODE_KEYS)
    c_name = mapping.get("name") or _match(cols, NAME_KEYS)
    c_qty = mapping.get("qty") or _match(cols, QTY_KEYS)
    if not (c_code and c_name and c_qty):
        raise ValueError(f"MISSING_COLUMNS:{cols}")
    c_val = mapping.get("val") or _match(cols, VAL_KEYS)
    c_sp = mapping.get("sp") or _match(cols, SP_KEYS)
    c_cost = mapping.get("cost") or _match(cols, CST_KEYS)

    def num(col):
        if col and col in t.columns:
            return pd.to_numeric(
                t[col].astype(str).str.replace(",", "", regex=False)
                .str.replace("\u2212", "-", regex=False), errors="coerce")
        return pd.Series([float("nan")] * len(t), index=t.index)

    t["qty"] = num(c_qty)
    t["val"] = num(c_val)
    t["sp"] = num(c_sp)
    t["cost"] = num(c_cost)
    if t["val"].isna().all():
        t["val"] = t["qty"] * t["cost"]

    grp = cat = None
    rows = []
    for r in t.to_dict("records"):
        code = str(r.get(c_code, "")).strip()
        name = str(r.get(c_name, "")).strip()
        if name in ("", "nan", "None"):
            if code in ("Stock", "Non Stock"):
                grp = code
            elif code not in ("", "nan", "None"):
                cat = code
            continue
        r["group"], r["category"] = grp, (cat or "UNCATEGORISED")
        r["Item Name"] = name
        r["bc"] = code
        rows.append(r)
    d = pd.DataFrame(rows)
    if d.empty:
        raise ValueError("NO_ROWS")
    d = d[d["qty"] < 0].reset_index(drop=True)
    if d.empty:
        raise ValueError("NO_NEGATIVES")
    d["val"] = d["val"].fillna(0)
    return d


# ==================== matching ====================
def toks(s, strip_mult=True):
    s = str(s).upper()
    sizes = {x.replace(" ", "") for x in SIZE.findall(s)}
    core = MULT.sub(" ", s) if strip_mult else s
    return {w for w in NOISE.sub(" ", core).split() if len(w) > 1}, sizes


def mult_of(s):
    s = str(s).upper()
    best, found = 1, False
    for mt in MULT.finditer(s):
        found = True
        for g in mt.groups():
            if g and 2 <= int(g) <= 200:
                best = max(best, int(g))
    return (best if best > 1 else (2 if found else 1)), found


@st.cache_data(show_spinner=False)
def find_breaks(m: pd.DataFrame, d: pd.DataFrame, threshold: float) -> pd.DataFrame:
    mm = m.copy()
    mm[["mult", "isml"]] = pd.DataFrame(
        mm["Item Name"].map(mult_of).tolist(), index=mm.index
    )
    par = mm[(mm["stock"] > 0) & mm["isml"] & (mm["mult"] > 1)]
    P, inv = [], defaultdict(list)
    for r in par[["Item Barcode", "Item Name", "stock", "mult", "cost"]].to_dict("records"):
        t, sz = toks(r["Item Name"])
        P.append((r["Item Barcode"], r["Item Name"], r["stock"], r["mult"],
                  t, sz, r["cost"]))
        for w in t:
            inv[w].append(len(P) - 1)

    out = []
    for r in d.to_dict("records"):
        ct, cs = toks(r["Item Name"])
        if not ct:
            continue
        best, bs, seen = None, 0.0, set()
        for w in ct:
            for i in inv.get(w, []):
                if i in seen:
                    continue
                seen.add(i)
                p = P[i]
                if p[0] == r["bc"]:
                    continue
                if cs and p[5] and not (cs & p[5]):
                    continue
                ov = len(ct & p[4]) / len(ct)
                if ov > bs:
                    bs, best = ov, p
        if best and bs >= threshold:
            need = math.ceil(abs(r["qty"]) / best[3])
            out.append(dict(
                neg_bc=r["bc"], neg_desc=r["Item Name"],
                category=r["category"], neg_qty=r["qty"], neg_val=r["val"],
                par_bc=best[0], par_desc=best[1], par_stock=best[2],
                par_cost=best[6], conv=best[3], outers_needed=need,
                covered=need <= best[2], score=round(bs, 2), kind="Bundle break",
            ))
    return pd.DataFrame(out)


@st.cache_data(show_spinner=False)
def find_swaps(m: pd.DataFrame, d: pd.DataFrame, lo: float, hi: float,
               price_tol: float) -> pd.DataFrame:
    mm = m.copy()
    mm["isml"] = mm["Item Name"].astype(str).str.upper().str.contains(MULT.pattern, regex=True, na=False)
    pos = mm[(mm["stock"] > 0) & (~mm["isml"])]
    P, inv = [], defaultdict(list)
    for r in pos[["Item Barcode", "Item Name", "stock", "mrp", "cost"]].to_dict("records"):
        t, sz = toks(r["Item Name"])
        P.append((r["Item Barcode"], r["Item Name"], r["stock"], r["mrp"],
                  t, sz, r["cost"]))
        for w in t:
            inv[w].append(len(P) - 1)

    out = []
    for r in d.to_dict("records"):
        ct, cs = toks(r["Item Name"])
        if len(ct) < 3:
            continue
        best, bs, seen = None, 0.0, set()
        for w in ct:
            for i in inv.get(w, []):
                if i in seen:
                    continue
                seen.add(i)
                p = P[i]
                if p[0] == r["bc"]:
                    continue
                if cs and p[5] and not (cs & p[5]):
                    continue
                ov = len(ct & p[4]) / max(len(ct | p[4]), 1)
                if ov < lo or ov > hi:
                    continue
                sp = r.get("sp")
                if sp and p[3] and abs(sp - p[3]) / max(sp, 0.01) > price_tol:
                    continue
                if p[2] < abs(r["qty"]):
                    continue
                if ov > bs:
                    bs, best = ov, p
        if best:
            out.append(dict(
                neg_bc=r["bc"], neg_desc=r["Item Name"], category=r["category"],
                neg_qty=r["qty"], neg_val=r["val"], par_bc=best[0], par_desc=best[1],
                par_stock=best[2], par_cost=best[6], conv=1,
                outers_needed=abs(r["qty"]), covered=True,
                score=round(bs, 2), kind="Wrong sale",
            ))
    return pd.DataFrame(out)


# ==================== sheet writer ====================
@dataclass
class Line:
    par_bc: str; par_desc: str; par_unit: str; par_qty: float; par_cost: float
    sng_bc: str; sng_desc: str; sng_unit: str; conv: float

    def rows(self):
        oq = -abs(float(self.par_qty))
        sq = round(abs(oq) * self.conv, 3)
        sc = round(float(self.par_cost) / self.conv, COST_DP)
        ov = round(oq * float(self.par_cost), 2)
        sv = round(sq * sc, 2)
        if round(ov + sv, 2) != 0:
            sv = round(-ov, 2)
        return [(self.par_bc, self.par_desc, self.par_unit, oq, self.par_cost, ov),
                (self.sng_bc, self.sng_desc, self.sng_unit, sq, sc, sv)]


def build_sheet(lines, remarks, date_str, prepared, checked, verified) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "ADJUSTMENT"
    ncol = len(HEADERS); last = get_column_letter(ncol)
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(f"A1:{last}1"); ws["A1"] = COMPANY
    ws["A1"].font = Font(size=14, bold=True, color="C00000")
    ws["A1"].alignment = CENTER
    ws.merge_cells(f"A2:{last}2"); ws["A2"] = "ADJUSTMENT"
    ws["A2"].font = Font(size=14, bold=True, color="C00000")
    ws["A2"].alignment = CENTER

    ws["A3"] = "BRANCH"; ws["A3"].fill = RED; ws["A3"].font = WHITE_BOLD
    ws["A3"].alignment = CENTER
    ws.merge_cells(f"B3:{last}3"); ws["B3"] = BRANCH
    ws["B3"].font = Font(size=12, bold=True)

    ws["A4"] = "Remarks"; ws["A4"].fill = RED; ws["A4"].font = WHITE_BOLD
    ws["A4"].alignment = CENTER
    ws.merge_cells("B4:F4"); ws["B4"] = remarks; ws["B4"].font = Font(bold=True)
    ws.merge_cells(f"G4:{last}4"); ws["G4"] = f"Date : {date_str}"
    ws["G4"].font = Font(bold=True)
    ws["G4"].alignment = Alignment(horizontal="right")
    for c in range(1, ncol + 1):
        for r in (3, 4):
            ws.cell(row=r, column=c).border = BOX

    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=6, column=c, value=h)
        cell.fill = RED; cell.font = WHITE_BOLD
        cell.alignment = CENTER; cell.border = BOX

    row = 7; first = row

    def rule(r):
        for c in range(1, ncol + 1):
            ws.cell(row=r, column=c).border = BOX
        ws.cell(row=r, column=2).number_format = TEXT_FMT
        ws.cell(row=r, column=3).number_format = TEXT_FMT
        ws.cell(row=r, column=5).alignment = CENTER
        ws.cell(row=r, column=6).number_format = "0.###"
        ws.cell(row=r, column=7).number_format = "0.#######"
        ws.cell(row=r, column=8).number_format = "0.00"

    sl = 0
    for ln in lines:
        sl += 1
        start = row
        (obc, odesc, ounit, oqty, ocost, oval), \
            (sbc, sdesc, sunit, sqty, scost, sval) = ln.rows()
        # outer row - barcode in the OUTER column only
        ws.cell(row=row, column=2, value=str(obc))
        ws.cell(row=row, column=4, value=odesc)
        ws.cell(row=row, column=5, value="OFR" if ln.conv != 1 else ounit)
        ws.cell(row=row, column=6, value=oqty)
        ws.cell(row=row, column=7, value=ocost)
        ws.cell(row=row, column=8, value=oval)
        rule(row); row += 1
        # single row - barcode in the SINGLE column only
        ws.cell(row=row, column=3, value=str(sbc))
        ws.cell(row=row, column=4, value=sdesc)
        ws.cell(row=row, column=5, value=sunit)
        ws.cell(row=row, column=6, value=sqty)
        ws.cell(row=row, column=7, value=scost)
        ws.cell(row=row, column=8, value=sval)
        rule(row); row += 1
        ws.merge_cells(start_row=start, start_column=1, end_row=row - 1, end_column=1)
        c = ws.cell(row=start, column=1, value=sl)
        c.alignment = CENTER; c.border = BOX

    # pad out with ruled empty pairs so the printed sheet always looks the same
    if BLANK_PAD:
        while sl < PAIRS_PER_FILE:
            sl += 1
            start = row
            for _ in range(2):
                ws.cell(row=row, column=8,
                        value=f"=IF(F{row}=\"\",0,ROUND(F{row}*G{row},2))")
                rule(row); row += 1
            ws.merge_cells(start_row=start, start_column=1,
                           end_row=row - 1, end_column=1)
            c = ws.cell(row=start, column=1, value=sl)
            c.alignment = CENTER; c.border = BOX

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    t = ws.cell(row=row, column=1, value="TOTAL")
    t.font = Font(bold=True); t.alignment = Alignment(horizontal="right")
    tv = ws.cell(row=row, column=8, value=f"=ROUND(SUM(H{first}:H{row-1}),2)")
    tv.font = Font(bold=True); tv.number_format = "0.00"; tv.border = BOX
    for c in range(1, ncol + 1):
        ws.cell(row=row, column=c).border = BOX

    row += 1
    ws.cell(row=row, column=1, value="Prepared By").font = Font(bold=True)
    ws.cell(row=row, column=4, value="Checked By").font = Font(bold=True)
    ws.cell(row=row, column=7, value="Verified By").font = Font(bold=True)
    ws.cell(row=row + 1, column=1, value=prepared).font = Font(bold=True)
    ws.cell(row=row + 1, column=4, value=checked).font = Font(bold=True)
    ws.cell(row=row + 1, column=7, value=verified).font = Font(bold=True)
    ws.row_dimensions[row + 2].height = 50

    ws.print_area = f"A1:{last}{row+2}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = "A7"

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def validate(row, master_idx, neg_map):
    out = []
    src = master_idx.get(row["par_bc"])
    if src is None:
        return ["source barcode not in masterlist"]
    qty_out = abs(row["outers_needed"])
    if src["stock"] < qty_out:
        out.append(f"source stock {src['stock']:.3f} < {qty_out:.3f} — would go negative")
    neg = neg_map.get(row["neg_bc"])
    if neg is None:
        out.append("target not in the negative report")
    else:
        gap = round(qty_out * row["conv"] - abs(neg), 3)
        if gap > 0.001:
            out.append(f"over-clears by {gap:g}")
        elif gap < -0.001:
            out.append(f"under-clears by {abs(gap):g}")
    if str(src.get("Is Active", "")).strip() == "Unchecked":
        out.append("source item inactive")
    return out


# ==================== UI ====================
st.title("📦 Negative Stock Adjustment Tool")

with st.sidebar:
    st.header("1. Data")
    f_master = st.file_uploader("Masterlist (CSV)", type=["csv"])
    f_neg = st.file_uploader("Negative stock report (XLSX)", type=["xlsx", "xls"])
    st.header("2. Sheet details")
    adj_date = st.text_input("Date", "11-09-26")
    prepared = st.text_input("Prepared by", "SWASTHIK")
    checked = st.text_input("Checked by", "IRSHAD")
    verified = st.text_input("Verified by", "THALLATH")
    st.header("3. Matching")
    br_thresh = st.slider("Bundle break strictness", 0.70, 1.00, 0.80, 0.01,
                          help="Higher = fewer but safer matches")
    sw_lo, sw_hi = st.slider("Wrong sale similarity window", 0.40, 1.00,
                             (0.55, 0.95), 0.05,
                             help="Similar but not identical")
    price_tol = st.slider("Wrong sale price tolerance", 0.05, 0.60, 0.25, 0.05)

if not (f_master and f_neg):
    st.info("Upload the masterlist and the negative stock report in the sidebar to start.")
    st.stop()

try:
    master = load_master(f_master.getvalue())
except Exception as e:
    st.error(f"Could not read the masterlist: {e}")
    st.stop()

raw_neg = f_neg.getvalue()
sheets = inspect_negatives(raw_neg)
sheet = sheets[0] if len(sheets) == 1 else st.selectbox(
    "Which sheet holds the negative stock?", sheets)
auto_row, auto_score, grid = detect_header(raw_neg, sheet)

neg = None
try:
    neg = load_negatives(raw_neg, sheet)
except Exception as err:
    st.warning("I could not read this export automatically — map the columns below.")
    with st.expander("First 15 rows of the file", expanded=True):
        st.dataframe(grid.head(15), use_container_width=True)
    hr = st.number_input(
        "Which row holds the column headings? (row 1 is the first row)",
        min_value=1, max_value=40,
        value=int(auto_row) + 1 if auto_row is not None else 1)
    try:
        cols = list(pd.read_excel(io.BytesIO(raw_neg), sheet_name=sheet,
                                  header=int(hr) - 1, dtype=str, nrows=5)
                    .dropna(axis=1, how="all").columns)
    except Exception:
        cols = []
    if cols:
        none = "— none —"
        c1, c2, c3 = st.columns(3)
        m_code = c1.selectbox("Barcode / item code", cols,
                              index=cols.index(_match(cols, CODE_KEYS))
                              if _match(cols, CODE_KEYS) else 0)
        m_name = c2.selectbox("Description", cols,
                              index=cols.index(_match(cols, NAME_KEYS))
                              if _match(cols, NAME_KEYS) else 0)
        m_qty = c3.selectbox("Quantity", cols,
                             index=cols.index(_match(cols, QTY_KEYS))
                             if _match(cols, QTY_KEYS) else 0)
        c4, c5, c6 = st.columns(3)
        opt = [none] + cols
        def pick(col, label, keys):
            g = _match(cols, keys)
            return col.selectbox(label, opt, index=opt.index(g) if g else 0)
        m_val = pick(c4, "Stock value (optional)", VAL_KEYS)
        m_cost = pick(c5, "Cost (optional)", CST_KEYS)
        m_sp = pick(c6, "Selling price (optional)", SP_KEYS)
        if st.button("Load with this mapping", type="primary"):
            try:
                neg = load_negatives(raw_neg, sheet, int(hr) - 1, {
                    "code": m_code, "name": m_name, "qty": m_qty,
                    "val": None if m_val == none else m_val,
                    "cost": None if m_cost == none else m_cost,
                    "sp": None if m_sp == none else m_sp,
                })
                st.session_state["neg_map_ok"] = True
            except Exception as e2:
                st.error(f"Still could not read it: {e2}")
    if neg is None:
        st.stop()

master_idx = master.set_index("Item Barcode").to_dict("index")
neg_map = dict(zip(neg["bc"], neg["qty"]))

tab1, tab2, tab3, tab4 = st.tabs(
    ["Overview", "Candidates", "Build sheets", "Manual pair"]
)

# ---------- Overview ----------
with tab1:
    c1, c2, c3 = st.columns(3)
    c1.metric("Negative lines", f"{len(neg):,}")
    c2.metric("Negative value", f"{neg['val'].sum():,.0f} AED")
    c3.metric("Master items", f"{len(master):,}")

    by_cat = (neg.groupby(["group", "category"])
              .agg(Lines=("val", "size"), Qty=("qty", "sum"), Value=("val", "sum"))
              .round(2).sort_values("Value").reset_index())
    st.subheader("By category")
    st.dataframe(by_cat, use_container_width=True, hide_index=True)

    st.subheader("Concentration")
    s = neg.sort_values("val")
    for n in (20, 50, 100, 200):
        if n <= len(s):
            v = s["val"].head(n).sum()
            st.write(f"Top {n} lines carry **{v:,.0f} AED** "
                     f"({v/neg['val'].sum()*100:.0f}% of the total)")

    st.subheader("Not in masterlist")
    miss = neg[~neg["bc"].isin(master["Item Barcode"])]
    st.write(f"{len(miss)} lines worth {miss['val'].sum():,.0f} AED — dead codes")
    if len(miss):
        st.dataframe(miss[["bc", "Item Name", "qty", "val"]],
                     use_container_width=True, hide_index=True)

# ---------- Candidates ----------
with tab2:
    cats = sorted(neg["category"].dropna().unique())
    default = [c for c in cats
               if c in ("GROCERY FOOD", "GROCERY NON FOOD",
                        "HEALTH AND BEAUTY", "GARMENTS")]
    pick_cats = st.multiselect("Categories", cats, default=default or cats)
    c1, c2 = st.columns(2)
    do_break = c1.checkbox("Find bundle breaks", True)
    do_swap = c2.checkbox("Find wrong sales", True)
    no_break_cats = st.multiselect(
        "No bundle breaks in these categories", cats,
        default=[c for c in cats if c == "GARMENTS"],
        help="Garments are wrong-sale only",
    )

    if st.button("Run matching", type="primary"):
        d = neg[neg["category"].isin(pick_cats)]
        parts = []
        with st.spinner("Matching…"):
            if do_break:
                b = find_breaks(master, d[~d["category"].isin(no_break_cats)], br_thresh)
                if len(b):
                    parts.append(b[b["covered"]])
            if do_swap:
                sw = find_swaps(master, d, sw_lo, sw_hi, price_tol)
                if len(sw):
                    parts.append(sw)
        if parts:
            cand = pd.concat(parts, ignore_index=True)
            cand = cand.drop_duplicates("neg_bc", keep="first")
            cand["problems"] = [
                "; ".join(validate(r, master_idx, neg_map))
                for r in cand.to_dict("records")
            ]
            cand["use"] = cand["problems"].eq("")
            st.session_state["cand"] = cand
        else:
            st.session_state["cand"] = pd.DataFrame()

    cand = st.session_state.get("cand")
    if cand is not None and len(cand):
        c1, c2, c3 = st.columns(3)
        c1.metric("Candidates", len(cand))
        c2.metric("Value covered", f"{cand['neg_val'].sum():,.0f} AED")
        c3.metric("Clean", int(cand["problems"].eq("").sum()))
        st.dataframe(
            cand.groupby(["category", "kind"])
            .agg(Pairs=("neg_val", "size"), Value=("neg_val", "sum"))
            .round(0).reset_index(),
            use_container_width=True, hide_index=True,
        )
        st.caption("Untick anything you do not want. Check the physical shelf first.")
        st.session_state["cand"] = st.data_editor(
            cand, use_container_width=True, hide_index=True,
            column_config={"use": st.column_config.CheckboxColumn("Use", width="small")},
            disabled=[c for c in cand.columns if c != "use"],
        )
    elif cand is not None:
        st.warning("No candidates found. Try loosening the sliders in the sidebar.")

# ---------- Build ----------
with tab3:
    cand = st.session_state.get("cand")
    manual = st.session_state.get("manual", [])
    if (cand is None or not len(cand)) and not manual:
        st.info("Run matching first, or add a pair by hand.")
    else:
        sel = cand[cand["use"]].to_dict("records") if cand is not None and len(cand) else []
        sel += manual
        st.write(f"**{len(sel)} pairs selected** — "
                 f"{sum(abs(r['neg_val']) for r in sel):,.0f} AED, "
                 f"{math.ceil(len(sel)/PAIRS_PER_FILE)} files")
        remarks = st.text_input(
            "Remarks",
            "OUTER BREAK FOR NEGATIVE STOCK",
        )
        if st.button("Generate adjustment sheets", type="primary") and sel:
            lines = [
                Line(r["par_bc"], r["par_desc"],
                     "KG" if r["conv"] > 1 else "PCS",
                     r["outers_needed"], r["par_cost"],
                     r["neg_bc"], r["neg_desc"], "PCS", r["conv"])
                for r in sel
            ]
            work = []
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                for i in range(0, len(lines), PAIRS_PER_FILE):
                    chunk = lines[i:i + PAIRS_PER_FILE]
                    name = f"ADJ_{i//PAIRS_PER_FILE + 1:03d}.xlsx"
                    z.writestr(name, build_sheet(chunk, remarks, adj_date,
                                                 prepared, checked, verified))
                    for ln, r in zip(chunk, sel[i:i + PAIRS_PER_FILE]):
                        a, b = ln.rows()
                        work.append({
                            "FILE": name, "OUTER BARCODE": a[0],
                            "OUTER DESCRIPTION": a[1], "SINGLE BARCODE": b[0],
                            "SINGLE DESCRIPTION": b[1], "CONV": ln.conv,
                            "OUTER QTY": a[3], "NEW SINGLE QTY": b[3],
                            "OLD NEGATIVE QTY": neg_map.get(b[0], 0),
                            "COST OUTER": a[4], "COST SINGLE": b[4],
                            "OUTER VALUE": a[5], "SINGLE VALUE": b[5],
                            "DIFFERENCE": round(a[5] + b[5], 2),
                            "REASON": r.get("kind", ""),
                        })
                wdf = pd.DataFrame(work)
                for c in ("OUTER BARCODE", "SINGLE BARCODE"):
                    wdf[c] = wdf[c].astype(str)
                wb = io.BytesIO()
                with pd.ExcelWriter(wb, engine="openpyxl") as xw:
                    wdf.to_excel(xw, sheet_name="WORKING", index=False)
                    sh = xw.sheets["WORKING"]
                    for col in ("C", "D"):
                        for cell in sh[col]:
                            cell.number_format = "@"
                    sh.column_dimensions["C"].width = 20
                    sh.column_dimensions["D"].width = 20
                z.writestr("WORKING_ALL.xlsx", wb.getvalue())

            bad = int((wdf["DIFFERENCE"].abs() > 0.01).sum())
            if bad:
                st.error(f"{bad} pairs do not net to zero — do not post these")
            else:
                st.success(f"{len(lines)} pairs, every file totals 0.00")
            st.download_button("⬇ Download all sheets (zip)", zbuf.getvalue(),
                               "adjustment_files.zip", "application/zip")
            st.dataframe(wdf, use_container_width=True, hide_index=True)

# ---------- Manual ----------
with tab4:
    st.caption("For pairs the matcher cannot find — repacking, or anything you know by eye.")
    with st.form("manual"):
        c1, c2 = st.columns(2)
        pbc = c1.text_input("Outer / source barcode")
        sbc = c2.text_input("Single / target barcode")
        c3, c4 = st.columns(2)
        oqty = c3.number_input("Outer qty (units leaving source)", min_value=0.0,
                               value=1.0, step=1.0)
        conv = c4.number_input("Conversion (singles per outer)", min_value=0.001,
                               value=1.0, step=0.5,
                               help="400GM packs from 1 KG = 2.5 · 18KG bag to loose KG = 18 · wrong sale = 1")
        if st.form_submit_button("Add pair"):
            src, dst = master_idx.get(pbc.strip()), master_idx.get(sbc.strip())
            if not src:
                st.error("Source barcode not in masterlist")
            elif not dst:
                st.error("Target barcode not in masterlist")
            else:
                st.session_state.setdefault("manual", []).append(dict(
                    par_bc=pbc.strip(), par_desc=src["Item Name"],
                    par_cost=src["cost"], par_stock=src["stock"],
                    neg_bc=sbc.strip(), neg_desc=dst["Item Name"],
                    neg_qty=neg_map.get(sbc.strip(), 0),
                    neg_val=0.0, conv=conv, outers_needed=oqty,
                    kind="Manual", category="MANUAL",
                ))
                st.success(f"Added — {src['Item Name']} → {dst['Item Name']}")

    if st.session_state.get("manual"):
        st.dataframe(pd.DataFrame(st.session_state["manual"]),
                     use_container_width=True, hide_index=True)
        if st.button("Clear manual pairs"):
            st.session_state["manual"] = []
            st.rerun()
