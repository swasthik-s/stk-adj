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
PAIRS_PER_FILE = 11       # 11 pairs = 22 rows per sheet
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
SIZE_JOIN_SIZE = re.compile(
    r"\d+(?:\.\d+)?\s?(?:GM|G|KG|ML|LTR|L)\s*(?:[+&]|PLUS|WITH|AND)\s*"
    r"\d+(?:\.\d+)?\s?(?:GM|G|KG|ML|LTR|L)", re.I)
FREE_QTY = re.compile(r"\b(?:\d+\s*)?(?:PCS|GM|G|ML)?\s*FREE\b", re.I)


def is_combo(desc):
    """A combo holds two DIFFERENT items in one pack: 100ML + 50ML.
    Breaking it releases both, so a one-target break would be wrong.

    Deliberately narrow: an ampersand inside a brand name (GLOW&LOVELY,
    REPAIR&REGENERATE) is NOT a combo. Two distinct size tokens is the test."""
    s = str(desc).upper()
    sizes = {x.replace(" ", "") for x in SIZE.findall(s)}
    if len(sizes) >= 2:
        return True
    if SIZE_JOIN_SIZE.search(s):
        return True
    if FREE_QTY.search(s) and len(sizes) >= 1 and re.search(r"[+&]", s):
        return True
    return False


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
def find_breaks(m: pd.DataFrame, d: pd.DataFrame, threshold: float,
                skip_combo: bool = True):
    mm = m.copy()
    mm[["mult", "isml"]] = pd.DataFrame(
        mm["Item Name"].map(mult_of).tolist(), index=mm.index
    )
    own_cost = dict(zip(mm["Item Barcode"], mm["cost"]))
    par = mm[(mm["stock"] > 0) & mm["isml"] & (mm["mult"] > 1)]
    P, inv = [], defaultdict(list)
    for r in par[["Item Barcode", "Item Name", "stock", "mult", "cost"]].to_dict("records"):
        t, sz = toks(r["Item Name"])
        P.append((r["Item Barcode"], r["Item Name"], r["stock"], r["mult"],
                  t, sz, r["cost"]))
        for w in t:
            inv[w].append(len(P) - 1)

    out, combos = [], []
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
            if skip_combo and (is_combo(best[1]) or is_combo(r["Item Name"])):
                combos.append(dict(
                    neg_bc=r["bc"], neg_desc=r["Item Name"], neg_qty=r["qty"],
                    neg_val=r["val"], par_bc=best[0], par_desc=best[1],
                    why="combo pack — breaking it releases more than one item"))
                continue
            need = math.ceil(abs(r["qty"]) / best[3])
            out.append(dict(
                neg_bc=r["bc"], neg_desc=r["Item Name"],
                category=r["category"], neg_qty=r["qty"], neg_val=r["val"],
                par_bc=best[0], par_desc=best[1], par_stock=best[2],
                par_cost=best[6], conv=best[3], outers_needed=need,
                covered=need <= best[2], score=round(bs, 2), kind="Bundle break",
                own_cost=own_cost.get(r["bc"], 0.0),
                derived_cost=round(best[6] / best[3], COST_DP),
                cost_drift_pct=(
                    round((best[6] / best[3] - own_cost.get(r["bc"], 0)) /
                          own_cost[r["bc"]] * 100, 1)
                    if own_cost.get(r["bc"]) else None),
            ))
    return pd.DataFrame(out), pd.DataFrame(combos)


@st.cache_data(show_spinner=False)
def find_swaps(m: pd.DataFrame, d: pd.DataFrame, lo: float, hi: float,
               price_tol: float) -> pd.DataFrame:
    mm = m.copy()
    mm["isml"] = mm["Item Name"].astype(str).str.upper().str.contains(MULT.pattern, regex=True, na=False)
    own_cost = dict(zip(mm["Item Barcode"], mm["cost"]))
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
                own_cost=own_cost.get(r["bc"], 0.0),
                derived_cost=round(best[6], COST_DP),
                cost_drift_pct=(
                    round((best[6] - own_cost.get(r["bc"], 0)) /
                          own_cost[r["bc"]] * 100, 1)
                    if own_cost.get(r["bc"]) else None),
            ))
    return pd.DataFrame(out), pd.DataFrame(combos)


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


def build_sheet(lines, remarks, date_str, prepared, checked, verified,
                live=True) -> bytes:
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
    if live:
        h = ws.cell(row=6, column=10, value="CONV")
        h.fill = RED; h.font = WHITE_BOLD; h.alignment = CENTER
        ws.column_dimensions["J"].width = 8

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
        orow = row
        ws.cell(row=row, column=8,
                value=f"=ROUND(F{orow}*G{orow},2)" if live else oval)
        rule(row); row += 1
        # single row - barcode in the SINGLE column only
        ws.cell(row=row, column=3, value=str(sbc))
        ws.cell(row=row, column=4, value=sdesc)
        ws.cell(row=row, column=5, value=sunit)
        cv = ln.conv
        if live:
            # conv lives in its own cell (column J, outside the print area) so a
            # wrong pack size is a number to fix, not a formula to rewrite.
            ws.cell(row=orow, column=10, value=cv).number_format = "0.####"
            ws.cell(row=row, column=6, value=f"=ROUND(ABS(F{orow})*J{orow},3)")
            ws.cell(row=row, column=7, value=f"=IF(J{orow}=0,0,ROUND(G{orow}/J{orow},7))")
            ws.cell(row=row, column=8, value=f"=-ROUND(F{orow}*G{orow},2)")
        else:
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


def validate(row, master_idx, neg_map, drift_tol=15.0):
    out = []
    dr = row.get("cost_drift_pct")
    if dr is not None and abs(dr) > drift_tol:
        out.append(f"cost drift {dr:+.0f}% — derived {row.get('derived_cost'):.4g} "
                   f"vs own {row.get('own_cost'):.4g}; likely a wrong pair")
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




# ==================== verification round-trip ====================
VERIFY_COLS = ["ROW", "KEY", "SECTION", "TYPE", "SINGLE BARCODE",
               "NEGATIVE ITEM", "NEG QTY", "OUTER BARCODE", "OUTER ITEM",
               "OUTER STOCK", "CONV", "OUTERS TO BREAK", "NEW SINGLE QTY",
               "COST DRIFT %", "SYSTEM FLAGS", "VERIFIED (Y/N)",
               "ACTUAL SHELF QTY", "STAFF REMARKS"]


def pair_key(r):
    return f"{r['neg_bc']}|{r['par_bc']}"


def make_verification_book(cand: pd.DataFrame) -> bytes:
    """One sheet per section, numbered serially, with a Y/N column for staff."""
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook(); wb.remove(wb.active)
    n = 0
    for section, grp in cand.groupby("category", sort=True):
        ws = wb.create_sheet(str(section)[:28] or "OTHER")
        ws.merge_cells(start_row=1, start_column=1, end_row=1,
                       end_column=len(VERIFY_COLS))
        ws["A1"] = f"{COMPANY} — NEGATIVE STOCK VERIFICATION — {section}"
        ws["A1"].font = Font(size=13, bold=True, color="C00000")
        ws["A1"].alignment = CENTER
        ws.merge_cells(start_row=2, start_column=1, end_row=2,
                       end_column=len(VERIFY_COLS))
        ws["A2"] = ("Check the shelf. Put Y only if the pairing is correct and the "
                    "outer is physically there. Put N or leave blank to reject. "
                    "Do not change any other column.")
        ws["A2"].font = Font(italic=True)
        for c, h in enumerate(VERIFY_COLS, start=1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.fill = RED; cell.font = WHITE_BOLD
            cell.alignment = Alignment("center", "center", wrap_text=True)
            cell.border = BOX
        widths = [7, 30, 18, 13, 18, 40, 10, 18, 40, 12, 7, 12, 13, 11, 30, 14, 14, 26]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[3].height = 34

        r = 4
        for rec in grp.to_dict("records"):
            n += 1
            vals = [n, pair_key(rec), rec.get("category", ""), rec.get("kind", ""),
                    str(rec["neg_bc"]), rec["neg_desc"], rec["neg_qty"],
                    str(rec["par_bc"]), rec["par_desc"], rec.get("par_stock", ""),
                    rec.get("conv", 1), rec.get("outers_needed", ""),
                    round(abs(rec.get("outers_needed", 0)) * rec.get("conv", 1), 3),
                    rec.get("cost_drift_pct", ""), rec.get("problems", ""), "", "", ""]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c, value=v); cell.border = BOX
                if c in (5, 8):
                    cell.number_format = "@"
            ws.cell(row=r, column=16).fill = PatternFill("solid", fgColor="FFF2CC")
            ws.cell(row=r, column=17).fill = PatternFill("solid", fgColor="FFF2CC")
            ws.cell(row=r, column=18).fill = PatternFill("solid", fgColor="FFF2CC")
            r += 1

        dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"P4:P{r-1}")
        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:{get_column_letter(len(VERIFY_COLS))}{r-1}"
        ws.column_dimensions["B"].hidden = True      # KEY, needed on re-upload
        ws.print_area = f"A1:{get_column_letter(len(VERIFY_COLS))}{r-1}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def read_verification(raw: bytes) -> pd.DataFrame:
    """Read back every sheet, keep the rows marked Y."""
    book = pd.read_excel(io.BytesIO(raw), sheet_name=None, header=2, dtype=str)
    out = []
    for name, t in book.items():
        if "KEY" not in t.columns or "VERIFIED (Y/N)" not in t.columns:
            continue
        t = t[t["KEY"].notna()]
        t["_verified"] = (t["VERIFIED (Y/N)"].astype(str).str.strip().str.upper()
                          .isin(["Y", "YES", "OK", "1", "TRUE"]))
        t["_sheet"] = name
        out.append(t[["ROW", "KEY", "_verified", "_sheet",
                      "ACTUAL SHELF QTY", "STAFF REMARKS"]])
    if not out:
        raise ValueError("No verification sheets found — is this the exported file?")
    v = pd.concat(out, ignore_index=True)
    v["ROW"] = pd.to_numeric(v["ROW"], errors="coerce")
    return v.sort_values("ROW")


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
    drift_tol = st.slider("Max cost drift %", 2.0, 60.0, 15.0, 1.0,
                          help="A real break barely moves the item's cost. "
                               "A big drift usually means the pair is wrong.")

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

tab1, tab2, tabV, tab3, tab4 = st.tabs(
    ["Overview", "Candidates", "Verify", "Build sheets", "Manual pair"]
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
@st.fragment
def candidates_tab(neg, master, master_idx, neg_map,
                   br_thresh, sw_lo, sw_hi, price_tol, drift_tol):
    counts = (neg.groupby("category")
              .agg(lines=("val", "size"), val=("val", "sum")))
    cats = sorted(neg["category"].dropna().unique())
    label = {c: f"{c}  ({int(counts.loc[c, 'lines'])} lines, "
                f"{counts.loc[c, 'val']:,.0f} AED)" for c in cats}

    c1, c2 = st.columns([3, 2])
    picked = c1.multiselect(
        "Categories — pick one at a time", [label[c] for c in cats], default=[],
        help="Nothing runs until you choose. One section at a time keeps the "
             "list short and the matches easier to check.")
    pick_cats = [c for c in cats if label[c] in picked]
    mode = c2.selectbox("What to look for",
                        ["Bundle breaks only", "Wrong sales only", "Both"],
                        index=0)
    c3, c4 = st.columns([3, 2])
    no_break_cats = c3.multiselect(
        "No bundle breaks in these categories", cats,
        default=[c for c in cats if c == "GARMENTS"],
        help="Garments are wrong-sale only")
    c4.write("")
    run = c4.button("Run matching", type="primary", use_container_width=True,
                    disabled=not pick_cats)
    if not pick_cats:
        st.info("Choose a category above to start.")

    if run and pick_cats:
        d = neg[neg["category"].isin(pick_cats)]
        parts, combos = [], pd.DataFrame()
        with st.status("Matching…", expanded=False) as status:
            if mode in ("Bundle breaks only", "Both"):
                status.update(label="Looking for bundle breaks…")
                b, combos = find_breaks(
                    master, d[~d["category"].isin(no_break_cats)], br_thresh)
                if len(b):
                    parts.append(b[b["covered"]])
            if mode in ("Wrong sales only", "Both"):
                status.update(label="Looking for wrong sales…")
                sw = find_swaps(master, d, sw_lo, sw_hi, price_tol)
                if len(sw):
                    parts.append(sw)
            status.update(label="Checking costs and stock…")
        if parts:
            cand = pd.concat(parts, ignore_index=True)
            cand = cand.drop_duplicates("neg_bc", keep="first")
            cand["abs_val"] = cand["neg_val"].abs()
            cand = (cand.sort_values("abs_val", ascending=False)
                    .drop(columns="abs_val").reset_index(drop=True))
            probs = [validate(r, master_idx, neg_map, drift_tol)
                     for r in cand.to_dict("records")]
            cand["problems"] = ["; ".join(p) for p in probs]
            # over/under-clear is normal odd-quantity overshoot, not a fault
            cand["blocked"] = [
                any(not (x.startswith("over-clears") or x.startswith("under-clears"))
                    for x in p) for p in probs]
            clean = ~cand["blocked"]
            # rank within the unblocked rows, so you always get a full sheet
            rank = cand["neg_val"].abs().where(clean).rank(ascending=False,
                                                           method="first")
            cand["use"] = clean & (rank <= PAIRS_PER_FILE)
            st.session_state["cand"] = cand
            st.session_state["combos"] = combos
            st.session_state["batch"] = 0
            status.update(label=f"{len(cand)} candidates", state="complete")
        else:
            status.update(label="No candidates", state="complete")
            st.session_state["cand"] = pd.DataFrame()
            st.session_state["combos"] = combos

    cand = st.session_state.get("cand")
    combos = st.session_state.get("combos")

    if combos is not None and len(combos):
        with st.expander(f"⚠ {len(combos)} combo packs skipped "
                         f"({combos['neg_val'].sum():,.0f} AED) — do these by hand"):
            st.caption("A combo holds two different items, e.g. 100ML + 50ML. "
                       "Breaking one releases both, so it needs two target lines. "
                       "Use the Manual pair tab for these.")
            st.dataframe(combos, use_container_width=True, hide_index=True)

    if cand is not None and len(cand):
        clean_mask = (~cand["blocked"] if "blocked" in cand.columns
                      else cand["problems"].eq(""))
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Candidates", len(cand))
        m2.metric("Value covered", f"{cand['neg_val'].sum():,.0f} AED")
        m3.metric("No blockers", int(clean_mask.sum()))
        m4.metric("Ticked", int(cand["use"].sum()))

        q1, q2, q3, q4 = st.columns([1.2, 1.2, 1.2, 1.4])
        topn = q1.number_input("Top N by value", min_value=1,
                               max_value=len(cand), value=min(11, len(cand)),
                               step=1, label_visibility="visible")
        q2.write(""); q3.write(""); q4.write("")
        if q2.button(f"Select top {int(topn)}", use_container_width=True):
            order = (cand["neg_val"].abs().where(clean_mask)
                     .rank(ascending=False, method="first"))
            cand["use"] = clean_mask & (order <= int(topn))
            st.session_state["cand"] = cand
        if q3.button("Select all clean", use_container_width=True):
            cand["use"] = clean_mask
            st.session_state["cand"] = cand
        if q4.button("Clear all", use_container_width=True):
            cand["use"] = False
            st.session_state["cand"] = cand

        st.caption(
            f"Sorted highest negative value first. "
            f"Top {int(topn)} unblocked pairs are worth "
            f"{cand.loc[clean_mask, 'neg_val'].abs().nlargest(int(topn)).sum():,.0f} AED "
            f"of {cand.loc[clean_mask, 'neg_val'].abs().sum():,.0f} unblocked total. "
            f"Over-clear notes are normal odd-quantity overshoot, not blockers."
        )

        BASIC = ["use", "neg_desc", "neg_qty", "neg_val", "par_desc", "conv",
                 "outers_needed", "cost_drift_pct", "problems"]
        extra_opts = [c for c in cand.columns if c not in BASIC]
        show_extra = st.multiselect("Add columns", extra_opts, default=[],
                                    help="The preview stays compact by default")
        cols = [c for c in BASIC if c in cand.columns] + show_extra

        with st.form("tick_form", border=False):
            st.caption("Tick freely — nothing reloads until you press Apply. "
                       "Check the physical shelf first.")
            edited = st.data_editor(
                cand[cols], use_container_width=True, hide_index=True, height=380,
                column_config={
                    "use": st.column_config.CheckboxColumn("✓", width="small"),
                    "neg_desc": st.column_config.TextColumn("Negative item", width="large"),
                    "par_desc": st.column_config.TextColumn("Outer / source", width="large"),
                    "neg_qty": st.column_config.NumberColumn("Neg qty", width="small"),
                    "neg_val": st.column_config.NumberColumn("Neg value", format="%.2f",
                                                             width="small"),
                    "conv": st.column_config.NumberColumn("Conv", width="small"),
                    "outers_needed": st.column_config.NumberColumn("Outers", width="small"),
                    "cost_drift_pct": st.column_config.NumberColumn(
                        "Drift %", format="%.0f%%", width="small"),
                    "problems": st.column_config.TextColumn("Problems", width="medium"),
                },
                disabled=[c for c in cols if c != "use"],
                key="cand_editor",
            )
            applied = st.form_submit_button("Apply ticks", type="primary")

        if applied:
            cand["use"] = edited["use"].values
            st.session_state["cand"] = cand
            st.session_state["batch"] = 0
    elif cand is not None:
        st.warning("No candidates. Loosen the sliders in the sidebar and run again.")

# ---------- Verify ----------
with tabV:
    st.caption("The app keeps nothing after you close it. This sheet is the record — "
               "export it, let the section staff check the shelf, upload it back.")
    cand = st.session_state.get("cand")

    st.subheader("1. Export for the sections")
    if cand is None or not len(cand):
        st.info("Run matching first.")
    else:
        st.write(f"{len(cand)} pairs across "
                 f"{cand['category'].nunique()} sections, numbered serially, "
                 f"one sheet per section.")
        st.download_button(
            "⬇ Download verification sheet",
            make_verification_book(cand),
            f"VERIFICATION_{adj_date.replace('-', '')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    st.divider()
    st.subheader("2. Upload it back once signed off")
    up = st.file_uploader("Verified sheet", type=["xlsx"], key="verified_up")
    if up is not None:
        try:
            v = read_verification(up.getvalue())
        except Exception as e:
            st.error(f"Could not read it: {e}")
            v = None
        if v is not None:
            yes = int(v["_verified"].sum())
            c1, c2, c3 = st.columns(3)
            c1.metric("Rows in sheet", len(v))
            c2.metric("Verified Y", yes)
            c3.metric("Rejected / blank", len(v) - yes)
            st.session_state["verified"] = v
            if cand is not None and len(cand):
                cand = cand.copy()
                cand["KEY"] = [pair_key(r) for r in cand.to_dict("records")]
                ok = set(v.loc[v["_verified"], "KEY"])
                order = dict(zip(v["KEY"], v["ROW"]))
                cand["use"] = cand["KEY"].isin(ok)
                cand["serial"] = cand["KEY"].map(order)
                cand = cand.sort_values("serial", na_position="last")
                st.session_state["cand"] = cand
                st.session_state["batch"] = 0
                st.success(f"{int(cand['use'].sum())} pairs ticked and put in "
                           f"serial order. Go to Build sheets.")
                st.dataframe(
                    cand.loc[cand["use"],
                             ["serial", "neg_desc", "neg_qty", "par_desc",
                              "conv", "outers_needed"]],
                    use_container_width=True, hide_index=True)
            else:
                st.warning("Run matching first so the verified rows can be matched "
                           "back to candidates.")

    st.divider()
    st.subheader("Where the tracking lives")
    st.markdown(
        "- The **verification sheet** is the saved state. It carries a hidden KEY "
        "column — do not delete it or the upload cannot match rows back.\n"
        "- Keep each dated sheet in a shared folder. That folder is your audit "
        "trail: who verified what, on which day, with shelf quantities and remarks.\n"
        "- Nothing is stored on the server. On Streamlit Cloud a database file "
        "would be wiped whenever the app sleeps or redeploys, so a file you hold "
        "is safer than one the app holds."
    )

# ---------- Build ----------
@st.fragment
def build_tab(adj_date, prepared, checked, verified):
    cand = st.session_state.get("cand")
    manual = st.session_state.get("manual", [])
    pool = []
    if cand is not None and len(cand):
        pool = cand[cand["use"]].to_dict("records")
    pool += manual

    if not pool:
        st.info("Run matching first, or add a pair by hand.")
    else:
        done = st.session_state.get("batch", 0)
        left = len(pool) - done
        c1, c2, c3 = st.columns(3)
        c1.metric("Pairs selected", len(pool))
        c2.metric("Already generated", done)
        c3.metric("Remaining", max(left, 0))

        o1, o2 = st.columns([2, 2])
        order = o1.selectbox(
            "Order the pairs by",
            ["Value, highest first", "Section then value", "As verified (serial)"],
            help="Value first clears the most money in the fewest sheets. "
                 "Section keeps one staff member in one aisle.")
        live_mode = o2.checkbox(
            "Live formulas (single side recalculates)", True,
            help="If staff correct the outer quantity, the single qty, cost and "
                 "value follow automatically and the total stays 0.00.")
        if order == "Value, highest first":
            pool = sorted(pool, key=lambda r: -abs(r.get("neg_val", 0)))
        elif order == "Section then value":
            pool = sorted(pool, key=lambda r: (str(r.get("category", "")),
                                               -abs(r.get("neg_val", 0))))
        elif "serial" in order:
            pool = sorted(pool, key=lambda r: (r.get("serial") or 1e9))

        remarks = st.text_input("Remarks", "OUTER BREAK FOR NEGATIVE STOCK")
        r1, r2, r3 = st.columns([1, 1, 2])
        start_at = r1.number_input("Start at serial", min_value=1,
                                   max_value=max(len(pool), 1),
                                   value=min(done + 1, max(len(pool), 1)))
        count = r2.number_input("How many pairs", min_value=1,
                                max_value=PAIRS_PER_FILE,
                                value=min(PAIRS_PER_FILE,
                                          max(len(pool) - int(start_at) + 1, 1)))
        r3.write(""); r3.write("")
        gen = r3.button(f"Generate serial {int(start_at)}–"
                        f"{int(start_at) + int(count) - 1}",
                        type="primary", use_container_width=True)
        if st.button("Reset counter"):
            st.session_state["batch"] = 0
            st.session_state.pop("last_sheet", None)
            st.rerun(scope="fragment")

        if gen:
            s0 = int(start_at) - 1
            chunk_rows = pool[s0:s0 + int(count)]
            lines = [Line(r["par_bc"], r["par_desc"],
                          "OFR" if r["conv"] != 1 else "PCS",
                          r["outers_needed"], r["par_cost"],
                          r["neg_bc"], r["neg_desc"], "PCS", r["conv"])
                     for r in chunk_rows]
            n = s0 // PAIRS_PER_FILE + 1
            data = build_sheet(lines, remarks, adj_date, prepared, checked,
                               verified, live=live_mode)
            prev = []
            for sl, (ln, r) in enumerate(zip(lines, chunk_rows), start=1):
                a_, b_ = ln.rows()
                prev.append({"SL": sl, "OUTER BARCODE": a_[0], "SINGLE BARCODE": "",
                             "DESCRIPTION": a_[1], "UNIT": a_[2],
                             "QTY": a_[3], "COST": a_[4], "VALUE": a_[5]})
                prev.append({"SL": "", "OUTER BARCODE": "", "SINGLE BARCODE": b_[0],
                             "DESCRIPTION": b_[1], "UNIT": b_[2],
                             "QTY": b_[3], "COST": b_[4], "VALUE": b_[5]})
            pdf = pd.DataFrame(prev)
            st.session_state["last_sheet"] = (n, data, pdf,
                                              round(pdf["VALUE"].sum(), 2))
            st.session_state["batch"] = max(done, s0 + len(lines))

        if st.session_state.get("last_sheet"):
            n, data, pdf, tot = st.session_state["last_sheet"]
            st.subheader(f"ADJ_{n:03d}  —  {len(pdf)} rows")
            if abs(tot) > 0.01:
                st.error(f"TOTAL {tot:.2f} — does not net to zero, do not post")
            else:
                st.success("TOTAL 0.00")
            st.dataframe(
                pdf, use_container_width=True, hide_index=True,
                column_config={
                    "OUTER BARCODE": st.column_config.TextColumn(width="medium"),
                    "SINGLE BARCODE": st.column_config.TextColumn(width="medium"),
                    "DESCRIPTION": st.column_config.TextColumn(width="large"),
                    "VALUE": st.column_config.NumberColumn(format="%.2f"),
                })
            st.download_button(f"⬇ Download ADJ_{n:03d}.xlsx", data,
                               f"ADJ_{n:03d}.xlsx",
                               "application/vnd.openxmlformats-officedocument."
                               "spreadsheetml.sheet",
                               use_container_width=True)

with tab2:
    candidates_tab(neg, master, master_idx, neg_map,
                   br_thresh, sw_lo, sw_hi, price_tol, drift_tol)

with tab3:
    build_tab(adj_date, prepared, checked, verified)


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
