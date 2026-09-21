"""
Negative Stock Adjustment Tool - Shams Al Madina
Streamlit app.  Run locally:  streamlit run app.py
"""

import io
import math
import re
from collections import defaultdict
from dataclasses import dataclass

import altair as alt
import pandas as pd
import streamlit as st

MONGO_IMPORT_ERR = None
mongo_store_mod = None
try:
    import mongo_store as mongo_store_mod
    from mongo_store import MongoStore
except Exception as _e:
    MongoStore = None
    MONGO_IMPORT_ERR = f"{type(_e).__name__}: {_e}"
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---- currency symbol -------------------------------------------------
# The Emirati dirham sign is U+20C3, accepted for Unicode 18.0 (Sept 2026).
# Most systems do not carry the glyph yet, so it renders as an empty box
# unless a font supplying it is loaded. Options:
#   "AED"      always safe, works everywhere including Excel
#   "\u20C3"   the new sign — needs DIRHAM_FONT_CSS below, and even then
#              Streamlit's data grid may not pick the font up
#   "\u062F.\u0625"  the Arabic د.إ, renders on any Arabic-capable font
AED = "AED"

# Paste a font URL here to load the glyph (e.g. the `dirham` npm package on a
# CDN). Leave empty to skip. Check it renders before switching AED above —
# a missing font shows a box, which is worse than the letters.
DIRHAM_FONT_URL = ""

# Official Emirati dirham sign, from the UAE design system. Inline SVG, so it
# inherits the surrounding text colour and needs no font or image file.
DIRHAM_SVG_PATH = (
    "M 7.0625 0.0820312 C 7.097656 0.128906 7.273438 0.347656 7.441406 0.554688 C 8.664062 2.019531 9.585938 4.402344 10.078125 7.402344 C 10.40625 9.375 10.425781 9.992188 10.425781 17.507812 L 10.425781 24.507812 L 7.078125 24.507812 C 4.023438 24.507812 3.671875 24.492188 3.070312 24.371094 C 2.128906 24.171875 1.152344 23.632812 0.496094 22.9375 C -0.0234375 22.382812 -0.0078125 22.351562 0.0234375 24.035156 C 0.0625 25.425781 0.078125 25.578125 0.28125 26.335938 C 0.601562 27.535156 1.039062 28.425781 1.703125 29.222656 C 2.609375 30.316406 3.527344 30.929688 4.839844 31.339844 C 5.121094 31.417969 5.710938 31.453125 7.808594 31.46875 L 10.425781 31.507812 L 10.425781 38.484375 L 6.734375 38.460938 L 3.03125 38.4375 L 2.390625 38.179688 C 1.632812 37.871094 1.289062 37.648438 0.542969 36.980469 L 0 36.488281 L 0.03125 38.023438 C 0.0703125 39.449219 0.078125 39.609375 0.28125 40.335938 C 0.976562 42.894531 2.65625 44.71875 4.871094 45.316406 C 5.425781 45.46875 5.640625 45.476562 7.953125 45.507812 L 10.425781 45.539062 L 10.425781 52.75 C 10.425781 57.101562 10.398438 60.3125 10.359375 60.859375 C 10.320312 61.359375 10.191406 62.292969 10.078125 62.945312 C 9.558594 65.945312 8.625 68.207031 7.28125 69.671875 L 7.007812 69.96875 L 20.535156 69.96875 C 28.625 69.96875 34.671875 69.9375 35.558594 69.894531 C 37.121094 69.816406 40.601562 69.46875 41.382812 69.300781 C 41.632812 69.25 42.097656 69.179688 42.398438 69.132812 C 43.046875 69.035156 44.121094 68.808594 45.664062 68.414062 C 47.839844 67.867188 49.824219 67.183594 51.769531 66.316406 C 52.375 66.042969 54.121094 65.148438 54.585938 64.867188 C 54.832031 64.722656 55.128906 64.542969 55.238281 64.488281 C 55.550781 64.320312 56.070312 63.980469 56.832031 63.433594 C 57.207031 63.160156 57.585938 62.894531 57.664062 62.839844 C 58 62.613281 59.160156 61.640625 59.6875 61.148438 C 61.695312 59.289062 63.375 57.222656 64.679688 55.011719 C 64.863281 54.6875 65.105469 54.285156 65.207031 54.117188 C 65.472656 53.667969 66.558594 51.414062 66.664062 51.074219 C 66.710938 50.921875 66.777344 50.761719 66.808594 50.730469 C 67.015625 50.457031 68.214844 46.660156 68.359375 45.828125 C 68.40625 45.5625 68.433594 45.523438 68.632812 45.484375 C 68.761719 45.460938 70.625 45.460938 72.777344 45.476562 C 77.078125 45.507812 77.078125 45.507812 78.03125 45.949219 C 78.566406 46.199219 78.726562 46.3125 79.320312 46.851562 C 80.097656 47.550781 80.023438 47.664062 79.976562 45.910156 C 79.945312 44.878906 79.902344 44.246094 79.832031 43.988281 C 79.558594 42.996094 79.496094 42.789062 79.257812 42.289062 C 78.472656 40.566406 77.160156 39.335938 75.480469 38.75 L 74.824219 38.507812 L 72.152344 38.476562 L 69.488281 38.4375 L 69.519531 37.496094 C 69.550781 36.253906 69.550781 33.800781 69.511719 32.539062 L 69.480469 31.523438 L 73.046875 31.507812 C 76.105469 31.492188 76.671875 31.507812 77.007812 31.597656 C 78.015625 31.878906 78.695312 32.265625 79.527344 33.027344 L 79.992188 33.464844 L 79.992188 32.273438 C 79.992188 30.855469 79.921875 30.230469 79.632812 29.296875 C 79.0625 27.40625 77.945312 25.996094 76.34375 25.128906 C 75.304688 24.5625 75.238281 24.546875 71.664062 24.523438 C 69.566406 24.507812 68.472656 24.476562 68.414062 24.425781 C 68.367188 24.378906 68.328125 24.300781 68.328125 24.234375 C 68.328125 24.171875 68.207031 23.664062 68.046875 23.117188 C 66.175781 16.460938 62.679688 11.175781 57.566406 7.257812 C 56.871094 6.71875 55.167969 5.582031 54.480469 5.199219 C 54.214844 5.042969 53.929688 4.882812 53.855469 4.835938 C 53.519531 4.652344 51.59375 3.699219 51.113281 3.5 C 50.824219 3.371094 50.449219 3.210938 50.28125 3.144531 C 47.457031 1.914062 42.71875 0.75 39.105469 0.386719 C 38.511719 0.328125 37.726562 0.242188 37.367188 0.210938 C 35.734375 0.0234375 33.472656 0 20.617188 0 C 9.753906 0 7.023438 0.0234375 7.0625 0.0820312 Z M 33.519531 3.5625 C 36.222656 3.726562 37.886719 3.933594 39.832031 4.410156 C 45.769531 5.824219 49.945312 8.820312 52.976562 13.824219 C 53.257812 14.289062 54.441406 16.71875 54.617188 17.210938 C 55.457031 19.488281 55.863281 20.839844 56.222656 22.625 C 56.3125 23.058594 56.433594 23.640625 56.488281 23.914062 C 56.542969 24.179688 56.566406 24.425781 56.542969 24.453125 C 56.503906 24.484375 48.472656 24.5 38.679688 24.492188 L 20.878906 24.476562 L 20.855469 14.136719 C 20.847656 8.457031 20.855469 3.734375 20.878906 3.644531 L 20.910156 3.492188 L 26.601562 3.492188 C 29.71875 3.492188 32.839844 3.523438 33.519531 3.5625 Z M 57.320312 31.75 C 57.375 32.09375 57.375 37.96875 57.320312 38.257812 L 57.273438 38.476562 L 39.070312 38.460938 L 20.878906 38.4375 L 20.863281 35.023438 C 20.847656 33.148438 20.863281 31.589844 20.878906 31.554688 C 20.902344 31.515625 28.65625 31.492188 39.097656 31.492188 L 57.273438 31.492188 Z M 56.503906 45.5625 C 56.542969 45.683594 56.351562 46.675781 55.960938 48.285156 C 55.511719 50.09375 54.902344 51.921875 54.289062 53.273438 C 53.984375 53.964844 53.222656 55.460938 53.039062 55.742188 C 52.953125 55.871094 52.695312 56.28125 52.472656 56.644531 C 51.03125 58.914062 48.976562 60.980469 46.632812 62.507812 C 45.777344 63.054688 44.015625 63.988281 43.542969 64.132812 C 43.449219 64.160156 43.34375 64.207031 43.304688 64.238281 C 43.246094 64.289062 42.519531 64.5625 41.671875 64.867188 C 40.113281 65.421875 37.144531 66.023438 34.761719 66.273438 C 33.214844 66.425781 32.96875 66.4375 27.023438 66.4375 L 20.871094 66.4375 L 20.871094 45.554688 L 38.542969 45.523438 C 48.265625 45.507812 56.273438 45.484375 56.335938 45.46875 C 56.40625 45.460938 56.480469 45.507812 56.503906 45.5625 Z M 56.503906 45.5625"
)

if DIRHAM_FONT_URL:
    st.markdown(
        f"""<style>
        @font-face {{
            font-family: 'DirhamSign';
            src: url('{DIRHAM_FONT_URL}') format('woff2');
            unicode-range: U+20C3;
            font-display: swap;
        }}
        html, body, [class*="st-"], [data-testid] {{
            font-family: 'DirhamSign', var(--font), sans-serif;
        }}
        </style>""", unsafe_allow_html=True)

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



XLSX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")


# ==================== icons ====================
# Lucide icons, inlined. Only usable where Streamlit renders raw HTML (the
# tiles and stat cards). Navigation, tabs and buttons only accept emoji or
# Material icons, so those use :material/…: — the rounded style is a close
# visual match to Lucide.
LUCIDE = {
    "folder-open": "<path d=\"m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2\" />",
    "pen-line": "<path d=\"M13 21h8\" /> <path d=\"M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622l4.353-1.32a2 2 0 0 0 .83-.497z\" />",
    "settings-2": "<path d=\"M14 17H5\" /> <path d=\"M19 7h-9\" /> <circle cx=\"17\" cy=\"17\" r=\"3\" /> <circle cx=\"7\" cy=\"7\" r=\"3\" />",
    "save": "<path d=\"M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z\" /> <path d=\"M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7\" /> <path d=\"M7 3v4a1 1 0 0 0 1 1h7\" />"
}


def lucide(name, size=26, colour="currentColor", stroke=1.8):
    inner = LUCIDE.get(name, "")
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" '
            f'height="{size}" viewBox="0 0 24 24" fill="none" '
            f'stroke="{colour}" stroke-width="{stroke}" stroke-linecap="round" '
            f'stroke-linejoin="round" style="flex-shrink:0">{inner}</svg>')


# ==================== display helpers ====================



def dirham_svg(height_em=0.82):
    """Inline SVG for the dirham sign. fill=currentColor, so it takes the
    colour of whatever text it sits in — red on negatives, green on positives."""
    return (f'<svg viewBox="0 0 80 70" height="{height_em}em" '
            f'style="vertical-align:-0.04em;margin-right:0.22em" '
            f'xmlns="http://www.w3.org/2000/svg" role="img" aria-label="AED">'
            f'<path fill="currentColor" d="{DIRHAM_SVG_PATH}"/></svg>')


def money_html(x, dp=2, colour=True):
    """Money with the real symbol, for places that render HTML."""
    try:
        v = float(x)
    except (TypeError, ValueError):
        return ""
    col = ("#ff6b6b" if v < 0 else "#2eb872" if v > 0 else "#9aa0a6") \
        if colour else "inherit"
    return (f'<span style="color:{col};white-space:nowrap">'
            f'{dirham_svg()}{v:,.{dp}f}</span>')


def stat_card(label, value_html, sub=""):
    return (f'<div style="padding:.7rem .9rem;border:1px solid rgba(250,250,250,.2);'
            f'border-radius:.5rem;background:rgba(250,250,250,.03)">'
            f'<div style="font-size:.78rem;opacity:.65;margin-bottom:.25rem">{label}</div>'
            f'<div style="font-size:1.5rem;font-weight:600;line-height:1.2">{value_html}</div>'
            f'<div style="font-size:.75rem;opacity:.55;margin-top:.15rem">{sub}</div></div>')


def money(x, dp=2, sign=False):
    """AED 1,234.56 — negatives keep their minus sign."""
    try:
        v = float(x)
    except (TypeError, ValueError):
        return ""
    s = f"{v:+,.{dp}f}" if sign else f"{v:,.{dp}f}"
    return f"{AED} {s}"


def colour_money(df, cols):
    """Red for negative, green for positive, grey for zero."""
    def style(v):
        try:
            v = float(v)
        except (TypeError, ValueError):
            return ""
        if v < 0:
            return "color: #ff6b6b; font-weight: 600"
        if v > 0:
            return "color: #2eb872; font-weight: 600"
        return "color: #888"
    live = [c for c in cols if c in df.columns]
    return df.style.map(style, subset=live) if live else df.style


# ==================== parsing ====================
@st.cache_data(show_spinner=False)
def load_master(raw: bytes) -> pd.DataFrame:
    want = ["Item Barcode", "Item No", "Item Name", "Stock", "Cost", "WAC",
            "Net MRP", "Is Active", "Category", "Group", "Brand"]
    head = pd.read_csv(io.BytesIO(raw), dtype=str, encoding="latin-1", nrows=0)
    use = [c for c in want if c in head.columns]
    m = pd.read_csv(io.BytesIO(raw), dtype=str, encoding="latin-1", usecols=use)
    m["Item Barcode"] = m["Item Barcode"].astype(str).str.strip()
    for src, dst in (("Stock", "stock"), ("Cost", "last_cost"),
                     ("WAC", "wac"), ("Net MRP", "mrp")):
        m[dst] = (
            pd.to_numeric(m[src].astype(str).str.replace(",", "", regex=False),
                          errors="coerce").fillna(0)
            if src in m.columns else 0.0
        )
    # iTrade values stock at the higher of last cost and weighted average.
    # Break at the lower figure and the adjustment under-clears the value.
    m["cost"] = m[["last_cost", "wac"]].max(axis=1)
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
        # code + qty is the minimum. Name is a bonus — it can be joined in
        # from the masterlist by barcode when the export leaves it out.
        essential = (bool(_match(cells, CODE_KEYS))
                     + bool(_match(cells, QTY_KEYS)))
        score = essential * 2 + bool(_match(cells, NAME_KEYS))
        if essential == 2 and score > best_score:
            best, best_score = i, score
    return best, best_score, grid


@st.cache_data(show_spinner=False)
def load_negatives(raw: bytes, sheet=0, header_row=None, mapping=None) -> pd.DataFrame:
    if header_row is None:
        header_row, score, _ = detect_header(raw, sheet)
        if header_row is None or score < 4:      # 4 = code + qty both found
            raise ValueError("HEADER_NOT_FOUND")
    t = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, header=header_row,
                      dtype=str).dropna(axis=1, how="all")
    cols = list(t.columns)
    mapping = mapping or {}
    c_code = mapping.get("code") or _match(cols, CODE_KEYS)
    c_name = mapping.get("name") or _match(cols, NAME_KEYS)   # may be absent
    c_qty = mapping.get("qty") or _match(cols, QTY_KEYS)
    if not (c_code and c_qty):
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
        name = (str(r.get(c_name, "")).strip() if c_name else "")
        if name in ("nan", "None"):
            name = ""
        if code in ("", "nan", "None"):
            continue

        if c_name:
            is_item = name != ""
        else:
            # no description column — a category band is a non-numeric code
            is_item = bool(re.fullmatch(r"[A-Za-z0-9\-_/]*\d[A-Za-z0-9\-_/]*",
                                        code))

        if not is_item:
            if code in ("Stock", "Non Stock"):
                grp = code
            else:
                cat = code
            continue

        r["group"] = grp or "STOCK"
        r["category"] = cat or "UNCATEGORISED"
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



# ==================== barcode variants ====================
# iTrade holds the same product under 070177178017 and 70177178017, and the
# POS accepts both. Stripping zeros blindly is not safe though: in this
# masterlist 01234 is a tailoring charge and 1234 is a ball needle. So a
# variant is only accepted when it resolves to exactly one item.

def bc_variants(code):
    """The forms a barcode might be stored under, most specific first."""
    c = str(code).strip()
    if not c:
        return []
    out = [c]
    bare = c.lstrip("0")
    if bare and bare != c:
        out.append(bare)
    for width in (12, 13, 14):
        if len(bare) < width:
            padded = bare.zfill(width)
            if padded != c:
                out.append(padded)
    return list(dict.fromkeys(out))


def resolve_bc(code, index):
    """Find code in index, trying zero variants. Returns (found, note).
    A variant that matches more than one item is refused."""
    c = str(code).strip()
    if c in index:
        return c, ""
    hits = [v for v in bc_variants(c)[1:] if v in index]
    if len(hits) == 1:
        return hits[0], f"matched {hits[0]} (leading zeros differ)"
    if len(hits) > 1:
        return None, f"ambiguous: {', '.join(hits)}"
    return None, ""


@st.cache_data(show_spinner=False)
def duplicate_barcodes(master: pd.DataFrame) -> pd.DataFrame:
    """Items held twice under zero-variant barcodes. Same product duplicated
    is itself a cause of negative stock — sales hit one code, receipts the
    other."""
    m = master[["Item Barcode", "Item Name", "stock", "cost"]].copy()
    m["canon"] = m["Item Barcode"].astype(str).str.strip().str.lstrip("0")
    m = m[m["canon"] != ""]
    grp = m.groupby("canon").filter(lambda g: g["Item Barcode"].nunique() > 1)
    if not len(grp):
        return pd.DataFrame()
    rows = []
    for canon, g in grp.groupby("canon"):
        names = g["Item Name"].astype(str).str.upper().str.replace(
            r"[^A-Z0-9]", "", regex=True)
        a = set(names.iloc[0])
        overlap = min(
            len(a & set(n)) / max(len(a | set(n)), 1) for n in names[1:])
        rows.append({
            "Canonical": canon,
            "Barcodes": " / ".join(g["Item Barcode"].astype(str)),
            "Items": " / ".join(g["Item Name"].astype(str)),
            "Stock": " / ".join(f"{v:g}" for v in g["stock"]),
            "Total stock": float(g["stock"].sum()),
            "Likely same product": overlap > 0.7,
        })
    return pd.DataFrame(rows).sort_values("Likely same product",
                                          ascending=False)



def cost_drift(derived, code, own_pair):
    """How far the derived single cost sits from the item's own cost.

    An item carries two: last purchase cost and weighted average. A break
    sets the WAC, a purchase sets the last cost, so either can legitimately
    match. Measure against the closer of the two — checking only one rejects
    correct pairs (TWININGS 20S: 0% against WAC, 25.6% against last cost)."""
    if not derived:
        return None
    costs = [c for c in own_pair.get(code, ()) if c and c > 0]
    if not costs:
        return None
    return round(min(((derived - c) / c * 100 for c in costs), key=abs), 1)


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
    own_pair = dict(zip(mm["Item Barcode"], zip(mm["last_cost"], mm["wac"])))
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
                cost_drift_pct=cost_drift(best[6] / best[3], r["bc"],
                                          own_pair),
            ))
    return pd.DataFrame(out), pd.DataFrame(combos)


@st.cache_data(show_spinner=False)
def find_swaps(m: pd.DataFrame, d: pd.DataFrame, lo: float, hi: float,
               price_tol: float) -> pd.DataFrame:
    mm = m.copy()
    mm["isml"] = mm["Item Name"].astype(str).str.upper().map(
        lambda x: bool(MULT.search(x)))
    own_cost = dict(zip(mm["Item Barcode"], mm["cost"]))
    own_pair = dict(zip(mm["Item Barcode"], zip(mm["last_cost"], mm["wac"])))
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
                cost_drift_pct=cost_drift(best[6], r["bc"], own_pair),
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


def write_adjustment(ws, lines, remarks, date_str, prepared, checked, verified,
                     live=True):
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
    return ws


def build_sheet(lines, remarks, date_str, prepared, checked, verified,
                live=True) -> bytes:
    """One adjustment as a standalone workbook."""
    wb = Workbook(); ws = wb.active; ws.title = "ADJUSTMENT"
    write_adjustment(ws, lines, remarks, date_str, prepared, checked,
                     verified, live)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def build_workbook(chunks, remarks, date_str, prepared, checked, verified,
                   live=True) -> bytes:
    """Every adjustment in one workbook, one worksheet per sheet."""
    wb = Workbook(); wb.remove(wb.active)
    for i, lines in enumerate(chunks, start=1):
        ws = wb.create_sheet(f"ADJ_{i:03d}")
        write_adjustment(ws, lines, remarks, date_str, prepared, checked,
                         verified, live)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _n(x):
    """Plain number, no trailing zeros, no thousands separator."""
    return f"{float(x):.7f}".rstrip("0").rstrip(".") or "0"


def make_txt(lines, prefix="SML"):
    """One line per stock movement, no header, no quotes:
        PREFIX,BARCODE,COST,QTY
    Outer side carries the negative qty, single side the positive.

    Returns (bytes, residual_report). The residual is computed from the
    numbers as actually written, not the full-precision ones, because that
    is what iTrade will multiply out."""
    out, report, net = [], [], 0.0
    for ln in lines:
        (obc, odesc, _, oqty, ocost, _), (sbc, sdesc, _, sqty, scost, _) = ln.rows()
        oc, oq = _n(ocost), _n(oqty)
        sc, sq = _n(scost), _n(sqty)
        out.append(f"{prefix},{obc},{oc},{oq}")
        out.append(f"{prefix},{sbc},{sc},{sq}")
        diff = round(float(oc) * float(oq) + float(sc) * float(sq), 2)
        net += diff
        if abs(diff) > 0.004:
            report.append({"Item": sdesc, "Conv": ln.conv,
                           "Written cost": sc, "Off by (AED)": diff})
    return (("\n".join(out) + "\n").encode("ascii", "ignore"),
            {"net": round(net, 2), "rows": report})


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



# ---- printable check sheet -------------------------------------------
# The full verification workbook has 18 columns and will not print. This is
# the paper version: nine narrow columns, A4 landscape, headings repeated on
# every page, one section per sheet.
PRINT_COLS = [
    ("OUTER BARCODE", 16), ("OUTER DESCRIPTION", 38),
    ("OUTER STOCK", 10), ("OUTER PRICE", 10),
    ("SINGLE BARCODE", 16), ("ITEM NAME", 38),
    ("NEGATIVE STOCK", 11),
    ("BREAK\n(outers)", 10), ("CONV\n(per outer)", 10),
    ("NEW QTY\n(singles)", 11), ("CONV. PRICE\n(per single)", 12),
    ("MOVES\n(AED)", 11), ("CLEARS\n(AED)", 11),
]


def make_print_sheet(cand: pd.DataFrame, date_str="", prepared="") -> bytes:
    """A4 landscape, one sheet per section, with the arithmetic spelled out
    so nobody has to work out what BREAK and CONV mean."""
    wb = Workbook()
    wb.remove(wb.active)
    for section, grp in cand.groupby("category", sort=True):
        ws = wb.create_sheet(str(section)[:28] or "OTHER")
        last = get_column_letter(len(PRINT_COLS))
        for i, (h, w) in enumerate(PRINT_COLS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        ws.merge_cells(f"A1:{last}1")
        ws["A1"] = f"STOCK CHECK — {section}"
        ws["A1"].font = Font(size=13, bold=True)
        ws["A1"].alignment = CENTER

        ws.merge_cells(f"A2:{last}2")
        ws["A2"] = ("Break BREAK outers  ×  CONV each  =  NEW QTY singles, "
                    "clearing the NEGATIVE STOCK. Single cost becomes CONV. "
                    "PRICE.    MOVES = value the adjustment shifts;  CLEARS = "
                    "negative value it removes. They differ when the single's "
                    "own cost is higher than the outer's.")
        ws["A2"].font = Font(size=9, italic=True)
        ws["A2"].alignment = Alignment("center")

        ws.merge_cells(f"A3:{last}3")
        ws["A3"] = (f"Date: {date_str}      Prepared by: {prepared}"
                    f"      Checked by: ____________________")
        ws["A3"].font = Font(size=9)

        for c, (h, _) in enumerate(PRINT_COLS, start=1):
            cell = ws.cell(row=4, column=c, value=h)
            cell.fill = GREY
            cell.font = Font(bold=True, size=9)
            cell.alignment = Alignment("center", "center", wrap_text=True)
            cell.border = BOX
        ws.row_dimensions[4].height = 30

        r = 5
        for rec in grp.sort_values("neg_val").to_dict("records"):
            conv = rec.get("conv", 1) or 1
            cost = rec.get("par_cost", 0) or 0
            outers = rec.get("outers_needed", 0) or 0
            vals = [str(rec.get("par_bc", "")), rec.get("par_desc", ""),
                    rec.get("par_stock", ""), cost,
                    str(rec.get("neg_bc", "")), rec.get("neg_desc", ""),
                    rec.get("neg_qty", 0),
                    outers, conv, round(abs(outers) * conv, 3),
                    round(cost / conv, 7) if conv else 0,
                    round(abs(outers) * cost, 2),
                    round(abs(rec.get("neg_val", 0) or 0), 2)]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BOX
                cell.font = Font(size=9)
                if c in (1, 5):
                    cell.number_format = "@"
                    cell.alignment = Alignment("center")
                elif c in (3, 7, 8, 9, 10):
                    cell.alignment = Alignment("center")
                elif c in (4, 11):
                    cell.number_format = "0.00####"
                    cell.alignment = Alignment("right")
                elif c in (12, 13):
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment("right")
                if c in (8, 9, 10):
                    cell.fill = PatternFill("solid", fgColor="EFEFEF")
            ws.row_dimensions[r].height = 16
            r += 1

        val = float(grp["neg_val"].abs().sum())
        moves = float(sum((g.get("outers_needed", 0) or 0)
                          * (g.get("par_cost", 0) or 0)
                          for g in grp.to_dict("records")))
        ws.cell(row=r, column=11, value="SHEET TOTAL").font = Font(
            bold=True, size=10)
        ws.cell(row=r, column=11).alignment = Alignment("right")
        for col, amount in ((12, moves), (13, val)):
            t = ws.cell(row=r, column=col, value=round(amount, 2))
            t.font = Font(bold=True, size=10)
            t.fill = GREY
            t.number_format = "#,##0.00"
            t.alignment = Alignment("right")
        ws.cell(row=r, column=1,
                value=f"{len(grp)} lines").font = Font(size=9, bold=True)
        for c in range(1, len(PRINT_COLS) + 1):
            ws.cell(row=r, column=c).border = BOX

        ws.cell(row=r + 2, column=1,
                value="Signature: ______________________").font = Font(size=9)

        ws.freeze_panes = "A5"
        ws.print_area = f"A1:{last}{r + 2}"
        ws.print_title_rows = "1:4"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = ws.page_margins.right = 0.25
        ws.page_margins.top = ws.page_margins.bottom = 0.35

    return _wb_bytes(wb)


def _wb_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def parse_serials(text, biggest):
    """'1,3,5-9 12' -> {1,3,5,6,7,8,9,12}. Ignores anything out of range."""
    out = set()
    for chunk in re.split(r"[,\s]+", str(text or "").strip()):
        if not chunk:
            continue
        m = re.fullmatch(r"(\d+)\s*[-–]\s*(\d+)", chunk)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.update(range(min(a, b), max(a, b) + 1))
        elif chunk.isdigit():
            out.add(int(chunk))
    return {n for n in out if 1 <= n <= biggest}


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


def store_diagnosis():
    """Returns (store_or_None, reason). The reason says which step failed."""
    if MongoStore is None:
        return None, ("`mongo_store.py` could not be imported. Either the file "
                      "is missing from the repo, or pymongo is not installed. "
                      f"Error: {MONGO_IMPORT_ERR}")
    try:
        cfg = st.secrets["mongo"]
    except Exception:
        try:
            keys = list(st.secrets.keys())
        except Exception:
            keys = []
        return None, ("No `[mongo]` section found in secrets. "
                      + (f"Sections present: {keys}. " if keys else
                         "No secrets are set at all. ")
                      + "On Streamlit Cloud: app menu -> Settings -> Secrets, "
                        "paste the block, press Save, then wait for the restart.")
    uri = cfg.get("uri", "")
    if not uri:
        return None, "The `[mongo]` section has no `uri` key."
    if not uri.startswith(("mongodb://", "mongodb+srv://")):
        return None, "The `uri` does not start with mongodb:// or mongodb+srv://"
    try:
        s = MongoStore(uri, cfg.get("db", "stockadj"))
        s.ensure_indexes()
        return s, "ok"
    except Exception as e:
        msg = f"{type(e).__name__}: {e}"
        if "dns" in msg.lower() or "srv" in msg.lower():
            msg += ("  —  a mongodb+srv:// URI needs dnspython. "
                    "Put `pymongo[srv]>=4.6` in requirements.txt.")
        return None, msg


@st.cache_resource(show_spinner=False)
def get_store():
    return store_diagnosis()[0]



# ==================== UI ====================
st.title(":material/inventory_2: Negative Stock Adjustment Tool")

# ---- controls as three large tiles ------------------------------------
def tile_head(icon, title, status, ok):
    """Lucide icon, title, and a one-line status in green or grey."""
    colour = "#2eb872" if ok else "#8b929c"
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:.75rem;'
        f'margin-bottom:.35rem">'
        f'<div style="color:var(--primary-color,#E23D3D);line-height:0">'
        f'{lucide(icon, 30)}</div>'
        f'<div><div style="font-size:1.15rem;font-weight:600">{title}</div>'
        f'<div style="font-size:.82rem;color:{colour}">{status}</div>'
        f'</div></div>', unsafe_allow_html=True)


# read current values so each tile can show its state before it is opened
_prev_date = st.session_state.get("adj_date", "11-09-26")
_prev_prep = st.session_state.get("prepared", "SWASTHIK")
_prev_thr = st.session_state.get("br_thresh", 0.80)
_prev_drift = st.session_state.get("drift_tol", 15.0)
_have_m = st.session_state.get("f_master") is not None
_have_n = st.session_state.get("f_neg") is not None

t1, t2, t3, t4 = st.columns([3, 3, 3, 2], gap="medium")

with t1.container(border=True):
    tile_head("folder-open", "Data",
              "Both files loaded" if (_have_m and _have_n)
              else "Masterlist and negative report needed",
              _have_m and _have_n)
    f_master = st.file_uploader("Masterlist (CSV)", type=["csv"],
                                key="f_master")
    f_neg = st.file_uploader("Negative stock report (XLSX)",
                             type=["xlsx", "xls"], key="f_neg")

with t2.container(border=True):
    tile_head("pen-line", "Sheet details",
              f"{_prev_date}  ·  {_prev_prep}", True)
    with st.popover("Edit", width="stretch"):
        adj_date = st.text_input("Date", "11-09-26", key="adj_date")
        prepared = st.text_input("Prepared by", "SWASTHIK", key="prepared")
        checked = st.text_input("Checked by", "IRSHAD", key="checked")
        verified = st.text_input("Verified by", "THALLATH", key="verified")
        txt_prefix = st.text_input("Import file prefix", "SML",
                                   key="txt_prefix",
                                   help="First field of every line in the .txt")

with t3.container(border=True):
    tile_head("settings-2", "Matching",
              f"Strictness {_prev_thr:.2f}  ·  drift ≤ {_prev_drift:.0f}%",
              True)
    with st.popover("Adjust", width="stretch"):
        br_thresh = st.slider("Bundle break strictness", 0.70, 1.00, 0.80,
                              0.01, key="br_thresh",
                              help="Higher = fewer but safer matches")
        sw_lo, sw_hi = st.slider("Wrong sale similarity window", 0.40, 1.00,
                                 (0.55, 0.95), 0.05, key="sw_window",
                                 help="Similar but not identical")
        price_tol = st.slider("Wrong sale price tolerance", 0.05, 0.60, 0.25,
                              0.05, key="price_tol")
        drift_tol = st.slider("Max cost drift %", 2.0, 60.0, 15.0, 1.0,
                              key="drift_tol",
                              help="A real break barely moves the item's "
                                   "cost. A big drift usually means the pair "
                                   "is wrong.")
st.write("")

with t4.container(border=True):
    _sid = st.session_state.get("session_id")
    tile_head("save", "Session",
              "Saved" if _sid else "Not saved yet", bool(_sid))
    _save_clicked = st.button("Save now", width="stretch",
                              type="primary", key="save_session_btn",
                              disabled=not (f_master and f_neg))

if not (f_master and f_neg):
    st.info("Drop the two files into the **Data** tile to start.")
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
        st.dataframe(grid.head(15), width="stretch")
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

store = get_store()
# A negative export without a description column is fine — join the names in
if neg is not None and master is not None:
    if "Item Name" not in neg.columns:
        neg["Item Name"] = ""
    blank = neg["Item Name"].astype(str).str.strip().eq("")
    if blank.any():
        lookup = dict(zip(master["Item Barcode"].astype(str).str.strip(),
                          master["Item Name"].astype(str)))
        def _name(code):
            hit, _ = resolve_bc(code, lookup)
            return lookup.get(hit, "") if hit else ""
        neg.loc[blank, "Item Name"] = neg.loc[blank, "bc"].map(_name)
        still = neg["Item Name"].astype(str).str.strip().eq("")
        neg.loc[still, "Item Name"] = "(not in masterlist) " + neg.loc[still, "bc"]
        st.caption(f"{int(blank.sum())} rows had no description in the export — "
                   f"{int(blank.sum() - still.sum())} filled in from the "
                   f"masterlist by barcode.")

if store is not None and neg is not None:
    # Optional. Never let the snapshot break the page — an older mongo_store.py
    # on the server will not have save_snapshot, and that must not be fatal.
    _snap = getattr(store, "save_snapshot", None)
    if _snap is None:
        st.session_state["snap_warn"] = (
            "History snapshots need the newer mongo_store.py. "
            "Push the updated file and redeploy — everything else works.")
    else:
        try:
            import hashlib
            sig = hashlib.md5(
                (str(len(neg)) + str(round(float(neg["val"].sum()), 2))
                 + str(getattr(f_neg, "name", ""))).encode()).hexdigest()
            if st.session_state.get("snap_sig") != sig:
                by = (neg.groupby("category", dropna=False)
                      .agg(lines=("val", "size"), qty=("qty", "sum"),
                           value=("val", "sum")).reset_index())
                okS, sid = _snap(
                    totals={"lines": int(len(neg)),
                            "value": round(float(neg["val"].sum()), 2),
                            "units": round(float(neg["qty"].sum()), 3),
                            "categories": int(neg["category"].nunique()),
                            "not_in_master": int((~neg["bc"].isin(
                                set(master["Item Barcode"]))).sum())},
                    by_category=[{"category": r["category"],
                                  "lines": int(r["lines"]),
                                  "qty": round(float(r["qty"]), 3),
                                  "value": round(float(r["value"]), 2)}
                                 for r in by.to_dict("records")],
                    source={"negative_file": getattr(f_neg, "name", ""),
                            "master_file": getattr(f_master, "name", "")})
                st.session_state["snap_sig"] = sig
                if okS:
                    st.session_state["snapshot_id"] = sid
                # new data loaded -> new session, saved automatically
                if hasattr(store, "save_session"):
                    okX, xid = store.save_session(
                        neg_df=neg,
                        source={"negative_file": getattr(f_neg, "name", ""),
                                "master_file": getattr(f_master, "name", "")},
                        settings={"break_threshold": br_thresh,
                                  "drift_tol": drift_tol,
                                  "price_tol": price_tol},
                        note="auto")
                    if okX:
                        st.session_state["session_id"] = xid
                        st.session_state["session_saved_at"] = "auto"
        except Exception as e:
            st.session_state["snap_warn"] = (
                f"Snapshot not saved ({type(e).__name__}). "
                f"Everything else is unaffected.")

if _save_clicked:
    if store is None or not hasattr(store, "save_session"):
        st.error("Saving needs the database — see the Archive tab.")
    else:
        _cand = st.session_state.get("cand")
        _settings = {"break_threshold": br_thresh, "drift_tol": drift_tol,
                     "price_tol": price_tol, "date": adj_date,
                     "prepared": prepared}
        if st.session_state.get("session_id"):
            ok_, _ = store.update_session(st.session_state["session_id"],
                                          cand_df=_cand, settings=_settings,
                                          note="saved")
        else:
            ok_, sid_ = store.save_session(
                neg_df=neg, cand_df=_cand, settings=_settings, note="saved",
                source={"negative_file": getattr(f_neg, "name", ""),
                        "master_file": getattr(f_master, "name", "")})
            if ok_:
                st.session_state["session_id"] = sid_
        (st.toast if hasattr(st, "toast") else st.success)(
            "Session saved" if ok_ else "Save failed")

tab1, tab2, tabV, tab3, tab4, tabA = st.tabs(
    ["Overview", "Candidates", "Verify", "Build sheets", "Manual pair", "Archive"]
)

# ---------- Overview ----------
with tab1:
    total = float(neg["val"].sum())
    units = float(neg["qty"].sum())
    _mcodes = set(master["Item Barcode"].astype(str).str.strip())
    in_master = neg["bc"].map(
        lambda c: resolve_bc(c, _mcodes)[0] is not None)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(stat_card("Negative lines", f"{len(neg):,}",
                          f"{neg['category'].nunique()} categories"),
                unsafe_allow_html=True)
    c2.markdown(stat_card("Negative value", money_html(total, 0)),
                unsafe_allow_html=True)
    c3.markdown(stat_card("Units short", f"{units:,.0f}"),
                unsafe_allow_html=True)
    c4.markdown(stat_card("Avg per line",
                          money_html(total / max(len(neg), 1), 1)),
                unsafe_allow_html=True)
    st.write("")

    # ---------- by group ----------
    if neg["group"].nunique() > 1:
        g = (neg.groupby("group", dropna=False)
             .agg(Lines=("val", "size"), Value=("val", "sum"))
             .reset_index().sort_values("Value"))
        cols = st.columns(len(g))
        for col, r in zip(cols, g.to_dict("records")):
            col.markdown(stat_card(r["group"], money_html(r["Value"], 0),
                                   f"{int(r['Lines'])} lines"),
                         unsafe_allow_html=True)
        st.write("")

    # ---------- by category ----------
    st.subheader("By category")
    by_cat = (neg.groupby("category", dropna=False)
              .agg(Lines=("val", "size"), Units=("qty", "sum"),
                   Value=("val", "sum"))
              .reset_index().sort_values("Value"))
    by_cat["Share %"] = (by_cat["Value"] / total * 100).round(1)
    by_cat["Avg/line"] = (by_cat["Value"] / by_cat["Lines"]).round(1)
    by_cat["Worst line"] = [
        neg.loc[neg["category"] == c, "val"].min() for c in by_cat["category"]]

    cdf = by_cat.assign(Negative=by_cat["Value"].abs()).copy()
    cdf["Units short"] = cdf["Units"].abs()
    cdf = cdf.sort_values("Negative", ascending=False).reset_index(drop=True)
    cdf["Cumulative %"] = (cdf["Negative"].cumsum() / cdf["Negative"].sum()
                           * 100).round(1)

    kind = st.radio("Chart", ["Linked", "Bars", "Line", "Area", "Pareto",
                              "Donut"],
                    horizontal=True, key="cat_chart_kind") or "Linked"

    TIP = [alt.Tooltip("category:N", title="Category"),
           alt.Tooltip("Lines:Q", title="Lines", format=",.0f"),
           alt.Tooltip("Negative:Q", title="Value (AED)", format=",.2f"),
           alt.Tooltip("Units short:Q", title="Units short", format=",.2f"),
           alt.Tooltip("Share %:Q", title="Share of total", format=".1f"),
           alt.Tooltip("Avg/line:Q", title="Avg per line", format=",.1f"),
           alt.Tooltip("Worst line:Q", title="Worst single line", format=",.2f")]
    GRID = alt.Axis(format=",.0f", grid=True, gridColor="#ffffff12")
    ORDER = cdf["category"].tolist()

    if kind == "Linked":
        # category bars on the left drive the item chart on the right
        items = neg[["bc", "Item Name", "category", "qty", "val"]].copy()
        items["Negative"] = items["val"].abs()
        items = (items.sort_values("Negative", ascending=False)
                 .groupby("category", group_keys=False).head(15))
        items["Item"] = items["Item Name"].str.slice(0, 42)

        click = alt.selection_point(fields=["category"], value=[
            {"category": cdf.iloc[0]["category"]}])

        left = (alt.Chart(cdf).mark_bar(cornerRadiusEnd=3)
                .encode(
                    x=alt.X("Negative:Q", title="Value (AED)",
                            axis=alt.Axis(format=",.0f", grid=True,
                                          gridColor="#ffffff12")),
                    y=alt.Y("category:N", sort="-x", title=None,
                            axis=alt.Axis(labelLimit=220, grid=False,
                                          labelOverlap=False, labelFontSize=11)),
                    color=alt.condition(click, alt.value("#ff5a5f"),
                                        alt.value("#3a3f4b")),
                    tooltip=TIP)
                .add_params(click)
                .properties(width=330, height=max(300, 26 * len(cdf)),
                            title="Click a category"))

        right = (alt.Chart(items).mark_bar(cornerRadiusEnd=3, color="#ff8f94")
                 .encode(
                     x=alt.X("Negative:Q", title="Value (AED)",
                             axis=alt.Axis(format=",.0f", grid=True,
                                           gridColor="#ffffff12")),
                     y=alt.Y("Item:N", sort="-x", title=None,
                             axis=alt.Axis(labelLimit=280, grid=False,
                                           labelOverlap=False, labelFontSize=11)),
                     tooltip=[alt.Tooltip("Item Name:N", title="Item"),
                              alt.Tooltip("bc:N", title="Barcode"),
                              alt.Tooltip("qty:Q", title="Qty", format=",.2f"),
                              alt.Tooltip("val:Q", title="Value (AED)",
                                          format=",.2f")])
                 .transform_filter(click)
                 .properties(width=420, height=max(300, 26 * len(cdf)),
                             title="Biggest items in it"))

        ch = alt.hconcat(left, right).resolve_scale(y="independent")

    elif kind == "Bars":
        sel = alt.selection_point(fields=["category"], on="click", empty=True)
        base = alt.Chart(cdf).encode(
            x=alt.X("Negative:Q", title="Negative value (AED)", axis=GRID),
            y=alt.Y("category:N", title=None, sort="-x",
                    axis=alt.Axis(labelLimit=220, grid=False,
                                  labelOverlap=False, labelFontSize=11)))
        ch = (base.mark_bar(cornerRadiusEnd=3, height=17).encode(
                  color=alt.condition(sel, alt.Color(
                      "Negative:Q", legend=None,
                      scale=alt.Scale(scheme="reds")), alt.value("#3a3f4b")),
                  opacity=alt.condition(sel, alt.value(.95), alt.value(.35)),
                  tooltip=TIP).add_params(sel)
              + base.mark_text(align="left", dx=4, color="#c9cdd4",
                               fontSize=11).encode(
                  text=alt.Text("Negative:Q", format=",.0f")))
        ch = ch.properties(height=max(280, 30 * len(cdf)))

    elif kind in ("Line", "Area"):
        base = alt.Chart(cdf).encode(
            x=alt.X("category:N", sort=ORDER, title=None,
                    axis=alt.Axis(labelAngle=-40, labelLimit=140, grid=False)),
            y=alt.Y("Negative:Q", title="Negative value (AED)", axis=GRID),
            tooltip=TIP)
        if kind == "Line":
            ch = (base.mark_line(color="#ff6b6b", strokeWidth=2,
                                 point=alt.OverlayMarkDef(
                                     color="#ff6b6b", size=60))
                  + base.mark_point(size=180, opacity=0))
        else:
            ch = (base.mark_area(
                      line={"color": "#ff6b6b", "strokeWidth": 2},
                      color=alt.Gradient(
                          gradient="linear",
                          stops=[alt.GradientStop(color="#ff6b6b00", offset=0),
                                 alt.GradientStop(color="#ff6b6b99", offset=1)],
                          x1=1, x2=1, y1=1, y2=0), interpolate="monotone")
                  + base.mark_point(size=180, opacity=0))
        ch = ch.properties(height=360)

    elif kind == "Pareto":
        base = alt.Chart(cdf).encode(
            x=alt.X("category:N", sort=ORDER, title=None,
                    axis=alt.Axis(labelAngle=-40, labelLimit=140, grid=False)))
        bar = base.mark_bar(cornerRadiusEnd=2, color="#ff6b6b", opacity=.75)\
            .encode(y=alt.Y("Negative:Q", title="Negative value (AED)",
                            axis=GRID), tooltip=TIP)
        line = base.mark_line(color="#2eb872", strokeWidth=2,
                              point=alt.OverlayMarkDef(color="#2eb872"))\
            .encode(y=alt.Y("Cumulative %:Q", title="Cumulative % of total",
                            axis=alt.Axis(format=".0f", grid=False)),
                    tooltip=[alt.Tooltip("category:N", title="Category"),
                             alt.Tooltip("Cumulative %:Q", format=".1f")])
        ch = alt.layer(bar, line).resolve_scale(y="independent")\
            .properties(height=380)

    else:  # Donut
        top = cdf.head(9).copy()
        rest = cdf.iloc[9:]
        if len(rest):
            top = pd.concat([top, pd.DataFrame([{
                "category": f"Other ({len(rest)})",
                "Negative": rest["Negative"].sum(),
                "Lines": rest["Lines"].sum(),
                "Units short": rest["Units short"].sum(),
                "Share %": round(rest["Negative"].sum() /
                                 cdf["Negative"].sum() * 100, 1),
                "Avg/line": 0, "Worst line": 0}])], ignore_index=True)
        ch = (alt.Chart(top).mark_arc(innerRadius=70, cornerRadius=2,
                                      stroke="#0e1117", strokeWidth=1)
              .encode(theta=alt.Theta("Negative:Q", stack=True),
                      color=alt.Color("category:N", title=None,
                                      scale=alt.Scale(scheme="reds"),
                                      sort=top["category"].tolist()),
                      tooltip=[alt.Tooltip("category:N", title="Category"),
                               alt.Tooltip("Negative:Q", title="Value (AED)",
                                           format=",.2f"),
                               alt.Tooltip("Share %:Q", format=".1f")])
              .properties(height=380))

    st.altair_chart(ch.configure_view(strokeWidth=0), width="stretch")
    st.caption({"Linked": "Click any category on the left — the right panel "
                          "switches to its biggest items. Hover either side "
                          "for barcode, quantity and value.",
                "Bars": "Hover for the full breakdown. Click a bar to isolate it.",
                "Line": "Categories ordered largest first — the drop-off shows "
                        "how few sections hold the money.",
                "Area": "Same ordering, filled to show accumulated weight.",
                "Pareto": "Bars are value, the green line is the running total. "
                          "Where it flattens, the rest stops mattering.",
                "Donut": "Top nine categories, everything else grouped."}
               .get(kind, ""))

    st.dataframe(
        by_cat.rename(columns={"category": "Category"}),
        width="stretch", hide_index=True,
        column_config={
            "Value": st.column_config.NumberColumn("Value", format="AED %.0f"),
            "Units": st.column_config.NumberColumn(format="%.0f"),
            "Share %": st.column_config.ProgressColumn(
                "Share", format="%.1f%%", min_value=0.0,
                max_value=float(max(by_cat["Share %"].max(), 1))),
            "Avg/line": st.column_config.NumberColumn(format="AED %.1f"),
            "Worst line": st.column_config.NumberColumn(format="AED %.0f"),
        })

    with st.expander("Lines against value — where effort pays off"):
        sc = by_cat.assign(Negative=by_cat["Value"].abs(),
                           AvgLine=by_cat["Avg/line"].abs()).copy()
        sc["Label"] = sc["category"].str.slice(0, 18)
        xmax = float(sc["Lines"].max()) * 1.6
        ymax = float(sc["Negative"].max()) * 1.6

        base = alt.Chart(sc).encode(
            x=alt.X("Lines:Q", title="Number of lines",
                    scale=alt.Scale(type="symlog", domain=[0, xmax],
                                    clamp=True),
                    axis=alt.Axis(grid=True, gridColor="#ffffff12")),
            y=alt.Y("Negative:Q", title="Negative value (AED)",
                    scale=alt.Scale(type="symlog", domain=[0, ymax],
                                    clamp=True),
                    axis=alt.Axis(format=",.0f", grid=True,
                                  gridColor="#ffffff12")),
            tooltip=[alt.Tooltip("category:N", title="Category"),
                     alt.Tooltip("Lines:Q", title="Lines", format=",.0f"),
                     alt.Tooltip("Negative:Q", title="Value (AED)",
                                 format=",.2f"),
                     alt.Tooltip("AvgLine:Q", title="Avg per line",
                                 format=",.1f")])

        pts = base.mark_circle(opacity=.85).encode(
            size=alt.Size("AvgLine:Q", title="Avg per line",
                          scale=alt.Scale(range=[70, 700])),
            color=alt.Color("Negative:Q", legend=None,
                            scale=alt.Scale(scheme="reds")))

        # size is FONT SIZE on a text mark — pin it, never inherit it
        txt = base.mark_text(align="left", dx=11, dy=-1, fontSize=10,
                             color="#c9cdd4").encode(
            text=alt.Text("Label:N"), size=alt.value(10),
            color=alt.value("#c9cdd4"))

        st.altair_chart((pts + txt).properties(height=430)
                        .configure_view(strokeWidth=0),
                        width="stretch")
        st.caption("Bubble size is the average per line. Top left is a few "
                   "lines holding a lot of money — worth doing by hand. "
                   "Bottom right is many small lines — bulk adjust.")

    # ---------- shape of the problem ----------
    st.subheader("Shape of the problem")
    bands = [(0, 10, "under 10"), (10, 50, "10 to 50"), (50, 200, "50 to 200"),
             (200, 1000, "200 to 1,000"), (1000, 1e12, "over 1,000")]
    rows = []
    av = neg["val"].abs()
    for lo, hi, name in bands:
        m = (av >= lo) & (av < hi)
        rows.append({"Value band (AED)": name, "Lines": int(m.sum()),
                     "Value": round(float(neg.loc[m, "val"].sum()), 2),
                     "% of value": round(
                         float(neg.loc[m, "val"].sum()) / total * 100, 1)})
    bands_df = pd.DataFrame(rows)
    bands_df["Negative"] = bands_df["Value"].abs()
    bc1, bc2 = st.columns([3, 2])
    with bc1:
        st.altair_chart(
            alt.Chart(bands_df).mark_bar(cornerRadiusEnd=3)
            .encode(
                x=alt.X("Value band (AED):N", sort=None, title=None,
                        axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Negative:Q", title="Value (AED)",
                        axis=alt.Axis(format=",.0f")),
                color=alt.Color("Negative:Q", legend=None,
                                scale=alt.Scale(scheme="reds")),
                tooltip=[alt.Tooltip("Value band (AED):N", title="Band"),
                         alt.Tooltip("Lines:Q", format=",.0f"),
                         alt.Tooltip("Negative:Q", title="Value (AED)",
                                     format=",.2f"),
                         alt.Tooltip("% of value:Q", format=".1f")])
            .properties(height=260).configure_view(strokeWidth=0),
            width="stretch")
    with bc2:
        st.dataframe(bands_df.drop(columns="Negative"),
                     width="stretch", hide_index=True,
                     column_config={
                         "Value": st.column_config.NumberColumn(format="AED %.0f"),
                         "% of value": st.column_config.ProgressColumn(
                             format="%.1f%%", min_value=0.0, max_value=100.0)})

    # ---------- data quality ----------
    st.subheader("Data quality")
    q1, q2, q3 = st.columns(3)
    miss = neg[~in_master]
    q1.metric("Not in masterlist", f"{len(miss):,}",
              money(miss["val"].sum(), 0), delta_color="off")
    frac = neg[(neg["qty"] % 1 != 0)]
    q2.metric("Fractional quantities", f"{len(frac):,}",
              money(frac["val"].sum(), 0), delta_color="off")
    nocost = neg[neg["cost"].isna() | (neg["cost"] == 0)] \
        if "cost" in neg.columns else neg.iloc[0:0]
    q3.metric("No cost on file", f"{len(nocost):,}",
              money(nocost["val"].sum(), 0), delta_color="off")

    # ---------- category drill-down ----------
    st.subheader("Categories")
    st.caption("Click a category to open its items.")

    f1, f2 = st.columns([2, 3])
    order = f1.selectbox("Order by", ["Value", "Name", "Lines"], index=0)
    find = f2.text_input("Search item or barcode (searches every category)", "")

    itemno_col = "Item No" if "Item No" in neg.columns else None
    hit = None
    if find.strip():
        t = find.strip().upper()
        hit = neg[neg["Item Name"].str.upper().str.contains(t, na=False)
                  | neg["bc"].str.upper().str.contains(t, na=False)]
        st.info(f"{len(hit)} matching lines · {money(hit['val'].sum(), 0)}")

    def copy_block(codes, key):
        """st.code carries a native copy icon, and copies the text verbatim —
        no thousands separators, no currency, no scientific notation."""
        codes = [str(c).strip() for c in codes if str(c).strip()]
        if not codes:
            return
        with st.expander(f"Copy barcodes ({len(codes)})"):
            sep = st.radio("Separator", ["One per line", "Comma", "Space"],
                           horizontal=True, key=f"sep_{key}",
                           label_visibility="collapsed") or "One per line"
            joiner = {"One per line": "\n", "Comma": ",", "Space": " "}[sep]
            st.code(joiner.join(codes), language=None)

    def item_table(df, key="x"):
        cols = ["bc", "Item Name"] + ([itemno_col] if itemno_col else []) + \
               ["qty", "val"]
        names = {"bc": "ItemCode", "Item Name": "Item Name",
                 "qty": "Quantity", "val": "Stock Value"}
        out = df.sort_values("val")[cols].rename(columns=names)
        st.dataframe(
            colour_money(out, ["Quantity", "Stock Value"]),
            width="stretch", hide_index=True,
            column_config={
                "ItemCode": st.column_config.TextColumn(width="medium"),
                "Item Name": st.column_config.TextColumn(width="large"),
                "Quantity": st.column_config.NumberColumn(format="%.2f"),
                "Stock Value": st.column_config.NumberColumn(format="AED %.2f")})
        copy_block(out["ItemCode"].tolist(), key)

    if hit is not None and len(hit):
        item_table(hit, key="search")
        st.divider()

    if order == "Value":
        cat_order = by_cat.sort_values("Value")["category"].tolist()
    elif order == "Lines":
        cat_order = by_cat.sort_values("Lines", ascending=False)["category"].tolist()
    else:
        cat_order = sorted(by_cat["category"])

    for c in cat_order:
        row = by_cat[by_cat["category"] == c].iloc[0]
        sub = neg[neg["category"] == c]
        with st.expander(
                f"{c}  ·  {int(row['Lines'])} lines  ·  "
                f"{row['Units']:,.2f} qty  ·  {money(row['Value'])}"):
            item_table(sub, key=f"cat_{c}")
            st.caption(f"Worst line {money(sub['val'].min())}  ·  "
                       f"average {money(row['Avg/line'], 1)} per line  ·  "
                       f"{row['Share %']:.1f}% of the store total")

    if len(miss):
        with st.expander(f"Not in the masterlist — {len(miss)} dead codes  ·  "
                         f"{money(miss['val'].sum())}"):
            st.caption("Leading-zero variants are already resolved, so these "
                       "genuinely do not exist under any form of the code.")
            item_table(miss, key="dead")

    dups = duplicate_barcodes(master)
    if len(dups):
        same = int(dups["Likely same product"].sum())
        with st.expander(f"One product under two barcodes — {len(dups)} pair(s), "
                         f"{same} look like the same item"):
            st.caption("The same product held under 0806149321941 and "
                       "806149321941 will take sales on one code and receipts "
                       "on the other, which drives one of them negative on its "
                       "own. Worth merging in the item master.")
            st.dataframe(dups, width="stretch", hide_index=True,
                         column_config={
                             "Barcodes": st.column_config.TextColumn(
                                 width="medium"),
                             "Items": st.column_config.TextColumn(
                                 width="large"),
                             "Likely same product":
                                 st.column_config.CheckboxColumn()})


# ---------- Candidates ----------
@st.fragment
def candidates_tab(neg, master, master_idx, neg_map,
                   br_thresh, sw_lo, sw_hi, price_tol, drift_tol):
    counts = (neg.groupby("category")
              .agg(lines=("val", "size"), val=("val", "sum")))
    cats = sorted(neg["category"].dropna().unique())
    label = {c: f"{c}  ({int(counts.loc[c, 'lines'])} lines, "
                f"{money(counts.loc[c, 'val'], 0)})" for c in cats}

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
    run = c4.button("Run matching", type="primary", width="stretch",
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
            # A pair either survives the checks or it is dropped here. Nothing
            # downstream re-decides, so every screen shows the same list.
            # Over/under-clear is normal odd-quantity overshoot, not a fault.
            fatal = [[x for x in p
                      if not (x.startswith("over-clears")
                              or x.startswith("under-clears"))] for p in probs]
            dropped = cand[[bool(f) for f in fatal]].copy()
            dropped["why"] = ["; ".join(f) for f in fatal if f]
            cand = cand[[not f for f in fatal]].reset_index(drop=True)
            cand["use"] = True
            st.session_state["dropped"] = dropped
            st.session_state["cand"] = cand
            st.session_state["combos"] = combos
            st.session_state["batch"] = 0
            st.session_state["sel_version"] = st.session_state.get(
                "sel_version", 0) + 1
            store = get_store()
            if (store is not None and st.session_state.get("session_id")
                    and hasattr(store, "update_session")):
                store.update_session(st.session_state["session_id"],
                                     cand_df=cand)
            if store is not None:
                okr, rid = store.save_run(
                    categories=pick_cats, mode=mode,
                    settings={"break_threshold": br_thresh,
                              "swap_window": [sw_lo, sw_hi],
                              "price_tol": price_tol, "drift_tol": drift_tol},
                    candidates_df=cand, combos_df=combos)
                if okr:
                    st.session_state["run_id"] = rid
            status.update(label=f"{len(cand)} candidates", state="complete")
        else:
            status.update(label="No candidates", state="complete")
            st.session_state["cand"] = pd.DataFrame()
            st.session_state["combos"] = combos
        st.rerun()               # one reload so Verify and Build see the result

    cand = st.session_state.get("cand")
    combos = st.session_state.get("combos")

    if combos is not None and len(combos):
        with st.expander(f":material/warning: {len(combos)} combo packs skipped "
                         f"({combos['neg_val'].sum():,.0f} AED) — do these by hand"):
            st.caption("A combo holds two different items, e.g. 100ML + 50ML. "
                       "Breaking one releases both, so it needs two target lines. "
                       "Use the Manual pair tab for these.")
            st.dataframe(combos, width="stretch", hide_index=True)

    if cand is not None and len(cand):
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Candidates", len(cand))
        m2.metric("Value covered", f"{cand['neg_val'].sum():,.0f} AED")
        m3.metric("Dropped by checks",
                  len(st.session_state.get("dropped", [])))
        m4.metric("Ticked", int(cand["use"].sum()))

        vals = cand["neg_val"].abs()
        vmax = float(vals.max()) if len(vals) else 0.0

        # Selection is computed from these controls every run — no widget holds
        # a stale copy of the ticks, so what you see is always what will build.
        q1, q2 = st.columns([2, 2])
        thresh = q1.number_input(
            "Only take pairs worth at least (AED)", min_value=0.0,
            max_value=max(vmax, 1.0), value=0.0, step=5.0, key="thresh",
            help="0 takes every pair. Raise it to skip the small ones.")
        q2.caption("Pairs that failed the cost, stock or barcode checks were "
                   "dropped when matching ran — they are not in this list.")
        eligible = cand[cand["neg_val"].abs() >= thresh]

        cand["use"] = cand.index.isin(eligible.index)
        st.session_state["cand"] = cand

        picked = cand.loc[cand["use"], "neg_val"].abs()
        per = st.session_state.get("per_sheet", PAIRS_PER_FILE)
        st.success(
            f"**{len(picked)} pairs selected**, {picked.sum():,.0f} AED — "
            f"{math.ceil(len(picked) / max(per, 1))} sheet(s) at {per} pairs each, "
            f"filled highest value first."
        )

        BASIC = ["neg_desc", "neg_qty", "neg_val", "par_desc", "conv",
                 "outers_needed", "cost_drift_pct", "problems"]
        extra_opts = [c for c in cand.columns
                      if c not in BASIC + ["use"]]
        show_extra = st.multiselect("Add columns", extra_opts, default=[],
                                    key="extra_cols")
        cols = [c for c in BASIC if c in cand.columns] + show_extra

        def cand_copy(df, key):
            codes = [str(c).strip() for c in df.get("neg_bc", []) if str(c).strip()]
            if not codes:
                return
            with st.expander(f"Copy barcodes ({len(codes)})"):
                which = st.radio("Which", ["Negative item", "Outer / source",
                                           "Both, one pair per line"],
                                 horizontal=True, key=f"cw_{key}",
                                 label_visibility="collapsed") or "Negative item"
                if which == "Negative item":
                    txt = "\n".join(str(x).strip() for x in df["neg_bc"])
                elif which == "Outer / source":
                    txt = "\n".join(str(x).strip() for x in df["par_bc"])
                else:
                    txt = "\n".join(f"{str(a).strip()},{str(b).strip()}"
                                     for a, b in zip(df["par_bc"], df["neg_bc"]))
                st.code(txt, language=None)

        view = st.radio("Show", ["Taken", "Left out", "Everything"],
                        horizontal=True, key="view_mode")
        shown = (cand[cand["use"]] if view == "Taken"
                 else cand[~cand["use"]] if view == "Left out" else cand)

        st.dataframe(
            colour_money(shown[cols], ["neg_val", "neg_qty", "cost_drift_pct"]),
            width="stretch", hide_index=True, height=380,
            column_config={
                "neg_desc": st.column_config.TextColumn("Negative item", width="large"),
                "par_desc": st.column_config.TextColumn("Outer / source", width="large"),
                "neg_qty": st.column_config.NumberColumn("Neg qty", width="small"),
                "neg_val": st.column_config.NumberColumn("Neg value",
                                                         format="AED %.2f",
                                                         width="small"),
                "conv": st.column_config.NumberColumn("Conv", width="small"),
                "outers_needed": st.column_config.NumberColumn("Outers", width="small"),
                "cost_drift_pct": st.column_config.NumberColumn(
                    "Drift %", format="%.0f%%", width="small"),
                "problems": st.column_config.TextColumn("Notes", width="medium"),
            })
        cand_copy(shown, "cand")

    elif cand is not None:
        st.warning("No candidates. Loosen the sliders in the sidebar and run again.")

# ---------- Verify ----------
with tabV:
    cand = st.session_state.get("cand")
    if cand is None or not len(cand):
        st.info("Run matching first.")
    else:
        cand = cand.copy()
        cand["serial"] = range(1, len(cand) + 1)
        st.session_state["cand"] = cand

        st.subheader("1. Print a check sheet")

        # The paper must match what Build sheets will actually use, or the
        # counts disagree and nobody knows which list is real.
        printable = cand[cand["use"]].copy() if "use" in cand.columns \
            else cand.copy()
        if not len(printable):
            st.warning("Nothing selected. Widen the value filter on the "
                       "Candidates tab.")
            st.stop()

        by_sec = (printable.groupby("category")
                  .agg(Lines=("neg_val", "size"),
                       Clears=("neg_val", lambda x: x.abs().sum()))
                  .reset_index().sort_values("Clears", ascending=False))
        st.caption(f"{len(printable)} pair(s) across "
                   f"{printable['category'].nunique()} section(s). A4 "
                   f"landscape, one sheet per section, headings repeat on "
                   f"every page.")
        v1, v2 = st.columns([1, 2])
        v1.markdown(stat_card("These sheets clear",
                              money_html(-printable["neg_val"].abs().sum(), 2),
                              f"{len(printable)} pairs"),
                    unsafe_allow_html=True)
        v2.dataframe(by_sec.rename(columns={"category": "Section"}),
                     width="stretch", hide_index=True,
                     column_config={"Clears": st.column_config.NumberColumn(
                         "Clears", format="AED %.2f")})

        dropped = st.session_state.get("dropped")
        if dropped is not None and len(dropped):
            with st.expander(f"{len(dropped)} pair(s) dropped when matching "
                             f"ran — not on this sheet"):
                st.dataframe(
                    dropped[["neg_bc", "neg_desc", "neg_val", "par_desc",
                             "why"]].rename(columns={
                                 "neg_bc": "Barcode", "neg_desc": "Item",
                                 "neg_val": "Value", "par_desc": "Outer",
                                 "why": "Reason"}),
                    width="stretch", hide_index=True,
                    column_config={"Barcode": st.column_config.TextColumn()})
                st.caption("Loosen the drift tolerance on Candidates if you "
                           "believe a pairing is right despite the cost gap.")

        # everything printed is what Build will use, unless staff cut some
        cand["use"] = cand.index.isin(printable.index)
        st.session_state["cand"] = cand
        st.session_state["batch"] = 0
        st.write("")

        p1, p2 = st.columns(2)
        p1.download_button(
            ":material/download: Check sheet for printing",
            make_print_sheet(printable, adj_date, prepared),
            f"CHECK_{adj_date.replace('-', '')}.xlsx", XLSX_MIME,
            width="stretch", type="primary")
        p2.download_button(
            ":material/download: Full workbook (every column)",
            make_verification_book(printable),
            f"VERIFICATION_{adj_date.replace('-', '')}.xlsx", XLSX_MIME,
            width="stretch")
        st.caption(f"Build sheets is now set to these {len(printable)} pair(s). "
                   f"If staff reject some, enter only the ones that passed "
                   f"below and it narrows to those.")

        st.divider()
        st.subheader("2. Enter what came back")
        st.caption("The sheet has no row numbers, so identify rows by their "
                   "SINGLE barcode. Paste or scan them — one per line, or "
                   "separated by commas or spaces.")
        t1, t2 = st.columns([3, 1])
        typed = t1.text_area("Single barcodes that passed", "", height=110,
                             key="passed_codes")
        t2.write("")
        if t2.button("All of them", width="stretch"):
            st.session_state["passed_codes"] = "\n".join(
                cand["neg_bc"].astype(str))
            st.rerun()
        if t2.button("Clear", width="stretch"):
            st.session_state["passed_codes"] = ""
            st.rerun()

        given = [c.strip() for c in re.split(r"[\s,;]+", typed or "")
                 if c.strip()]
        known = set(cand["neg_bc"].astype(str))
        picked, unknown, fuzzy = set(), set(), []
        for c in given:
            hit, note = resolve_bc(c, known)
            if hit:
                picked.add(hit)
                if note:
                    fuzzy.append(f"{c} → {hit}")
            else:
                unknown.add(c)
        if fuzzy:
            st.caption("Matched despite different leading zeros: "
                       + "; ".join(fuzzy[:6])
                       + (" …" if len(fuzzy) > 6 else ""))
        if unknown:
            st.warning(f"{len(unknown)} barcode(s) are not on this list — "
                       f"ignored: {', '.join(sorted(unknown)[:8])}"
                       + (" …" if len(unknown) > 8 else ""))

        if picked:
            cand["use"] = cand["neg_bc"].astype(str).isin(picked)
            st.session_state["cand"] = cand
            st.session_state["batch"] = 0
            sel = cand[cand["use"]]
            c1, c2, c3 = st.columns(3)
            c1.metric("Passed", len(sel))
            c2.metric("Left out", len(cand) - len(sel))
            c3.metric("Value covered", money(sel["neg_val"].sum(), 0))
            st.success("Go to Build sheets — these are the pairs it will use.")
            st.dataframe(
                sel[["neg_bc", "neg_desc", "neg_qty", "par_bc", "par_desc",
                     "outers_needed", "conv"]].rename(columns={
                         "neg_bc": "Single barcode", "neg_desc": "Item name",
                         "neg_qty": "Negative stock",
                         "par_bc": "Outer barcode", "par_desc": "Outer",
                         "outers_needed": "Breaking", "conv": "Conv"}),
                width="stretch", hide_index=True, height=320,
                column_config={
                    "Single barcode": st.column_config.TextColumn(),
                    "Outer barcode": st.column_config.TextColumn()})

        with st.expander("Or upload the full workbook back instead"):
            st.caption("Only if staff filled the Y/N column in Excel rather "
                       "than on paper. Do not delete the hidden KEY column.")
            up = st.file_uploader("Verified workbook", type=["xlsx"],
                                  key="verified_up")
            if up is not None:
                try:
                    v = read_verification(up.getvalue())
                except Exception as e:
                    st.error(f"Could not read it: {e}")
                    v = None
                if v is not None:
                    cand["KEY"] = [pair_key(r) for r in cand.to_dict("records")]
                    ok = set(v.loc[v["_verified"], "KEY"])
                    cand["use"] = cand["KEY"].isin(ok)
                    st.session_state["cand"] = cand
                    st.session_state["batch"] = 0
                    st.success(f"{int(cand['use'].sum())} of {len(v)} rows "
                               f"marked Y. Go to Build sheets.")

# ---------- Build ----------
@st.fragment
def build_tab(adj_date, prepared, checked, verified, txt_prefix, store):
    cand = st.session_state.get("cand")
    manual = st.session_state.get("manual", [])
    pool = []
    if cand is not None and len(cand):
        pool = cand[cand["use"]].to_dict("records")
    pool += manual

    if not pool:
        n_all = len(cand) if cand is not None else 0
        if n_all:
            st.warning(f"None of the {n_all} candidate(s) are selected. "
                       f"Go to the Verify tab — it sets the list — or widen "
                       f"the value filter on Candidates.")
        else:
            st.info("Run matching first, or add a pair by hand.")
    else:
        done = st.session_state.get("batch", 0)
        left = len(pool) - done
        c1, c2, c3 = st.columns(3)
        c1.metric("Pairs selected", len(pool),
                  f"of {len(cand)} candidates" if cand is not None
                  and len(cand) else None, delta_color="off")
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
        per = st.number_input("Pairs per sheet", min_value=1, max_value=60,
                              value=PAIRS_PER_FILE, step=1, key="per_sheet",
                              help="11 pairs = 22 rows. Pair 12 starts a new sheet.")
        st.caption(f"{len(pool)} pairs → {math.ceil(len(pool)/int(per))} sheets")

        st.markdown("#### All sheets at once")
        if st.button(f"Generate all {math.ceil(len(pool)/int(per))} sheets",
                     type="primary", width="stretch"):
            chunks, work, bad = [], [], 0
            txt_net, txt_rows, alltxt = 0.0, [], []
            for i in range(0, len(pool), int(per)):
                rows_ = pool[i:i + int(per)]
                lns = [Line(r["par_bc"], r["par_desc"],
                            "OFR" if r["conv"] != 1 else "PCS",
                            r["outers_needed"], r["par_cost"],
                            r["neg_bc"], r["neg_desc"], "PCS", r["conv"])
                       for r in rows_]
                chunks.append(lns)
                t, rep = make_txt(lns, txt_prefix)
                alltxt.append(t.decode())
                txt_net += rep["net"]; txt_rows += rep["rows"]
                k = i // int(per) + 1
                for ln in lns:
                    a_, b_ = ln.rows()
                    work.append({"FILE": f"ADJ_{k:03d}",
                                 "OUTER BARCODE": a_[0], "SINGLE BARCODE": b_[0],
                                 "OUTER ITEM": a_[1], "SINGLE ITEM": b_[1],
                                 "CONV": ln.conv, "OUTER QTY": a_[3],
                                 "SINGLE QTY": b_[3], "COST OUTER": a_[4],
                                 "COST SINGLE": b_[4], "OUTER VALUE": a_[5],
                                 "SINGLE VALUE": b_[5],
                                 "DIFFERENCE": round(a_[5] + b_[5], 2)})
                    bad += abs(a_[5] + b_[5]) > 0.01

            wdf = pd.DataFrame(work)
            for c in ("OUTER BARCODE", "SINGLE BARCODE"):
                wdf[c] = wdf[c].astype(str)
            wbuf = io.BytesIO()
            with pd.ExcelWriter(wbuf, engine="openpyxl") as xw:
                wdf.to_excel(xw, sheet_name="WORKING", index=False)

            st.session_state["bulk"] = {
                "xlsx": build_workbook(chunks, remarks, adj_date, prepared,
                                       checked, verified, live=live_mode),
                "txts": [(f"ADJ_{i:03d}.txt",
                          make_txt(c, txt_prefix)[0])
                         for i, c in enumerate(chunks, start=1)],
                "txt": "".join(alltxt).encode("ascii", "ignore"),
                "working": wbuf.getvalue(),
                "sheets": len(chunks), "pairs": len(pool),
                "bad": bad, "net": round(txt_net, 2), "rows": txt_rows,
            }

        bulk = st.session_state.get("bulk")
        if bulk:
            if bulk["bad"]:
                st.error(f"{bulk['bad']} pairs do not net to zero")
            else:
                st.success(f"{bulk['pairs']} pairs in {bulk['sheets']} sheets, "
                           f"every total 0.00")
            if bulk["rows"]:
                st.warning(f"Import file residual {bulk['net']:+.2f} AED across "
                           f"{len(bulk['rows'])} pair(s) — uneven conversions. "
                           f"The Excel sheets are exact.")
                with st.expander("Which pairs, and by how much"):
                    st.dataframe(pd.DataFrame(bulk["rows"]),
                                 width="stretch", hide_index=True)

            stamp = adj_date.replace("-", "")
            XL = ("application/vnd.openxmlformats-officedocument."
                  "spreadsheetml.sheet")

            g1, g2 = st.columns(2)
            g1.download_button(
                f":material/download: Excel — {bulk['sheets']} sheets in one file",
                bulk["xlsx"], f"ADJUSTMENTS_{stamp}.xlsx", XL,
                width="stretch", type="primary")
            g2.download_button(
                ":material/download: Txt — every line in one file", bulk["txt"],
                f"ADJUSTMENTS_{stamp}.txt", "text/plain",
                width="stretch")

            st.caption(f"Or one txt per sheet, named to match the tabs "
                       f"in the workbook:")
            tcols = st.columns(min(4, len(bulk["txts"])) or 1)
            for i, (name, data_) in enumerate(bulk["txts"]):
                tcols[i % len(tcols)].download_button(
                    f":material/download: {name}", data_, name, "text/plain",
                    key=f"bulktxt_{name}", width="stretch")

            with st.expander("Working file and previews"):
                st.download_button(
                    ":material/download: Working file (reconciliation)", bulk["working"],
                    f"WORKING_{stamp}.xlsx", XL, width="stretch")
                st.caption("Combined import file:")
                st.code(bulk["txt"].decode(), language=None)

            if store and st.button(":material/save: Save this batch to the database",
                                   width="stretch"):
                ok, res = store.save_batch(
                    label=f"{adj_date} · {bulk['pairs']} pairs · "
                          f"{bulk['sheets']} sheets",
                    files={f"ADJUSTMENTS_{stamp}.xlsx": bulk["xlsx"],
                           f"ADJUSTMENTS_{stamp}.txt": bulk["txt"],
                           f"WORKING_{stamp}.xlsx": bulk["working"]},
                    meta={"remarks": remarks, "pairs": bulk["pairs"],
                          "sheets": bulk["sheets"], "per_sheet": int(per)},
                    run_id=st.session_state.get("run_id"))
                (st.success if ok else st.error)(
                    "Saved." if ok else f"Not saved: {res}")

        st.divider()
        st.markdown("#### One sheet at a time")
        st.caption("For printing and handing to a section.")
        r1, r2, r3 = st.columns([1, 1, 2])
        start_at = r1.number_input("Start at serial", min_value=1,
                                   max_value=max(len(pool), 1),
                                   value=min(done + 1, max(len(pool), 1)))
        count = r2.number_input("How many pairs", min_value=1,
                                max_value=int(per),
                                value=min(int(per),
                                          max(len(pool) - int(start_at) + 1, 1)))
        r3.write(""); r3.write("")
        gen = r3.button(f"Generate serial {int(start_at)}–"
                        f"{int(start_at) + int(count) - 1}",
                        type="primary", width="stretch")
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
            _t, _rep = make_txt(lines, txt_prefix)
            st.session_state["last_txt"] = _t
            st.session_state["last_txt_rep"] = _rep
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
                colour_money(pdf, ["QTY", "VALUE"]),
                width="stretch", hide_index=True,
                column_config={
                    "OUTER BARCODE": st.column_config.TextColumn(width="medium"),
                    "SINGLE BARCODE": st.column_config.TextColumn(width="medium"),
                    "DESCRIPTION": st.column_config.TextColumn(width="large"),
                    "VALUE": st.column_config.NumberColumn(format="AED %.2f"),
                })
            with st.expander("Copy barcodes from this sheet"):
                st.code("\n".join(
                    [str(x).strip() for x in pdf["OUTER BARCODE"] if str(x).strip()]
                    + [str(x).strip() for x in pdf["SINGLE BARCODE"]
                       if str(x).strip()]), language=None)
            txt = st.session_state.get("last_txt", b"")
            d1, d2 = st.columns(2)
            d1.download_button(
                f":material/download: Excel — ADJ_{n:03d}.xlsx", data, f"ADJ_{n:03d}.xlsx",
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet", width="stretch", type="primary")
            d2.download_button(
                f":material/download: Txt — ADJ_{n:03d}.txt", txt, f"ADJ_{n:03d}.txt",
                "text/plain", width="stretch")
            rep = st.session_state.get("last_txt_rep", {"net": 0, "rows": []})
            if rep["rows"]:
                st.warning(f"Import file residual {rep['net']:+.2f} AED on "
                           f"{len(rep['rows'])} pair(s) — uneven conversion. "
                           f"The Excel sheet itself is exact.")
            if store:
                if st.button(f":material/save: Save ADJ_{n:03d} to the database",
                             width="stretch"):
                    ok, res = store.save_batch(
                        label=f"{adj_date} · ADJ_{n:03d}",
                        files={f"ADJ_{n:03d}.xlsx": data,
                               f"ADJ_{n:03d}.txt": txt},
                        meta={"remarks": remarks, "pairs": len(pdf) // 2},
                        run_id=st.session_state.get("run_id"))
                    (st.success if ok else st.error)(
                        "Saved." if ok else f"Not saved: {res}")
            with st.expander("Preview the import file"):
                st.code(txt.decode(), language="text")
                if rep["rows"]:
                    st.dataframe(pd.DataFrame(rep["rows"]),
                                 width="stretch", hide_index=True)

with tab2:
    candidates_tab(neg, master, master_idx, neg_map,
                   br_thresh, sw_lo, sw_hi, price_tol, drift_tol)

with tab3:
    build_tab(adj_date, prepared, checked, verified, txt_prefix, store)


with tabA:
    if store is None:
        _, why = store_diagnosis()
        st.error(f"**Database not connected.**  {why}")
        st.info("Every run and every generated sheet can be kept once this "
                "is set up. The secrets block should look like:")
        st.code('[mongo]\nuri = "mongodb+srv://user:password@cluster0.xxxxx.'
                'mongodb.net/?retryWrites=true&w=majority"\ndb  = "stockadj"',
                language="toml")
        st.caption("MongoDB Atlas M0 is free and gives 512 MB — a generated "
                   "sheet is about 7 KB. Streamlit Cloud has no fixed outbound "
                   "IP, so Atlas Network Access has to allow 0.0.0.0/0. That "
                   "makes the password the only barrier, so use a long one and "
                   "give the user access to this database only.")
    else:
        if st.session_state.get("snap_warn"):
            st.warning(st.session_state["snap_warn"])
        ok, msg = store.check()
        (st.success if ok else st.error)(msg)
        if getattr(mongo_store_mod, "STORE_VERSION", 1) < 2:
            st.warning("mongo_store.py is out of date — push the current one.")
        if ok:
            sub0, sub1, sub2 = st.tabs(["History by date", "Saved batches",
                                        "Run history"])

            with sub0:
                if not hasattr(store, "list_snapshots"):
                    st.warning("`mongo_store.py` in this deployment is older "
                               "than `app.py`. Push the current mongo_store.py "
                               "to the repo and the history will start "
                               "recording.")
                    snaps = []
                else:
                    snaps = store.list_snapshots()
                if not snaps:
                    st.caption("No snapshots yet. One is saved automatically "
                               "each time you upload a negative stock file.")
                else:
                    hist = pd.DataFrame([{
                        "When": s_["at"].strftime("%Y-%m-%d %H:%M"),
                        "Date": s_["at"].date(),
                        "Lines": s_["totals"].get("lines"),
                        "Value": s_["totals"].get("value"),
                        "Units": s_["totals"].get("units"),
                        "Categories": s_["totals"].get("categories"),
                        "Dead codes": s_["totals"].get("not_in_master"),
                        "_id": str(s_["_id"]),
                    } for s_ in snaps]).sort_values("When")

                    m1, m2, m3 = st.columns(3)
                    first, last = hist.iloc[0], hist.iloc[-1]
                    m1.markdown(stat_card("First snapshot",
                                          money_html(first["Value"], 0),
                                          str(first["When"])),
                                unsafe_allow_html=True)
                    m2.markdown(stat_card("Latest",
                                          money_html(last["Value"], 0),
                                          str(last["When"])),
                                unsafe_allow_html=True)
                    delta = last["Value"] - first["Value"]
                    m3.markdown(stat_card(
                        "Change", money_html(delta, 0),
                        "lower is better" if delta > 0 else "gone the wrong way"),
                        unsafe_allow_html=True)
                    st.write("")

                    plot = hist.assign(Negative=hist["Value"].abs())
                    st.altair_chart(
                        alt.Chart(plot).mark_line(
                            color="#ff6b6b", strokeWidth=2,
                            point=alt.OverlayMarkDef(color="#ff6b6b"))
                        .encode(
                            x=alt.X("When:T", title=None),
                            y=alt.Y("Negative:Q", title="Negative value (AED)",
                                    axis=alt.Axis(format=",.0f", grid=True,
                                                  gridColor="#ffffff12")),
                            tooltip=[alt.Tooltip("When:T"),
                                     alt.Tooltip("Negative:Q", format=",.2f"),
                                     alt.Tooltip("Lines:Q", format=",.0f"),
                                     alt.Tooltip("Dead codes:Q")])
                        .properties(height=280).configure_view(strokeWidth=0),
                        width="stretch")

                    st.dataframe(hist.drop(columns=["_id", "Date"])
                                 .sort_values("When", ascending=False),
                                 width="stretch", hide_index=True,
                                 column_config={
                                     "Value": st.column_config.NumberColumn(
                                         format="AED %.2f"),
                                     "Units": st.column_config.NumberColumn(
                                         format="%.2f")})

                    pickd = st.selectbox("Open a snapshot",
                                         hist["When"].tolist()[::-1])
                    row = hist[hist["When"] == pickd].iloc[0]
                    full = (store.get_snapshot(row["_id"])
                            if hasattr(store, "get_snapshot") else None)
                    if full and full.get("by_category"):
                        bysnap = pd.DataFrame(full["by_category"])
                        bysnap = bysnap.sort_values("value")
                        st.caption(
                            f"{full.get('source', {}).get('negative_file', '')}")
                        st.dataframe(bysnap, width="stretch",
                                     hide_index=True,
                                     column_config={
                                         "value": st.column_config.NumberColumn(
                                             "Value", format="AED %.2f"),
                                         "qty": st.column_config.NumberColumn(
                                             "Units", format="%.2f")})

            with sub1:
                batches = store.list_batches()
                if not batches:
                    st.caption("Nothing saved yet. Generate sheets, then press "
                               "Save on the Build tab.")
                for b in batches:
                    when = b["at"].strftime("%Y-%m-%d %H:%M")
                    with st.expander(f"{b['label']}  ·  {when}  ·  "
                                     f"{b['n_files']} file(s), "
                                     f"{b['bytes']/1024:.0f} KB"):
                        meta = b.get("meta", {})
                        if meta:
                            st.caption(" · ".join(f"{k}: {v}"
                                                  for k, v in meta.items()))
                        for f in b["files"]:
                            data = store.get_file(str(b["_id"]), f["name"])
                            if data is None:
                                continue
                            mime = ("text/plain" if f["name"].endswith(".txt")
                                    else "application/zip"
                                    if f["name"].endswith(".zip")
                                    else "application/vnd.openxmlformats-"
                                         "officedocument.spreadsheetml.sheet")
                            st.download_button(
                                f":material/download: {f['name']}  ({f['size']/1024:.0f} KB)",
                                data, f["name"], mime,
                                key=f"dl_{b['_id']}_{f['name']}",
                                width="stretch")
                            if f["name"].endswith(".txt"):
                                st.code(data.decode("ascii", "ignore"),
                                        language="text")

            with sub2:
                runs = store.list_runs()
                if not runs:
                    st.caption("No runs recorded yet.")
                else:
                    st.dataframe(pd.DataFrame([{
                        "When": r["at"].strftime("%Y-%m-%d %H:%M"),
                        "Categories": ", ".join(r.get("categories", [])),
                        "Mode": r.get("mode", ""),
                        "Candidates": r.get("n_candidates", 0),
                        "Value": round(r.get("value", 0), 2),
                        "Drift tol": r.get("settings", {}).get("drift_tol"),
                    } for r in runs]), width="stretch", hide_index=True)
                    pick = st.selectbox(
                        "Reopen a run",
                        [f"{r['at'].strftime('%Y-%m-%d %H:%M')} — "
                         f"{r.get('n_candidates', 0)} candidates" for r in runs])
                    if st.button("Load this run's candidates"):
                        r = runs[[f"{x['at'].strftime('%Y-%m-%d %H:%M')} — "
                                  f"{x.get('n_candidates', 0)} candidates"
                                  for x in runs].index(pick)]
                        full = store.get_run(str(r["_id"]))
                        if full and full.get("candidates"):
                            st.session_state["cand"] = pd.DataFrame(
                                full["candidates"])
                            st.session_state["batch"] = 0
                            st.success("Loaded. Go to Build sheets.")
                        else:
                            st.warning("That run has no candidates stored.")


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
                     width="stretch", hide_index=True)
        if st.button("Clear manual pairs"):
            st.session_state["manual"] = []
            st.rerun()
