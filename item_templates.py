"""
item_templates.py — the three item request sheets, as PDF.

    CREATION      new items the purchaser has to create in iTrade
    ACTIVATION    items that exist but are switched off
    DESCRIPTION   a name that is wrong and has to be corrected

All three are approval paper: they go up with a signature from the person who
prepared them, the purchaser who owns the vendor, and whoever approves. The
layout follows the sheets already in use, so nobody has to learn a new form.

No Streamlit in here on purpose — the page imports it, and so does the test.

GP% is always computed, never typed:

    GP% = (RSP - COST) / RSP * 100

on the raw figures, with no VAT adjustment. That matches every row on the
existing sheets (35/55 -> 36.36, 29.16/49 -> 40.49, 10.83/17.5 -> 38.11).
"""

from datetime import date
from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether,
                                PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

# ==================== house colours ====================
# Taken off the sheets in use: maroon banner and header row, blue purpose line.
MAROON = colors.HexColor("#C00000")
BLUE = colors.HexColor("#0B5394")
GREY_LINE = colors.HexColor("#808080")
SOFT = colors.HexColor("#F2F2F2")

PAGE = landscape(A4)
MARGIN = 12 * mm
USABLE = PAGE[0] - 2 * MARGIN          # 273 mm

COMPANY = "AL MADINA GROUP(SHAMS AL MADINA, SALEM MALL)"


# ==================== the three sheet definitions ====================
# Each column: (key, heading, width in mm, alignment, kind)
#   kind "index" the row number, filled in for you
#        "text"  wraps
#        "num"   right aligned, 2 dp
#        "calc"  GP%, computed from COST and RSP
#
# fields  = the header lines above the table, in order
# blanks  = what an empty cell prints as. Creation sheets go to the purchaser
#           before the barcode exists, and the sheets in use say NEED BARCODE
#           rather than leaving a gap, so it is clear it was not forgotten.

_ITEM_COLS = [
    ("sno", "S.No", 12, "c", "index"),
    ("subcat", "Sub Category", 34, "l", "text"),
    ("single", "Single Barcode", 28, "c", "text"),
    ("outer", "Outer Barcode", 28, "c", "text"),
    ("desc", "Description", 79, "l", "text"),
    ("unit", "UNIT", 16, "c", "text"),
    ("packing", "PACKING", 18, "c", "text"),
    ("cost", "COST", 20, "r", "num"),
    ("rsp", "RSP", 20, "r", "num"),
    ("gp", "GP%", 18, "r", "calc"),
]

TYPES = {
    "creation": {
        "label": "Creation",
        "purpose": "CREATION",
        "prefix": "CRE",
        # The store's own creation sheet carries a Remark line as well.
        "fields": ["date", "vendor", "maingrp", "remark"],
        "cols": _ITEM_COLS,
        "blanks": {"single": "NEED BARCODE", "outer": "NEED BARCODE"},
    },
    "activation": {
        "label": "Activation",
        "purpose": "ACTIVATION",
        "prefix": "ACT",
        "fields": ["date", "vendor", "maingrp", "remark"],
        "cols": _ITEM_COLS,
        "blanks": {},
    },
    "description": {
        "label": "Description update",
        "purpose": "DESCRIPTION UPDATE",
        "prefix": "DSC",
        "fields": ["date", "reason"],
        "cols": [
            ("sno", "S.No", 11, "c", "index"),
            ("single", "Single Barcode", 26, "c", "text"),
            ("outer", "Outer Barcode", 26, "c", "text"),
            ("old_desc", "Old Description", 68, "l", "text"),
            ("unit", "Unit", 14, "c", "text"),
            ("new_desc", "New Description", 68, "l", "text"),
            ("cost", "COST", 21, "r", "num"),
            ("rsp", "RSP", 21, "r", "num"),
            ("gp", "GP%", 18, "r", "calc"),
        ],
        "blanks": {},
    },
}

FIELD_LABELS = {
    "date": "DATE",
    "vendor": "VENDOR",
    "maingrp": "Main Grp",
    "reason": "REASON",
    "remark": "Remark",
}

SIGN_SLOTS = [("prepared", "Prepared By"),
              ("purchaser", "Concerned Purchaser"),
              ("approved", "Approved By")]


# ==================== numbers ====================
def num(x):
    """A figure or None. Blank cells, stray spaces and commas all become None
    rather than raising — the person is typing into a grid, not a form."""
    if x is None:
        return None
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        f = float(x)
        # A cleared cell in a number column comes back as NaN, which would
        # otherwise print as "nan" on the sheet.
        return None if f != f else f
    s = str(x).strip().replace(",", "")
    if not s or s.lower() in ("nan", "none", "nat", "<na>"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def gp_pct(cost, rsp):
    """Margin on selling price. None when it cannot be worked out, so the
    cell stays empty instead of showing a wrong 0.00 or 100.00."""
    c, r = num(cost), num(rsp)
    if c is None or r is None or r == 0:
        return None
    return (r - c) / r * 100.0


def fmt(v, dp=2):
    return "" if v is None else f"{v:,.{dp}f}"


def row_problems(kind, row, i):
    """What is wrong with one row, in words a storekeeper can act on. Returns
    a list — a row can have more than one thing missing."""
    spec = TYPES[kind]
    out = []
    desc_key = "new_desc" if kind == "description" else "desc"
    if is_blank(row.get(desc_key)):
        out.append(f"row {i}: {FIELD_LABELS.get(desc_key, 'description')} "
                   f"is empty")
    c, r = num(row.get("cost")), num(row.get("rsp"))
    # On a description sheet the prices are only there to identify the item;
    # the change being asked for is the name, so blank prices are not a fault.
    if kind != "description":
        if c is None:
            out.append(f"row {i}: COST is empty")
        if r is None:
            out.append(f"row {i}: RSP is empty")
    if c is not None and r is not None:
        if r == 0:
            out.append(f"row {i}: RSP is zero, so GP% cannot be worked out")
        elif r < c:
            out.append(f"row {i}: RSP {r:,.2f} is below COST {c:,.2f} — "
                       f"selling at a loss")
    if kind == "description":
        if is_blank(row.get("single")) and is_blank(row.get("outer")):
            out.append(f"row {i}: no barcode — iTrade needs one to find the "
                       f"item being renamed")
        old = "" if is_blank(row.get("old_desc")) else str(row["old_desc"]).strip()
        new = "" if is_blank(row.get("new_desc")) else str(row["new_desc"]).strip()
        if old and new and old.upper() == new.upper():
            out.append(f"row {i}: old and new description are the same")
    elif kind == "activation":
        # An activation is for an item that already exists, so a missing
        # barcode means the row cannot be acted on at all.
        if is_blank(row.get("single")) and is_blank(row.get("outer")):
            out.append(f"row {i}: no barcode — an activation needs the "
                       f"existing barcode")
    return out


def is_blank(v):
    """True for None, empty text, and the NaN a cleared grid cell produces.
    str(nan) is "nan", which is truthy — so a plain truth test is not enough."""
    if v is None:
        return True
    if isinstance(v, float) and v != v:
        return True
    return str(v).strip().lower() in ("", "nan", "none", "nat", "<na>")


def filled_rows(kind, rows):
    """Drop rows the person left blank. A grid always has empty rows at the
    bottom; printing them as numbered lines would be wrong."""
    spec = TYPES[kind]
    keys = [k for k, *_ in spec["cols"] if k not in ("sno", "gp")]
    return [r for r in rows if any(not is_blank(r.get(k)) for k in keys)]


# ==================== styles ====================
def _styles(base=7.4):
    return {
        "cell": ParagraphStyle("cell", fontName="Helvetica", fontSize=base,
                               leading=base + 1.6, alignment=TA_LEFT),
        "cellc": ParagraphStyle("cellc", fontName="Helvetica", fontSize=base,
                                leading=base + 1.6, alignment=TA_CENTER),
        "cellr": ParagraphStyle("cellr", fontName="Helvetica", fontSize=base,
                                leading=base + 1.6, alignment=TA_RIGHT),
        "head": ParagraphStyle("head", fontName="Helvetica-Bold",
                               fontSize=base, leading=base + 1.6,
                               alignment=TA_CENTER,
                               textColor=colors.white),
    }


def _cell(text, align, sty):
    """Wrapped text for the long columns, plain strings for the narrow ones —
    a Paragraph in every cell makes long sheets noticeably slower to build."""
    s = "" if text is None else str(text)
    s = (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    key = {"c": "cellc", "r": "cellr"}.get(align, "cell")
    return Paragraph(s, sty[key])


# ==================== the sheet ====================
def build_pdf(kind, header, rows, *, ref="", logo=None):
    """Return the PDF as bytes.

    kind    one of TYPES
    header  date, vendor, maingrp, reason, remark, prepared, purchaser,
            approved — whatever the type asks for
    rows    list of dicts keyed by the column keys
    ref     the reference printed top right, e.g. CRE-260926-1
    logo    optional image bytes for the banner
    """
    if kind not in TYPES:
        raise ValueError(f"unknown template type {kind!r}")
    spec = TYPES[kind]
    cols = spec["cols"]
    rows = filled_rows(kind, rows)

    buf = BytesIO()
    doc = BaseDocTemplate(buf, pagesize=PAGE,
                          leftMargin=MARGIN, rightMargin=MARGIN,
                          topMargin=10 * mm, bottomMargin=10 * mm,
                          title=f"{spec['purpose']} {ref}".strip(),
                          author=header.get("prepared", ""),
                          subject=f"Item {spec['label'].lower()} request")
    frame = Frame(MARGIN, 10 * mm, USABLE,
                  PAGE[1] - 20 * mm, id="body",
                  leftPadding=0, rightPadding=0,
                  topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate("p", [frame],
                                       onPage=_page_furniture(ref))])

    sty = _styles()
    flow = []

    # ---- banner -------------------------------------------------------
    band = Table([[_banner_cell(logo),
                   Paragraph(COMPANY, ParagraphStyle(
                       "co", fontName="Helvetica-Bold", fontSize=13,
                       leading=15, alignment=TA_CENTER,
                       textColor=colors.white))]],
                 colWidths=[32 * mm, USABLE - 32 * mm], rowHeights=[14 * mm])
    band.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), MAROON),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("BOX", (0, 0), (-1, -1), 0.7, MAROON),
    ]))
    flow += [band, Spacer(1, 3.5 * mm)]

    # ---- purpose ------------------------------------------------------
    flow.append(Paragraph(
        f"PURPOSE&nbsp;:&nbsp; {spec['purpose']}",
        ParagraphStyle("pp", fontName="Helvetica-Bold", fontSize=11.5,
                       leading=13, textColor=BLUE)))
    flow.append(Spacer(1, 2.5 * mm))

    # ---- header fields ------------------------------------------------
    lab = ParagraphStyle("lab", fontName="Helvetica-Bold", fontSize=8.6,
                         leading=11)
    val = ParagraphStyle("val", fontName="Helvetica", fontSize=8.6,
                         leading=11)
    fcells = []
    for f in spec["fields"]:
        v = str(header.get(f) or "").strip()
        if not v and f != "date":
            continue
        fcells.append([Paragraph(f"{FIELD_LABELS[f]} :", lab),
                       Paragraph(v.replace("&", "&amp;"), val)])
    if fcells:
        ft = Table(fcells, colWidths=[26 * mm, USABLE - 26 * mm])
        ft.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ]))
        flow += [ft, Spacer(1, 3 * mm)]

    # ---- the table ----------------------------------------------------
    widths = [w * mm for _, _, w, _, _ in cols]
    head = [Paragraph(h, sty["head"]) for _, h, _, _, _ in cols]
    body = [head]
    for i, r in enumerate(rows, 1):
        line = []
        for key, _h, _w, align, kindc in cols:
            if kindc == "index":
                line.append(_cell(i, align, sty))
            elif kindc == "calc":
                line.append(_cell(fmt(gp_pct(r.get("cost"), r.get("rsp"))),
                                  align, sty))
            elif kindc == "num":
                line.append(_cell(fmt(num(r.get(key))), align, sty))
            else:
                v = r.get(key)
                txt = "" if is_blank(v) else str(v).strip()
                if not txt:
                    txt = spec["blanks"].get(key, "")
                line.append(_cell(txt, align, sty))
        body.append(line)

    t = Table(body, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), MAROON),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, GREY_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, SOFT]),
    ]))
    flow.append(t)

    # ---- signatures ---------------------------------------------------
    # Kept with nothing after it, so it never lands alone on a second page
    # without at least some of the table above it.
    flow.append(Spacer(1, 8 * mm))
    flow.append(_signatures(header))

    doc.build(flow)
    return buf.getvalue()


def _banner_cell(logo):
    """The logo if one was supplied, otherwise the branch initials — an empty
    white square on a red band looks like a missing image."""
    if logo:
        try:
            from reportlab.platypus import Image
            img = Image(BytesIO(logo))
            scale = min((28 * mm) / img.imageWidth, (11 * mm) / img.imageHeight)
            img.drawWidth = img.imageWidth * scale
            img.drawHeight = img.imageHeight * scale
            return img
        except Exception:
            pass
    return Paragraph("AL MADINA", ParagraphStyle(
        "lg", fontName="Helvetica-Bold", fontSize=10, leading=12,
        alignment=TA_CENTER, textColor=colors.white))


def _signatures(header):
    sty = ParagraphStyle("sg", fontName="Helvetica-Bold", fontSize=8.6,
                         leading=13)
    cells = []
    for key, label in SIGN_SLOTS:
        name = str(header.get(key) or "").strip()
        cells.append(Paragraph(f"{label} :&nbsp; {name}", sty))
    w = (USABLE - 8 * mm) / 3
    t = Table([cells], colWidths=[w, w, w], rowHeights=[16 * mm])
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (0, 0), 0.7, GREY_LINE),
        ("BOX", (1, 0), (1, 0), 0.7, GREY_LINE),
        ("BOX", (2, 0), (2, 0), 0.7, GREY_LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
    ]))
    return KeepTogether(t)


def _page_furniture(ref):
    """Reference top right, page number bottom right. Both on every page, so a
    sheet that runs to two pages can still be filed."""
    def draw(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(GREY_LINE)
        if ref:
            canvas.drawRightString(PAGE[0] - MARGIN, PAGE[1] - 7 * mm, ref)
        canvas.drawRightString(PAGE[0] - MARGIN, 6 * mm,
                               f"Page {canvas.getPageNumber()}")
        canvas.drawString(MARGIN, 6 * mm,
                          f"Generated {date.today().strftime('%d-%m-%Y')}")
        canvas.restoreState()
    return draw
