"""
pages/templates.py — item request sheets.

Three forms go to the purchaser: create an item, switch a dead one back on,
or correct a wrong description. They were being typed into Excel by hand each
time; this fills the same layout from a grid, works out GP% so it cannot be
mistyped, and keeps a copy so a sheet can be found again months later.

Nothing here touches stock. It is paperwork.
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

# The PIN gate runs here too, not only in app.py. A page script can be reached
# without the router having run — Streamlit's own pages/ discovery does exactly
# that — and a gate that only guards the front door is not a gate.
try:
    from auth import require_pin
    require_pin()
except ImportError:
    pass


import xlsx_template as xt
from item_templates import (FIELD_LABELS, SIGN_SLOTS, TYPES, build_pdf,
                            filled_rows, gp_pct, is_blank, num, row_problems)

XLSX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")

# ==================== icons ====================
# Lucide, inlined. Only usable where Streamlit renders raw HTML — tabs and
# buttons take Material icons instead.
LUCIDE = {
    "file-plus": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M9 15h6"/><path d="M12 18v-6"/>',
    "list-checks": '<path d="m3 17 2 2 4-4"/><path d="m3 7 2 2 4-4"/><path d="M13 6h8"/><path d="M13 12h8"/><path d="M13 18h8"/>',
    "pen-line": '<path d="M13 21h8"/><path d="M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622l4.353-1.32a2 2 0 0 0 .83-.497z"/>',
    "signature": '<path d="M20 20a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2"/><path d="M3 10c1.5 0 2.5-1 3-2 .8 3.2 2 5 3.5 5 1 0 1.5-1 1.5-2 0-2 1-3 2-3s2 1 2 3c0 1.2.8 2 2 2h4"/>',
}


def lucide(name, size=28, stroke=1.8):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" '
            f'height="{size}" viewBox="0 0 24 24" fill="none" '
            f'stroke="currentColor" stroke-width="{stroke}" '
            f'stroke-linecap="round" stroke-linejoin="round" '
            f'style="flex-shrink:0">{LUCIDE.get(name, "")}</svg>')


def tile_head(icon, title, status, ok=True):
    colour = "#2eb872" if ok else "#8b929c"
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:.75rem;'
        f'margin-bottom:.35rem">'
        f'<div style="color:var(--primary-color,#E23D3D);line-height:0">'
        f'{lucide(icon, 30)}</div>'
        f'<div><div style="font-size:1.15rem;font-weight:600">{title}</div>'
        f'<div style="font-size:.82rem;color:{colour}">{status}</div>'
        f'</div></div>', unsafe_allow_html=True)


# ==================== storage ====================
def _diagnose():
    """(store, reason). Each failure says what is actually wrong — a single
    shared message used to blame the secrets when the real cause was an old
    file on the server."""
    try:
        import mongo_store
    except Exception as e:
        return None, (f"`mongo_store.py` could not be imported "
                      f"({type(e).__name__}: {e}).")
    if not hasattr(mongo_store.MongoStore, "save_template"):
        return None, ("The server is running an **old `mongo_store.py`** — it "
                      "has no template storage. Push the current file, then "
                      "**Manage app → Reboot**: Streamlit keeps imported "
                      "files in memory, so a push alone is not enough.")
    try:
        cfg = st.secrets["mongo"]
        uri = cfg["uri"]
    except Exception:
        return None, ("No `[mongo]` block with a `uri` in your secrets. "
                      "Sheets will still generate and download — they just "
                      "will not be kept.")
    try:
        s = mongo_store.MongoStore(uri, cfg.get("db", "stockadj"))
        ok, msg = s.check()
        if not ok:
            return None, f"Database unreachable: {msg}"
        return s, "ok"
    except Exception as e:
        return None, f"Could not connect: {type(e).__name__}: {e}"


@st.cache_resource(show_spinner=False)
def get_store():
    return _diagnose()[0]


store, why = _diagnose()
APP = (store.get_settings() if store is not None else {})


def default(key, fallback):
    v = APP.get(key)
    return v if v not in (None, "") else fallback


# Used only when the Settings page has never been filled in.
SIGN_FALLBACK = {"prepared": "SWASTHIK", "purchaser": "IQBAL",
                 "approved": "THALLATH"}


def seed(key, value):
    """Put a starting value in place before the widget is drawn. Passing both
    `value=` and a `key` that already exists makes Streamlit warn and ignore
    one of them; seeding first avoids the ambiguity altogether."""
    if key not in st.session_state:
        st.session_state[key] = value


def layout_for(kind):
    """The uploaded workbook for this sheet type, or None to use the built-in
    layout. Read fresh each run — a layout saved on the other tab must take
    effect on the next generate, not after a reload."""
    if store is None or not hasattr(store, "get_layout"):
        return None
    try:
        doc = store.get_layout(kind)
    except Exception:
        return None
    return doc if doc and doc.get("workbook") else None


# ==================== what gets typed per item ====================
# Items are entered one at a time and collected in a list, not a spreadsheet
# grid. A grid looks efficient but hides what it is doing: it defers every
# value until the cell loses focus, so GP% lags a keystroke behind, and a
# stray row of whitespace is invisible. One item at a time, with the sheet
# building up underneath, is what this job actually looks like.
ITEM_FIELDS = {
    "creation": ["single", "outer", "desc", "unit", "packing", "cost", "rsp"],
    "activation": ["single", "outer", "desc", "unit", "packing", "cost",
                   "rsp"],
    "description": ["single", "outer", "old_desc", "new_desc"],
}

# Shown in an expander rather than the main row: a description sheet is about
# the name, but the template has price columns, so they stay reachable.
EXTRA_FIELDS = {"description": ["unit", "cost", "rsp"]}

# What must be there before an item can be added.
REQUIRED = {
    "creation": ["desc"],
    "activation": ["desc"],
    "description": ["new_desc"],
}


HEADINGS = {k: h for spec in TYPES.values()
            for k, h, *_ in spec["cols"]}

HELP = {
    "single": "The piece barcode. Leave blank on a creation — it prints as "
              "NEED BARCODE.",
    "outer": "The case or outer barcode, if the item has one.",
    "packing": "Pieces per outer.",
    "cost": "Landed cost per piece, excluding VAT.",
    "rsp": "Shelf price per piece. GP% is worked out from this.",
    "old_desc": "The wrong description, exactly as it reads in iTrade now.",
    "new_desc": "What it should say.",
}


# Header field -> the middle part of its widget key, so a sheet loaded back
# from the archive can refill the boxes it was typed into.
FIELD_WIDGET = {"vendor": "vendor", "maingrp": "grp",
                "reason": "reason", "remark": "remark"}


def items_key(kind):
    return f"tpl_items_{kind}"


def items(kind):
    """The items typed so far for this sheet type. Each type keeps its own
    list, so switching between them mid-job loses nothing."""
    return st.session_state.setdefault(items_key(kind), [])


def entry_key(kind, field):
    return f"tpl_in_{kind}_{field}"


def entry_value(kind, field):
    v = st.session_state.get(entry_key(kind, field))
    return None if is_blank(v) else v


def all_entry_fields(kind):
    return ITEM_FIELDS[kind] + EXTRA_FIELDS.get(kind, [])


def clear_entry(kind):
    """Empty the input boxes. Only safe before the widgets are drawn, so it
    is requested with a flag and carried out at the top of the next run."""
    for f in all_entry_fields(kind):
        st.session_state.pop(entry_key(kind, f), None)


def entry_widget(kind, field, label=None):
    """One input box. Barcodes and packing are text, not numbers: a barcode
    in a number box loses its leading zero and turns into 1.00071E+10."""
    key = entry_key(kind, field)
    lbl = label or HEADINGS.get(field, field)
    if field in ("cost", "rsp"):
        return st.number_input(lbl, key=key, min_value=0.0, step=0.01,
                               format="%.2f", value=None,
                               placeholder="0.00", help=HELP.get(field))
    return st.text_input(lbl, key=key, help=HELP.get(field),
                         placeholder=PLACEHOLDER.get(field))


PLACEHOLDER = {
    "single": "leave blank if not created yet",
    "outer": "optional",
    "desc": "LADIES LONG DRESS DYN-4412",
    "old_desc": "as it reads in iTrade now",
    "new_desc": "what it should say",
    "unit": "PCS",
    "packing": "1",
}


# ==================== page ====================
st.title(":material/description: Item Templates")
st.caption("Creation, activation and description update sheets for the "
           "purchaser. Nothing here changes stock.")

tab_new, tab_saved, tab_layout = st.tabs([
    ":material/edit_document: New sheet",
    ":material/inventory: Saved sheets",
    ":material/table_chart: Layouts"])

# -------------------------------------------------- new sheet
with tab_new:
    kinds = list(TYPES)
    labels = [TYPES[k]["label"] for k in kinds]

    # Loading a sheet out of the archive has to happen here, at the top of a
    # fresh run, and nowhere else: Streamlit refuses to let a widget's value
    # be set once that widget has been drawn, and the Load button sits in the
    # other tab — which runs *after* every box on this one. So the button
    # only leaves a note, and the note is applied now.
    pend = st.session_state.pop("tpl_pending", None)
    if pend:
        k = pend["kind"]
        st.session_state["tpl_kind"] = TYPES[k]["label"]
        st.session_state[items_key(k)] = pend["items"]
        clear_entry(k)
        for wkey, val in pend["widgets"].items():
            st.session_state[wkey] = val
        st.session_state.pop("tpl_last", None)
        st.success(f"{pend['ref']} loaded — edit and generate again.")

    # Emptying the input boxes after an item is added is the same problem:
    # requested when Add was pressed, carried out here before they exist.
    _clr = st.session_state.pop("tpl_clear", None)
    if _clr:
        clear_entry(_clr)

    # "Edit it" takes an item off the sheet and puts it back in the boxes.
    # Same rule: the boxes do not exist yet on the run the button was
    # pressed, so the item travels here and is unpacked before they are made.
    _res = st.session_state.pop("tpl_restore", None)
    if _res:
        rk, ritem = _res
        for f in all_entry_fields(rk):
            v = ritem.get(f)
            key = entry_key(rk, f)
            if is_blank(v):
                st.session_state.pop(key, None)
            elif f in ("cost", "rsp"):
                st.session_state[key] = num(v)
            else:
                st.session_state[key] = str(v)

    choice = st.radio("Sheet type", labels, horizontal=True, key="tpl_kind",
                      label_visibility="collapsed")
    kind = kinds[labels.index(choice)]
    spec = TYPES[kind]

    left, right = st.columns([1.15, 1], gap="large")

    with left:
        tile_head("file-plus", "Sheet details",
                  f"{spec['purpose']} — goes to the purchaser")
        today = date.today()
        seed("tpl_date", today)
        d = st.date_input("Date", format="DD-MM-YYYY", key="tpl_date")
        header = {"date": d.strftime("%d-%m-%Y")}

        if "vendor" in spec["fields"]:
            header["vendor"] = st.text_input(
                FIELD_LABELS["vendor"], key=f"tpl_vendor_{kind}",
                placeholder="DIYANA FASHION LLC")
        if "maingrp" in spec["fields"]:
            header["maingrp"] = st.text_input(
                FIELD_LABELS["maingrp"], key=f"tpl_grp_{kind}",
                placeholder="LIFESTYLE → GARMENTS → Ladies Wear",
                help="The group path the item should sit under.")
        if "reason" in spec["fields"]:
            seed(f"tpl_reason_{kind}", "Description Mismatch")
            header["reason"] = st.text_input(
                FIELD_LABELS["reason"], key=f"tpl_reason_{kind}")
        if "remark" in spec["fields"]:
            header["remark"] = st.text_input(
                FIELD_LABELS["remark"], key=f"tpl_remark_{kind}",
                placeholder="For LPO")

    with right:
        tile_head("signature", "Signatories",
                  "Printed on the sheet, filled from Settings")
        sg = st.columns(3)
        for col, (key, label) in zip(sg, SIGN_SLOTS):
            with col:
                seed(f"tpl_sign_{key}", default(key, SIGN_FALLBACK[key]))
                header[key] = st.text_input(label, key=f"tpl_sign_{key}")
        st.caption("Change these once on the Settings page and they will be "
                   "the default here.")

    _lay = layout_for(kind)
    if _lay:
        st.caption(f":material/table_chart: Using your uploaded "
                   f"**{_lay.get('filename', 'template')}** — output will be "
                   f"that workbook, plus a PDF of it.")

    st.divider()
    tile_head("list-checks", "Add an item",
              "Type it, check the GP%, add it. The sheet builds up below.")

    ifields = ITEM_FIELDS[kind]
    if kind == "description":
        a1, a2 = st.columns(2)
        with a1:
            entry_widget(kind, "single")
            entry_widget(kind, "old_desc")
        with a2:
            entry_widget(kind, "outer")
            entry_widget(kind, "new_desc")
        with st.expander("Price columns (optional)"):
            e1, e2, e3 = st.columns(3)
            with e1:
                entry_widget(kind, "unit")
            with e2:
                entry_widget(kind, "cost")
            with e3:
                entry_widget(kind, "rsp")
    else:
        b1, b2, b3 = st.columns([1, 1, 2])
        with b1:
            entry_widget(kind, "single")
        with b2:
            entry_widget(kind, "outer")
        with b3:
            entry_widget(kind, "desc", label="Item name")
        c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 1.1])
        with c1:
            entry_widget(kind, "unit")
        with c2:
            entry_widget(kind, "packing")
        with c3:
            entry_widget(kind, "cost")
        with c4:
            entry_widget(kind, "rsp")
        with c5:
            # Live, because it is read off the widgets on this run rather
            # than from anything stored. Typing an RSP moves it immediately.
            g = gp_pct(entry_value(kind, "cost"), entry_value(kind, "rsp"))
            st.metric("GP%", "—" if g is None else f"{g:.2f}",
                      help="(RSP − COST) ÷ RSP × 100")
            if g is not None and g < 0:
                st.caption(":red[selling below cost]")

    typed = {f: entry_value(kind, f) for f in all_entry_fields(kind)}
    has_required = all(not is_blank(typed.get(f)) for f in REQUIRED[kind])

    add1, add2, add3 = st.columns([1.2, 1, 2.4])
    with add1:
        add = st.button(":material/add: Add to sheet", type="primary",
                        width="stretch", disabled=not has_required,
                        key=f"tpl_add_{kind}")
    with add2:
        if st.button(":material/backspace: Clear boxes", width="stretch",
                     key=f"tpl_clr_{kind}"):
            st.session_state["tpl_clear"] = kind
            st.rerun()
    with add3:
        if not has_required:
            need = ", ".join(HEADINGS.get(f, f) for f in REQUIRED[kind])
            st.caption(f"{need} is needed before an item can be added.")

    if add:
        items(kind).append(typed)
        st.session_state["tpl_clear"] = kind
        st.rerun()

    # -------- the sheet so far -------------------------------------------
    live = filled_rows(kind, items(kind))
    st.divider()

    if not live:
        st.info("No items yet. Add the first one above and it will appear "
                "here as it will print.")
    else:
        desc_key = "new_desc" if kind == "description" else "desc"
        tile_head("list-checks", f"On this sheet — {len(live)} item(s)",
                  "Exactly what will print, in this order")

        cols = [c for c in spec["cols"] if c[0] != "sno"]
        prev = pd.DataFrame([{
            "#": i,
            **{h: ("" if is_blank(r.get(k)) else str(r[k]).strip())
               for k, h, _w, _a, t in cols if t == "text"},
            **{h: num(r.get(k)) for k, h, _w, _a, t in cols if t == "num"},
            "GP%": gp_pct(r.get("cost"), r.get("rsp")),
        } for i, r in enumerate(live, 1)])
        # Blank barcodes print as NEED BARCODE on a creation sheet, so show
        # that here too rather than an empty cell that looks forgotten.
        for k, fill_with in spec["blanks"].items():
            h = HEADINGS.get(k)
            if h in prev.columns:
                prev[h] = prev[h].replace("", fill_with)

        st.dataframe(
            prev, hide_index=True, width="stretch",
            column_config={c: st.column_config.NumberColumn(format="%.2f")
                           for c in ("COST", "RSP", "GP%") if c in prev})

        s1, s2, s3 = st.columns([1.3, 1.3, 2])
        gps = prev["GP%"].dropna()
        if not gps.empty:
            s1.metric("Average GP%", f"{gps.mean():.2f}")
            s2.caption(f"Lowest {gps.min():.2f}%, highest {gps.max():.2f}%")
        with s3:
            pick = st.selectbox(
                "Item to change", range(1, len(live) + 1),
                format_func=lambda i: f"{i}. "
                + (str(live[i - 1].get(desc_key) or "").strip() or "(no name)"),
                key=f"tpl_pick_{kind}", label_visibility="collapsed")

        r1, r2, r3 = st.columns([1, 1, 2])
        with r1:
            if st.button(":material/edit: Edit it", width="stretch",
                         key=f"tpl_edit_{kind}",
                         help="Puts it back in the boxes above and takes it "
                              "off the sheet, so you can retype and add it."):
                back = items(kind).pop(pick - 1)
                st.session_state["tpl_restore"] = (kind, back)
                st.rerun()
        with r2:
            if st.button(":material/delete: Remove it", width="stretch",
                         key=f"tpl_rm_{kind}"):
                items(kind).pop(pick - 1)
                st.rerun()
        with r3:
            if st.button(":material/delete_sweep: Start the sheet again",
                         key=f"tpl_wipe_{kind}"):
                st.session_state[items_key(kind)] = []
                st.rerun()

        problems = []
        for i, r in enumerate(live, 1):
            problems += row_problems(kind, r, i)
        if problems:
            with st.expander(f":material/warning: {len(problems)} thing(s) "
                             f"to check before this goes up", expanded=True):
                for p in problems:
                    st.write("- " + p)
                st.caption("These are warnings, not blocks — a creation sheet "
                           "with no barcode yet is normal.")

    # -------- generate ---------------------------------------------------
    st.divider()
    gen_col, opt_col = st.columns([1, 2])
    with gen_col:
        go = st.button(":material/picture_as_pdf: Generate PDF",
                       type="primary", width="stretch",
                       disabled=not live, key="tpl_go")
    with opt_col:
        note = st.text_input("Note (kept with the saved copy, not printed)",
                             key="tpl_note", placeholder="optional")

    if go:
        ref = ""
        if store is not None and hasattr(store, "next_template_ref"):
            try:
                ref = store.next_template_ref(kind, d.strftime("%d%m%y"))
            except Exception:
                ref = ""
        if not ref:
            ref = (f"{spec['prefix']}-{d.strftime('%d%m%y')}-"
                   f"{datetime.now().strftime('%H%M')}")

        out = {"ref": ref, "kind": kind, "header": header, "rows": live,
               "note": note, "files": {}, "source": "built-in"}
        lay_doc = layout_for(kind)

        if lay_doc:
            # The store's own workbook. The .xlsx is the real output — the
            # PDF is that same file put through LibreOffice, so the two
            # cannot disagree.
            try:
                gps = [gp_pct(r.get("cost"), r.get("rsp")) for r in live]
                xlsx = xt.fill(lay_doc["workbook"], lay_doc["layout"],
                               header, live, gp_values=gps,
                               blanks=spec["blanks"],
                               page_fit=lay_doc.get("page_fit", True))
                out["files"][f"{ref}.xlsx"] = xlsx
                out["source"] = lay_doc.get("filename", "uploaded template")
                pdf, err = xt.to_pdf(xlsx)
                if pdf:
                    out["files"][f"{ref}.pdf"] = pdf
                else:
                    out["pdf_error"] = err
            except Exception as e:
                st.error(f"Could not fill your template: "
                         f"{type(e).__name__}: {e}")
                st.caption("The built-in layout was used instead. Check the "
                           "mapping on the **Layouts** tab.")
                lay_doc = None

        if not lay_doc:
            try:
                out["files"][f"{ref}.pdf"] = build_pdf(kind, header, live,
                                                       ref=ref)
            except Exception as e:
                st.error(f"Could not build the PDF: {type(e).__name__}: {e}")

        if out["files"]:
            main = f"{ref}.pdf" if f"{ref}.pdf" in out["files"] \
                else next(iter(out["files"]))
            out["filename"] = main
            st.session_state["tpl_last"] = out
            # Keep the copy first, so a failed save is reported before the
            # person walks off with the download.
            if store is not None:
                try:
                    ok, res = store.save_template(
                        kind=kind, ref=ref, header=header, rows=live,
                        pdf=out["files"].get(f"{ref}.pdf"),
                        filename=main, note=note)
                    out["saved"] = (ok, res)
                except Exception as e:
                    out["saved"] = (False, f"{type(e).__name__}: {e}")
            else:
                out["saved"] = (None, why)

    last = st.session_state.get("tpl_last")
    if last:
        st.success(f"**{last['ref']}** — {len(last['rows'])} row(s), "
                   f"{TYPES[last['kind']]['label'].lower()} sheet, "
                   f"from {last.get('source', 'built-in')}.")
        ok, res = last.get("saved", (None, ""))
        if ok is True:
            st.caption(":material/cloud_done: Copy kept in the database.")
        elif ok is False:
            st.warning(f"Generated, but **not saved**: {res}")
        else:
            st.caption(f":material/cloud_off: Not saved — {res}")
        if last.get("pdf_error"):
            st.warning(f"The .xlsx is ready, but no PDF: {last['pdf_error']}")

        dls = st.columns(max(len(last["files"]), 1))
        for col, (name, blob) in zip(dls, last["files"].items()):
            col.download_button(
                f":material/download: {name}", data=blob, file_name=name,
                mime=("application/pdf" if name.endswith(".pdf")
                      else XLSX_MIME),
                width="stretch", key=f"tpl_dl_{name}")

        pdf_blob = next((b for n, b in last["files"].items()
                         if n.endswith(".pdf")), None)
        if pdf_blob:
            with st.expander("Preview", expanded=True):
                try:
                    import pypdfium2 as pdfium
                    doc = pdfium.PdfDocument(pdf_blob)
                    for i in range(min(len(doc), 3)):
                        st.image(doc[i].render(scale=2).to_pil(),
                                 width="stretch")
                    if len(doc) > 3:
                        st.caption(f"{len(doc) - 3} more page(s) in the file.")
                except Exception:
                    st.caption("Preview needs pypdfium2 — the download above "
                               "is the real file either way.")

# -------------------------------------------------- saved sheets
with tab_saved:
    tile_head("pen-line", "Saved sheets",
              "Every generated sheet, newest first",
              ok=store is not None)

    if store is None:
        st.info(why)
    else:
        f1, f2 = st.columns([1, 2])
        with f1:
            pick = st.selectbox(
                "Type", ["All"] + [TYPES[k]["label"] for k in TYPES],
                key="tpl_filter")
        kind_f = None
        if pick != "All":
            kind_f = next(k for k in TYPES if TYPES[k]["label"] == pick)

        try:
            docs = store.list_templates(limit=200, kind=kind_f)
        except Exception as e:
            docs = []
            st.error(f"Could not read the list: {type(e).__name__}: {e}")

        if not docs:
            st.info("No sheets saved yet.")
        else:
            table = pd.DataFrame([{
                "Ref": t.get("ref", ""),
                "Type": TYPES.get(t.get("kind"), {}).get("label",
                                                         t.get("kind", "")),
                "Date": (t.get("header") or {}).get("date", ""),
                "Rows": t.get("n_rows", 0),
                "Vendor": (t.get("header") or {}).get("vendor", ""),
                "Note": t.get("note", ""),
                "Saved": t["at"].strftime("%d-%m-%Y %H:%M")
                if isinstance(t.get("at"), datetime) else "",
            } for t in docs])
            st.dataframe(table, hide_index=True, width="stretch")

            refs = [f'{t.get("ref", "?")}  ·  {t.get("n_rows", 0)} row(s)'
                    for t in docs]
            sel = st.selectbox("Open a sheet", refs, key="tpl_open")
            doc = docs[refs.index(sel)]

            b1, b2, b3 = st.columns(3)
            with b1:
                full = None
                try:
                    full = store.get_template(doc["id"])
                except Exception as e:
                    st.error(f"{type(e).__name__}: {e}")
                if full and full.get("pdf"):
                    st.download_button(
                        ":material/download: Download PDF",
                        data=full["pdf"],
                        file_name=full.get("filename",
                                           f'{doc.get("ref", "sheet")}.pdf'),
                        mime="application/pdf", width="stretch",
                        key="tpl_dl_saved")
                else:
                    st.caption("No PDF stored on this one.")
            with b2:
                if st.button(":material/edit: Load into the editor",
                             width="stretch", key="tpl_reload",
                             help="Copy this sheet back into the form so it "
                                  "can be corrected and generated again."):
                    k = doc.get("kind")
                    if k not in TYPES or not full:
                        st.error("This sheet cannot be reopened — it was "
                                 "saved by an older version.")
                    else:
                        keep = all_entry_fields(k)
                        loaded = [{f: it.get(f) for f in keep}
                                  for it in full.get("rows", [])]
                        h = full.get("header", {})
                        widgets = {f"tpl_{wid}_{k}": h[f]
                                   for f, wid in FIELD_WIDGET.items()
                                   if h.get(f)}
                        widgets.update({f"tpl_sign_{key}": h[key]
                                        for key, _lbl in SIGN_SLOTS
                                        if h.get(key)})
                        st.session_state["tpl_pending"] = {
                            "kind": k, "items": loaded,
                            "widgets": widgets,
                            "ref": doc.get("ref", "Sheet"),
                        }
                        st.rerun()
            with b3:
                if st.button(":material/delete: Delete", width="stretch",
                             key="tpl_del"):
                    if store.delete_template(doc["id"]):
                        st.success(f'{doc.get("ref")} deleted.')
                        st.rerun()
                    else:
                        st.error("Could not delete it.")

# -------------------------------------------------- layouts
with tab_layout:
    tile_head("file-plus", "Use your own Excel template",
              "Upload the sheet you already type into, per type",
              ok=store is not None)

    st.markdown(
        "The built-in layout is a rebuild from screenshots — close, but not "
        "your file. Upload the real `.xlsx` and the app writes the rows "
        "straight into it, so the logo, widths and formatting are yours. "
        "You get that workbook back, plus a PDF of it.")

    conv = xt.converter()
    if not conv:
        st.warning(
            "LibreOffice is not installed on this server, so an uploaded "
            "template can only come back as **.xlsx**, not PDF. To turn PDFs "
            "on, add `libreoffice-calc` to `packages.txt` and reboot the app. "
            "It is a large package and slows the first boot — if that is a "
            "problem, keep the built-in layout, which makes PDFs directly.")

    if store is None:
        st.info(why)
        st.caption("A layout has to be kept somewhere, so this needs the "
                   "database. The built-in layout still works without it.")
        st.stop()

    lk_labels = [TYPES[k]["label"] for k in TYPES]
    lk_choice = st.selectbox("Which sheet type is this template for?",
                             lk_labels, key="lay_kind")
    lkind = next(k for k in TYPES if TYPES[k]["label"] == lk_choice)

    existing = None
    try:
        existing = store.get_layout(lkind)
    except Exception as e:
        st.error(f"Could not read the saved layout: {type(e).__name__}: {e}")

    if existing:
        c1, c2 = st.columns([3, 1])
        c1.success(f"**{existing.get('filename', 'template')}** is in use for "
                   f"{lk_choice.lower()} sheets "
                   f"({existing.get('bytes', 0) / 1024:.0f} KB, saved "
                   f"{existing['at'].strftime('%d-%m-%Y')})."
                   if isinstance(existing.get("at"), datetime)
                   else f"**{existing.get('filename', 'template')}** is in use.")
        if c2.button(":material/delete: Stop using it", width="stretch",
                     key="lay_del",
                     help="Go back to the built-in layout. The saved sheets "
                          "already generated are not affected."):
            store.delete_layout(lkind)
            st.rerun()

    up = st.file_uploader(
        f"{lk_choice} template (.xlsx)", type=["xlsx", "xlsm"],
        key=f"lay_up_{lkind}",
        help="A blank copy is best — any rows already in it are treated as "
             "the space your items go into.")

    if up is not None:
        raw = up.getvalue()
        names = xt.sheet_names(raw)
        sheet = names[0] if names else None
        if len(names) > 1:
            sheet = st.selectbox("Which sheet?", names, key="lay_sheet")

        try:
            det = xt.scan(raw, sheet=sheet)
        except Exception as e:
            st.error(f"Could not read that file: {type(e).__name__}: {e}")
            det = None

        if det and not det.get("header_row"):
            for w in det["warnings"]:
                st.error(w)
        elif det:
            for w in det["warnings"]:
                st.warning(w)

            st.divider()
            st.markdown("#### Check what was found")
            st.caption("This is a guess. Correct anything that is wrong "
                       "before saving — a template mapped to the wrong "
                       "column produces a sheet that looks right and says "
                       "the wrong thing.")

            m1, m2, m3 = st.columns(3)
            m1.metric("Heading row", det["header_row"])
            m2.metric("Rows start at", det["data_start"])
            m3.metric("Room for", f'{det["data_rows"]} rows')

            # ---- columns -------------------------------------------------
            st.markdown("**Columns**")
            wanted = [k for k, *_ in TYPES[lkind]["cols"]]
            col_opts = ["— not used —"] + wanted
            chosen = {}
            grid = st.columns(3)
            for i, (cnum, text) in enumerate(sorted(det["headings"].items())):
                guess = next((k for k, v in det["columns"].items()
                              if v == cnum), None)
                if guess not in wanted:
                    guess = None
                with grid[i % 3]:
                    pickd = st.selectbox(
                        f'Col {xt.get_column_letter(cnum)} — "{text}"',
                        col_opts,
                        index=col_opts.index(guess) if guess else 0,
                        key=f"lay_col_{lkind}_{cnum}")
                    if pickd != "— not used —":
                        chosen[pickd] = cnum

            taken = list(chosen.values())
            dup = sorted({xt.get_column_letter(c) for c in taken
                          if taken.count(c) > 1})
            if dup:
                st.error("Two fields point at the same column ("
                         + ", ".join(dup) + "). One would overwrite the "
                         "other, so fix this before saving.")
            missing = [k for k in ("cost", "rsp",
                                   "new_desc" if lkind == "description"
                                   else "desc") if k not in chosen]
            if missing:
                st.error("Not mapped: "
                         + ", ".join(HEADINGS.get(m, m) for m in missing)
                         + ". These are needed for every row.")

            # ---- header fields and signatures ----------------------------
            f1, f2 = st.columns(2)
            with f1:
                st.markdown("**Header lines**")
                for f in TYPES[lkind]["fields"]:
                    pos = det["fields"].get(f)
                    st.write(f"- {FIELD_LABELS[f]}: "
                             + (f"`{xt.get_column_letter(pos[1])}{pos[0]}`"
                                if pos else ":red[not found — will be left "
                                            "blank]"))
            with f2:
                st.markdown("**Signature lines**")
                for key, label in SIGN_SLOTS:
                    pos = det["signs"].get(key)
                    st.write(f"- {label}: "
                             + (f"`{xt.get_column_letter(pos[1])}{pos[0]}`"
                                if pos else ":red[not found]"))

            fit = st.checkbox(
                "Force onto one page across when making the PDF", value=True,
                key=f"lay_fit_{lkind}",
                help="Leave on unless your template already prints correctly. "
                     "Anything the workbook already specifies is kept.")

            st.divider()
            t1, t2 = st.columns(2)
            with t1:
                test = st.button(":material/play_arrow: Test it with sample "
                                 "rows", width="stretch", key="lay_test",
                                 disabled=bool(missing or dup))
            with t2:
                save = st.button(":material/save: Save this layout",
                                 type="primary", width="stretch",
                                 key="lay_save", disabled=bool(missing or dup))

            layout = dict(det)
            layout["columns"] = chosen
            layout["sheet"] = sheet or det["sheet"]
            layout.pop("unmapped", None)
            layout.pop("warnings", None)
            layout.pop("headings", None)

            if test:
                sample = [{"subcat": "LADIES WEAR", "single": "",
                           "outer": "", "desc": "SAMPLE ITEM ONE",
                           "old_desc": "OLD SAMPLE NAME",
                           "new_desc": "NEW SAMPLE NAME", "unit": "PCS",
                           "packing": "1", "cost": 35.0, "rsp": 55.0},
                          {"subcat": "LADIES WEAR", "single": "1234567890123",
                           "outer": "", "desc": "SAMPLE ITEM TWO",
                           "old_desc": "ANOTHER OLD NAME",
                           "new_desc": "ANOTHER NEW NAME", "unit": "PCS",
                           "packing": "6", "cost": 29.16, "rsp": 49.0}]
                try:
                    xlsx = xt.fill(
                        raw, layout,
                        {"date": date.today().strftime("%d-%m-%Y"),
                         "vendor": "TEST VENDOR LLC",
                         "maingrp": "LIFESTYLE > GARMENTS",
                         "reason": "Test", "remark": "Test",
                         "prepared": "SWASTHIK", "purchaser": "IQBAL",
                         "approved": "THALLATH"},
                        sample,
                        gp_values=[gp_pct(r["cost"], r["rsp"])
                                   for r in sample],
                        blanks=TYPES[lkind]["blanks"], page_fit=fit)
                    st.session_state["lay_test_out"] = xlsx
                    pdf, err = xt.to_pdf(xlsx)
                    st.session_state["lay_test_pdf"] = pdf
                    st.session_state["lay_test_err"] = err
                except Exception as e:
                    st.error(f"Filling failed: {type(e).__name__}: {e}")

            if st.session_state.get("lay_test_out"):
                st.markdown("**Test sheet** — two sample rows, GP% worked out.")
                dl1, dl2 = st.columns(2)
                dl1.download_button(":material/download: test.xlsx",
                                    data=st.session_state["lay_test_out"],
                                    file_name="test.xlsx", mime=XLSX_MIME,
                                    width="stretch", key="lay_dl_x")
                tp = st.session_state.get("lay_test_pdf")
                if tp:
                    dl2.download_button(":material/download: test.pdf",
                                        data=tp, file_name="test.pdf",
                                        mime="application/pdf",
                                        width="stretch", key="lay_dl_p")
                    try:
                        import pypdfium2 as pdfium
                        doc = pdfium.PdfDocument(tp)
                        st.caption(f"{len(doc)} page(s). If that is more than "
                                   f"you expect, leave the one-page option on.")
                        st.image(doc[0].render(scale=2).to_pil(),
                                 width="stretch")
                    except Exception:
                        pass
                elif st.session_state.get("lay_test_err"):
                    st.info(st.session_state["lay_test_err"])

            if save:
                ok, res = store.save_layout(
                    kind=lkind, layout={**layout, "page_fit": fit},
                    workbook=raw, filename=up.name)
                if ok:
                    st.success(f"Saved. {lk_choice} sheets will now come from "
                               f"**{up.name}**.")
                    st.session_state.pop("lay_test_out", None)
                    st.session_state.pop("lay_test_pdf", None)
                else:
                    st.error(f"Could not save it: {res}")
