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

from item_templates import (FIELD_LABELS, SIGN_SLOTS, TYPES, build_pdf,
                            filled_rows, gp_pct, is_blank, num, row_problems)

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


# ==================== grid ====================
# The editor holds only what is typed. GP% is never an editable column — it
# is worked out and shown in the preview below, so the printed figure and the
# figure on screen cannot drift apart.
GRID_COLS = {
    "creation": ["subcat", "single", "outer", "desc", "unit", "packing",
                 "cost", "rsp"],
    "activation": ["subcat", "single", "outer", "desc", "unit", "packing",
                   "cost", "rsp"],
    "description": ["single", "old_desc", "unit", "new_desc", "cost", "rsp"],
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


def blank_frame(kind, n=6):
    cols = GRID_COLS[kind]
    return pd.DataFrame(
        {c: pd.Series([None] * n, dtype="float64" if c in ("cost", "rsp")
                      else "object") for c in cols})


def grid_key(kind):
    return f"tpl_rows_{kind}"


def column_config(kind):
    cfg = {}
    for c in GRID_COLS[kind]:
        common = dict(label=HEADINGS.get(c, c), help=HELP.get(c))
        if c in ("cost", "rsp"):
            cfg[c] = st.column_config.NumberColumn(
                format="%.2f", min_value=0.0, step=0.01, **common)
        elif c in ("desc", "old_desc", "new_desc"):
            cfg[c] = st.column_config.TextColumn(width="large", **common)
        elif c in ("single", "outer"):
            # Text, not number: barcodes have leading zeros and are long
            # enough that a float would round them.
            cfg[c] = st.column_config.TextColumn(width="medium", **common)
        else:
            cfg[c] = st.column_config.TextColumn(width="small", **common)
    return cfg


def rows_from(df, kind):
    """Editor frame -> the list of dicts the PDF builder wants."""
    if df is None or df.empty:
        return []
    out = []
    for _, r in df.iterrows():
        d = {}
        for c in GRID_COLS[kind]:
            v = r.get(c)
            # NaN comes back from any cell the person cleared. It must not
            # reach the PDF, where it would print as "nan".
            if not isinstance(v, str):
                try:
                    if pd.isna(v):
                        v = None
                except (TypeError, ValueError):
                    pass
            d[c] = v
        out.append(d)
    return out


# ==================== page ====================
st.title(":material/description: Item Templates")
st.caption("Creation, activation and description update sheets for the "
           "purchaser. Nothing here changes stock.")

tab_new, tab_saved = st.tabs([":material/edit_document: New sheet",
                              ":material/inventory: Saved sheets"])

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
        st.session_state[grid_key(k)] = pend["frame"]
        st.session_state.pop(f"tpl_editor_{k}", None)
        for wkey, val in pend["widgets"].items():
            st.session_state[wkey] = val
        st.session_state.pop("tpl_last", None)
        st.success(f"{pend['ref']} loaded — edit and generate again.")

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

    st.divider()
    tile_head("list-checks", "Items",
              "Add a row for each item. Blank rows are dropped.")

    gk = grid_key(kind)
    if gk not in st.session_state:
        st.session_state[gk] = blank_frame(kind)

    edited = st.data_editor(
        st.session_state[gk], key=f"tpl_editor_{kind}",
        num_rows="dynamic", width="stretch", hide_index=True,
        column_config=column_config(kind))

    rows = rows_from(edited, kind)
    live = filled_rows(kind, rows)

    # -------- what will actually print ----------------------------------
    if live:
        desc_key = "new_desc" if kind == "description" else "desc"
        prev = pd.DataFrame([{
            "S.No": i,
            HEADINGS.get(desc_key): ("" if is_blank(r.get(desc_key))
                                     else str(r[desc_key]).strip()),
            "COST": num(r.get("cost")),
            "RSP": num(r.get("rsp")),
            "GP%": gp_pct(r.get("cost"), r.get("rsp")),
        } for i, r in enumerate(live, 1)])

        c1, c2 = st.columns([1.6, 1], gap="large")
        with c1:
            st.markdown("**GP% as it will print**")
            st.dataframe(prev, hide_index=True, width="stretch",
                         column_config={
                             "COST": st.column_config.NumberColumn(
                                 format="%.2f"),
                             "RSP": st.column_config.NumberColumn(
                                 format="%.2f"),
                             "GP%": st.column_config.NumberColumn(
                                 format="%.2f"),
                         })
        with c2:
            gps = prev["GP%"].dropna()
            st.metric("Rows", len(live))
            if not gps.empty:
                st.metric("Average GP%", f"{gps.mean():.2f}")
                lo, hi = gps.min(), gps.max()
                st.caption(f"Lowest {lo:.2f}%, highest {hi:.2f}% — "
                           f"GP% = (RSP − COST) ÷ RSP × 100")

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
    else:
        st.info("Nothing typed yet. Add at least one row to generate a sheet.")

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
        try:
            pdf = build_pdf(kind, header, rows, ref=ref)
        except Exception as e:
            st.error(f"Could not build the PDF: {type(e).__name__}: {e}")
            pdf = None

        if pdf:
            st.session_state["tpl_last"] = {
                "pdf": pdf, "ref": ref, "kind": kind,
                "filename": f"{ref}.pdf",
                "header": header, "rows": live, "note": note,
            }
            # Keep the copy first, so a failed save is reported before the
            # person walks off with the download.
            if store is not None:
                try:
                    ok, res = store.save_template(
                        kind=kind, ref=ref, header=header, rows=live,
                        pdf=pdf, filename=f"{ref}.pdf", note=note)
                    st.session_state["tpl_last"]["saved"] = (ok, res)
                except Exception as e:
                    st.session_state["tpl_last"]["saved"] = (
                        False, f"{type(e).__name__}: {e}")
            else:
                st.session_state["tpl_last"]["saved"] = (None, why)

    last = st.session_state.get("tpl_last")
    if last:
        st.success(f"**{last['ref']}** — {len(last['rows'])} row(s), "
                   f"{TYPES[last['kind']]['label'].lower()} sheet.")
        ok, res = last.get("saved", (None, ""))
        if ok is True:
            st.caption(":material/cloud_done: Copy kept in the database.")
        elif ok is False:
            st.warning(f"Generated, but **not saved**: {res}")
        else:
            st.caption(f":material/cloud_off: Not saved — {res}")

        st.download_button(
            ":material/download: Download " + last["filename"],
            data=last["pdf"], file_name=last["filename"],
            mime="application/pdf", width="stretch",
            key="tpl_dl_last")

        with st.expander("Preview", expanded=True):
            try:
                import pypdfium2 as pdfium
                doc = pdfium.PdfDocument(last["pdf"])
                for i in range(min(len(doc), 3)):
                    st.image(doc[i].render(scale=2).to_pil(), width="stretch")
                if len(doc) > 3:
                    st.caption(f"{len(doc) - 3} more page(s) in the file.")
            except Exception:
                st.caption("Preview needs pypdfium2 — the download above is "
                           "the real file either way.")

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
                        df = pd.concat(
                            [blank_frame(k, 0),
                             pd.DataFrame(full.get("rows", []))],
                            ignore_index=True)
                        for c in GRID_COLS[k]:
                            if c not in df.columns:
                                df[c] = None
                        h = full.get("header", {})
                        widgets = {f"tpl_{wid}_{k}": h[f]
                                   for f, wid in FIELD_WIDGET.items()
                                   if h.get(f)}
                        widgets.update({f"tpl_sign_{key}": h[key]
                                        for key, _lbl in SIGN_SLOTS
                                        if h.get(key)})
                        st.session_state["tpl_pending"] = {
                            "kind": k, "frame": df[GRID_COLS[k]],
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
