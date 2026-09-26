"""
History — every saved session, newest first, with the sheets rebuilt on
demand. Works without uploading anything: the saved candidates carry every
cost and conversion the sheets need.

The sheet builders live in negative_stock.py. Rather than copy them here and
let the two drift apart, this page lifts just their definitions out of that
file, so there is one source of truth.
"""

import ast
from pathlib import Path

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


st.title(":material/history: History")

# ---------------------------------------------------------------- store
def _diagnose():
    """Return (store, reason). Each failure gets its own message — they used
    to share one, which blamed secrets when the real cause was an old file."""
    try:
        import mongo_store
    except Exception as e:
        return None, (f"`mongo_store.py` could not be imported "
                      f"({type(e).__name__}: {e}). Check it is in the repo "
                      f"root beside `app.py`.")
    if not hasattr(mongo_store.MongoStore, "list_sessions"):
        return None, ("The server is running an **old `mongo_store.py`** — "
                      "it has no settings support. Push the new file, then "
                      "**Manage app → Reboot**. A reboot is needed even after "
                      "pushing: Streamlit keeps imported files in memory.")
    try:
        cfg = st.secrets["mongo"]
        uri = cfg["uri"]
    except Exception:
        return None, ("No `[mongo]` block with a `uri` in your secrets. Add it "
                      "under Manage app → Settings → Secrets.")
    try:
        s = mongo_store.MongoStore(uri, cfg.get("db", "stockadj"))
        ok, msg = s.check()
        if not ok:
            return None, f"Database unreachable: {msg}"
        s.ensure_indexes()
        return s, "ok"
    except Exception as e:
        return None, f"Could not connect: {type(e).__name__}: {e}"


store, _why = _diagnose()
if store is None:
    st.error(_why)
    st.stop()


# ---------------------------------------------------------------- builders
@st.cache_resource(show_spinner=False)
def _builders():
    """Pull the sheet builders out of negative_stock.py without running its
    page code."""
    src_path = Path(__file__).with_name("negative_stock.py")
    src = src_path.read_text()
    tree = ast.parse(src)
    want = {"Line", "write_adjustment", "build_sheet", "build_workbook",
            "make_txt", "_n", "make_print_sheet", "_wb_bytes"}
    parts = []
    for node in tree.body:
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            parts.append(ast.get_source_segment(src, node))
        elif isinstance(node, (ast.FunctionDef, ast.ClassDef)) \
                and node.name in want:
            seg = ast.get_source_segment(src, node)
            deco = "".join(f"@{ast.get_source_segment(src, d)}\n"
                           for d in node.decorator_list)
            parts.append(deco + seg)
        elif isinstance(node, ast.Assign):
            names = [t.id for t in node.targets if isinstance(t, ast.Name)]
            # module constants only: BOX, GREY, PRINT_COLS, COST_DP …
            if names and all(n.upper() == n for n in names):
                parts.append(ast.get_source_segment(src, node))
    ns = {}
    exec(compile("\n".join(p for p in parts if p), str(src_path), "exec"), ns)
    return ns


try:
    B = _builders()
except Exception as e:
    B = None
    st.warning(f"Sheet rebuild unavailable: {type(e).__name__}: {e}")

XL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------- list
sessions = store.list_sessions()
if not sessions:
    st.info("Nothing saved yet. A session is saved automatically each time "
            "you load a new negative stock report, and whenever you press "
            "Save on the Negative stock page.")
    st.stop()

rows = []
for x in sessions:
    src = x.get("source", {}) or {}
    rows.append({
        "When": x["at"].strftime("%d-%b %H:%M"),
        "File": src.get("negative_file", ""),
        "Lines": x.get("neg_lines", 0),
        "Negative (AED)": x.get("neg_value", 0.0),
        "Pairs": x.get("n_candidates", 0),
        "Clears (AED)": x.get("cleared", 0.0),
        "How": x.get("note", ""),
        "_id": str(x["_id"]),
    })
hist = pd.DataFrame(rows)

c1, c2, c3 = st.columns(3)
c1.metric("Sessions", len(hist))
c2.metric("Latest negative", f"AED {hist['Negative (AED)'].iloc[0]:,.0f}")
if len(hist) > 1:
    delta = hist["Negative (AED)"].iloc[0] - hist["Negative (AED)"].iloc[-1]
    c3.metric("Since first", f"AED {delta:+,.0f}",
              "better" if delta > 0 else "worse", delta_color="off")

st.dataframe(hist.drop(columns="_id"), width="stretch",
             hide_index=True, height=min(420, 38 + 35 * len(hist)),
             column_config={
                 "Negative (AED)": st.column_config.NumberColumn(
                     format="%.2f"),
                 "Clears (AED)": st.column_config.NumberColumn(format="%.2f")})

# ---------------------------------------------------------------- open one
st.divider()
labels = [f"{r['When']}  ·  {r['File'] or 'no file name'}  ·  "
          f"{r['Pairs']} pairs" for r in rows]
pick = st.selectbox("Open a session", labels)
chosen = rows[labels.index(pick)]
full = store.get_session(chosen["_id"])
if not full:
    st.error("Could not load that session.")
    st.stop()

cand = pd.DataFrame(full.get("candidates") or [])
negs = pd.DataFrame(full.get("negatives") or [])
settings = full.get("settings") or {}

def _load_into_workspace():
    """Hand the saved session to the Negative stock page and switch to it."""
    c = cand.copy()
    if len(c):
        c["use"] = c["use"].fillna(True).astype(bool) \
            if "use" in c.columns else True
    st.session_state["restored"] = {
        "id": chosen["_id"], "when": chosen["When"],
        "source": full.get("source", {}),
        "negatives": full.get("negatives") or [],
        "settings": settings,
    }
    st.session_state["cand"] = c if len(c) else None
    st.session_state["session_id"] = chosen["_id"]
    st.session_state["batch"] = 0
    st.session_state.pop("dropped", None)
    st.session_state.pop("bulk", None)
    st.session_state.pop("last_sheet", None)


lc1, lc2 = st.columns([1, 3])
if lc1.button(":material/upload_file: Load into workspace", type="primary",
              width="stretch"):
    _load_into_workspace()
    try:
        st.switch_page("pages/negative_stock.py")
    except Exception:
        st.success("Loaded. Open **Negative stock** from the top bar.")
lc2.caption("Opens this session on the Negative stock page — overview, "
            "check sheet and build sheets all work without uploading. "
            "Add the masterlist there only if you want to re-run matching.")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Negative lines", f"{full.get('neg_lines', 0):,}")
m2.metric("Negative value", f"AED {full.get('neg_value', 0):,.0f}")
m3.metric("Pairs", len(cand))
m4.metric("Clears", f"AED {full.get('cleared', 0):,.0f}")

if len(negs):
    with st.expander("Negative stock by category"):
        st.dataframe(
            negs.groupby("category", dropna=False)
            .agg(Lines=("val", "size"), Value=("val", "sum"))
            .sort_values("Value").reset_index(),
            width="stretch", hide_index=True,
            column_config={"Value": st.column_config.NumberColumn(
                format="%.2f")})

if not len(cand):
    st.caption("No matching was run in this session, so there are no "
               "sheets to rebuild.")
else:
    with st.expander(f"Pairs ({len(cand)})"):
        show = [c for c in ["neg_bc", "neg_desc", "neg_qty", "par_bc",
                            "par_desc", "outers_needed", "conv", "par_cost"]
                if c in cand.columns]
        st.dataframe(cand[show], width="stretch", hide_index=True,
                     column_config={
                         "neg_bc": st.column_config.TextColumn("Single"),
                         "par_bc": st.column_config.TextColumn("Outer")})

    if B is not None:
        use = cand[cand["use"].fillna(True).astype(bool)] \
            if "use" in cand.columns else cand
        date = settings.get("date", "")
        prep = settings.get("prepared", "")
        lines = [B["Line"](r["par_bc"], r["par_desc"],
                           "OFR" if r["conv"] != 1 else "PCS",
                           r["outers_needed"], r["par_cost"],
                           r["neg_bc"], r["neg_desc"], "PCS", r["conv"])
                 for r in use.to_dict("records")]
        chunks = [lines[i:i + 11] for i in range(0, len(lines), 11)]
        stamp = chosen["When"].replace(" ", "_").replace(":", "")

        st.caption("Rebuilt from the saved session — identical to what the "
                   "Negative stock page produced at the time.")
        d1, d2, d3 = st.columns(3)
        d1.download_button(
            ":material/download: Check sheet", B["make_print_sheet"](use, date, prep),
            f"CHECK_{stamp}.xlsx", XL, width="stretch",
            type="primary")
        d2.download_button(
            f":material/download: Adjustments ({len(chunks)} sheets)",
            B["build_workbook"](chunks, "OUTER BREAK FOR NEGATIVE STOCK",
                                date, prep, "IRSHAD", "THALLATH"),
            f"ADJUSTMENTS_{stamp}.xlsx", XL, width="stretch")
        txt = "".join(B["make_txt"](c, "SML")[0].decode() for c in chunks)
        d3.download_button(":material/download: iTrade import", txt.encode(),
                           f"ADJUSTMENTS_{stamp}.txt", "text/plain",
                           width="stretch")

with st.expander("Delete this session"):
    st.caption("Removes it from history permanently.")
    if st.button("Delete", key=f"del_{chosen['_id']}"):
        store.delete_session(chosen["_id"])
        st.rerun()
