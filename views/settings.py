"""
Settings — what the app saves, how long it keeps it, and the defaults every
sheet starts from. Stored in the database, so a change here applies to every
browser and survives restarts.
"""

from datetime import datetime, timezone

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


st.title(":material/settings: Settings")

def _diagnose():
    """Return (store, reason). Each failure gets its own message — they used
    to share one, which blamed secrets when the real cause was an old file."""
    try:
        import mongo_store
    except Exception as e:
        return None, (f"`mongo_store.py` could not be imported "
                      f"({type(e).__name__}: {e}). Check it is in the repo "
                      f"root beside `app.py`.")
    if not hasattr(mongo_store.MongoStore, "get_settings"):
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

S = store.get_settings()
tab_store, tab_save, tab_defaults, tab_access = st.tabs([
    ":material/database: Storage", ":material/save: Saving",
    ":material/tune: Defaults", ":material/lock: Access"])


def _age(ts):
    if not ts:
        return "—"
    if ts.tzinfo is None:
        ts = ts.replace(tzinfo=timezone.utc)
    d = (datetime.now(timezone.utc) - ts).days
    return "today" if d == 0 else f"{d} day{'s' if d != 1 else ''} ago"


# =================================================================== storage
with tab_store:
    rows, total = store.storage()
    used_mb = (total or 0) / 1e6
    c1, c2 = st.columns([1, 2])
    c1.metric("Used", f"{used_mb:.1f} MB", "of 512 MB free tier",
              delta_color="off")
    c2.progress(min(used_mb / 512, 1.0),
                text=f"{used_mb / 512 * 100:.1f}% of the free tier")

    st.dataframe(pd.DataFrame([{
        "What": r["label"], "Records": r["count"],
        "Oldest": _age(r["oldest"]), "Newest": _age(r["newest"]),
        "Size": (f"{r['bytes'] / 1e6:.2f} MB" if r["bytes"] is not None
                 else "—")} for r in rows]),
        width="stretch", hide_index=True)

    st.subheader("Clean up old records")
    st.caption("Retention is set per kind of record. 0 keeps it forever. "
               "Nothing is deleted until you press Prune.")

    keep = {}
    # One column per collection — a fixed count would silently drop any
    # collection added later.
    cols = st.columns(max(len(rows), 1))
    for col, r in zip(cols, rows):
        k = f"keep_{r['key']}"
        keep[k] = col.number_input(f"{r['label']} — days", 0, 3650,
                                   int(S.get(k, 30)), 1, key=f"in_{k}")

    preview = {r["key"]: store.count_older(r["key"], keep[f"keep_{r['key']}"])
               for r in rows}
    n_total = sum(preview.values())
    if n_total:
        st.warning("Pruning would delete " + ", ".join(
            f"**{v}** {next(r['label'] for r in rows if r['key'] == k).lower()}"
            for k, v in preview.items() if v) + ".")
    else:
        st.success("Nothing is older than these limits.")

    # A widget's value can't be changed after it is drawn in the same run,
    # so the reset after a prune is requested now and applied next run —
    # before the checkbox exists.
    if st.session_state.pop("_reset_prune", False):
        st.session_state.pop("prune_confirm", None)
    if st.session_state.get("_prune_result"):
        st.success(st.session_state.pop("_prune_result"))

    p1, p2, p3 = st.columns([2, 2, 3])
    if p1.button(":material/save: Save these limits", width="stretch"):
        ok = store.save_settings(keep)
        (st.toast if hasattr(st, "toast") else st.success)(
            "Retention saved" if ok else "Could not save")
    sure = p3.checkbox("I understand this cannot be undone",
                       key="prune_confirm")
    if p2.button(":material/delete_sweep: Prune now", type="primary",
                 width="stretch", disabled=not (n_total and sure)):
        store.save_settings(keep)
        done = {k: store.prune(k, keep[f"keep_{k}"]) for k in preview}
        st.session_state["_prune_result"] = (
            "Deleted " + ", ".join(f"{v} {k}" for k, v in done.items() if v)
            + "." if any(done.values()) else "Nothing needed deleting.")
        st.session_state["_reset_prune"] = True
        st.rerun()

    auto = st.toggle("Prune automatically once a day, using the limits above",
                     value=bool(S.get("auto_prune")), key="auto_prune_t")
    if auto != bool(S.get("auto_prune")):
        store.save_settings({"auto_prune": auto, **keep})
        st.rerun()

# ==================================================================== saving
with tab_save:
    st.caption("What the app writes to the database on its own.")
    a = st.toggle("Save a session automatically when a new negative report "
                  "is loaded", value=bool(S["autosave_session"]))
    b = st.toggle("Save an upload snapshot (totals only — feeds the history "
                  "trend)", value=bool(S["save_snapshots"]))
    c = st.toggle("Save every matching run", value=bool(S["save_runs"]))
    st.caption("Turning these off stops new records. It does not delete "
               "existing ones — use Storage for that. The Save now button on "
               "the Negative stock page always works regardless.")
    if st.button(":material/save: Save", key="save_saving", type="primary"):
        ok = store.save_settings({"autosave_session": a,
                                  "save_snapshots": b, "save_runs": c})
        (st.toast if hasattr(st, "toast") else st.success)(
            "Saved" if ok else "Could not save")

# ================================================================== defaults
with tab_defaults:
    st.caption("What every new sheet starts with. You can still change any "
               "of them on the Negative stock page for a single run.")
    st.markdown("**Signatories and sheets**")
    d1, d2, d3 = st.columns(3)
    prepared = d1.text_input("Prepared by", S["prepared"])
    checked = d2.text_input("Checked by", S["checked"])
    verified = d3.text_input("Verified by", S["verified"])
    st.markdown("**Item template signatories**")
    t1, t2 = st.columns(3)[:2]
    purchaser = t1.text_input("Concerned purchaser", S.get("purchaser", ""))
    approved = t2.text_input("Approved by", S.get("approved", ""),
                             help="Printed on creation, activation and "
                                  "description sheets.")

    st.markdown("**Sheets**")
    e1, e2, e3 = st.columns(3)
    prefix = e1.text_input("Import file prefix", S["txt_prefix"])
    per = e2.number_input("Pairs per sheet", 1, 60, int(S["pairs_per_sheet"]))
    remarks = e3.text_input("Remarks", S["remarks"])

    st.markdown("**Matching**")
    f1, f2, f3 = st.columns(3)
    br = f1.slider("Bundle break strictness", 0.70, 1.00,
                   float(S["br_thresh"]), 0.01)
    dt = f2.slider("Max cost drift %", 2.0, 60.0, float(S["drift_tol"]), 1.0)
    pt = f3.slider("Wrong sale price tolerance", 0.05, 0.60,
                   float(S["price_tol"]), 0.05)

    g1, g2 = st.columns([1, 3])
    if g1.button(":material/save: Save defaults", type="primary",
                 width="stretch"):
        ok = store.save_settings({
            "prepared": prepared, "checked": checked, "verified": verified,
            "purchaser": purchaser, "approved": approved,
            "txt_prefix": prefix, "pairs_per_sheet": int(per),
            "remarks": remarks, "br_thresh": br, "drift_tol": dt,
            "price_tol": pt})
        (st.toast if hasattr(st, "toast") else st.success)(
            "Defaults saved — they apply from the next page load"
            if ok else "Could not save")
    if g2.button(":material/restart_alt: Reset to factory defaults"):
        store.save_settings({k: v for k, v in store.DEFAULTS.items()
                             if k not in ("pin",)})
        st.rerun()

# ==================================================================== access
with tab_access:
    has_db_pin = bool(S.get("pin"))
    st.caption("The PIN set here overrides the one in Streamlit secrets. "
               "Clear it to fall back to secrets.")
    st.info("Current source: **" + ("this page" if has_db_pin
                                     else "Streamlit secrets / auth.py") + "**")
    with st.form("pin_change"):
        new1 = st.text_input("New PIN", type="password")
        new2 = st.text_input("Repeat new PIN", type="password")
        go = st.form_submit_button("Change PIN", type="primary")
    if go:
        if not new1 or len(new1) < 4:
            st.error("Use at least 4 characters.")
        elif new1 != new2:
            st.error("The two entries do not match.")
        else:
            store.save_settings({"pin": new1})
            st.success("PIN changed. Anyone signed in stays signed in; new "
                       "sessions need the new PIN.")
    if has_db_pin and st.button("Clear PIN set here"):
        store.save_settings({"pin": ""})
        st.rerun()

    st.divider()
    if st.button(":material/logout: Sign out of this browser"):
        st.session_state["authed"] = False
        st.rerun()
