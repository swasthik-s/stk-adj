"""
Shams Al Madina — stock tools
=============================
Entry point. Declares the pages explicitly with st.Page / st.navigation
rather than relying on folder discovery, so navigation behaves the same
locally and on Streamlit Cloud.

    app.py                     this router
    pages/negative_stock.py    negative stock adjustments
    pages/pdf_tools.py         invoices and any other PDF -> Excel / txt

st.set_page_config belongs here and nowhere else — a page script that
calls it again raises an error.
"""

from pathlib import Path

import streamlit as st

st.set_page_config(page_title="Al Madina Stock Tools", page_icon=":material/inventory_2:",
                   layout="wide")

# PIN gate. Keeps a passer-by out; it is not security — see auth.py.
try:
    from auth import require_pin
    require_pin()
except ImportError:
    pass          # auth.py not deployed — run open rather than break

HERE = Path(__file__).parent
PAGES = [
    ("pages/negative_stock.py", "Negative stock", ":material/inventory_2:", True),
    ("pages/pdf_tools.py", "PDF to Excel / txt", ":material/picture_as_pdf:", False),
    ("pages/templates.py", "Item templates", ":material/description:", False),
    ("pages/history.py", "History", ":material/history:", False),
    ("pages/settings.py", "Settings", ":material/settings:", False),
]

# Daily housekeeping: prune old records if Settings has it switched on.
# Runs at most once per browser session and never blocks the page.
if not st.session_state.get("_pruned_today"):
    st.session_state["_pruned_today"] = True
    try:
        from mongo_store import MongoStore
        _cfg = st.secrets["mongo"]
        _s = MongoStore(_cfg["uri"], _cfg.get("db", "stockadj"),
                        timeout_ms=3000)
        _set = _s.get_settings()
        if _set.get("auto_prune"):
            _s.prune_all(_set)
    except Exception:
        pass

found, missing = [], []
for rel, title, icon, default in PAGES:
    if (HERE / rel).exists():
        found.append(st.Page(rel, title=title, icon=icon, default=default))
    else:
        missing.append(rel)

if not found:
    st.error("No page files found. They must sit beside this file:")
    st.code("\n".join(rel for rel, *_ in PAGES), language=None)
    st.caption(f"This file is at {Path(__file__).resolve()}")
    st.stop()

if missing:
    st.sidebar.warning("Not deployed: " + ", ".join(missing))

# Older Streamlit builds have no st.navigation, and position="top" only
# arrived in 1.44 — fall back cleanly rather than showing a stack trace.
if hasattr(st, "navigation"):
    try:
        st.navigation(found, position="top").run()
    except TypeError:
        st.navigation(found).run()      # older signature, sidebar nav
else:
    st.sidebar.info("This Streamlit version has no multipage navigation — "
                    "showing the negative stock tool only.")
    main = HERE / PAGES[0][0]
    exec(compile(main.read_text(), str(main), "exec"),
         {"__name__": "__main__", "__file__": str(main)})
