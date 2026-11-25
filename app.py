"""PDF to Excel converter — single-page Streamlit application."""

from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="PDF to Excel",
    page_icon=":material/picture_as_pdf:",
    layout="wide",
)

# Keep the app focused on the PDF converter only.
page = Path(__file__).parent / "views" / "pdf_tools.py"
if not page.is_file():
    st.error("The PDF conversion module could not be found.")
    st.stop()

exec(compile(page.read_text(encoding="utf-8"), str(page), "exec"),
     {"__name__": "__main__", "__file__": str(page)})
