"Standalone image and PDF extraction tools."
from pathlib import Path
import streamlit as st

st.set_page_config(page_title="Document Converter", page_icon=":material/document_scanner:", layout="wide")
root = Path(__file__).parent
pages = [
    st.Page(root / "views" / "pdf_tools.py", title="PDF to Excel", icon=":material/picture_as_pdf:", default=True),
    st.Page(root / "views" / "image_tools.py", title="Image to Excel", icon=":material/image:"),
]
st.navigation(pages, position="top").run()
