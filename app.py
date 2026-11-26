"""Standalone PDF and image extraction tools with a tabbed interface."""
import streamlit as st
from views.pdf_tools import render as render_pdf_tools
from views.image_tools import render as render_image_tools

st.set_page_config(
    page_title="Document Converter",
    page_icon=":material/document_scanner:",
    layout="wide",
)

st.title(":material/document_scanner: Document Converter")
pdf_tab, image_tab = st.tabs([
    ":material/picture_as_pdf: PDF to Excel",
    ":material/image: Image to Excel",
])

with pdf_tab:
    render_pdf_tools()

with image_tab:
    render_image_tools()
