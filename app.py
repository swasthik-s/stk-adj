"""Document Converter — a polished, responsive Streamlit workspace."""
import streamlit as st

st.set_page_config(
    page_title="Document Converter",
    page_icon=":material/description:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Tailwind utility support (the app's core layout remains resilient if the
# optional injector is unavailable).
try:
    import st_tailwind as tw
    tw.initialize_tailwind()
except Exception:
    tw = None

if "ui_theme" not in st.session_state:
    st.session_state.ui_theme = "Light"

top_left, top_right = st.columns([5, 1], vertical_alignment="center")
with top_left:
    st.markdown(
        '<div class="brand-lockup"><div class="brand-mark">DC</div>'
        '<div><div class="brand-name">Document Converter</div>'
        '<div class="brand-subtitle">A simple workspace for turning documents into data</div></div></div>',
        unsafe_allow_html=True,
    )
with top_right:
    theme_choice = st.selectbox(
        "Appearance", ["Light", "Dark"],
        index=0 if st.session_state.ui_theme == "Light" else 1,
        label_visibility="collapsed",
        key="appearance_selector",
    )
    st.session_state.ui_theme = theme_choice

dark = st.session_state.ui_theme == "Dark"
theme_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root {
  --app-bg: __BG__; --surface: __SURFACE__; --surface-2: __SURFACE2__;
  --text: __TEXT__; --muted: __MUTED__; --border: __BORDER__;
  --accent: #8424e8; --accent-hover: #701bd0; --accent-soft: rgba(132,36,232,.10);
  --shadow: __SHADOW__;
}
html, body, [data-testid="stAppViewContainer"] {background:var(--app-bg)!important;color:var(--text)!important;}
[data-testid="stHeader"] {background:transparent!important;}
[data-testid="stAppViewContainer"] * {font-family:'DM Sans',sans-serif;}
.block-container {max-width:1180px;padding-top:1.2rem;padding-bottom:3.5rem;}
.brand-lockup {display:flex;align-items:center;gap:12px;padding:6px 0 22px;}
.brand-mark {width:42px;height:42px;border-radius:13px;display:grid;place-items:center;background:linear-gradient(135deg,#9b4dff,#6d20c9);color:#fff;font:800 13px Manrope,sans-serif;box-shadow:0 6px 18px rgba(132,36,232,.25);}
.brand-name {font:800 17px Manrope,sans-serif;letter-spacing:-.4px;color:var(--text);}
.brand-subtitle {font-size:12px;color:var(--muted);margin-top:2px;}
[data-testid="stSelectbox"] {margin:0!important;}
[data-testid="stSelectbox"] label {display:none!important;}
[data-testid="stSelectbox"] > div > div {background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:11px!important;}
h1,h2,h3 {font-family:Manrope,sans-serif!important;letter-spacing:-.6px!important;color:var(--text)!important;}
p, label, .stMarkdown {color:var(--text);}
div[data-testid="stTabs"] {margin-top:8px;}
div[data-testid="stTabs"] [role="tablist"] {gap:8px;border-bottom:1px solid var(--border);padding-bottom:10px;}
div[data-testid="stTabs"] button[role="tab"] {border:1px solid transparent;border-radius:10px;padding:10px 17px;background:transparent;color:var(--muted);font-weight:600;}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {background:var(--accent-soft);color:var(--accent);border-color:rgba(132,36,232,.22);}
div[data-testid="stTabs"] button[role="tab"] p {font-size:14px;font-weight:700;}
div[data-testid="stFileUploader"] {background:var(--surface);border:1px solid var(--border);border-radius:18px;padding:18px;box-shadow:var(--shadow);}
div[data-testid="stFileUploader"] section {background:var(--surface-2)!important;border:1.5px dashed rgba(132,36,232,.42)!important;border-radius:14px!important;padding:28px 18px!important;}
div[data-testid="stFileUploader"] section small {color:var(--muted)!important;}
div[data-testid="stFileUploader"] button, .stDownloadButton button {border-radius:10px!important;font-weight:700!important;}
.stButton button[kind="primary"], .stDownloadButton button[kind="primary"] {background:var(--accent)!important;border:1px solid var(--accent)!important;color:#fff!important;}
.stButton button[kind="primary"]:hover, .stDownloadButton button[kind="primary"]:hover {background:var(--accent-hover)!important;border-color:var(--accent-hover)!important;}
.stButton button, .stDownloadButton button {border-radius:10px!important;min-height:42px;font-weight:600!important;}
[data-testid="stTextInput"] input,[data-testid="stTextArea"] textarea,[data-testid="stSelectbox"] input {background:var(--surface)!important;color:var(--text)!important;border-color:var(--border)!important;border-radius:10px!important;}
[data-testid="stDataFrame"],[data-testid="stDataEditor"] {border:1px solid var(--border);border-radius:12px;overflow:hidden;}
[data-testid="stMetric"] {background:var(--surface);border:1px solid var(--border);padding:14px 16px;border-radius:13px;}
[data-testid="stAlert"] {border-radius:12px;}
hr {border-color:var(--border)!important;}
div[data-testid="stRadio"] [role="radiogroup"] {gap:10px;}
div[data-testid="stRadio"] label {background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:9px 13px;}
.caption, [data-testid="stCaptionContainer"] {color:var(--muted)!important;}
@media(max-width:700px) {
 .block-container {padding-left:1rem;padding-right:1rem;padding-top:.7rem;}
 .brand-subtitle {display:none;}
 div[data-testid="stTabs"] button[role="tab"] {padding:9px 11px;}
}
</style>
""".replace("__BG__", "#0d1018" if dark else "#f5f7fc").replace(
    "__SURFACE__", "#171b26" if dark else "#ffffff"
).replace("__SURFACE2__", "#111621" if dark else "#fbfaff").replace(
    "__TEXT__", "#edf0f7" if dark else "#171923"
).replace("__MUTED__", "#9ba5b7" if dark else "#737b8c").replace(
    "__BORDER__", "#2a3140" if dark else "#e7e9f0"
).replace("__SHADOW__", "0 10px 30px rgba(0,0,0,.20)" if dark else "0 10px 28px rgba(24,30,50,.05)")
st.markdown(theme_css, unsafe_allow_html=True)

st.markdown(
    '<div class="workspace-intro"><div class="eyebrow">WORKSPACE</div>'
    '<h1>What would you like to convert?</h1>'
    '<p>Choose a tool to extract tables and information from your files.</p></div>',
    unsafe_allow_html=True,
)
st.markdown("""
<style>
.workspace-intro {padding:12px 0 16px;}
.workspace-intro .eyebrow {font-size:10px;letter-spacing:1.8px;font-weight:800;color:#9b62dc;margin-bottom:8px;}
.workspace-intro h1 {font-size:clamp(25px,3vw,34px);margin:0 0 6px;font-weight:800;}
.workspace-intro p {margin:0;color:var(--muted);font-size:14px;}
</style>
""", unsafe_allow_html=True)

from views.pdf_tools import render as render_pdf_tools
from views.image_tools import render as render_image_tools

pdf_tab, image_tab = st.tabs([
    ":material/picture_as_pdf:  PDF to Excel",
    ":material/image:  Image to Excel",
])
with pdf_tab:
    render_pdf_tools()
with image_tab:
    render_image_tools()
