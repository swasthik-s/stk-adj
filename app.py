"""Document Converter — responsive Tailwind-powered Streamlit workspace."""
import streamlit as st
import st_tailwind as tw

st.set_page_config(page_title="Document Converter", page_icon=":material/description:", layout="wide", initial_sidebar_state="collapsed")
tw.initialize_tailwind()

if "ui_theme" not in st.session_state:
    st.session_state.ui_theme = "Light"

left, right = st.columns([5, 1], vertical_alignment="center")
with left:
    st.markdown('<div class="brand-lockup"><div class="brand-mark">DC</div><div><div class="brand-name">Document Converter</div><div class="brand-subtitle">A simple workspace for turning documents into data</div></div></div>', unsafe_allow_html=True)
with right:
    choice = st.selectbox("Appearance", ["Light", "Dark"], index=0 if st.session_state.ui_theme == "Light" else 1, label_visibility="collapsed", key="appearance_selector")
    st.session_state.ui_theme = choice

dark = choice == "Dark"
css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root{--bg:__BG__;--surface:__SURFACE__;--surface2:__SURFACE2__;--fg:__FG__;--muted:__MUTED__;--border:__BORDER__;--accent:#8424e8;--accent-hover:#701bd0;--soft:__SOFT__;--shadow:__SHADOW__;color-scheme:__SCHEME__}
html,body,[data-testid="stAppViewContainer"],[data-testid="stMain"]{background:var(--bg)!important;color:var(--fg)!important}
[data-testid="stHeader"]{background:transparent!important}
[data-testid="stAppViewContainer"] *{font-family:'DM Sans',sans-serif}
.block-container{max-width:1320px!important;padding:1.5rem clamp(1rem,3vw,2.5rem) 4rem!important}
.brand-lockup{display:flex;align-items:center;gap:13px;padding:3px 0 24px}
.brand-mark{width:44px;height:44px;border-radius:14px;display:grid;place-items:center;background:linear-gradient(135deg,#a855f7,#6d28d9);color:#fff;font:800 13px Manrope,sans-serif;box-shadow:0 7px 20px rgba(132,36,232,.25)}
.brand-name{font:800 17px Manrope,sans-serif;letter-spacing:-.4px;color:var(--fg)}
.brand-subtitle{font-size:12px;color:var(--muted);margin-top:3px}
[data-testid="stSelectbox"]{margin:0!important}[data-testid="stSelectbox"] label{display:none!important}
[data-testid="stSelectbox"]>div>div{background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:12px!important;min-height:42px}
[data-testid="stSelectbox"] *{color:var(--fg)!important}
h1,h2,h3,h4{font-family:Manrope,sans-serif!important;letter-spacing:-.55px!important;color:var(--fg)!important}
p,label,.stMarkdown{color:var(--fg)}
[data-testid="stCaptionContainer"] p{color:var(--muted)!important}
div[data-testid="stTabs"]{margin-top:26px}
div[data-testid="stTabs"] [role="tablist"]{display:flex;gap:9px;border-bottom:1px solid var(--border);padding:0 0 13px;overflow-x:auto}
div[data-testid="stTabs"] button[role="tab"]{border:1px solid var(--border)!important;border-radius:12px!important;padding:12px 22px!important;min-height:49px;background:var(--surface)!important;color:var(--muted)!important;font-weight:700!important;transition:all .18s ease}
div[data-testid="stTabs"] button[role="tab"]:hover{background:var(--soft)!important;border-color:rgba(132,36,232,.4)!important;color:var(--accent)!important}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"]{background:var(--accent)!important;color:#fff!important;border-color:var(--accent)!important;box-shadow:0 5px 14px rgba(132,36,232,.2)}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] p,div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] span{color:#fff!important}
div[data-testid="stTabs"] button[role="tab"] p{font-size:14px;font-weight:700;margin:0}
div[data-testid="stTabs"] [data-baseweb="tab-highlight"],div[data-testid="stTabs"] [data-baseweb="tab-border"]{display:none!important}
div[data-testid="stFileUploader"]{background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:18px!important;padding:18px!important;box-shadow:var(--shadow)}
div[data-testid="stFileUploader"] label,div[data-testid="stFileUploader"] small{color:var(--muted)!important}
div[data-testid="stFileUploader"] section{background:var(--surface2)!important;border:1.5px dashed rgba(132,36,232,.45)!important;border-radius:14px!important;padding:30px 22px!important;min-height:120px}
div[data-testid="stFileUploader"] section *{color:var(--muted)!important}
div[data-testid="stFileUploader"] button{background:var(--surface)!important;color:var(--fg)!important;border:1px solid var(--border)!important;border-radius:10px!important;font-weight:700!important}
div[data-testid="stFileUploader"] button:hover{border-color:var(--accent)!important;color:var(--accent)!important}
.stButton button,.stDownloadButton button{border-radius:11px!important;min-height:44px!important;font-weight:700!important;transition:all .18s ease!important}
.stButton button[kind="primary"],.stDownloadButton button[kind="primary"]{background:var(--accent)!important;border:1px solid var(--accent)!important;color:#fff!important}
.stButton button[kind="primary"]:hover,.stDownloadButton button[kind="primary"]:hover{background:var(--accent-hover)!important;border-color:var(--accent-hover)!important;box-shadow:0 5px 14px rgba(132,36,232,.22)}
.stButton button[kind="secondary"],.stDownloadButton button[kind="secondary"]{background:var(--surface)!important;border:1px solid var(--border)!important;color:var(--fg)!important}
.stButton button[kind="secondary"]:hover,.stDownloadButton button[kind="secondary"]:hover{background:var(--soft)!important;border-color:var(--accent)!important;color:var(--accent)!important}
.stButton button:focus-visible,.stDownloadButton button:focus-visible{outline:3px solid rgba(132,36,232,.3)!important;outline-offset:2px}
[data-testid="stTextInput"] input,[data-testid="stTextArea"] textarea,[data-testid="stSelectbox"] input{background:var(--surface)!important;color:var(--fg)!important;border-color:var(--border)!important;border-radius:10px!important}
[data-testid="stRadio"] label,[data-testid="stCheckbox"] label{color:var(--fg)!important}
div[data-testid="stRadio"] [role="radiogroup"]{gap:10px}
div[data-testid="stRadio"] label{background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:11px!important;padding:10px 14px!important}
div[data-testid="stRadio"] label p{color:var(--fg)!important}
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{border:1px solid var(--border);border-radius:13px;overflow:hidden}
[data-testid="stMetric"]{background:var(--surface);border:1px solid var(--border);padding:15px 17px;border-radius:13px}
[data-testid="stAlert"]{border-radius:12px!important}
hr{border-color:var(--border)!important}
.tool-heading{display:flex;align-items:center;gap:14px;padding:22px 0 18px}
.tool-icon{width:48px;height:48px;flex:0 0 48px;border-radius:15px;display:grid;place-items:center;background:var(--soft);color:var(--accent);font-size:19px;font-weight:800}
.tool-heading h2{font-size:25px!important;margin:0 0 4px!important;font-weight:800}
.tool-heading p{margin:0;color:var(--muted)!important;font-size:14px}
.paste-label{font-size:10px;font-weight:800;letter-spacing:1.4px;color:var(--muted);margin:0 0 8px}
.result-status{display:inline-flex;align-items:center;border:1px solid rgba(22,163,74,.25);background:rgba(22,163,74,.09);color:#16a34a;border-radius:999px;padding:7px 11px;font-size:12px;font-weight:700}
@media(max-width:700px){.block-container{padding:.8rem 1rem 3rem!important}.brand-lockup{padding-bottom:16px}.brand-subtitle{display:none}div[data-testid="stTabs"]{margin-top:16px}div[data-testid="stTabs"] button[role="tab"]{padding:10px 13px!important;min-height:44px}}
</style>
""".replace("__BG__","#0b0e14" if dark else "#f5f7fc").replace("__SURFACE__","#151a24" if dark else "#fff").replace("__SURFACE2__","#10151f" if dark else "#fbfaff").replace("__FG__","#f3f4f8" if dark else "#171923").replace("__MUTED__","#a0a9b8" if dark else "#737b8c").replace("__BORDER__","#2b3342" if dark else "#e5e8f0").replace("__SOFT__","rgba(168,85,247,.16)" if dark else "rgba(132,36,232,.09)").replace("__SHADOW__","0 12px 32px rgba(0,0,0,.24)" if dark else "0 12px 32px rgba(24,30,50,.055)").replace("__SCHEME__","dark" if dark else "light")
st.markdown(css, unsafe_allow_html=True)

# Tailwind utility classes handle the workspace typography and spacing;
# native Streamlit tabs preserve keyboard access and their built-in state.
with tw.container(classes="w-full mx-auto max-w-6xl px-0"):
    tw.write("WORKSPACE", classes="text-[10px] tracking-[0.2em] font-extrabold text-purple-500 mb-3")
    tw.write("What would you like to convert?", classes="text-3xl md:text-4xl font-extrabold tracking-tight mb-2")
    tw.write("Choose a tool to extract tables and information from your files", classes="text-sm md:text-base text-slate-500 leading-relaxed")

from views.pdf_tools import render as render_pdf_tools
from views.image_tools import render as render_image_tools

pdf_tab, image_tab = st.tabs([":material/picture_as_pdf:  PDF to Excel", ":material/image:  Image to Excel"])
with pdf_tab:
    render_pdf_tools()
with image_tab:
    render_image_tools()
