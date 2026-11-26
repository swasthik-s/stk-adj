"""Document Converter — compact Tailwind-styled workspace."""
import streamlit as st
import st_tailwind as tw

st.set_page_config(
    page_title="Document Converter",
    page_icon=":material/description:",
    layout="wide",
    initial_sidebar_state="collapsed",
)
tw.initialize_tailwind()

# Dark is the first-run default. A separate widget key ensures old saved
# selectbox state does not silently force the previous Light default.
if "ui_theme" not in st.session_state:
    st.session_state.ui_theme = "Dark"

header_left, header_right = st.columns([5, 1], vertical_alignment="center")
with header_left:
    st.markdown(
        '<div class="brand-lockup"><div class="brand-mark">DC</div>'
        '<div><div class="brand-name">Document Converter</div>'
        '<div class="brand-subtitle">A simple workspace for turning documents into data</div></div></div>',
        unsafe_allow_html=True,
    )
with header_right:
    choice = st.selectbox(
        "Appearance",
        ["Dark", "Light"],
        index=0 if st.session_state.ui_theme == "Dark" else 1,
        label_visibility="collapsed",
        key="appearance_selector_v2",
    )
    st.session_state.ui_theme = choice

dark = choice == "Dark"
css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root {
 --bg:__BG__; --surface:__SURFACE__; --surface2:__SURFACE2__;
 --fg:__FG__; --muted:__MUTED__; --border:__BORDER__;
 --accent:#a855f7; --accent-hover:#9333ea; --soft:__SOFT__;
 --shadow:__SHADOW__; color-scheme:__SCHEME__;
}
html,body,[data-testid="stAppViewContainer"],[data-testid="stMain"] {
 background:var(--bg)!important; color:var(--fg)!important;
 font-family:'DM Sans',sans-serif!important;
}
[data-testid="stHeader"] {background:transparent!important;}
[data-testid="stAppViewContainer"] p,
[data-testid="stAppViewContainer"] label,
[data-testid="stAppViewContainer"] input,
[data-testid="stAppViewContainer"] textarea,
[data-testid="stAppViewContainer"] button {font-family:'DM Sans',sans-serif;}
.block-container {max-width:1240px!important;padding:1.25rem clamp(1rem,3vw,2.5rem) 4rem!important;}
.brand-lockup {display:flex;align-items:center;gap:13px;padding:4px 0 18px;}
.brand-mark {width:44px;height:44px;flex:0 0 44px;border-radius:14px;display:grid;place-items:center;background:linear-gradient(135deg,#a855f7,#6d28d9);color:#fff;font:800 13px Manrope,sans-serif;box-shadow:0 8px 22px rgba(168,85,247,.28);}
.brand-name {font:800 17px Manrope,sans-serif;letter-spacing:-.4px;color:var(--fg);}
.brand-subtitle {font-size:12px;color:var(--muted);margin-top:3px;}
[data-testid="stSelectbox"] {margin:0!important;}
[data-testid="stSelectbox"] label {display:none!important;}
[data-testid="stSelectbox"]>div>div {background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:12px!important;min-height:42px;}
[data-testid="stSelectbox"] * {color:var(--fg)!important;}
h1,h2,h3,h4 {font-family:Manrope,sans-serif!important;letter-spacing:-.45px!important;color:var(--fg)!important;}
[data-testid="stCaptionContainer"] p {color:var(--muted)!important;}
/* Navigation tabs */
div[data-testid="stTabs"] {margin-top:12px;}
div[data-testid="stTabs"] [role="tablist"] {display:flex;gap:8px;border-bottom:1px solid var(--border);padding:0 0 12px;overflow-x:auto;}
div[data-testid="stTabs"] button[role="tab"] {border:1px solid var(--border)!important;border-radius:12px!important;padding:12px 20px!important;min-height:48px;background:var(--surface)!important;color:var(--muted)!important;font-weight:700!important;transition:all .18s ease;}
div[data-testid="stTabs"] button[role="tab"]:hover {background:var(--soft)!important;border-color:rgba(168,85,247,.45)!important;color:var(--fg)!important;}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {background:var(--soft)!important;color:var(--accent)!important;border-color:rgba(168,85,247,.55)!important;box-shadow:inset 0 -2px 0 var(--accent);}
div[data-testid="stTabs"] button[role="tab"] p {font-size:14px;font-weight:700;margin:0;color:inherit!important;}
div[data-testid="stTabs"] [data-baseweb="tab-highlight"],div[data-testid="stTabs"] [data-baseweb="tab-border"] {display:none!important;}
/* Upload cards. Do not apply the body font to Streamlit's icon glyphs. */
div[data-testid="stFileUploader"] {background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:18px!important;padding:16px!important;box-shadow:var(--shadow);}
div[data-testid="stFileUploader"] label,div[data-testid="stFileUploader"] small {color:var(--muted)!important;}
div[data-testid="stFileUploader"] section {background:var(--surface2)!important;border:1.5px dashed rgba(168,85,247,.55)!important;border-radius:14px!important;padding:28px 22px!important;min-height:118px;}
div[data-testid="stFileUploader"] section * {color:var(--muted)!important;}
div[data-testid="stFileUploader"] section button {background:var(--surface)!important;color:var(--fg)!important;border:1px solid var(--border)!important;border-radius:10px!important;font-weight:700!important;}
div[data-testid="stFileUploader"] section button:hover {border-color:var(--accent)!important;color:var(--accent)!important;}
/* The upload icon is a Material Symbols ligature. Hide it rather than allowing
   the page font to render the word "upload" as visible text. */
div[data-testid="stFileUploader"] section button [data-testid="stIconMaterial"] {display:none!important;}
.material-icons,.material-symbols-rounded,[data-testid="stIconMaterial"] {font-family:'Material Symbols Rounded'!important;font-weight:normal!important;font-style:normal!important;}
.stButton button,.stDownloadButton button {border-radius:11px!important;min-height:44px!important;font-weight:700!important;transition:all .18s ease!important;}
.stButton button[kind="primary"],.stDownloadButton button[kind="primary"] {background:var(--accent)!important;border:1px solid var(--accent)!important;color:#fff!important;}
.stButton button[kind="primary"]:hover,.stDownloadButton button[kind="primary"]:hover {background:var(--accent-hover)!important;border-color:var(--accent-hover)!important;box-shadow:0 5px 14px rgba(168,85,247,.22);}
.stButton button[kind="secondary"],.stDownloadButton button[kind="secondary"] {background:var(--surface)!important;border:1px solid var(--border)!important;color:var(--fg)!important;}
.stButton button[kind="secondary"]:hover,.stDownloadButton button[kind="secondary"]:hover {background:var(--soft)!important;border-color:var(--accent)!important;color:var(--accent)!important;}
.stButton button:focus-visible,.stDownloadButton button:focus-visible {outline:3px solid rgba(168,85,247,.3)!important;outline-offset:2px;}
[data-testid="stTextInput"] input,[data-testid="stTextArea"] textarea,[data-testid="stSelectbox"] input {background:var(--surface)!important;color:var(--fg)!important;border-color:var(--border)!important;border-radius:10px!important;}
[data-testid="stRadio"] label {color:var(--fg)!important;background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:11px!important;padding:10px 14px!important;}
[data-testid="stRadio"] label p {color:var(--fg)!important;}
[data-testid="stDataFrame"],[data-testid="stDataEditor"] {border:1px solid var(--border);border-radius:13px;overflow:hidden;}
[data-testid="stMetric"] {background:var(--surface);border:1px solid var(--border);padding:15px 17px;border-radius:13px;}
[data-testid="stAlert"] {border-radius:12px!important;}
hr {border-color:var(--border)!important;}
.tool-heading {display:flex;align-items:center;gap:14px;padding:18px 0 18px;}
.tool-icon {width:48px;height:48px;flex:0 0 48px;border-radius:15px;display:grid;place-items:center;background:var(--soft);color:var(--accent);font-size:17px;font-weight:800;}
.tool-heading h2 {font-size:25px!important;margin:0 0 4px!important;font-weight:800;}
.tool-heading p {margin:0;color:var(--muted)!important;font-size:14px;}
.paste-zone {padding:8px 0 2px;text-align:center;}
.paste-zone-note {font-size:12px;color:var(--muted);text-align:center;margin:5px 0 0;}
.result-status {display:inline-flex;align-items:center;border:1px solid rgba(34,197,94,.28);background:rgba(34,197,94,.1);color:#22c55e;border-radius:999px;padding:7px 11px;font-size:12px;font-weight:700;}
@media(max-width:700px) {
 .block-container {padding:.8rem 1rem 3rem!important;}
 .brand-lockup {padding-bottom:12px;}
 .brand-subtitle {display:none;}
 div[data-testid="stTabs"] {margin-top:8px;}
 div[data-testid="stTabs"] button[role="tab"] {padding:10px 13px!important;min-height:44px;}
 .tool-heading h2 {font-size:21px!important;}
}
</style>
""".replace("__BG__","#090d16" if dark else "#f5f7fc").replace(
 "__SURFACE__","#111827" if dark else "#ffffff"
).replace("__SURFACE2__","#0d1422" if dark else "#fbfaff").replace(
 "__FG__","#f1f5f9" if dark else "#171923"
).replace("__MUTED__","#94a3b8" if dark else "#737b8c").replace(
 "__BORDER__","#273449" if dark else "#e5e8f0"
).replace("__SOFT__","rgba(168,85,247,.14)" if dark else "rgba(132,36,232,.08)"
).replace("__SHADOW__","0 14px 34px rgba(0,0,0,.25)" if dark else "0 12px 32px rgba(24,30,50,.055)"
).replace("__SCHEME__","dark" if dark else "light")
st.markdown(css, unsafe_allow_html=True)

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
