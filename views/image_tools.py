"""Image to Excel with local OCR and a streamlined upload interface."""
import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image

try:
    import numpy as np
    from rapidocr import RapidOCR
    OCR_ERROR = None
except Exception as exc:
    np, RapidOCR = None, None
    OCR_ERROR = f"{type(exc).__name__}: {exc}"

try:
    from streamlit_paste_button import paste_image_button
except Exception:
    paste_image_button = None


@st.cache_resource(show_spinner=False)
def engine():
    return RapidOCR()


def visual_rows(result, tolerance=14):
    boxes = getattr(result, "boxes", None)
    texts = getattr(result, "txts", None)
    if boxes is None or texts is None:
        return []
    tokens = []
    for box, text in zip(boxes, texts):
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        tokens.append((sum(ys) / len(ys), min(xs), max(xs), str(text)))
    tokens.sort(key=lambda item: (item[0], item[1]))
    grouped = []
    for token in tokens:
        row = next((r for r in reversed(grouped) if abs(r[0] - token[0]) <= tolerance), None)
        if row is None:
            grouped.append([token[0], [token]])
        else:
            row[1].append(token)
            row[0] = sum(t[0] for t in row[1]) / len(row[1])
    output = []
    for _, row in grouped:
        row.sort(key=lambda item: item[1])
        widths = [(item[2] - item[1]) / max(1, len(item[3])) for item in row]
        gap = max(18, min(70, sorted(widths)[len(widths) // 2] * 2.8))
        cells, current, right = [], [], None
        for token in row:
            if current and right is not None and token[1] - right > gap:
                cells.append(" ".join(t[3] for t in current).strip())
                current = []
            current.append(token)
            right = token[2]
        if current:
            cells.append(" ".join(t[3] for t in current).strip())
        if cells:
            output.append(cells)
    return output


def make_frame(rows, mode):
    if not rows:
        return pd.DataFrame()
    if mode == "Only Table":
        rows = [row for row in rows if len([v for v in row if str(v).strip()]) >= 2]
    if not rows:
        return pd.DataFrame()
    width = max(map(len, rows))
    rows = [row + [""] * (width - len(row)) for row in rows]
    first = rows[0]
    words = sum(bool(re.search(r"[A-Za-z]", value)) for value in first)
    if len(first) >= 2 and words >= max(2, len(first) // 2):
        headers, seen = [], {}
        for i, value in enumerate(first):
            name = re.sub(r"\s+", " ", str(value)).strip() or f"Column {i + 1}"
            seen[name] = seen.get(name, 0) + 1
            headers.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
        return pd.DataFrame(rows[1:], columns=headers)
    return pd.DataFrame(rows, columns=[f"Column {i + 1}" for i in range(width)])


def xlsx_bytes(frame):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name="Extracted Data")
        ws = writer.sheets["Extracted Data"]
        for col in ws.columns:
            letter = col[0].column_letter
            width = max((len(str(cell.value or "")) for cell in col[:100]), default=10)
            ws.column_dimensions[letter].width = min(max(width + 2, 12), 48)
    return buf.getvalue()


def render():
    st.markdown("""
    <style>
      .block-container {max-width:1180px; padding-top:1.7rem; padding-bottom:3rem;}
      .converter-heading {text-align:center; margin:0 auto 1.05rem auto;}
      .converter-heading h1 {font-size:2rem; font-weight:650; margin:0 0 .35rem 0; color:var(--text-color);}
      .converter-heading p {font-size:1.05rem; margin:0; color:var(--text-color); opacity:.82;}
      .upload-shell {background:var(--secondary-background-color); border:1px solid rgba(120,120,140,.13);
        border-radius:18px; padding:16px; box-shadow:0 8px 22px rgba(20,20,40,.05);}
      .upload-inner {border:2px dashed #e7d2ff; border-radius:13px; min-height:235px;
        padding:26px 20px 12px 20px; text-align:center; background:rgba(255,255,255,.24);}
      .upload-title {font-size:1.08rem; font-weight:600; margin:10px 0 5px 0;}
      .upload-subtitle {font-size:.88rem; color:#9299a8; margin-bottom:14px;}
      .privacy-note {font-size:.78rem; color:#9299a8; font-style:italic; margin:18px 0 0 0;}
      div[data-testid="stFileUploader"] {max-width:650px; margin:0 auto;}
      div[data-testid="stFileUploader"] section {border:0 !important; padding:0 !important; background:transparent !important;}
      div[data-testid="stFileUploader"] section > div {justify-content:center;}
      div[data-testid="stFileUploader"] button {background:#8424e8 !important; color:white !important;
        border:1px solid #8424e8 !important; border-radius:9px !important; padding:.55rem 1.5rem !important;}
      div[data-testid="stRadio"] {max-width:460px; margin:0 auto;}
      div[data-testid="stRadio"] > label {display:none;}
      div[data-testid="stRadio"] [role="radiogroup"] {justify-content:center; gap:10px;}
      div[data-testid="stRadio"] label {border:1px solid rgba(130,130,150,.22); border-radius:10px;
        padding:8px 13px; background:var(--secondary-background-color);}
      div[data-testid="stButton"] button {border-radius:9px;}
      @media(max-width:640px){.converter-heading h1{font-size:1.55rem}.upload-inner{padding:18px 10px}}
    </style>
    <div class="converter-heading">
      <h1>JPG To Excel</h1>
      <p>Easily convert images to Excel in seconds with our image to Excel converter.</p>
    </div>
    <div class="upload-shell"><div class="upload-inner">
      <div class="upload-title">Drag &amp; Drop, Upload or Paste image</div>
      <div class="upload-subtitle">Supported formats: JPG, JPEG, PNG, WEBP, BMP, TIFF</div>
    </div></div>
    """, unsafe_allow_html=True)

    if OCR_ERROR:
        st.error(f"Image OCR is unavailable: {OCR_ERROR}")
        return

    left, right, spacer = st.columns([1, 1, 3])
    with left:
        uploaded = st.file_uploader(
            "Browse image", type=["jpg", "jpeg", "png", "webp", "bmp", "tif", "tiff"],
            label_visibility="collapsed", key="image_file_upload"
        )
    pasted_image = None
    with right:
        if paste_image_button:
            try:
                pasted = paste_image_button("Paste image", key="paste_image")
                pasted_image = getattr(pasted, "image_data", None)
            except Exception:
                st.caption("Clipboard paste is not available in this browser.")
        else:
            st.caption("Clipboard paste unavailable")

    mode = st.radio(
        "Extraction mode", ["Only Table", "Extract All Data"],
        horizontal=True, captions=["Only tabular data", "Tables, headings, text"],
        key="image_extract_mode",
    )
    st.markdown(
        '<div class="privacy-note">*Your privacy is protected. Files are processed for conversion during this session.</div>',
        unsafe_allow_html=True
    )

    image = None
    source_name = "pasted_image"
    if uploaded is not None:
        try:
            image = Image.open(uploaded).convert("RGB")
            source_name = Path(uploaded.name).stem
        except Exception as exc:
            st.error(f"Could not open image: {exc}")
    elif pasted_image is not None:
        image = pasted_image.convert("RGB")

    if image is not None:
        st.image(image, caption="Image ready for extraction", use_container_width=True)
        if st.button("Extract data", type="primary", use_container_width=True):
            with st.spinner("Reading image and detecting text…"):
                try:
                    rows = visual_rows(engine()(np.array(image)))
                    frame = make_frame(rows, mode)
                    st.session_state["image_ocr_frame"] = frame
                    st.session_state["image_ocr_raw"] = "\n".join("\t".join(r) for r in rows)
                    st.session_state["image_ocr_name"] = source_name
                except Exception as exc:
                    st.error(f"OCR failed: {type(exc).__name__}: {exc}")

    frame = st.session_state.get("image_ocr_frame")
    if isinstance(frame, pd.DataFrame):
        st.divider()
        st.subheader("Extracted data")
        st.caption(f"{len(frame):,} rows · {len(frame.columns):,} columns. Review and edit before exporting.")
        edited = st.data_editor(frame, num_rows="dynamic", use_container_width=True, key="image_ocr_editor")
        stem = re.sub(r"[^A-Za-z0-9_-]+", "_", st.session_state.get("image_ocr_name", "image"))[:48] or "image"
        a, b, c, d = st.columns(4)
        a.download_button("Download Excel", xlsx_bytes(edited), f"{stem}.xlsx",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          type="primary", use_container_width=True)
        b.download_button("Download CSV", edited.to_csv(index=False).encode("utf-8-sig"),
                          f"{stem}.csv", "text/csv", use_container_width=True)
        c.download_button("Download text", edited.to_csv(index=False, sep="\t").encode("utf-8"),
                          f"{stem}.txt", "text/plain", use_container_width=True)
        d.download_button("Raw OCR text", st.session_state.get("image_ocr_raw", "").encode("utf-8"),
                          f"{stem}_raw.txt", "text/plain", use_container_width=True)
    else:
        st.info("Your extracted table will appear here. OCR works best on clear, high-resolution images. Review figures before using exported data.")
