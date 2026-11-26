"""Image to Excel with local OCR and clipboard-assisted image input."""
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
    st.markdown(
        '<div class="tool-heading"><div class="tool-icon">▧</div>'
        '<div><h2>Image to Excel</h2><p>Extract tables from a photo, screenshot or scanned image.</p></div></div>',
        unsafe_allow_html=True,
    )
    if OCR_ERROR:
        st.error(f"Image OCR is unavailable: {OCR_ERROR}")
        return

    uploaded = st.file_uploader(
        "Drop an image here or browse files",
        type=["jpg", "jpeg", "png", "webp", "bmp", "tif", "tiff", "gif", "jfif", "heic"],
        key="image_file_upload_v2",
        help="Supported: JPG, JPEG, PNG, WEBP, BMP, TIFF and GIF. HEIC depends on installed Pillow codecs.",
    )

    pasted_image = None
    if paste_image_button:
        st.markdown('<div class="paste-zone">', unsafe_allow_html=True)
        left, center, right = st.columns([1, 1.4, 1])
        with center:
            try:
                pasted = paste_image_button(
                    "Paste image from clipboard",
                    text_color="#ffffff",
                    background_color="#9333ea",
                    hover_background_color="#7e22ce",
                    key="paste_image_centered_v2",
                    errors="raise",
                )
                pasted_image = getattr(pasted, "image_data", None)
            except Exception as exc:
                st.warning(f"Clipboard access is unavailable: {exc}")
        st.markdown(
            '<div class="paste-zone-note">Copy an image first, then select Paste image from clipboard. '
            'Clipboard access requires HTTPS and browser permission.</div></div>',
            unsafe_allow_html=True,
        )
    else:
        st.info("Clipboard paste is unavailable. You can still upload an image above.")

    mode = st.radio(
        "Extraction mode",
        ["Only Table", "Extract All Data"],
        horizontal=True,
        captions=["Only tabular data", "Tables, headings and text"],
        key="image_extract_mode",
    )
    st.caption("Image processing runs locally in this app. Review extracted values before exporting.")

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
        st.markdown("#### Image preview")
        preview_col, details_col = st.columns([1.5, 1], vertical_alignment="center")
        with preview_col:
            st.image(image, use_container_width=True)
        with details_col:
            st.markdown("**Ready to extract**")
            st.write(f"{image.width:,} × {image.height:,} pixels")
            st.caption("For best results, use a sharp, well-lit image with visible column spacing.")
            if st.button("Extract data", type="primary", use_container_width=True, key="image_extract_button_v2"):
                with st.spinner("Reading image and detecting table structure…"):
                    try:
                        rows = visual_rows(engine()(np.array(image)))
                        frame = make_frame(rows, mode)
                        st.session_state["image_ocr_frame"] = frame
                        st.session_state["image_ocr_raw"] = "\n".join("\t".join(r) for r in rows)
                        st.session_state["image_ocr_name"] = source_name
                        st.session_state["image_ocr_editor"] = frame
                    except Exception as exc:
                        st.error(f"OCR failed: {type(exc).__name__}: {exc}")

    frame = st.session_state.get("image_ocr_frame")
    if isinstance(frame, pd.DataFrame):
        st.divider()
        head_left, head_right = st.columns([3, 1], vertical_alignment="center")
        with head_left:
            st.markdown("### Extracted data")
            st.caption(f"{len(frame):,} rows · {len(frame.columns):,} columns · Editable preview")
        with head_right:
            st.markdown('<div class="result-status">✓ Extraction complete</div>', unsafe_allow_html=True)
        edited = st.data_editor(
            frame, num_rows="dynamic", use_container_width=True,
            key="image_ocr_editor_v2",
            column_config={column: st.column_config.TextColumn(column) for column in frame.columns},
        )
        stem = re.sub(r"[^A-Za-z0-9_-]+", "_", st.session_state.get("image_ocr_name", "image"))[:48] or "image"
        st.markdown("#### Export your data")
        a, b, c, d = st.columns(4)
        a.download_button("Excel (.xlsx)", xlsx_bytes(edited), f"{stem}.xlsx",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          type="primary", use_container_width=True)
        b.download_button("CSV", edited.to_csv(index=False).encode("utf-8-sig"),
                          f"{stem}.csv", "text/csv", use_container_width=True)
        c.download_button("Text", edited.to_csv(index=False, sep="\t").encode("utf-8"),
                          f"{stem}.txt", "text/plain", use_container_width=True)
        d.download_button("Raw OCR", st.session_state.get("image_ocr_raw", "").encode("utf-8"),
                          f"{stem}_raw.txt", "text/plain", use_container_width=True)
