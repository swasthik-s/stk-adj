"""Image to Excel, CSV and text using local OCR."""
import io, re
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
st.title(":material/image: Image to Excel")
st.caption("Drop an image, browse for a file, or paste a screenshot from your clipboard. Extracted rows can be exported to Excel, CSV or text.")
if OCR_ERROR:
    st.error(f"Image OCR is unavailable: {OCR_ERROR}")
    st.stop()
@st.cache_resource(show_spinner=False)
def engine():
    return RapidOCR()
def visual_rows(result, tolerance=14):
    boxes, texts = getattr(result, "boxes", None), getattr(result, "txts", None)
    if boxes is None or texts is None: return []
    tokens = []
    for box, text in zip(boxes, texts):
        xs, ys = [float(p[0]) for p in box], [float(p[1]) for p in box]
        tokens.append((sum(ys)/len(ys), min(xs), max(xs), str(text)))
    tokens.sort(key=lambda t: (t[0], t[1]))
    grouped = []
    for token in tokens:
        row = next((r for r in reversed(grouped) if abs(r[0]-token[0]) <= tolerance), None)
        if row is None: grouped.append([token[0], [token]])
        else:
            row[1].append(token)
            row[0] = sum(t[0] for t in row[1]) / len(row[1])
    output = []
    for _, row in grouped:
        row.sort(key=lambda t: t[1])
        widths = [(t[2]-t[1])/max(1,len(t[3])) for t in row]
        gap = max(18, min(70, sorted(widths)[len(widths)//2]*2.8))
        cells, current, right = [], [], None
        for token in row:
            if current and right is not None and token[1]-right > gap:
                cells.append(" ".join(t[3] for t in current).strip()); current=[]
            current.append(token); right=token[2]
        if current: cells.append(" ".join(t[3] for t in current).strip())
        if cells: output.append(cells)
    return output
def dataframe(rows):
    if not rows: return pd.DataFrame()
    width=max(map(len,rows)); rows=[r+[""]*(width-len(r)) for r in rows]
    first=rows[0]
    words=sum(bool(re.search(r"[A-Za-z]",v)) for v in first)
    if len(first)>=2 and words>=max(2,len(first)//2):
        headers=[]; seen={}
        for i,value in enumerate(first):
            name=re.sub(r"\s+"," ",value).strip() or f"Column {i+1}"
            seen[name]=seen.get(name,0)+1
            headers.append(name if seen[name]==1 else f"{name}_{seen[name]}")
        return pd.DataFrame(rows[1:],columns=headers)
    return pd.DataFrame(rows,columns=[f"Column {i+1}" for i in range(width)])
def xlsx_bytes(df):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as writer:
        df.to_excel(writer,index=False,sheet_name="Extracted Data")
        ws=writer.sheets["Extracted Data"]
        for col in ws.columns:
            letter=col[0].column_letter
            ws.column_dimensions[letter].width=min(max(max((len(str(c.value or "")) for c in col[:100]),default=10)+2,12),48)
    return buf.getvalue()
left,right=st.columns([3,2])
with left:
    uploaded=st.file_uploader("Drag and drop an image here, or browse",type=["png","jpg","jpeg","webp","bmp","tif","tiff"],key="image_file_upload")
with right:
    pasted_image=None
    if paste_image_button:
        paste_result=paste_image_button("Paste image from clipboard",key="paste_image")
        pasted_image=paste_result.image_data
    else:
        st.info("Clipboard paste component unavailable. File upload is ready.")
st.caption("For clipboard paste, copy an image first and click Paste. Browser clipboard permission and HTTPS are required.")
image=None; source_name="pasted_image"
if uploaded is not None:
    try: image=Image.open(uploaded).convert("RGB"); source_name=Path(uploaded.name).stem
    except Exception as exc: st.error(f"Could not open image: {exc}")
elif pasted_image is not None: image=pasted_image.convert("RGB")
if image is not None:
    st.image(image,caption="Image ready for extraction",use_container_width=True)
    if st.button("Extract data",type="primary",use_container_width=True):
        with st.spinner("Reading image and detecting rows…"):
            try:
                result=engine()(np.array(image))
                rows=visual_rows(result)
                st.session_state["image_ocr_frame"]=dataframe(rows)
                st.session_state["image_ocr_raw"]="\n".join("\t".join(r) for r in rows)
                st.session_state["image_ocr_name"]=source_name
            except Exception as exc: st.error(f"OCR failed: {type(exc).__name__}: {exc}")
frame=st.session_state.get("image_ocr_frame")
if isinstance(frame,pd.DataFrame):
    st.divider(); st.subheader("Extracted data")
    st.caption(f"{len(frame):,} rows · {len(frame.columns):,} columns. Review and edit before exporting.")
    edited=st.data_editor(frame,num_rows="dynamic",use_container_width=True,key="image_ocr_editor")
    stem=re.sub(r"[^A-Za-z0-9_-]+","_",st.session_state.get("image_ocr_name","image"))[:48] or "image"
    a,b,c,d=st.columns(4)
    a.download_button("Download Excel",xlsx_bytes(edited),f"{stem}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary",use_container_width=True)
    b.download_button("Download CSV",edited.to_csv(index=False).encode("utf-8-sig"),f"{stem}.csv","text/csv",use_container_width=True)
    c.download_button("Download text",edited.to_csv(index=False,sep="\t").encode("utf-8"),f"{stem}.txt","text/plain",use_container_width=True)
    d.download_button("Raw OCR text",st.session_state.get("image_ocr_raw","").encode("utf-8"),f"{stem}_raw.txt","text/plain",use_container_width=True)
else:
    st.info("Your extracted table will appear here. OCR works best on clear, high-resolution images. Review figures before using exported data.")