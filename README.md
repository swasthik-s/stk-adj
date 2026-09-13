# Negative Stock Adjustment Tool

Streamlit app for preparing iTrade negative-stock adjustment sheets
(bundle break, wrong sale, repacking) for Shams Al Madina Hypermarket.

## Run locally
    pip install -r requirements.txt
    streamlit run app.py

## Deploy on Streamlit Community Cloud
1. Push `app.py`, `requirements.txt` and this README to a **private** GitHub repo.
2. share.streamlit.io -> New app -> pick the repo -> main file `app.py`.
3. In the app's Settings -> Sharing, set it to **private** and add only the
   people who should see it. The default is public to anyone with the link.

No data is stored by the app. Files are uploaded per session and disappear
when the session ends.

## Use
Sidebar: upload the masterlist CSV and the negative stock XLSX, set the date
and signatory names, adjust matching strictness.

- **Overview** — totals, value by category, concentration, dead barcodes
- **Candidates** — run matching, review each proposed pair, untick bad ones
- **Build sheets** — generates ADJ_001.xlsx ... plus WORKING_ALL.xlsx, as a zip
- **Manual pair** — enter a pair by hand for repacking or anything matched by eye

Conversion field: 400GM packs from 1 KG = 2.5, 18KG bag to loose KG = 18,
wrong sale = 1.

Every pair nets to 0.00 by construction: single qty = |outer qty| x conv,
single cost = outer cost / conv.
