"""Interactive audit of an iTrade master CSV."""
import io
import re
import pandas as pd
import streamlit as st

try:
    from auth import require_pin
    require_pin()
except ImportError:
    pass

st.title("Master Sheet Audit")
st.caption("Upload an iTrade master CSV to inspect stock, pricing and data quality. This page never posts or changes stock.")

ALIASES = {
    "Item": ["item name", "item description", "description", "product name", "name", "item"],
    "Barcode": ["item barcode", "barcode", "bar code", "ean", "item code", "item no", "sku"],
    "Stock": ["stock", "current stock", "stock qty", "stock quantity", "qty", "quantity", "closing stock", "balance"],
    "Cost": ["cost", "cost price", "last cost", "wac", "weighted average cost", "purchase price", "unit cost"],
    "Selling price": ["net mrp", "lp price", "selling price", "sell price", "retail price", "mrp", "sale price"],
    "Category": ["class", "category", "department", "group", "item category"],
    "Pack / unit": ["pack qty", "packing", "pack", "unit", "uom"],
}
def norm(x):
    return re.sub(r"[^a-z0-9]+", " ", str(x).lower()).strip()
def detect(cols, aliases):
    index = {norm(c): c for c in cols}
    return next((index[norm(a)] for a in aliases if norm(a) in index), None)
def number(s):
    s = s.astype(str).str.strip().str.replace(",", "", regex=False)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    s = s.str.replace(r"[^0-9.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")
@st.cache_data(show_spinner=False, max_entries=2)
def parse_csv(raw):
    errors = []
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            frame = pd.read_csv(io.BytesIO(raw), dtype=str, encoding=enc,
                                sep=",", engine="c", on_bad_lines="warn")
            # iTrade exports are normally comma-separated. Fall back to delimiter
            # sniffing only when the whole header was read as one column.
            if len(frame.columns) == 1:
                frame = pd.read_csv(io.BytesIO(raw), dtype=str, encoding=enc,
                                    sep=None, engine="python", on_bad_lines="warn")
            return frame
        except Exception as exc:
            errors.append(f"{enc}: {exc}")
    raise ValueError("Unable to parse CSV: " + " | ".join(errors))
def to_xlsx(frame):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name="Audit")
        ws = writer.sheets["Audit"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cells in ws.columns:
            ws.column_dimensions[cells[0].column_letter].width = min(
                48, max(12, max(len(str(c.value or "")) for c in cells) + 2))
    return out.getvalue()

uploaded = st.file_uploader("Upload master sheet CSV", type=["csv"],
                            help="20 MB+ files are supported; processing depends on available app memory.")
if not uploaded:
    st.info("Upload your iTrade master CSV to begin.")
    st.markdown("🔴 Negative stock / selling below cost · 🟡 Missing data / duplicate barcodes · 🟢 No detected issue")
    st.stop()
raw = uploaded.getvalue()
if not raw:
    st.error("The uploaded file is empty.")
    st.stop()
try:
    with st.spinner("Reading and auditing CSV…"):
        source = parse_csv(raw)
except Exception as exc:
    st.error(str(exc))
    st.stop()
if source.empty:
    st.warning("This CSV contains no data rows.")
    st.stop()

cols = list(source.columns)
auto = {name: detect(cols, aliases) for name, aliases in ALIASES.items()}
with st.expander("Column mapping", expanded=not (auto["Stock"] and auto["Item"])):
    st.caption("Detected columns are selected automatically. Adjust them if your export uses different headings.")
    options = ["— Not available —"] + cols
    c1, c2 = st.columns(2)
    mapping = {}
    for i, name in enumerate(ALIASES):
        parent = c1 if i % 2 == 0 else c2
        default = auto[name] if auto[name] in cols else "— Not available —"
        choice = parent.selectbox(name, options, index=options.index(default), key=f"master_audit_{name}")
        mapping[name] = None if choice == "— Not available —" else choice
if not mapping["Stock"]:
    st.error("Map the Stock column to continue.")
    st.stop()

df = pd.DataFrame(index=source.index)
df["Item"] = source[mapping["Item"]].fillna("").astype(str).str.strip() if mapping["Item"] else "(Item name unavailable)"
df["Barcode"] = source[mapping["Barcode"]].fillna("").astype(str).str.strip() if mapping["Barcode"] else ""
df["Stock"] = number(source[mapping["Stock"]])
df["Cost"] = number(source[mapping["Cost"]]) if mapping["Cost"] else float("nan")
df["Selling price"] = number(source[mapping["Selling price"]]) if mapping["Selling price"] else float("nan")
df["Category"] = source[mapping["Category"]].fillna("Uncategorised").astype(str).str.strip() if mapping["Category"] else "Uncategorised"
df["Category"] = df["Category"].replace("", "Uncategorised")
df["Pack / unit"] = source[mapping["Pack / unit"]].fillna("").astype(str) if mapping["Pack / unit"] else ""
df["Gross profit / unit"] = df["Selling price"] - df["Cost"]
df["Gross margin %"] = df["Gross profit / unit"] / df["Selling price"].where(df["Selling price"] != 0) * 100
df["Negative stock"] = df["Stock"] < 0
df["Below cost"] = df["Cost"].notna() & df["Selling price"].notna() & (df["Cost"] > 0) & (df["Selling price"] < df["Cost"])
df["Missing cost"] = df["Cost"].isna() | (df["Cost"] <= 0) if mapping["Cost"] else True
df["Missing selling price"] = df["Selling price"].isna() | (df["Selling price"] <= 0) if mapping["Selling price"] else True
df["Duplicate barcode"] = df["Barcode"].ne("") & df["Barcode"].duplicated(keep=False)
issue_flags = [
    (df["Negative stock"], "Negative stock"),
    (df["Below cost"], "Selling below cost"),
    (df["Missing cost"], "Missing / zero cost"),
    (df["Missing selling price"], "Missing / zero selling price"),
    (df["Duplicate barcode"], "Duplicate barcode"),
]
df["Issues"] = [
    "; ".join(label for mask, label in issue_flags if bool(mask.iloc[i])) or "No detected issue"
    for i in range(len(df))
]
df["Status"] = [
    "🔴 Critical" if n or b else "🟡 Investigate" if mc or mp or dup else "🟢 Healthy"
    for n, b, mc, mp, dup in zip(df["Negative stock"], df["Below cost"], df["Missing cost"], df["Missing selling price"], df["Duplicate barcode"])
]
df["Suggested action"] = [
    "Investigate stock ledger / physical count" if n else
    "Review selling price or cost" if b else
    "Verify cost in master data" if mc else
    "Verify selling price in master data" if mp else
    "Check duplicate barcode records" if dup else "No action flagged"
    for n, b, mc, mp, dup in zip(df["Negative stock"], df["Below cost"], df["Missing cost"], df["Missing selling price"], df["Duplicate barcode"])
]
neg = df[df["Negative stock"]]
below = df[df["Below cost"]]
data_issues = df[df["Missing cost"] | df["Missing selling price"] | df["Duplicate barcode"]]
healthy = df[df["Status"] == "🟢 Healthy"]
st.caption(f"**{uploaded.name}** · {len(df):,} rows · {len(cols)} columns · {len(raw)/(1024*1024):.1f} MB")
m = st.columns(5)
for col, label, value in zip(m, ["📦 Total rows", "🔴 Negative stock", "💸 Below cost", "🟡 Data issues", "🟢 Healthy"],
                             [len(df), len(neg), len(below), len(data_issues), len(healthy)]):
    col.metric(label, f"{value:,}")
st.markdown("""
<style>
.audit-red{background:rgba(220,38,38,.12);border:1px solid #dc2626;border-radius:8px;padding:9px;margin:5px 0}
.audit-yellow{background:rgba(217,119,6,.12);border:1px solid #d97706;border-radius:8px;padding:9px;margin:5px 0}
.audit-green{background:rgba(22,163,74,.10);border:1px solid #16a34a;border-radius:8px;padding:9px;margin:5px 0}
</style>
<div class="audit-red">🔴 <b>Critical:</b> negative stock or selling below cost</div>
<div class="audit-yellow">🟡 <b>Investigate:</b> missing / zero values or duplicate barcodes</div>
<div class="audit-green">🟢 <b>Healthy:</b> no issue detected in the mapped fields</div>
""", unsafe_allow_html=True)

tabs = st.tabs(["📊 Overview", "🔴 Negative Stock", "💸 Below Cost", "🟡 Data Issues", "🟢 Healthy Items", "🔎 All Items", "⬇️ Reports"])
def view_for(key):
    return {"negative": neg, "below": below, "issues": data_issues, "healthy": healthy, "all": df}[key].copy()
def filtered_table(frame, key):
    a, b, c = st.columns([2, 1, 1])
    query = a.text_input("Search item / barcode", key=f"{key}_q", placeholder="Item name or barcode")
    categories = sorted(frame["Category"].dropna().astype(str).unique())
    selected_categories = b.multiselect("Category", categories, key=f"{key}_cat")
    statuses = frame["Status"].unique().tolist()
    selected_status = c.multiselect("Status", statuses, default=statuses, key=f"{key}_status")
    out = frame
    if query:
        q = query.casefold()
        out = out[out["Item"].str.casefold().str.contains(q, na=False) | out["Barcode"].str.casefold().str.contains(q, na=False)]
    if selected_categories:
        out = out[out["Category"].isin(selected_categories)]
    out = out[out["Status"].isin(selected_status)] if selected_status else out.iloc[:0]
    a, b, c = st.columns([2, 1, 1])
    sort_opts = ["Item A–Z", "Barcode", "Stock: smallest first", "Stock: largest first", "Cost: low to high", "Selling price: low to high", "Gross profit/unit: lowest first", "Gross margin: lowest first"]
    sort = a.selectbox("Sort by", sort_opts, key=f"{key}_sort")
    limit = b.selectbox("Rows to display", [100, 500, 1000, 5000], index=1, key=f"{key}_limit")
    direction = c.selectbox("Text order", ["Ascending", "Descending"], key=f"{key}_dir")
    col = {"Item A–Z":"Item", "Barcode":"Barcode", "Stock: smallest first":"Stock", "Stock: largest first":"Stock",
           "Cost: low to high":"Cost", "Selling price: low to high":"Selling price",
           "Gross profit/unit: lowest first":"Gross profit / unit", "Gross margin: lowest first":"Gross margin %"}[sort]
    asc = direction == "Ascending"
    if sort == "Stock: largest first":
        asc = False
    out = out.sort_values(col, ascending=asc, na_position="first", kind="stable")
    st.caption(f"{len(out):,} matching rows · showing up to {min(len(out), limit):,}. Exports contain all matching rows.")
    show_cols = ["Status", "Item", "Barcode", "Category", "Stock", "Cost", "Selling price", "Gross profit / unit", "Gross margin %", "Issues", "Suggested action"]
    shown = out[[x for x in show_cols if x in out.columns]].head(limit)
    if len(shown):
        def paint(row):
            color = "rgba(220,38,38,.12)" if str(row.get("Status","")).startswith("🔴") else "rgba(217,119,6,.12)" if str(row.get("Status","")).startswith("🟡") else "rgba(22,163,74,.07)"
            return [f"background-color: {color}"] * len(row)
        st.dataframe(shown.style.apply(paint, axis=1), width="stretch", hide_index=True,
                     height=min(650, 38*(len(shown)+1)),
                     column_config={"Cost":st.column_config.NumberColumn(format="AED %.2f"),
                                    "Selling price":st.column_config.NumberColumn(format="AED %.2f"),
                                    "Gross profit / unit":st.column_config.NumberColumn(format="AED %.2f"),
                                    "Gross margin %":st.column_config.NumberColumn(format="%.2f%%"),
                                    "Stock":st.column_config.NumberColumn(format="%.3f")})
    else:
        st.success("No items match the selected filters.")
    return out

with tabs[0]:
    left, right = st.columns(2)
    with left:
        st.markdown("#### 🔴 Critical checks")
        st.write(f"Negative stock: **{len(neg):,}**")
        st.write(f"Selling below cost: **{len(below):,}**")
        if mapping["Cost"] and mapping["Selling price"]:
            gap = (below["Cost"] - below["Selling price"]).sum()
            st.write(f"Below-cost gap per unit, summed: **AED {gap:,.2f}**")
        else:
            st.warning("Map both cost and selling-price columns to calculate below-cost items.")
    with right:
        st.markdown("#### 🟡 Data quality")
        st.write(f"Missing / zero cost: **{int(df['Missing cost'].sum()):,}**")
        st.write(f"Missing / zero selling price: **{int(df['Missing selling price'].sum()):,}**")
        st.write(f"Rows with duplicate barcode: **{int(df['Duplicate barcode'].sum()):,}**")
    st.markdown("#### Largest stock deficits")
    st.dataframe(neg.sort_values("Stock").head(20)[["Item","Barcode","Category","Stock","Suggested action"]], width="stretch", hide_index=True)
    st.markdown("#### Largest below-cost gaps")
    if len(below):
        tmp = below.assign(**{"Loss / unit": below["Cost"] - below["Selling price"]})
        st.dataframe(tmp.sort_values("Loss / unit", ascending=False).head(20)[["Item","Barcode","Cost","Selling price","Loss / unit","Suggested action"]], width="stretch", hide_index=True)
    else:
        st.success("No below-cost items detected from the mapped columns.")
with tabs[1]:
    st.caption("Investigate transactions and physical stock before posting any adjustment.")
    filtered_table(view_for("negative"), "negative")
with tabs[2]:
    if not mapping["Cost"] or not mapping["Selling price"]:
        st.warning("Map both Cost and Selling price in Column mapping to enable this check.")
    else:
        st.caption("Master-data comparison only; it does not prove the price charged on an individual sale.")
        filtered_table(view_for("below"), "below")
with tabs[3]:
    filtered_table(view_for("issues"), "issues")
with tabs[4]:
    st.caption("Green means no issue was detected by these checks; it does not verify physical stock.")
    filtered_table(view_for("healthy"), "healthy")
with tabs[5]:
    filtered_table(view_for("all"), "all")
with tabs[6]:
    report_name = st.selectbox("Report to export", ["Complete audit", "Negative stock", "Selling below cost", "Data issues", "Healthy items"])
    report_key = {"Complete audit":"all", "Negative stock":"negative", "Selling below cost":"below", "Data issues":"issues", "Healthy items":"healthy"}[report_name]
    report = view_for(report_key).drop(columns=["Negative stock","Below cost","Missing cost","Missing selling price","Duplicate barcode"], errors="ignore")
    st.metric("Rows in report", f"{len(report):,}")
    c1, c2 = st.columns(2)
    c1.download_button("⬇️ Download CSV", report.to_csv(index=False).encode("utf-8-sig"),
                       f"master_audit_{report_key}.csv", "text/csv", width="stretch")
    try:
        xlsx = to_xlsx(report)
        c2.download_button("⬇️ Download Excel", xlsx, f"master_audit_{report_key}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    except Exception as exc:
        st.warning(f"Excel export unavailable: {exc}. CSV export is available.")
    st.dataframe(report.head(100), width="stretch", hide_index=True)

st.divider()
st.caption("A master file can flag issues but cannot establish why stock went negative. Check sales, receiving, transfers and physical counts before adjustments.")
