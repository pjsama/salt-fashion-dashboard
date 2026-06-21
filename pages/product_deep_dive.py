import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Salt Fashion — Product Deep Dive",
    page_icon="🔍", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container{padding:1.5rem 2rem}
/* KPI cards */
.kpi{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;text-align:center}
.kpi-val{font-size:26px;font-weight:700;margin:0;line-height:1.1}
.kpi-lbl{font-size:11px;color:#6b7280;margin:4px 0 0}
/* Verdict banner */
.verdict{border-radius:10px;padding:14px 18px;margin:14px 0;font-size:14px;font-weight:500}
.verdict-reorder{background:#dcfce7;border-left:5px solid #16a34a;color:#166534}
.verdict-watch{background:#fef9c3;border-left:5px solid #d97706;color:#92400e}
.verdict-pause{background:#fee2e2;border-left:5px solid #dc2626;color:#991b1b}
/* Section headers */
.sec{font-size:13px;font-weight:700;color:#1F3864;text-transform:uppercase;
     letter-spacing:.08em;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin:18px 0 10px}
/* STR badges */
.sf{background:#1B5E20;color:#fff;padding:2px 9px;border-radius:8px;font-size:11px;font-weight:700}
.fa{background:#43A047;color:#fff;padding:2px 9px;border-radius:8px;font-size:11px;font-weight:700}
.me{background:#F9A825;color:#000;padding:2px 9px;border-radius:8px;font-size:11px;font-weight:700}
.sl{background:#E53935;color:#fff;padding:2px 9px;border-radius:8px;font-size:11px;font-weight:700}
.de{background:#424242;color:#fff;padding:2px 9px;border-radius:8px;font-size:11px;font-weight:700}
</style>
""", unsafe_allow_html=True)

# ── Drive IDs ─────────────────────────────────────────────────────────────────
GDRIVE_MAIN_ID    = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_VARIANT_ID = "1LPeoGXDDd3ZAppTiuLskzY4q-71CJWfJ"
GDRIVE_STORE_ID   = "1B8_Ml_tAL59MSPrEDwKUR93ruFEC1m23"

# ── Loaders ───────────────────────────────────────────────────────────────────
def _gdrive(file_id):
    try:
        from google.oauth2.service_account import Credentials
        import googleapiclient.discovery
        from googleapiclient.http import MediaIoBaseDownload
        import json as _j
        creds = Credentials.from_service_account_info(
            _j.loads(_j.dumps(dict(st.secrets["gcp_service_account"]))),
            scopes=["https://www.googleapis.com/auth/drive"])
        svc = googleapiclient.discovery.build("drive","v3",credentials=creds)
        buf = BytesIO()
        dl  = MediaIoBaseDownload(buf, svc.files().get_media(fileId=file_id))
        done = False
        while not done: _, done = dl.next_chunk()
        buf.seek(0); return buf
    except: return None

@st.cache_data(ttl=600, show_spinner=False)
def load_products():
    buf = _gdrive(GDRIVE_MAIN_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl")
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base+r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"), reverse=True) if Path(d).exists() else []
            if files: df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl"); break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    for col in ["On Hand Qty","Total Units Sold","Revenue","Sell-Through %","Sales Price","Cost Price","Days of Cover"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Sell-Through %" in df.columns and df["Sell-Through %"].max() <= 1.0:
        df["Sell-Through %"] *= 100
    for col in ["Product Name","Brand","Category","Sub Category","STR Status","ABC Class","DOC Status","Color","Size","SKU / Variant"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df

@st.cache_data(ttl=600, show_spinner=False)
def load_variants():
    buf = _gdrive(GDRIVE_VARIANT_ID)
    if buf:
        try:
            size_df  = pd.read_excel(buf, sheet_name="Size Breakdown",  engine="openpyxl")
            buf.seek(0)
            color_df = pd.read_excel(buf, sheet_name="Color Breakdown", engine="openpyxl")
        except: size_df = color_df = None
    else:
        size_df = color_df = None
    if size_df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        local = Path(base) / "variant_analysis.xlsx"
        if local.exists():
            size_df  = pd.read_excel(local, sheet_name="Size Breakdown",  engine="openpyxl")
            color_df = pd.read_excel(local, sheet_name="Color Breakdown", engine="openpyxl")
    if size_df is None: return None, None
    for df in [size_df, color_df]:
        df.columns = [c.strip() for c in df.columns]
        for col in ["Units Sold","In Stock","STR %"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        # Strip attribute prefixes "Size: XL" → "XL", "Color: Black" → "Black"
        for attr_col in ["Size","Color","Brand"]:
            if attr_col in df.columns:
                df[attr_col] = df[attr_col].astype(str).str.replace(rf"^{attr_col}:\s*","",regex=True).str.strip()
        for col in ["Product Name","Brand","Category","Sub Category","Status"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()
    return size_df, color_df

@st.cache_data(ttl=600, show_spinner=False)
def load_store_top_products():
    buf = _gdrive(GDRIVE_STORE_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="🏆 Top Products by Store", engine="openpyxl")
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export\exports"
        files = sorted(Path(base).glob("store_analysis*.xlsx"), reverse=True) if Path(base).exists() else []
        if files:
            try: df = pd.read_excel(files[0], sheet_name="🏆 Top Products by Store", engine="openpyxl")
            except: pass
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    for col in ["Revenue (NPR)","Units Sold"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in ["Store","Product","Brand","Category"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    # Forward-fill Store column (store name only on first rank row)
    if "Store" in df.columns:
        df["Store"] = df["Store"].replace("", pd.NA).ffill()
    return df

# ── Helpers ───────────────────────────────────────────────────────────────────
SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL",
              "36","37","38","39","40","41","42","43","44","ONE SIZE","FREE SIZE"]

def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

def str_badge(status):
    cls = {"Super Fast":"sf","Fast":"fa","Medium":"me","Slow":"sl","Dead":"de"}.get(status,"de")
    return f'<span class="{cls}">{status}</span>'

def str_color(status):
    return {"Super Fast":"#1B5E20","Fast":"#43A047","Medium":"#F9A825",
            "Slow":"#E53935","Dead":"#424242"}.get(status,"#9E9E9E")

def reorder_qty(units_sold, in_stock):
    """Suggest reorder = bring stock back to sold level (Fast/SuperFast only)."""
    return max(0, round(units_sold - in_stock))

# ── Load data ─────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod = load_products()
    size_df, color_df = load_variants()
    df_store = load_store_top_products()

if df_prod is None:
    st.error("Could not load product data. Check Google Drive connection.")
    st.stop()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Product Deep Dive")
    st.markdown("---")

    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("","nan","True","False")])
    sel_brand = st.selectbox("Brand", brands)

    prod_brand = df_prod[df_prod["Brand"] == sel_brand]

    # Category filter to narrow product list
    cats = ["All"] + sorted([c for c in prod_brand["Category"].unique()
                              if c and c not in ("","nan")])
    sel_cat = st.selectbox("Category (to narrow list)", cats)

    if sel_cat != "All":
        prod_list_df = prod_brand[prod_brand["Category"] == sel_cat]
    else:
        prod_list_df = prod_brand

    # Product search
    products = sorted(prod_list_df["Product Name"].unique())
    products = [p for p in products if p and len(p) > 3]

    search = st.text_input("Search product", placeholder="Type to filter…")
    if search.strip():
        products = [p for p in products if search.lower() in p.lower()]

    if not products:
        st.warning("No products found.")
        st.stop()

    sel_product = st.selectbox("Select Product", products)

    st.markdown("---")
    target_weeks = st.slider("Target weeks of stock", 2, 12, 4,
        help="How many weeks of selling you want in stock. Affects suggested reorder qty.")
    st.markdown("---")

    st.caption(f"{len(products)} products shown")
    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ── Pull data for selected product ───────────────────────────────────────────
prod_rows = df_prod[df_prod["Product Name"] == sel_product].copy()

# Aggregate product-level metrics (product may have multiple size/color rows)
total_sold    = prod_rows["Total Units Sold"].sum()
total_stock   = prod_rows["On Hand Qty"].sum()
total_revenue = prod_rows["Revenue"].sum()
avg_price     = prod_rows["Sales Price"].mean() if "Sales Price" in prod_rows.columns else 0
avg_str       = prod_rows["Sell-Through %"].mean() if "Sell-Through %" in prod_rows.columns else 0
str_status    = prod_rows["STR Status"].mode()[0] if "STR Status" in prod_rows.columns and len(prod_rows) else "—"
abc_class     = prod_rows["ABC Class"].mode()[0]  if "ABC Class"  in prod_rows.columns and len(prod_rows) else "—"
doc_status    = prod_rows["DOC Status"].mode()[0] if "DOC Status"  in prod_rows.columns and len(prod_rows) else "—"
category      = prod_rows["Category"].iloc[0]      if len(prod_rows) else ""
sub_cat       = prod_rows["Sub Category"].iloc[0]  if len(prod_rows) else ""

# Variant data for this product
p_sizes  = size_df[size_df["Product Name"] == sel_product].copy()  if size_df  is not None else pd.DataFrame()
p_colors = color_df[color_df["Product Name"] == sel_product].copy() if color_df is not None else pd.DataFrame()

# Store data for this product
p_stores = df_store[df_store["Product"] == sel_product].copy() if df_store is not None else pd.DataFrame()

# ── Reorder verdict ───────────────────────────────────────────────────────────
weekly_rate = total_sold / 52 if total_sold > 0 else 0  # rough all-time weekly rate
target_stock_needed = weekly_rate * target_weeks
pool_order = max(0, round(target_stock_needed - total_stock))

if str_status in ("Super Fast","Fast") and pool_order > 0:
    verdict_class = "verdict-reorder"
    verdict_icon  = "✅"
    verdict_text  = f"<strong>Reorder recommended — {pool_order} units</strong> to reach {target_weeks}-week target. Selling fast (STR {avg_str:.0f}%)."
elif str_status in ("Super Fast","Fast") and pool_order == 0:
    verdict_class = "verdict-watch"
    verdict_icon  = "📦"
    verdict_text  = f"Stock level OK for now ({target_weeks}-week target met). But strong seller — watch closely."
elif str_status == "Medium":
    verdict_class = "verdict-watch"
    verdict_icon  = "⚠️"
    verdict_text  = f"Medium performer (STR {avg_str:.0f}%). Monitor — reorder only if specific sizes are running out."
else:
    verdict_class = "verdict-pause"
    verdict_icon  = "🛑"
    verdict_text  = f"Slow/Dead seller (STR {avg_str:.0f}%). Do <strong>not</strong> reorder — focus on clearing existing stock first."

# ── Page header ───────────────────────────────────────────────────────────────
st.title("🔍 Product Deep Dive")
st.markdown(
    f"**{sel_product}** &nbsp;·&nbsp; {category}"
    + (f" › {sub_cat}" if sub_cat else "")
    + f" &nbsp;·&nbsp; {sel_brand}",
    unsafe_allow_html=True
)

# Verdict banner
st.markdown(
    f'<div class="verdict {verdict_class}">{verdict_icon} {verdict_text}</div>',
    unsafe_allow_html=True
)

# ── KPI strip ─────────────────────────────────────────────────────────────────
c1,c2,c3,c4,c5,c6 = st.columns(6)
for col, val, lbl, clr in [
    (c1, f"{int(total_sold):,}",       "Total Units Sold",    "#1d4ed8"),
    (c2, f"{int(total_stock):,}",      "In Stock Now",        "#374151"),
    (c3, f"{avg_str:.0f}%",            "Sell-Through Rate",   str_color(str_status)),
    (c4, fmt_npr(total_revenue),       "Total Revenue",       "#374151"),
    (c5, fmt_npr(avg_price),           "Avg Selling Price",   "#374151"),
    (c6, f"{pool_order:,} units",      f"Suggest Reorder ({target_weeks}wk)", "#16a34a" if pool_order > 0 else "#6b7280"),
]:
    with col:
        st.markdown(
            f'<div class="kpi"><p class="kpi-val" style="color:{clr}">{val}</p>'
            f'<p class="kpi-lbl">{lbl}</p></div>',
            unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Size breakdown ────────────────────────────────────────────────────────────
st.markdown('<div class="sec">📏 Size Performance — which sizes sell vs which are stuck</div>', unsafe_allow_html=True)

if not p_sizes.empty:
    # Sort by SIZE_ORDER
    p_sizes["_sk"] = p_sizes["Size"].apply(
        lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
    p_sizes = p_sizes.sort_values("_sk").drop(columns=["_sk"])

    # Add reorder suggestion per size
    p_sizes["Suggest Reorder"] = p_sizes.apply(
        lambda r: reorder_qty(r["Units Sold"], r["In Stock"])
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)

    p_sizes["Suggest Reorder"] = p_sizes["Suggest Reorder"].astype(int)

    # Color-code the status column
    def style_status(val):
        colors = {"Super Fast":"background-color:#1B5E20;color:white",
                  "Fast":"background-color:#43A047;color:white",
                  "Medium":"background-color:#F9A825;color:black",
                  "Slow":"background-color:#E53935;color:white",
                  "Dead":"background-color:#424242;color:white"}
        return colors.get(val,"")

    def style_reorder(val):
        if isinstance(val, (int,float)) and val > 0:
            return "background-color:#dcfce7;color:#166534;font-weight:700"
        return ""

    display_sizes = p_sizes[["Size","Units Sold","In Stock","STR %","Status","Suggest Reorder"]].copy()
    display_sizes["STR %"] = display_sizes["STR %"].round(1)

    styled = display_sizes.style\
        .applymap(style_status, subset=["Status"])\
        .applymap(style_reorder, subset=["Suggest Reorder"])\
        .format({"STR %": "{:.1f}%", "Units Sold": "{:,.0f}",
                 "In Stock": "{:,.0f}", "Suggest Reorder": "{:,.0f}"})

    st.dataframe(styled, use_container_width=True, hide_index=True)

    # Insight callout
    dead_sizes = p_sizes[p_sizes["Status"].isin(["Dead","Slow"])]["Size"].tolist()
    fast_sizes = p_sizes[p_sizes["Status"].isin(["Super Fast","Fast"])]["Size"].tolist()
    total_suggest_size = int(p_sizes["Suggest Reorder"].sum())

    if fast_sizes or dead_sizes:
        insight_parts = []
        if fast_sizes:
            insight_parts.append(f"🟢 <strong>Fast-moving sizes: {', '.join(fast_sizes)}</strong> — reorder these")
        if dead_sizes:
            insight_parts.append(f"🔴 <strong>Stuck sizes: {', '.join(dead_sizes)}</strong> — these aren't selling, hold off")
        if total_suggest_size > 0:
            insight_parts.append(f"📦 <strong>Total suggested reorder from sizes: {total_suggest_size} units</strong>")
        st.info("  ·  ".join(insight_parts))
else:
    st.info("Size breakdown not available — variant_analysis.xlsx needed. Run `python variant_export.py`.")

# ── Color breakdown ───────────────────────────────────────────────────────────
st.markdown('<div class="sec">🎨 Color Performance — which colors customers want vs which are sitting</div>', unsafe_allow_html=True)

if not p_colors.empty:
    p_colors = p_colors.sort_values("Units Sold", ascending=False)
    p_colors["Suggest Reorder"] = p_colors.apply(
        lambda r: reorder_qty(r["Units Sold"], r["In Stock"])
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)
    p_colors["Suggest Reorder"] = p_colors["Suggest Reorder"].astype(int)

    display_colors = p_colors[["Color","Units Sold","In Stock","STR %","Status","Suggest Reorder"]].copy()
    display_colors["STR %"] = display_colors["STR %"].round(1)

    styled_c = display_colors.style\
        .applymap(style_status, subset=["Status"])\
        .applymap(style_reorder, subset=["Suggest Reorder"])\
        .format({"STR %": "{:.1f}%", "Units Sold": "{:,.0f}",
                 "In Stock": "{:,.0f}", "Suggest Reorder": "{:,.0f}"})

    st.dataframe(styled_c, use_container_width=True, hide_index=True)

    top_color    = p_colors.iloc[0]["Color"] if len(p_colors) > 0 else "—"
    dead_colors  = p_colors[p_colors["Status"].isin(["Dead","Slow"])]["Color"].tolist()
    fast_colors  = p_colors[p_colors["Status"].isin(["Super Fast","Fast"])]["Color"].tolist()

    if fast_colors or dead_colors:
        parts = []
        if fast_colors: parts.append(f"🟢 <strong>Top colors: {', '.join(fast_colors[:4])}</strong>")
        if dead_colors: parts.append(f"🔴 <strong>Not moving: {', '.join(dead_colors[:4])}</strong>")
        st.info("  ·  ".join(parts))
else:
    st.info("Color breakdown not available — variant_analysis.xlsx needed.")

# ── Store performance ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">🏪 Store Performance — where this product sells most</div>', unsafe_allow_html=True)

if not p_stores.empty:
    p_stores_display = p_stores[["Store","Units Sold","Revenue (NPR)"]].copy()
    p_stores_display = p_stores_display.sort_values("Units Sold", ascending=False)
    p_stores_display["Revenue (NPR)"] = p_stores_display["Revenue (NPR)"].apply(fmt_npr)

    col_s, col_chart = st.columns([2,3])
    with col_s:
        st.dataframe(p_stores_display, use_container_width=True, hide_index=True)

    with col_chart:
        # Simple horizontal bar via metric cards
        max_units = p_stores_display["Units Sold"].max() if len(p_stores_display) > 0 else 1
        for _, row in p_stores_display.iterrows():
            pct = row["Units Sold"] / max_units * 100 if max_units > 0 else 0
            st.markdown(
                f'<div style="margin-bottom:6px">'
                f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">'
                f'<span><strong>{row["Store"]}</strong></span>'
                f'<span style="color:#6b7280">{int(row["Units Sold"]):,} units</span></div>'
                f'<div style="background:#e2e8f0;border-radius:4px;height:8px">'
                f'<div style="background:#1d4ed8;width:{pct:.0f}%;height:8px;border-radius:4px"></div></div>'
                f'</div>',
                unsafe_allow_html=True
            )
else:
    st.info("Store breakdown not available — store_analysis.xlsx needed.")

# ── Size × Color full grid (from odoo_products.xlsx rows) ────────────────────
st.markdown('<div class="sec">📋 Full SKU Breakdown — every size × color combination</div>', unsafe_allow_html=True)

if not prod_rows.empty:
    sku_cols = [c for c in ["Color","Size","SKU / Variant","On Hand Qty",
                             "Total Units Sold","Sell-Through %","STR Status",
                             "Sales Price","DOC Status"] if c in prod_rows.columns]
    sku_display = prod_rows[sku_cols].copy()
    if "Sell-Through %" in sku_display.columns:
        sku_display["Sell-Through %"] = sku_display["Sell-Through %"].round(1)

    if "Size" in sku_display.columns:
        sku_display["_sk"] = sku_display["Size"].apply(
            lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        sku_display = sku_display.sort_values(["Color","_sk"]).drop(columns=["_sk"])

    def style_str_status(val):
        colors = {"Super Fast":"background-color:#1B5E20;color:white",
                  "Fast":"background-color:#43A047;color:white",
                  "Medium":"background-color:#F9A825;color:black",
                  "Slow":"background-color:#E53935;color:white",
                  "Dead":"background-color:#424242;color:white"}
        return colors.get(val,"")
    def style_doc(val):
        colors = {"Reorder Now":"background-color:#B71C1C;color:white",
                  "Watch":"background-color:#F57F17;color:white",
                  "OK":"background-color:#2E7D32;color:white"}
        return colors.get(val,"")

    fmt_dict = {}
    if "Sell-Through %" in sku_display.columns: fmt_dict["Sell-Through %"] = "{:.1f}%"
    if "On Hand Qty"    in sku_display.columns: fmt_dict["On Hand Qty"]    = "{:,.0f}"
    if "Total Units Sold" in sku_display.columns: fmt_dict["Total Units Sold"] = "{:,.0f}"
    if "Sales Price"    in sku_display.columns: fmt_dict["Sales Price"]    = "NPR {:,.0f}"

    apply_cols = [c for c in ["STR Status"] if c in sku_display.columns]
    doc_cols   = [c for c in ["DOC Status"] if c in sku_display.columns]

    styled_sku = sku_display.style.format(fmt_dict)
    if apply_cols: styled_sku = styled_sku.applymap(style_str_status, subset=apply_cols)
    if doc_cols:   styled_sku = styled_sku.applymap(style_doc,        subset=doc_cols)

    st.dataframe(styled_sku, use_container_width=True, hide_index=True)

# ── Similar products in same category ────────────────────────────────────────
st.markdown('<div class="sec">📊 How this product compares — similar products in same category</div>', unsafe_allow_html=True)

cat_peers = df_prod[
    (df_prod["Brand"] == sel_brand) &
    (df_prod["Category"] == category) &
    (df_prod["Product Name"] != sel_product)
].groupby("Product Name").agg(
    Total_Sold  =("Total Units Sold","sum"),
    In_Stock    =("On Hand Qty","sum"),
    Revenue     =("Revenue","sum"),
    STR_Pct     =("Sell-Through %","mean"),
    STR_Status  =("STR Status", lambda x: x.mode()[0] if len(x) else "—"),
).reset_index().sort_values("Total_Sold", ascending=False).head(15)

if not cat_peers.empty:
    # Add current product as highlighted row
    current_row = pd.DataFrame([{
        "Product Name": f"➡️ {sel_product} (current)",
        "Total_Sold":   total_sold,
        "In_Stock":     total_stock,
        "Revenue":      total_revenue,
        "STR_Pct":      avg_str,
        "STR_Status":   str_status,
    }])
    combined = pd.concat([current_row, cat_peers], ignore_index=True)
    combined["Revenue"] = combined["Revenue"].apply(fmt_npr)
    combined["STR_Pct"] = combined["STR_Pct"].round(1)
    combined = combined.rename(columns={
        "Product Name":"Product","Total_Sold":"Units Sold",
        "In_Stock":"In Stock","STR_Pct":"STR %","STR_Status":"Status"})

    def highlight_current(row):
        if str(row["Product"]).startswith("➡️"):
            return ["background-color:#eff6ff;font-weight:600"] * len(row)
        return [""] * len(row)

    styled_peers = combined.style\
        .apply(highlight_current, axis=1)\
        .applymap(style_status, subset=["Status"])\
        .format({"STR %": "{:.1f}%", "Units Sold": "{:,.0f}", "In Stock": "{:,.0f}"})
    st.dataframe(styled_peers, use_container_width=True, hide_index=True)
    rank = (cat_peers["Total_Sold"] > total_sold).sum() + 1
    st.caption(f"Ranked #{rank} by units sold within {category} ({sel_brand})")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    # Sheet 1: summary
    summary = pd.DataFrame([{
        "Product": sel_product, "Category": category, "Sub Category": sub_cat,
        "Brand": sel_brand, "Total Sold": total_sold, "In Stock": total_stock,
        "STR %": round(avg_str,1), "STR Status": str_status,
        "Revenue": total_revenue, "Avg Price": round(avg_price,0),
        "Suggested Reorder": pool_order,
    }])
    summary.to_excel(writer, sheet_name="Summary", index=False)
    # Sheet 2: sizes
    if not p_sizes.empty:
        p_sizes[["Size","Units Sold","In Stock","STR %","Status","Suggest Reorder"]]\
            .to_excel(writer, sheet_name="By Size", index=False)
    # Sheet 3: colors
    if not p_colors.empty:
        p_colors[["Color","Units Sold","In Stock","STR %","Status","Suggest Reorder"]]\
            .to_excel(writer, sheet_name="By Color", index=False)
    # Sheet 4: stores
    if not p_stores.empty:
        p_stores[["Store","Units Sold","Revenue (NPR)"]]\
            .to_excel(writer, sheet_name="By Store", index=False)
    # Sheet 5: full SKUs
    if not prod_rows.empty:
        prod_rows[sku_cols].to_excel(writer, sheet_name="All SKUs", index=False)

out.seek(0)
st.download_button(
    f"⬇️ Download {sel_product} — full report",
    data=out,
    file_name=f"deep_dive_{sel_product[:40].replace(' ','_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
