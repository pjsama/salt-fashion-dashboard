import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pathlib import Path
from datetime import datetime

st.set_page_config(
    page_title="Salt Fashion — Bulk Reorder",
    page_icon="🛒", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container{padding:1.5rem 2rem}
.kpi{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;text-align:center}
.kpi-val{font-size:26px;font-weight:700;margin:0;line-height:1.1}
.kpi-lbl{font-size:11px;color:#6b7280;margin:4px 0 0}
.sec{font-size:13px;font-weight:700;color:#1F3864;text-transform:uppercase;
     letter-spacing:.08em;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin:18px 0 10px}
</style>
""", unsafe_allow_html=True)

# ── Google Drive IDs ──────────────────────────────────────────────────────────
GDRIVE_MAIN_ID      = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_VARIANT_ID   = "1LPeoGXDDd3ZAppTiuLskzY4q-71CJWfJ"
GDRIVE_PRODSTORE_ID = "10ZvRKu4icGDw_g95PplVVdKmj_m-Zpo4"

LOCATION_ORDER = ["Baneshwor","Lazimpat","Kumaripati","Chitwan","Pokhara","Online",
                  "Baneshwor Lush","Chitwan Lush","Pokhara Lush"]

SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL","5XL","ONE SIZE","FREE SIZE",
              "36","37","38","39","40","41","42","43","44",
              "7 (2-4 Y)","9 (4-5 Y)","11 (5-7 Y)","13 (7-9 Y)","5 (18-24 M)"]

WINTER_CATS = {"Coat","Jacket","Sweater","Cardigan","Sweatshirt","Hoodie","Waistcoat",
               "Pajamas Set","Vest","Knitted","Beanie","Boots","Scarves & Mufflers","Gloves"}
SUMMER_CATS = {"T-Shirts","Shorts","Tops","Dress","Co-Ord Set","Tank Top","Swim Wear",
               "Skirt","Skort","Sundress","Basic Top"}
JUNK_CATS   = {"all","saleable","pos","","nan","none","true","false"}

def cat_season(cat):
    if cat in WINTER_CATS: return "Winter"
    if cat in SUMMER_CATS: return "Summer"
    return "All-Season"

def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

def str_status(p):
    if p >= 95: return "Super Fast"
    if p >= 70: return "Fast"
    if p >= 30: return "Medium"
    if p > 0:   return "Slow"
    return "Dead"

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

@st.cache_resource(show_spinner=False)
def load_products():
    buf = _gdrive(GDRIVE_MAIN_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl",
                                usecols=lambda c: c not in ("Image_Base64","Image"))
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base+r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"), reverse=True) if Path(d).exists() else []
            if files:
                df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl",
                                   usecols=lambda c: c not in ("Image_Base64","Image"))
                break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    for col in ["On Hand Qty","Total Units Sold","Revenue","Sell-Through %","Sales Price","Cost Price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Sell-Through %" in df.columns and df["Sell-Through %"].max() <= 1.0:
        df["Sell-Through %"] *= 100
    for col in ["Brand","Category","Sub Category","STR Status","Product Name","Color","Size"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    if "Create Date" in df.columns:
        df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce")
    SKIP = {"All","Saleable","PoS",""}
    if "Category" in df.columns and df["Category"].str.contains("/", na=False).any():
        def split_cat(raw):
            parts = [p.strip() for p in str(raw).split("/") if p.strip() not in SKIP]
            if not parts: return "", ""
            return (parts[0], parts[1]) if len(parts) > 1 else (parts[0], "")
        sp = df["Category"].apply(split_cat)
        df["Category"]     = sp.apply(lambda x: x[0])
        df["Sub Category"] = sp.apply(lambda x: x[1])
    return df

@st.cache_resource(show_spinner=False)
def load_variants():
    buf = _gdrive(GDRIVE_VARIANT_ID)
    size_df = color_df = None
    if buf:
        try:
            size_df  = pd.read_excel(buf, sheet_name="Size Breakdown",  engine="openpyxl")
            buf.seek(0)
            color_df = pd.read_excel(buf, sheet_name="Color Breakdown", engine="openpyxl")
        except: pass
    if size_df is None:
        local = Path(r"C:\Users\Legion\Desktop\odoo_export") / "variant_analysis.xlsx"
        if local.exists():
            size_df  = pd.read_excel(local, sheet_name="Size Breakdown",  engine="openpyxl")
            color_df = pd.read_excel(local, sheet_name="Color Breakdown", engine="openpyxl")
    if size_df is None: return None, None

    def _prep(df):
        df = df.copy()
        df.columns = [c.strip() for c in df.columns]
        df["Product Name"] = (df["Product Name"].fillna("").astype(str)
                              .str.replace('\n',' ',regex=False)
                              .str.replace('\t',' ',regex=False)
                              .str.replace(r'\s+',' ',regex=True)
                              .str.strip()
                              .str.strip('"'))
        for col in ["Units Sold","In Stock","STR %"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in ["Brand","Category","Sub Category","Size","Color","Status"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.replace(
                    r"^(Size|Color|Brand):\s*","",regex=True).str.strip()
        return df

    size_df  = _prep(size_df)
    color_df = _prep(color_df)
    SIZE_SET = set(s.upper() for s in SIZE_ORDER)

    def parse_name_color(name):
        name = re.sub(r'^\[[^\]]+\]\s*', '', name).strip()
        if '/' in name:
            parts = name.rsplit('/', 1)
            suffix = parts[1].strip()
            if suffix.upper() in SIZE_SET:
                return parts[0].strip(), None
            return parts[0].strip(), suffix.strip()
        return name, None

    parsed = size_df["Product Name"].apply(parse_name_color)
    size_df["_base"]  = parsed.apply(lambda x: x[0])
    size_df["_color"] = parsed.apply(lambda x: x[1])

    grp = [c for c in ["_base","Brand","Category","Sub Category","Size"] if c in size_df.columns]
    size_agg = size_df.groupby(grp, as_index=False).agg(
        **{"Units Sold":("Units Sold","sum"), "In Stock":("In Stock","sum")}
    ).rename(columns={"_base":"Product Name"})
    total = size_agg["Units Sold"] + size_agg["In Stock"]
    size_agg["STR %"]  = (size_agg["Units Sold"] / total.replace(0,float("nan")) * 100).fillna(0).round(1)
    size_agg["Status"] = size_agg["STR %"].apply(str_status)
    size_df = size_agg

    # Build synthetic color rows for products that store color in name
    color_df["Product Name"] = color_df["Product Name"].apply(
        lambda n: re.sub(r"/[^/]+$", "", n).strip())
    existing_colors = set(color_df["Product Name"].str.lower())

    syn_src = size_df[size_df.get("_color", pd.Series(dtype=str)).notna()].copy() if "_color" in size_df.columns else pd.DataFrame()
    # Re-extract from raw (before groupby) — use the _color before aggregation
    buf2 = _gdrive(GDRIVE_VARIANT_ID)
    if buf2:
        try:
            raw_sz = pd.read_excel(buf2, sheet_name="Size Breakdown", engine="openpyxl")
            raw_sz = _prep(raw_sz)
            parsed3 = raw_sz["Product Name"].apply(parse_name_color)
            raw_sz["_base"]  = parsed3.apply(lambda x: x[0])
            raw_sz["_color"] = parsed3.apply(lambda x: x[1])
            syn_src = raw_sz[raw_sz["_color"].notna() & ~raw_sz["_base"].str.lower().isin(existing_colors)]
            if len(syn_src) > 0:
                grp_c = [c for c in ["_base","Brand","Category","Sub Category","_color"] if c in syn_src.columns]
                syn_agg = syn_src.groupby(grp_c, as_index=False).agg(
                    **{"Units Sold":("Units Sold","sum"), "In Stock":("In Stock","sum")}
                ).rename(columns={"_base":"Product Name","_color":"Color"})
                total_c = syn_agg["Units Sold"] + syn_agg["In Stock"]
                syn_agg["STR %"]  = (syn_agg["Units Sold"] / total_c.replace(0,float("nan")) * 100).fillna(0).round(1)
                syn_agg["Status"] = syn_agg["STR %"].apply(str_status)
                color_df = pd.concat([color_df, syn_agg], ignore_index=True)
        except: pass

    return size_df, color_df

@st.cache_resource(show_spinner=False)
def load_product_store():
    buf = _gdrive(GDRIVE_PRODSTORE_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Product × Store", engine="openpyxl")
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export\exports"
        files = sorted(Path(base).glob("product_store_sales_*.xlsx"), reverse=True) if Path(base).exists() else []
        if files:
            try: df = pd.read_excel(files[0], sheet_name="Product × Store", engine="openpyxl")
            except: pass
    if df is None or df.empty: return None
    df.columns = [str(c).strip() for c in df.columns]
    for col in ["Units Sold","Revenue (NPR)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in ["Product Name","Brand","Category","Store"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df


# ── Load ──────────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod           = load_products()
    size_df, color_df = load_variants()
    df_prodstore      = load_product_store()

if df_prod is None:
    st.error("Could not load product data."); st.stop()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🛒 Bulk Reorder Tool")
    st.markdown("---")

    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("","nan","True","False")])
    sel_brand = st.selectbox("Brand", brands)

    cats = ["All"] + sorted([c for c in df_prod[df_prod["Brand"]==sel_brand]["Category"].unique()
                              if c.strip().lower() not in JUNK_CATS])
    sel_cat = st.selectbox("Category", cats)

    # Sub-category cascades from category
    sel_sub = "All"
    if sel_cat != "All" and "Sub Category" in df_prod.columns:
        subs = sorted([s for s in df_prod[
            (df_prod["Brand"]==sel_brand) &
            (df_prod["Category"]==sel_cat)
        ]["Sub Category"].unique() if s and s not in ("","nan")])
        if subs:
            sel_sub = st.selectbox("Sub Category", ["All"] + subs)

    # Search by product name
    search = st.text_input("🔍 Search product", placeholder="Type to filter products…")

    st.markdown("---")
    st.markdown("**Season**")
    SEASON_OPTS = ["All", "Summer (+ All-Season)", "Winter (+ All-Season)", "All-Season only"]
    sel_season_raw = st.selectbox("Season filter", SEASON_OPTS, index=1)
    sel_season = {
        "All": "All",
        "Summer (+ All-Season)": "Summer",
        "Winter (+ All-Season)": "Winter",
        "All-Season only": "All-Season",
    }[sel_season_raw]

    st.markdown("---")
    st.markdown("**Reorder Settings**")
    min_str_pct = st.slider("Min STR % to include", 0, 100, 50,
        help="Only show products at or above this sell-through rate")
    target_weeks = st.slider("Target weeks of cover", 2, 12, 4)
    show_zero = st.checkbox("Show products with 0 order qty", value=False)

    st.markdown("---")
    st.markdown("**📅 Date Added Filter**")
    date_opts = ["All time","Last 30 days","Last 60 days","Last 90 days",
                 "Older than 30 days","Older than 60 days","Older than 90 days"]
    sel_date = st.selectbox("Date filter", date_opts, index=0,
        help="'Older than X' excludes new arrivals that haven't had time to sell")

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_resource.clear(); st.rerun()

# ── Filter products ───────────────────────────────────────────────────────────
bdf = df_prod[df_prod["Brand"] == sel_brand].copy()
bdf = bdf[~bdf["Category"].str.strip().str.lower().isin(JUNK_CATS)]
if sel_cat != "All":
    bdf = bdf[bdf["Category"] == sel_cat]
if sel_sub != "All" and "Sub Category" in bdf.columns:
    bdf = bdf[bdf["Sub Category"] == sel_sub]
if search.strip():
    bdf = bdf[bdf["Product Name"].str.contains(search.strip(), case=False, na=False)]

# Season filter
if sel_season != "All":
    season_map = bdf["Category"].apply(cat_season)
    bdf = bdf[season_map.isin([sel_season, "All-Season"])]

# Date filter
today = pd.Timestamp.today().normalize()
if sel_date != "All time" and "Create Date" in bdf.columns:
    cd = bdf["Create Date"]
    if sel_date == "Last 30 days":       bdf = bdf[cd >= today - pd.Timedelta(days=30)]
    elif sel_date == "Last 60 days":     bdf = bdf[cd >= today - pd.Timedelta(days=60)]
    elif sel_date == "Last 90 days":     bdf = bdf[cd >= today - pd.Timedelta(days=90)]
    elif "Older than 30" in sel_date:    bdf = bdf[cd < today - pd.Timedelta(days=30)]
    elif "Older than 60" in sel_date:    bdf = bdf[cd < today - pd.Timedelta(days=60)]
    elif "Older than 90" in sel_date:    bdf = bdf[cd < today - pd.Timedelta(days=90)]

# ── Build product-level summary ───────────────────────────────────────────────
grp_cols = ["Product Name","Category"]
if "Sub Category" in bdf.columns: grp_cols.append("Sub Category")

prod_sum = bdf.groupby(grp_cols).agg(
    Total_Sold  = ("Total Units Sold","sum"),
    Total_Stock = ("On Hand Qty",     "sum"),
    Avg_Price   = ("Sales Price",     "mean"),
    Total_Rev   = ("Revenue",         "sum"),
).reset_index()

prod_sum["STR_Pct"]    = (prod_sum["Total_Sold"] /
    (prod_sum["Total_Sold"] + prod_sum["Total_Stock"]).replace(0, float("nan")) * 100
).fillna(0).round(1)
prod_sum["STR_Status"] = prod_sum["STR_Pct"].apply(str_status)
prod_sum["Season"]     = prod_sum["Category"].apply(cat_season)

# Weekly rate from create date
if "Create Date" in bdf.columns:
    dates = bdf.groupby("Product Name")["Create Date"].min().reset_index()
    dates["Create Date"] = pd.to_datetime(dates["Create Date"], errors="coerce")
    prod_sum = prod_sum.merge(dates, on="Product Name", how="left")
    prod_sum["weeks_live"] = ((today - prod_sum["Create Date"]).dt.days / 7).fillna(52).clip(lower=4)
else:
    prod_sum["weeks_live"] = 52

prod_sum["Weekly_Rate"]  = (prod_sum["Total_Sold"] / prod_sum["weeks_live"]).round(2)
prod_sum["Target_Stock"] = (prod_sum["Weekly_Rate"] * target_weeks).round(0)
prod_sum["Reorder_Wk"]   = (prod_sum["Target_Stock"] - prod_sum["Total_Stock"]).clip(lower=0).round().astype(int)
prod_sum["Reorder_STR"]  = (prod_sum["Total_Sold"]   - prod_sum["Total_Stock"]).clip(lower=0).round().astype(int)
prod_sum["Est_Value"]    = prod_sum["Reorder_Wk"] * prod_sum["Avg_Price"]

# Apply STR filter
prod_sum = prod_sum[prod_sum["STR_Pct"] >= min_str_pct]
if not show_zero:
    prod_sum = prod_sum[(prod_sum["Reorder_Wk"] > 0) | (prod_sum["Reorder_STR"] > 0)]

prod_sum = prod_sum.sort_values("Total_Sold", ascending=False)

total_units_wk  = int(prod_sum["Reorder_Wk"].sum())
total_units_str = int(prod_sum["Reorder_STR"].sum())
total_value     = prod_sum["Est_Value"].sum()
n_products      = len(prod_sum)
fast_count      = prod_sum["STR_Status"].isin(["Super Fast","Fast"]).sum()

# ── Page header ───────────────────────────────────────────────────────────────
st.title("🛒 Bulk Reorder Tool")
filter_parts = [sel_brand, sel_cat]
if sel_sub != "All": filter_parts.append(sel_sub)
if search.strip(): filter_parts.append(f'"{search}"')
filter_parts.append(sel_season_raw)
if sel_date != "All time": filter_parts.append(sel_date)
st.markdown(f"**{'  ·  '.join(filter_parts)}** · STR ≥ {min_str_pct}% · {target_weeks}-week target · {today.strftime('%b %d, %Y')}")

# ── KPIs ──────────────────────────────────────────────────────────────────────
c1,c2,c3,c4,c5 = st.columns(5)
for col, val, lbl, clr in [
    (c1, f"{n_products:,}",            "Products",               "#374151"),
    (c2, f"{fast_count:,}",            "Fast / Super Fast",      "#16a34a"),
    (c3, f"{total_units_wk:,}",        f"Order Qty (Wk/{target_weeks}wk)", "#1d4ed8"),
    (c4, f"{total_units_str:,}",       "Order Qty (STR restore)","#7c3aed"),
    (c5, fmt_npr(total_value),         "Est. Value (Wk)",        "#374151"),
]:
    with col:
        st.markdown(f'<div class="kpi"><p class="kpi-val" style="color:{clr}">{val}</p>'
                    f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Category Summary ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">📊 Category Summary</div>', unsafe_allow_html=True)

has_sub = "Sub Category" in prod_sum.columns and prod_sum["Sub Category"].str.strip().ne("").any()
cat_grp = ["Category","Sub Category"] if has_sub else ["Category"]

cat_sum = prod_sum.groupby(cat_grp).agg(
    Products    = ("Product Name","count"),
    Units_Sold  = ("Total_Sold",  "sum"),
    In_Stock    = ("Total_Stock", "sum"),
    Avg_STR     = ("STR_Pct",     "mean"),
    Order_Wk    = ("Reorder_Wk",  "sum"),
    Order_STR   = ("Reorder_STR", "sum"),
    Est_Value   = ("Est_Value",   "sum"),
).reset_index().sort_values(["Category","Order_Wk"], ascending=[True,False])

cat_sum["Avg_STR"]   = cat_sum["Avg_STR"].round(1)
cat_sum["Est_Value"] = cat_sum["Est_Value"].apply(fmt_npr)
cat_sum = cat_sum.rename(columns={
    "Products":"# Products","Units_Sold":"Units Sold","In_Stock":"In Stock",
    "Avg_STR":"Avg STR %","Order_Wk":f"Order (Wk/{target_weeks}wk)",
    "Order_STR":"Order (STR)","Est_Value":"Est. Value"
})

def _cat_style(val):
    if isinstance(val,(int,float)) and val > 0:
        return "background-color:#dbeafe;color:#1e40af;font-weight:700"
    return ""

disp_cat_cols = (["Category","Sub Category"] if has_sub else ["Category"]) + \
    ["# Products","Units Sold","In Stock","Avg STR %",
     f"Order (Wk/{target_weeks}wk)","Order (STR)","Est. Value"]
disp_cat_cols = [c for c in disp_cat_cols if c in cat_sum.columns]

st.dataframe(
    cat_sum[disp_cat_cols].style
        .map(_cat_style, subset=[f"Order (Wk/{target_weeks}wk)","Order (STR)"])
        .format({"Avg STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                 f"Order (Wk/{target_weeks}wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}),
    width='stretch', hide_index=True)
st.caption("Sorted by Category A–Z, then by Order Qty descending within each category")

# ── Product-Level Table ───────────────────────────────────────────────────────
st.markdown('<div class="sec">📋 Product-Level Reorder Plan</div>', unsafe_allow_html=True)

show_cols = ["Product Name","Category"] + \
    (["Sub Category"] if has_sub else []) + \
    ["STR_Status","STR_Pct","Total_Sold","Total_Stock",
     "Weekly_Rate","Reorder_Wk","Reorder_STR","Avg_Price","Est_Value"]
show_cols = [c for c in show_cols if c in prod_sum.columns]

disp = prod_sum[show_cols].copy().rename(columns={
    "STR_Status":"Status","STR_Pct":"STR %","Total_Sold":"Units Sold",
    "Total_Stock":"In Stock","Weekly_Rate":"Rate/wk",
    "Reorder_Wk":"Order (Wk)","Reorder_STR":"Order (STR)",
    "Avg_Price":"Avg Price","Est_Value":"Est. Value"
})

def _style_status(val):
    return {"Super Fast":"background-color:#1B5E20;color:white","Fast":"background-color:#43A047;color:white",
            "Medium":"background-color:#F9A825;color:black","Slow":"background-color:#E53935;color:white",
            "Dead":"background-color:#424242;color:white"}.get(val,"")

def _style_order(val):
    if isinstance(val,(int,float)) and val > 0:
        return "background-color:#dbeafe;color:#1e40af;font-weight:700"
    return ""

fmt_d = {"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
         "Rate/wk":"{:.2f}","Avg Price":"NPR {:,.0f}","Est. Value":"{:,.0f}",
         "Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}
_st = disp.style.map(_style_status, subset=["Status"])
if "Order (Wk)"  in disp.columns: _st = _st.map(_style_order, subset=["Order (Wk)"])
if "Order (STR)" in disp.columns: _st = _st.map(_style_order, subset=["Order (STR)"])
st.dataframe(_st.format(fmt_d), width='stretch', hide_index=True)
st.caption(f"{len(disp):,} products · 🔵 Order (Wk) = {target_weeks}-week buffer · 🟣 Order (STR) = restore original stock")

# ── Size Breakdown — per product, all Fast+ items ────────────────────────────
st.markdown('<div class="sec">📏 Size Breakdown by Product</div>', unsafe_allow_html=True)

if size_df is None:
    st.info("Variant data not available. Run `variant_export.py` first.")
else:
    # Get fast/superfast products from prod_sum
    fast_products = set(prod_sum[prod_sum["STR_Status"].isin(["Super Fast","Fast"])]["Product Name"].str.strip())

    # Filter size_df to brand + category + fast products
    sz = size_df[size_df["Brand"].str.strip() == sel_brand].copy()
    if sel_cat != "All":
        sz = sz[sz["Category"].str.strip() == sel_cat]
    if sel_sub != "All" and "Sub Category" in sz.columns:
        sz = sz[sz["Sub Category"].str.strip() == sel_sub]
    if search.strip():
        sz = sz[sz["Product Name"].str.contains(search.strip(), case=False, na=False)]

    # Only show products that are in our filtered prod_sum (respects STR filter)
    filtered_products = set(prod_sum["Product Name"].str.strip())
    sz = sz[sz["Product Name"].str.strip().isin(filtered_products)]

    if sz.empty:
        st.info(f"No size data for the current filters.")
    else:
        # Sort sizes correctly
        sz["_sk"] = sz["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        sz = sz.sort_values(["Product Name","_sk"]).drop(columns=["_sk"])

        # Add weekly rate from prod_sum for reorder calculation
        rate_map = prod_sum.set_index("Product Name")["Weekly_Rate"].to_dict()
        sz["_prod_rate"] = sz["Product Name"].map(rate_map).fillna(0)

        # Size share = this size's sold / total sold for the product
        prod_total_sold = sz.groupby("Product Name")["Units Sold"].transform("sum")
        sz["_size_share"] = sz["Units Sold"] / prod_total_sold.replace(0, float("nan"))
        sz["_size_share"] = sz["_size_share"].fillna(1.0 / sz.groupby("Product Name")["Units Sold"].transform("count"))

        # Reorder qty per size using two methods
        sz["Weekly Rate"] = (sz["_prod_rate"] * sz["_size_share"]).round(2)
        sz["Order (Wk)"]  = (sz["Weekly Rate"] * target_weeks - sz["In Stock"]).clip(lower=0).round().astype(int)
        sz["Order (STR)"] = (sz["Units Sold"]  - sz["In Stock"]).clip(lower=0).round().astype(int)

        disp_sz = sz[["Product Name","Size","Units Sold","In Stock","STR %","Status","Weekly Rate","Order (Wk)","Order (STR)"]].copy()

        def _style_sz_status(val):
            return {"Super Fast":"background-color:#1B5E20;color:white","Fast":"background-color:#43A047;color:white",
                    "Medium":"background-color:#F9A825;color:black","Slow":"background-color:#E53935;color:white",
                    "Dead":"background-color:#424242;color:white"}.get(val,"")

        def _style_sz_order(val):
            if isinstance(val,(int,float)) and val > 0:
                return "background-color:#dbeafe;color:#1e40af;font-weight:700"
            return ""

        def _style_sz_str(val):
            if isinstance(val,(int,float)) and val > 0:
                return "background-color:#ede9fe;color:#5b21b6;font-weight:700"
            return ""

        _sst = (disp_sz.style
            .map(_style_sz_status, subset=["Status"])
            .map(_style_sz_order,  subset=["Order (Wk)"])
            .map(_style_sz_str,    subset=["Order (STR)"])
            .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                     "Weekly Rate":"{:.2f}","Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}))
        st.dataframe(_sst, width='stretch', hide_index=True)
        st.caption(
            f"{len(sz):,} size rows across {sz['Product Name'].nunique():,} products · "
            f"🔵 Order (Wk) = {target_weeks}-week buffer per size · "
            f"🟣 Order (STR) = restore to original level"
        )

# ── Color Breakdown ───────────────────────────────────────────────────────────
if color_df is not None:
    st.markdown('<div class="sec">🎨 Color Breakdown by Product</div>', unsafe_allow_html=True)

    cl = color_df[color_df["Brand"].str.strip() == sel_brand].copy()
    if sel_cat != "All":
        cl = cl[cl["Category"].str.strip() == sel_cat]
    if sel_sub != "All" and "Sub Category" in cl.columns:
        cl = cl[cl["Sub Category"].str.strip() == sel_sub]
    if search.strip():
        cl = cl[cl["Product Name"].str.contains(search.strip(), case=False, na=False)]

    filtered_products = set(prod_sum["Product Name"].str.strip())
    cl = cl[cl["Product Name"].str.strip().isin(filtered_products)]
    cl = cl[cl["Status"].isin(["Super Fast","Fast"])]  # only fast colors

    if cl.empty:
        st.info("No Fast/Super Fast colors for current filters.")
    else:
        cl = cl.sort_values(["Product Name","Units Sold"], ascending=[True, False])
        cl["Order (STR)"] = (cl["Units Sold"] - cl["In Stock"]).clip(lower=0).round().astype(int)

        disp_cl = cl[["Product Name","Color","Units Sold","In Stock","STR %","Status","Order (STR)"]].copy()
        _cst = (disp_cl.style
            .map(_style_sz_status, subset=["Status"])
            .map(_style_sz_str,    subset=["Order (STR)"])
            .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}","Order (STR)":"{:,.0f}"}))
        st.dataframe(_cst, width='stretch', hide_index=True)

# ── Store Distribution ────────────────────────────────────────────────────────
st.markdown('<div class="sec">🏪 Reorder Distribution by Store</div>', unsafe_allow_html=True)

if df_prodstore is None:
    st.info("Store sales data not available. Run `fetch_product_store_sales.py` and set GDRIVE_PRODSTORE_ID.")
else:
    ps = df_prodstore[df_prodstore["Brand"].str.strip() == sel_brand].copy()
    if sel_cat != "All" and "Category" in ps.columns:
        ps = ps[ps["Category"].str.strip() == sel_cat]
    if search.strip() and "Product Name" in ps.columns:
        ps = ps[ps["Product Name"].str.contains(search.strip(), case=False, na=False)]

    if ps.empty:
        st.info(f"No store sales data for **{sel_brand}** / {sel_cat}.")
    else:
        tab_store, tab_catstore = st.tabs(["📍 By Store", "📊 Category × Store"])

        store_totals = ps.groupby("Store").agg(
            Units_Sold = ("Units Sold","sum"),
            Revenue    = ("Revenue (NPR)","sum"),
        ).reset_index()
        store_totals["_order"] = store_totals["Store"].apply(
            lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
        store_totals = store_totals.sort_values("_order").drop(columns=["_order"])
        grand_sold   = store_totals["Units_Sold"].sum()
        store_totals["Share_%"]   = (store_totals["Units_Sold"] / grand_sold * 100).round(1) if grand_sold > 0 else 0
        store_totals["Order_Wk"]  = (store_totals["Share_%"] / 100 * total_units_wk ).round().astype(int)
        store_totals["Order_STR"] = (store_totals["Share_%"] / 100 * total_units_str).round().astype(int)
        store_totals = store_totals[store_totals["Units_Sold"] > 0]

        def _style_ord(val):
            return "background-color:#dbeafe;color:#1e40af;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""
        def _style_str_ord(val):
            return "background-color:#ede9fe;color:#5b21b6;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""

        with tab_store:
            col_tbl, col_bar = st.columns([2, 3])
            with col_tbl:
                disp_st = store_totals[["Store","Units_Sold","Share_%","Order_Wk","Order_STR"]].rename(columns={
                    "Units_Sold":"Units Sold","Share_%":"Share %",
                    "Order_Wk":f"Order (Wk)","Order_STR":"Order (STR)"})
                st.dataframe(
                    disp_st.style
                        .map(_style_ord,     subset=["Order (Wk)"])
                        .map(_style_str_ord, subset=["Order (STR)"])
                        .format({"Units Sold":"{:,.0f}","Share %":"{:.1f}%",
                                 "Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}),
                    width='stretch', hide_index=True)
                st.caption(f"Total: {store_totals['Order_Wk'].sum():,} units (Wk) / "
                           f"{store_totals['Order_STR'].sum():,} units (STR) · "
                           f"{len(store_totals)} stores")
            with col_bar:
                max_u = store_totals["Units_Sold"].max() or 1
                for _, row in store_totals.iterrows():
                    pct = row["Units_Sold"] / max_u * 100
                    st.markdown(
                        f'<div style="margin-bottom:6px">'
                        f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">'
                        f'<span><strong>{row["Store"]}</strong></span>'
                        f'<span style="color:#6b7280">{int(row["Units_Sold"]):,} units · '
                        f'{row["Share_%"]:.0f}% · <span style="color:#1d4ed8">Order {int(row["Order_Wk"])} (Wk)</span></span>'
                        f'</div>'
                        f'<div style="background:#e2e8f0;border-radius:4px;height:8px">'
                        f'<div style="background:#1d4ed8;width:{pct:.0f}%;height:8px;border-radius:4px"></div>'
                        f'</div></div>', unsafe_allow_html=True)

        with tab_catstore:
            grp_key = ["Category","Sub Category","Store"] if "Sub Category" in ps.columns else ["Category","Store"]
            cat_store = ps.groupby(grp_key).agg(Units_Sold=("Units Sold","sum")).reset_index()
            stores_present = [s for s in LOCATION_ORDER if s in cat_store["Store"].unique()]
            pivot_cols = ["Category","Sub Category"] if "Sub Category" in cat_store.columns else ["Category"]
            pivot = cat_store.pivot_table(
                index=pivot_cols, columns="Store", values="Units_Sold",
                aggfunc="sum", fill_value=0
            ).reset_index()
            pivot.columns.name = None
            store_cols_present = [c for c in stores_present if c in pivot.columns]
            pivot["Total"] = pivot[store_cols_present].sum(axis=1)
            pivot = pivot.sort_values("Total", ascending=False)
            st.dataframe(
                pivot.style.format({c: "{:,.0f}" for c in store_cols_present + ["Total"]}),
                width='stretch', hide_index=True)
            st.caption("Units sold per category per store")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    cat_sum.to_excel(writer, sheet_name="Category Summary", index=False)

    full = prod_sum[["Product Name","Category"] +
                   (["Sub Category"] if "Sub Category" in prod_sum.columns else []) +
                   ["STR_Status","STR_Pct","Total_Sold","Total_Stock",
                    "Weekly_Rate","Reorder_Wk","Reorder_STR","Avg_Price","Est_Value"]].copy()
    full = full.rename(columns={"STR_Status":"Status","STR_Pct":"STR %",
                                "Total_Sold":"Units Sold","Total_Stock":"In Stock",
                                "Weekly_Rate":"Rate/wk","Reorder_Wk":f"Order (Wk/{target_weeks}wk)",
                                "Reorder_STR":"Order (STR)","Avg_Price":"Avg Price NPR",
                                "Est_Value":"Est. Value NPR"})
    full.to_excel(writer, sheet_name="Product Reorder Plan", index=False)

    if size_df is not None and "sz" in dir() and not sz.empty:
        sz_exp = sz[["Product Name","Size","Units Sold","In Stock","STR %","Status","Weekly Rate","Order (Wk)","Order (STR)"]].copy()
        sz_exp.to_excel(writer, sheet_name="By Size", index=False)

    if color_df is not None and "cl" in dir() and not cl.empty:
        cl_exp = cl[["Product Name","Color","Units Sold","In Stock","STR %","Status","Order (STR)"]].copy()
        cl_exp.to_excel(writer, sheet_name="By Color", index=False)

    if df_prodstore is not None and "store_totals" in dir() and not store_totals.empty:
        store_totals.rename(columns={"Units_Sold":"Units Sold","Share_%":"Share %",
                                     "Order_Wk":"Order (Wk)","Order_STR":"Order (STR)"})\
            .to_excel(writer, sheet_name="By Store", index=False)

out.seek(0)
fname = f"reorder_{sel_brand.replace(' ','_')}_{(sel_cat if sel_cat!='All' else 'AllCats').replace(' ','_')}.xlsx"
st.download_button(
    f"⬇️ Download Full Reorder Plan — {sel_brand} / {sel_cat}",
    data=out, file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("Download includes: Category Summary · Product Plan · Size Breakdown · Color Breakdown · Store Distribution")