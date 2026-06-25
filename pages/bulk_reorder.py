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

# ── Product popup — select any product to open a full detail dialog ───────────
product_names_list = ["— select a product to inspect —"] + list(prod_sum["Product Name"])
sel_popup = st.selectbox(
    "🔍 Open product detail popup",
    product_names_list,
    index=0,
    help="Select a product to see its full size, color and store breakdown in a popup",
    key="popup_sel"
)

@st.dialog("📦 Product Detail", width="large")
def show_product_dialog(pname, prow, sz_df, cl_df, ps_all, target_wks, fmt_npr_fn):
    STATUS_COLOR = {"Super Fast":"#1B5E20","Fast":"#43A047","Medium":"#F9A825",
                    "Slow":"#E53935","Dead":"#424242"}
    pstatus  = prow["STR_Status"]
    pstr     = prow["STR_Pct"]
    psold    = int(prow["Total_Sold"])
    pstock   = int(prow["Total_Stock"])
    p_ord_wk = int(prow["Reorder_Wk"])
    p_ord_st = int(prow["Reorder_STR"])
    pcat     = prow.get("Category","")
    psub     = prow.get("Sub Category","") if "Sub Category" in prow.index else ""
    avg_p    = prow.get("Avg_Price", 0)
    sc = STATUS_COLOR.get(pstatus,"#6b7280")

    # Header
    st.markdown(
        f"<h3 style='margin:0'>{pname}</h3>"
        f"<p style='color:#6b7280;margin:2px 0 12px'>"
        f"{pcat}{' › ' + psub if psub else ''}</p>",
        unsafe_allow_html=True)

    # KPI strip
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    for kc, kv, kl, kclr in [
        (k1, f"<span style='color:{sc}'>{pstatus}</span>", "STR Status",    "#374151"),
        (k2, f"{pstr:.0f}%",      "Sell-Through",   sc),
        (k3, f"{psold:,}",        "Units Sold",      "#374151"),
        (k4, f"{pstock:,}",       "In Stock",        "#374151"),
        (k5, f"🔵 {p_ord_wk:,}", f"Order (Wk/{target_wks}wk)", "#1d4ed8"),
        (k6, f"🟣 {p_ord_st:,}", "Order (STR)",     "#7c3aed"),
    ]:
        kc.markdown(
            f'<div style="text-align:center;padding:10px 6px;background:#f8fafc;'
            f'border-radius:8px;border:1px solid #e2e8f0;margin-bottom:4px">'
            f'<div style="font-size:17px;font-weight:700">{kv}</div>'
            f'<div style="font-size:10px;color:#6b7280;margin-top:2px">{kl}</div></div>',
            unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Tabs
    tab_sz, tab_cl, tab_st = st.tabs(["📏 Sizes", "🎨 Colors", "🏪 Stores"])

    def _s_status(v):
        return {"Super Fast":"background-color:#1B5E20;color:white",
                "Fast":"background-color:#43A047;color:white",
                "Medium":"background-color:#F9A825;color:black",
                "Slow":"background-color:#E53935;color:white",
                "Dead":"background-color:#424242;color:white"}.get(v,"")
    def _s_wk(v):
        return "background-color:#dbeafe;color:#1e40af;font-weight:700" if isinstance(v,(int,float)) and v>0 else ""
    def _s_str(v):
        return "background-color:#ede9fe;color:#5b21b6;font-weight:700" if isinstance(v,(int,float)) and v>0 else ""

    with tab_sz:
        if sz_df is not None and not sz_df.empty and pname in sz_df["Product Name"].values:
            p_sz = sz_df[sz_df["Product Name"] == pname][
                ["Size","Units Sold","In Stock","STR %","Status","Weekly Rate","Order (Wk)","Order (STR)"]
            ].copy()
            st.dataframe(
                p_sz.style
                    .map(_s_status, subset=["Status"])
                    .map(_s_wk,     subset=["Order (Wk)"])
                    .map(_s_str,    subset=["Order (STR)"])
                    .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                             "Weekly Rate":"{:.2f}","Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}),
                width='stretch', hide_index=True)
            t_wk  = int(p_sz["Order (Wk)"].sum())
            t_str = int(p_sz["Order (STR)"].sum())
            st.markdown(
                f'<div style="padding:10px;background:#eff6ff;border-radius:8px;margin-top:8px">'
                f'<b>Total across all sizes:</b> &nbsp; '
                f'🔵 <b style="color:#1d4ed8">{t_wk} units</b> (Wk) &nbsp;·&nbsp; '
                f'🟣 <b style="color:#7c3aed">{t_str} units</b> (STR) &nbsp;·&nbsp; '
                f'Est. value: <b>{fmt_npr_fn(t_wk * avg_p)}</b></div>',
                unsafe_allow_html=True)
        else:
            st.info("No size breakdown available for this product in variant_analysis.xlsx")

    with tab_cl:
        if cl_df is not None and not cl_df.empty and pname in cl_df["Product Name"].values:
            p_cl = cl_df[cl_df["Product Name"] == pname][
                ["Color","Units Sold","In Stock","STR %","Status","Order (STR)"]
            ].copy()
            st.dataframe(
                p_cl.style
                    .map(_s_status, subset=["Status"])
                    .map(_s_str,    subset=["Order (STR)"])
                    .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}",
                             "In Stock":"{:,.0f}","Order (STR)":"{:,.0f}"}),
                width='stretch', hide_index=True)
        else:
            st.info("No color breakdown available for this product.")

    with tab_st:
        if ps_all is not None:
            p_store = ps_all[ps_all["Product Name"].str.strip() == pname].copy()
            if p_store.empty:
                pn_w = set(pname.lower().split())
                mask = ps_all["Product Name"].str.lower().apply(
                    lambda n: set(n.split()).issuperset(pn_w) and len(pn_w) >= 2)
                p_store = ps_all[mask].copy()
            if p_store.empty:
                st.info("No store sales data recorded for this product.")
            else:
                p_st_agg = p_store.groupby("Store").agg(
                    Units_Sold=("Units Sold","sum"),
                    Revenue=("Revenue (NPR)","sum")
                ).reset_index()
                p_st_agg["_o"] = p_st_agg["Store"].apply(
                    lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
                p_st_agg = p_st_agg.sort_values("_o").drop(columns=["_o"])
                grand = p_st_agg["Units_Sold"].sum()
                p_st_agg["Share %"] = (p_st_agg["Units_Sold"] / grand * 100).round(1) if grand > 0 else 0
                p_st_agg["Order (Wk)"]  = (p_st_agg["Share %"] / 100 * p_ord_wk ).round().astype(int)
                p_st_agg["Order (STR)"] = (p_st_agg["Share %"] / 100 * p_ord_st).round().astype(int)
                p_st_agg = p_st_agg[p_st_agg["Units_Sold"] > 0]
                p_st_agg["Revenue (NPR)"] = p_st_agg["Revenue"].apply(fmt_npr_fn)
                p_st_agg = p_st_agg.rename(columns={"Units_Sold":"Units Sold"})
                st.dataframe(
                    p_st_agg[["Store","Units Sold","Share %","Order (Wk)","Order (STR)","Revenue (NPR)"]].style
                        .map(_s_wk,  subset=["Order (Wk)"])
                        .map(_s_str, subset=["Order (STR)"])
                        .format({"Units Sold":"{:,.0f}","Share %":"{:.1f}%",
                                 "Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}),
                    width='stretch', hide_index=True)

                # Visual bar chart
                st.markdown("**Sales share by store**")
                max_u = p_st_agg["Units Sold"].max() or 1
                for _, row in p_st_agg.iterrows():
                    pct = row["Units Sold"] / max_u * 100
                    st.markdown(
                        f'<div style="margin-bottom:5px">'
                        f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">'
                        f'<span><b>{row["Store"]}</b></span>'
                        f'<span style="color:#6b7280">{int(row["Units Sold"]):,} units &nbsp;·&nbsp; '
                        f'<span style="color:#1d4ed8">🔵{int(row["Order (Wk)"])}</span> &nbsp; '
                        f'<span style="color:#7c3aed">🟣{int(row["Order (STR)"])}</span></span>'
                        f'</div>'
                        f'<div style="background:#e2e8f0;border-radius:4px;height:8px">'
                        f'<div style="background:#1d4ed8;width:{pct:.0f}%;height:8px;border-radius:4px"></div>'
                        f'</div></div>', unsafe_allow_html=True)
        else:
            st.info("Store data not loaded. Set GDRIVE_PRODSTORE_ID.")

# Trigger the dialog when a product is selected
if sel_popup != "— select a product to inspect —":
    prow_sel = prod_sum[prod_sum["Product Name"] == sel_popup].iloc[0]
    ps_all_ref = _ps_all if df_prodstore is not None else None
    show_product_dialog(
        sel_popup, prow_sel,
        sz if not sz.empty else None,
        cl if not cl.empty else None,
        ps_all_ref,
        target_weeks, fmt_npr
    )

# ── Pre-compute size / color / store lookups for the expanders ────────────────
# Size lookup: product_name -> DataFrame of sizes with reorder qtys
sz = pd.DataFrame()
if size_df is not None:
    _sz = size_df[size_df["Brand"].str.strip() == sel_brand].copy()
    if sel_cat != "All" and "Category" in _sz.columns:
        _sz = _sz[_sz["Category"].str.strip() == sel_cat]
    filtered_products_set = set(prod_sum["Product Name"].str.strip())
    _sz = _sz[_sz["Product Name"].str.strip().isin(filtered_products_set)]
    if not _sz.empty:
        _sz["_sk"] = _sz["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        _sz = _sz.sort_values(["Product Name","_sk"]).drop(columns=["_sk"])
        rate_map = prod_sum.set_index("Product Name")["Weekly_Rate"].to_dict()
        _sz["_prod_rate"] = _sz["Product Name"].map(rate_map).fillna(0)
        prod_total_sold = _sz.groupby("Product Name")["Units Sold"].transform("sum")
        _sz["_size_share"] = (_sz["Units Sold"] / prod_total_sold.replace(0, float("nan"))).fillna(
            1.0 / _sz.groupby("Product Name")["Units Sold"].transform("count"))
        _sz["Weekly Rate"] = (_sz["_prod_rate"] * _sz["_size_share"]).round(2)
        _sz["Order (Wk)"]  = (_sz["Weekly Rate"] * target_weeks - _sz["In Stock"]).clip(lower=0).round().astype(int)
        _sz["Order (STR)"] = (_sz["Units Sold"]  - _sz["In Stock"]).clip(lower=0).round().astype(int)
        sz = _sz

# Color lookup: product_name -> DataFrame of colors
cl = pd.DataFrame()
if color_df is not None:
    _cl = color_df[color_df["Brand"].str.strip() == sel_brand].copy()
    if sel_cat != "All" and "Category" in _cl.columns:
        _cl = _cl[_cl["Category"].str.strip() == sel_cat]
    filtered_products_set = set(prod_sum["Product Name"].str.strip())
    _cl = _cl[_cl["Product Name"].str.strip().isin(filtered_products_set)]
    if not _cl.empty:
        _cl = _cl.sort_values(["Product Name","Units Sold"], ascending=[True,False])
        _cl["Order (STR)"] = (_cl["Units Sold"] - _cl["In Stock"]).clip(lower=0).round().astype(int)
        cl = _cl

# Store lookup: product_name -> DataFrame of stores
store_by_product = {}
if df_prodstore is not None:
    _ps_all = df_prodstore[df_prodstore["Brand"].str.strip() == sel_brand].copy()
    if sel_cat != "All" and "Category" in _ps_all.columns:
        _ps_all = _ps_all[_ps_all["Category"].str.strip() == sel_cat]

# ── Style helpers for expanders ───────────────────────────────────────────────
def _style_sz_status(val):
    return {"Super Fast":"background-color:#1B5E20;color:white","Fast":"background-color:#43A047;color:white",
            "Medium":"background-color:#F9A825;color:black","Slow":"background-color:#E53935;color:white",
            "Dead":"background-color:#424242;color:white"}.get(val,"")

def _style_sz_order(val):
    return "background-color:#dbeafe;color:#1e40af;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""

def _style_sz_str(val):
    return "background-color:#ede9fe;color:#5b21b6;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""

# ── Per-product expanders ─────────────────────────────────────────────────────
st.markdown('<div class="sec">📦 Product Detail — click any product to expand</div>', unsafe_allow_html=True)
st.caption("Each row shows sizes, colors and store distribution for that product.")

# Limit to top N to avoid rendering thousands of expanders
MAX_EXPANDERS = 100
products_to_show = prod_sum.head(MAX_EXPANDERS)
if len(prod_sum) > MAX_EXPANDERS:
    st.info(f"Showing {MAX_EXPANDERS} of {len(prod_sum)} products. Use search or category filters to narrow down.")

for _, prow in products_to_show.iterrows():
    pname    = prow["Product Name"]
    pstatus  = prow["STR_Status"]
    pstr     = prow["STR_Pct"]
    psold    = int(prow["Total_Sold"])
    pstock   = int(prow["Total_Stock"])
    p_ord_wk = int(prow["Reorder_Wk"])
    p_ord_st = int(prow["Reorder_STR"])
    pcat     = prow.get("Category","")
    psub     = prow.get("Sub Category","") if "Sub Category" in prow.index else ""

    STATUS_COLOR = {"Super Fast":"#1B5E20","Fast":"#43A047","Medium":"#F9A825","Slow":"#E53935","Dead":"#424242"}
    sc = STATUS_COLOR.get(pstatus,"#6b7280")

    # Build expander label with key info visible without opening
    label = (f"**{pname}** &nbsp; "
             f"<span style='color:{sc};font-weight:700'>{pstatus}</span> &nbsp; "
             f"STR {pstr:.0f}% &nbsp;·&nbsp; "
             f"Sold {psold:,} &nbsp;·&nbsp; "
             f"Stock {pstock:,} &nbsp;·&nbsp; "
             f"🔵 Order {p_ord_wk:,} (Wk) &nbsp;·&nbsp; "
             f"🟣 Order {p_ord_st:,} (STR)")

    # Auto-expand if search is active (user is looking for a specific product)
    auto_expand = bool(search.strip()) and search.strip().lower() in pname.lower()

    with st.expander(f"{pname}  ·  {pstatus}  ·  STR {pstr:.0f}%  ·  Stock {pstock:,}  ·  🔵{p_ord_wk:,}  🟣{p_ord_st:,}", expanded=auto_expand):

        # Header row with key numbers
        h1, h2, h3, h4, h5 = st.columns(5)
        for hcol, hval, hlbl, hclr in [
            (h1, f"{pstr:.0f}%",       "STR %",          sc),
            (h2, f"{psold:,}",         "Units Sold",      "#374151"),
            (h3, f"{pstock:,}",        "In Stock",        "#374151"),
            (h4, f"{p_ord_wk:,} units",f"Order (Wk/{target_weeks}wk)", "#1d4ed8"),
            (h5, f"{p_ord_st:,} units","Order (STR)",     "#7c3aed"),
        ]:
            hcol.markdown(f'<div style="text-align:center;padding:8px;background:#f8fafc;border-radius:8px;border:1px solid #e2e8f0">'
                          f'<div style="font-size:18px;font-weight:700;color:{hclr}">{hval}</div>'
                          f'<div style="font-size:10px;color:#6b7280">{hlbl}</div></div>',
                          unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # Tabs: Sizes | Colors | Stores
        has_sizes  = not sz.empty  and pname in sz["Product Name"].values
        has_colors = not cl.empty  and pname in cl["Product Name"].values
        has_stores = df_prodstore is not None

        tab_labels = []
        if has_sizes:  tab_labels.append("📏 Sizes")
        if has_colors: tab_labels.append("🎨 Colors")
        if has_stores: tab_labels.append("🏪 Stores")

        if not tab_labels:
            st.caption("No variant or store data found for this product.")
        else:
            tabs = st.tabs(tab_labels)
            tab_idx = 0

            if has_sizes:
                with tabs[tab_idx]:
                    p_sz = sz[sz["Product Name"] == pname][
                        ["Size","Units Sold","In Stock","STR %","Status","Weekly Rate","Order (Wk)","Order (STR)"]
                    ].copy()
                    _t = (p_sz.style
                        .map(_style_sz_status, subset=["Status"])
                        .map(_style_sz_order,  subset=["Order (Wk)"])
                        .map(_style_sz_str,    subset=["Order (STR)"])
                        .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                                 "Weekly Rate":"{:.2f}","Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}))
                    st.dataframe(_t, width='stretch', hide_index=True)
                    total_wk  = int(p_sz["Order (Wk)"].sum())
                    total_str = int(p_sz["Order (STR)"].sum())
                    st.caption(f"Total across all sizes: 🔵 {total_wk} units (Wk) · 🟣 {total_str} units (STR)")
                tab_idx += 1

            if has_colors:
                with tabs[tab_idx]:
                    p_cl = cl[cl["Product Name"] == pname][
                        ["Color","Units Sold","In Stock","STR %","Status","Order (STR)"]
                    ].copy()
                    _tc = (p_cl.style
                        .map(_style_sz_status, subset=["Status"])
                        .map(_style_sz_str,    subset=["Order (STR)"])
                        .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}",
                                 "In Stock":"{:,.0f}","Order (STR)":"{:,.0f}"}))
                    st.dataframe(_tc, width='stretch', hide_index=True)
                tab_idx += 1

            if has_stores:
                with tabs[tab_idx]:
                    p_store = _ps_all[_ps_all["Product Name"].str.strip() == pname].copy()
                    if p_store.empty:
                        # fuzzy: word overlap fallback
                        pn_words = set(pname.lower().split())
                        mask = _ps_all["Product Name"].str.lower().apply(
                            lambda n: set(n.split()).issuperset(pn_words) and len(pn_words) >= 2)
                        p_store = _ps_all[mask].copy()

                    if p_store.empty:
                        st.caption("This product has no store sales recorded.")
                    else:
                        p_store_agg = p_store.groupby("Store").agg(
                            Units_Sold=("Units Sold","sum"),
                            Revenue=("Revenue (NPR)","sum")
                        ).reset_index()
                        p_store_agg["_o"] = p_store_agg["Store"].apply(
                            lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
                        p_store_agg = p_store_agg.sort_values("_o").drop(columns=["_o"])
                        grand = p_store_agg["Units_Sold"].sum()
                        p_store_agg["Share %"] = (p_store_agg["Units_Sold"] / grand * 100).round(1) if grand > 0 else 0
                        p_store_agg["Order (Wk)"]  = (p_store_agg["Share %"] / 100 * p_ord_wk ).round().astype(int)
                        p_store_agg["Order (STR)"] = (p_store_agg["Share %"] / 100 * p_ord_st).round().astype(int)
                        p_store_agg = p_store_agg[p_store_agg["Units_Sold"] > 0]
                        p_store_agg["Revenue"] = p_store_agg["Revenue"].apply(fmt_npr)
                        p_store_agg = p_store_agg.rename(columns={"Units_Sold":"Units Sold","Revenue":"Revenue (NPR)"})

                        _ts = (p_store_agg[["Store","Units Sold","Share %","Order (Wk)","Order (STR)","Revenue (NPR)"]].style
                            .map(_style_sz_order, subset=["Order (Wk)"])
                            .map(_style_sz_str,   subset=["Order (STR)"])
                            .format({"Units Sold":"{:,.0f}","Share %":"{:.1f}%",
                                     "Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}))
                        st.dataframe(_ts, width='stretch', hide_index=True)
                tab_idx += 1


# ── Overall Store Distribution (category-level summary) ──────────────────────
st.markdown('<div class="sec">🏪 Overall Store Distribution</div>', unsafe_allow_html=True)
st.caption("Total reorder split across stores for the current filters. Per-product store breakdown is inside each product expander above.")

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