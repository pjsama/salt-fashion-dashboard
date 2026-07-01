import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pathlib import Path
from datetime import datetime

st.set_page_config(
    page_title="Salt Fashion — Product Deep Dive",
    page_icon="🔍", layout="wide",
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
.insight{background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;
         padding:10px 14px;font-size:13px;color:#1e40af;margin-bottom:8px}
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

JUNK_CATS = {"all","saleable","pos","","nan","none","true","false"}

SIZE_SUFFIXES = {"XS","S","M","L","XL","2XL","3XL","4XL","5XL",
                 "ONE SIZE","FREE SIZE","36","37","38","39","40","41","42","43","44"}
_dash_re = re.compile(r'\s[-–]\s([A-Z0-9]{1,4})$')

NEW_PRODUCT_DAYS = 90

def str_status(p):
    if p >= 95: return "Super Fast"
    if p >= 70: return "Fast"
    if p >= 30: return "Medium"
    if p > 0:   return "Slow"
    return "Dead"

def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

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

def _fix_name_size(name, size):
    """Strip size/dash suffix from product name, return (clean_name, size)."""
    if "/" in name:
        parts = name.rsplit("/", 1)
        suffix = parts[1].strip()
        if suffix.upper() in SIZE_SUFFIXES:
            return parts[0].strip(), size if size else suffix
    m = _dash_re.search(name)
    if m and m.group(1).upper() in SIZE_SUFFIXES:
        return name[:m.start()].strip(), size if size else m.group(1)
    return name, size

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
    for col in ["Brand","Category","Sub Category","Product Name","Color","Size"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    # Fix size in name
    if "Product Name" in df.columns and "Size" in df.columns:
        fixed = df.apply(lambda r: _fix_name_size(r["Product Name"], r["Size"]), axis=1, result_type="expand")
        df["Product Name"] = fixed[0]
        df["Size"]         = fixed[1]
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

    SIZE_SUFFIXES_SET = {"XS","S","M","L","XL","2XL","3XL","4XL","5XL",
                         "ONE SIZE","FREE SIZE","36","37","38","39","40","41","42","43","44"}
    _dash_re2 = re.compile(r'\s[-–]\s([A-Z0-9]{1,4})$')

    def _prep(df):
        df = df.copy()
        df.columns = [c.strip() for c in df.columns]
        df["Product Name"] = (df["Product Name"].fillna("").astype(str)
                              .str.replace('\n',' ',regex=False).str.replace('\t',' ',regex=False)
                              .str.replace(r'\s+',' ',regex=True).str.strip().str.strip('"'))
        # Strip [SKU] prefix
        df["Product Name"] = df["Product Name"].apply(
            lambda n: re.sub(r'^\[[^\]]+\]\s*', '', n).strip())
        # Strip size from name (both / and - patterns)
        def _fix_vn(row):
            name = row["Product Name"]
            size = str(row.get("Size","")).strip() if "Size" in row.index else ""
            if "/" in name:
                parts = name.rsplit("/", 1)
                suffix = parts[1].strip()
                if suffix.upper() in SIZE_SUFFIXES_SET:
                    return parts[0].strip(), size if size else suffix
            m = _dash_re2.search(name)
            if m and m.group(1).upper() in SIZE_SUFFIXES_SET:
                return name[:m.start()].strip(), size if size else m.group(1)
            return name, size
        if "Size" in df.columns:
            fixed = df.apply(_fix_vn, axis=1, result_type="expand")
            df["Product Name"] = fixed[0]
            df["Size"]         = fixed[1]
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

    # Aggregate size_df by (Product Name, Brand, Category, Sub Category, Size)
    # Also extract color from name suffix for synthetic color rows
    SIZE_SET = set(s.upper() for s in SIZE_ORDER)

    def parse_name_color(name):
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

    # Build synthetic color rows
    color_df["Product Name"] = color_df["Product Name"].apply(
        lambda n: re.sub(r"/[^/]+$", "", n).strip())
    existing_colors = set(color_df["Product Name"].str.lower())
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
    # Fix: strip size suffix from product names (same as load_products)
    if "Product Name" in df.columns:
        df["Product Name"] = df["Product Name"].apply(
            lambda n: _fix_name_size(n, "")[0])
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
    st.markdown("### 🔍 Product Deep Dive")
    st.markdown("---")

    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("","nan","True","False")])
    sel_brand = st.selectbox("Brand", brands)

    cats = ["All"] + sorted([c for c in df_prod[df_prod["Brand"]==sel_brand]["Category"].unique()
                              if c.strip().lower() not in JUNK_CATS])
    sel_cat = st.selectbox("Category", cats)

    # Product search
    search = st.text_input("🔍 Search product", placeholder="Type product name…")

    # Product selector
    bdf_all = df_prod[df_prod["Brand"]==sel_brand].copy()
    if sel_cat != "All":
        bdf_all = bdf_all[bdf_all["Category"]==sel_cat]
    if search.strip():
        bdf_all = bdf_all[bdf_all["Product Name"].str.contains(search.strip(), case=False, na=False)]

    products = sorted(bdf_all["Product Name"].dropna().unique().tolist())
    if not products:
        st.warning("No products found for these filters.")
        st.stop()

    sel_product = st.selectbox("Select Product", products)

    st.markdown("---")
    cover_days = st.slider("Days of cover", 30, 120, 60, step=15,
        help="velocity × cover_days − stock = reorder qty")

    st.markdown("---")
    n_cat = len(bdf_all["Product Name"].unique())
    st.caption(f"{n_cat} products · {len(bdf_all)} variants in category")

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_resource.clear(); st.rerun()

# ── Get product data ──────────────────────────────────────────────────────────
today    = pd.Timestamp.today().normalize()
prod_rows = df_prod[df_prod["Product Name"] == sel_product].copy()

total_sold  = prod_rows["Total Units Sold"].sum()
total_stock = prod_rows["On Hand Qty"].sum()
avg_price   = prod_rows["Sales Price"].mean()
str_pct     = round(total_sold / (total_sold + total_stock) * 100, 1) if (total_sold + total_stock) > 0 else 0
status      = str_status(str_pct)
has_recent  = "Recent Sold 60d" in prod_rows.columns
recent_60   = prod_rows["Recent Sold 60d"].sum() if has_recent else 0
net_lbl     = "Net Sales (60d)" if has_recent else "Total Sold"
net_val     = recent_60 if has_recent else total_sold

# Velocity — same 3-tier logic as bulk reorder
if "Create Date" in prod_rows.columns:
    min_date  = prod_rows["Create Date"].min()
    days_live = max(7, (today - pd.Timestamp(min_date)).days) if pd.notna(min_date) else 365
else:
    days_live = 365

lifetime_vel = total_sold / min(days_live, 365)

if has_recent and recent_60 > 0:
    recent_vel = recent_60 / 60
    if days_live >= NEW_PRODUCT_DAYS:
        daily_vel   = min(recent_vel, lifetime_vel)
        reorder_vel = min(recent_vel, lifetime_vel)
        if recent_vel < lifetime_vel * 0.8:   trend = "📉 Slowing"
        elif recent_vel > lifetime_vel * 1.2:  trend = "📈 Trending"
        else:                                  trend = "✅ Stable"
    else:
        daily_vel   = recent_vel
        reorder_vel = recent_vel
        trend       = "🆕 New (<90d)"
elif has_recent and recent_60 == 0:
    daily_vel   = lifetime_vel
    reorder_vel = 0
    trend       = "🔴 No demand"
else:
    daily_vel   = lifetime_vel
    reorder_vel = lifetime_vel
    trend       = "—"

weekly_rate  = daily_vel * 7
reorder_qty  = max(0, round(reorder_vel * cover_days - total_stock))
est_value    = reorder_qty * avg_price

# ── Page header ───────────────────────────────────────────────────────────────
st.title(f"🔍 {sel_product}")
cat_row = prod_rows.iloc[0] if len(prod_rows) > 0 else {}
st.markdown(
    f"**{sel_brand}** · {cat_row.get('Category','')} "
    f"{'› '+cat_row.get('Sub Category','') if cat_row.get('Sub Category','') else ''} · "
    f"Trend: {trend} · {today.strftime('%b %d, %Y')}"
)

# ── KPIs ──────────────────────────────────────────────────────────────────────
c1,c2,c3,c4,c5,c6 = st.columns(6)
STATUS_COLORS = {
    "Super Fast":"#1B5E20","Fast":"#43A047","Medium":"#F9A825",
    "Slow":"#E53935","Dead":"#424242"
}
for col, val, lbl, clr in [
    (c1, f"{str_pct:.1f}%",       "STR %",           STATUS_COLORS.get(status,"#374151")),
    (c2, status,                   "Status",          STATUS_COLORS.get(status,"#374151")),
    (c3, f"{int(total_sold):,}",   "Units Sold",      "#374151"),
    (c4, f"{int(net_val):,}",      net_lbl,           "#0f766e"),
    (c5, f"{int(total_stock):,}",  "In Stock",        "#374151"),
    (c6, f"{reorder_qty:,}",       f"Order ({cover_days}d)", "#1d4ed8"),
]:
    with col:
        st.markdown(f'<div class="kpi"><p class="kpi-val" style="color:{clr}">{val}</p>'
                    f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Velocity insight ──────────────────────────────────────────────────────────
vel_line = (
    f"**Velocity:** {daily_vel:.3f} u/day · {weekly_rate:.2f} u/wk · "
    f"**Lifetime:** {lifetime_vel:.3f} u/day ({days_live}d) · "
    f"**Recent (60d):** {recent_60:.0f} units · "
    f"**Order ({cover_days}d):** {reorder_qty} units · "
    f"**Est. Value:** {fmt_npr(est_value)}"
)
st.markdown(f'<div class="insight">{vel_line}</div>', unsafe_allow_html=True)

def _style_status(val):
    colors = {"Super Fast":("background-color:#1B5E20;color:white"),
              "Fast":"background-color:#43A047;color:white",
              "Medium":"background-color:#F9A825;color:black",
              "Slow":"background-color:#E53935;color:white",
              "Dead":"background-color:#424242;color:white"}
    return colors.get(val,"")

def _style_order(val):
    return "background-color:#dbeafe;color:#1e40af;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""

def _style_stock(val):
    return "background-color:#fee2e2;color:#991b1b" if isinstance(val,(int,float)) and val == 0 else ""

def _style_str(val):
    if not isinstance(val,(int,float)): return ""
    if val >= 70: return "background-color:#dcfce7;color:#166534"
    if val >= 30: return "background-color:#fef9c3;color:#854d0e"
    return "background-color:#fee2e2;color:#991b1b"

# ── Size Performance ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">📏 Size Performance</div>', unsafe_allow_html=True)

if size_df is not None:
    sz = size_df[size_df["Product Name"].str.strip() == sel_product].copy()
    if not sz.empty:
        sz["_sk"] = sz["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        sz = sz.sort_values("_sk").drop(columns=["_sk"])
        # Velocity per size — proportional to size share of total product sales
        prod_total_sold_sz = sz["Units Sold"].sum()
        sz["Size Share %"] = (sz["Units Sold"] / prod_total_sold_sz * 100).round(1) if prod_total_sold_sz > 0 else 0
        sz["Rate/wk"]      = (weekly_rate * sz["Size Share %"] / 100).round(2)
        sz[f"Order ({cover_days}d)"] = (
            reorder_qty * sz["Size Share %"] / 100
        ).round().astype(int)

        disp_sz = sz[["Size","Units Sold","In Stock","STR %","Status","Size Share %",
                       "Rate/wk",f"Order ({cover_days}d)"]].copy()
        st.dataframe(
            disp_sz.style
                .map(_style_status, subset=["Status"])
                .map(_style_stock,  subset=["In Stock"])
                .map(_style_str,    subset=["STR %"])
                .map(_style_order,  subset=[f"Order ({cover_days}d)"])
                .format({"STR %":"{:.1f}%","Size Share %":"{:.1f}%",
                         "Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                         "Rate/wk":"{:.2f}",f"Order ({cover_days}d)":"{:,.0f}"}),
            width='stretch', hide_index=True)

        # Fast sizes insight
        fast_sizes = sz[sz["Status"].isin(["Super Fast","Fast"])]["Size"].tolist()
        out_sizes  = sz[(sz["In Stock"]==0) & (sz["Units Sold"]>0)]["Size"].tolist()
        if fast_sizes:
            st.markdown(f'<div class="insight">✅ Fast sizes: {", ".join(fast_sizes)}'
                        + (f' &nbsp;|&nbsp; 🔴 Sold out: {", ".join(out_sizes)}' if out_sizes else '')
                        + '</div>', unsafe_allow_html=True)
    else:
        st.info("No size data available for this product in variant_analysis.xlsx.")
else:
    st.info("Variant data not loaded.")

# ── Color Performance ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">🎨 Color Performance</div>', unsafe_allow_html=True)

if color_df is not None:
    cl = color_df[color_df["Product Name"].str.strip() == sel_product].copy()
    if not cl.empty:
        cl = cl.sort_values("Units Sold", ascending=False)
        # Distribute reorder by color share
        prod_total_sold_cl = cl["Units Sold"].sum()
        cl["Color Share %"] = (cl["Units Sold"] / prod_total_sold_cl * 100).round(1) if prod_total_sold_cl > 0 else 0
        cl["Rate/wk"]       = (weekly_rate * cl["Color Share %"] / 100).round(2)
        cl[f"Order ({cover_days}d)"] = (
            reorder_qty * cl["Color Share %"] / 100
        ).round().astype(int)

        disp_cl = cl[["Color","Units Sold","In Stock","STR %","Status",
                       "Color Share %","Rate/wk",f"Order ({cover_days}d)"]].copy()
        st.dataframe(
            disp_cl.style
                .map(_style_status, subset=["Status"])
                .map(_style_stock,  subset=["In Stock"])
                .map(_style_str,    subset=["STR %"])
                .map(_style_order,  subset=[f"Order ({cover_days}d)"])
                .format({"STR %":"{:.1f}%","Color Share %":"{:.1f}%",
                         "Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                         "Rate/wk":"{:.2f}",f"Order ({cover_days}d)":"{:,.0f}"}),
            width='stretch', hide_index=True)

        top_colors = cl[cl["Status"].isin(["Super Fast","Fast"])]["Color"].tolist()
        if top_colors:
            st.markdown(f'<div class="insight">✅ Top colors: {", ".join(top_colors)}</div>',
                        unsafe_allow_html=True)
    else:
        st.info("No color data for this product.")

# ── Store Performance ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">🏪 Store Performance</div>', unsafe_allow_html=True)

if df_prodstore is not None:
    ps = df_prodstore[df_prodstore["Product Name"].str.strip() == sel_product].copy()
    if not ps.empty:
        ps["_order"] = ps["Store"].apply(
            lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
        ps = ps.sort_values("_order").drop(columns=["_order"])
        grand = ps["Units Sold"].sum()
        ps["Share %"] = (ps["Units Sold"] / grand * 100).round(1) if grand > 0 else 0
        ps[f"Order ({cover_days}d)"] = (
            reorder_qty * ps["Share %"] / 100
        ).round().astype(int)

        col_tbl, col_bar = st.columns([2,3])
        with col_tbl:
            disp_ps = ps[["Store","Units Sold","Revenue (NPR)","Share %",
                           f"Order ({cover_days}d)"]].copy()
            st.dataframe(
                disp_ps.style
                    .map(_style_order, subset=[f"Order ({cover_days}d)"])
                    .format({"Units Sold":"{:,.0f}","Revenue (NPR)":"NPR {:,.0f}",
                             "Share %":"{:.1f}%",f"Order ({cover_days}d)":"{:,.0f}"}),
                width='stretch', hide_index=True)
            st.caption(f"Source: product_store_sales.xlsx (all-time)")

        with col_bar:
            max_u = ps["Units Sold"].max() or 1
            for _, row in ps.iterrows():
                pct = row["Units Sold"] / max_u * 100
                st.markdown(
                    f'<div style="margin-bottom:5px">'
                    f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">'
                    f'<span><strong>{row["Store"]}</strong></span>'
                    f'<span style="color:#6b7280">{int(row["Units Sold"]):,} units · {row["Share %"]:.0f}%</span>'
                    f'</div>'
                    f'<div style="background:#e2e8f0;border-radius:4px;height:7px">'
                    f'<div style="background:#1d4ed8;width:{pct:.0f}%;height:7px;border-radius:4px"></div>'
                    f'</div></div>', unsafe_allow_html=True)
    else:
        st.info(f"**{sel_product}** has no POS sales recorded in product_store_sales.xlsx.")
else:
    st.info("Store data not available. Run `fetch_product_store_sales.py`.")

# ── Full SKU Breakdown ────────────────────────────────────────────────────────
st.markdown('<div class="sec">📋 Full SKU Breakdown — Every Size × Color from Odoo</div>',
            unsafe_allow_html=True)

sku_rows = prod_rows[["Color","Size","SKU / Variant","On Hand Qty",
                       "Total Units Sold","Sell-Through %","STR Status","Sales Price",
                       "Days of Cover","DOC Status"]].copy()
sku_rows = sku_rows.rename(columns={
    "SKU / Variant":"SKU","On Hand Qty":"On Hand",
    "Total Units Sold":"Units Sold","Sell-Through %":"STR %",
    "STR Status":"Status","Sales Price":"Price","Days of Cover":"DOC","DOC Status":"DOC Status"
})
sku_rows["_sk"] = sku_rows["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
sku_rows = sku_rows.sort_values(["Color","_sk"]).drop(columns=["_sk"])

st.dataframe(
    sku_rows.style
        .map(_style_status, subset=["Status"])
        .map(_style_stock,  subset=["On Hand"])
        .map(_style_str,    subset=["STR %"])
        .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","On Hand":"{:,.0f}",
                 "Price":"NPR {:,.0f}"}),
    width='stretch', hide_index=True)
st.caption(f"{len(sku_rows)} variants total")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    # Summary
    summary = pd.DataFrame([{
        "Product":sel_product,"Brand":sel_brand,
        "Total Sold":int(total_sold),"In Stock":int(total_stock),
        "STR %":str_pct,"Status":status,
        net_lbl:int(net_val),"Trend":trend,
        "Velocity (u/day)":round(daily_vel,4),
        "Rate/wk":round(weekly_rate,2),
        f"Order ({cover_days}d)":reorder_qty,
        "Avg Price":round(avg_price,0),
        "Est. Value":round(est_value,0),
    }])
    summary.to_excel(writer, sheet_name="Summary", index=False)

    if size_df is not None:
        sz_exp = size_df[size_df["Product Name"].str.strip()==sel_product].copy()
        if not sz_exp.empty:
            sz_exp["_sk"] = sz_exp["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
            sz_exp = sz_exp.sort_values("_sk").drop(columns=["_sk"])
            sz_exp.to_excel(writer, sheet_name="By Size", index=False)

    if color_df is not None:
        cl_exp = color_df[color_df["Product Name"].str.strip()==sel_product].copy()
        if not cl_exp.empty:
            cl_exp.to_excel(writer, sheet_name="By Color", index=False)

    sku_rows.to_excel(writer, sheet_name="SKU Breakdown", index=False)

    if df_prodstore is not None:
        ps_exp = df_prodstore[df_prodstore["Product Name"].str.strip()==sel_product].copy()
        if not ps_exp.empty:
            ps_exp.to_excel(writer, sheet_name="By Store", index=False)

out.seek(0)
safe_name = re.sub(r'[^\w\s-]','',sel_product)[:40].strip().replace(' ','_')
st.download_button(
    f"⬇️ Download {sel_product} — Full Report",
    data=out,
    file_name=f"deep_dive_{safe_name}_{today.strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("Download includes: Summary · By Size · By Color · SKU Breakdown · By Store")