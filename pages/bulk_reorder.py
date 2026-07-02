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
GDRIVE_LOCSTK_ID    = "1zgTBhh7vOTjxEIz-LO3YSM-TXJeDUrBT"

LOCATION_ORDER = ["Baneshwor","Lazimpat","Kumaripati","Chitwan","Pokhara","Online",
                  "Baneshwor Lush","Chitwan Lush","Pokhara Lush"]

# Stores excluded from clothing/accessories reorder by default
# (cosmetics-only locations that skew the data)
DEFAULT_EXCLUDED_STORES = {"Baneshwor Lush","Chitwan Lush","Pokhara Lush"}

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

    # Fix: some products have size embedded in name ("Dress/S", "Top/XL")
    # but Size column is empty — strip it out and populate Size
    SIZE_SUFFIXES = {"XS","S","M","L","XL","2XL","3XL","4XL","5XL",
                     "ONE SIZE","FREE SIZE","36","37","38","39","40","41","42","43","44"}
    _dash_re = re.compile(r'\s[-–]\s([A-Z0-9]{1,4})$')

    def _fix_name_size(row):
        name = row["Product Name"]
        size = row["Size"]
        # Pattern 1: "Product/S"
        if "/" in name:
            parts = name.rsplit("/", 1)
            suffix = parts[1].strip()
            if suffix.upper() in SIZE_SUFFIXES:
                return parts[0].strip(), size if size else suffix
        # Pattern 2: "Product - M"
        m = _dash_re.search(name)
        if m and m.group(1).upper() in SIZE_SUFFIXES:
            return name[:m.start()].strip(), size if size else m.group(1)
        return name, size

    if "Product Name" in df.columns and "Size" in df.columns:
        fixed = df.apply(_fix_name_size, axis=1, result_type="expand")
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
                         "ONE SIZE","FREE SIZE","36","37","38","39","40",
                         "41","42","43","44"}

    def _prep(df):
        df = df.copy()
        df.columns = [c.strip() for c in df.columns]
        df["Product Name"] = (df["Product Name"].fillna("").astype(str)
                              .str.replace('\n',' ',regex=False)
                              .str.replace('\t',' ',regex=False)
                              .str.replace(r'\s+',' ',regex=True)
                              .str.strip()
                              .str.strip('"'))

        # Strip size suffix from product name where it's embedded (e.g. "Dress/S" or "Dress - M")
        # Always strip if the suffix is a known size — even if Size column is already populated
        _dash_re2 = re.compile(r'\s[-–]\s([A-Z0-9]{1,4})$')

        def _fix_variant_name(row):
            name = row["Product Name"]
            size = str(row.get("Size","")).strip() if "Size" in row.index else ""
            # Pattern 1: "Product/S"
            if "/" in name:
                parts = name.rsplit("/", 1)
                suffix = parts[1].strip()
                if suffix.upper() in SIZE_SUFFIXES_SET:
                    return parts[0].strip(), size if size else suffix
            # Pattern 2: "Product - M"
            m = _dash_re2.search(name)
            if m and m.group(1).upper() in SIZE_SUFFIXES_SET:
                return name[:m.start()].strip(), size if size else m.group(1)
            return name, size

        if "Size" in df.columns:
            fixed = df.apply(_fix_variant_name, axis=1, result_type="expand")
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
    # Fix: strip size suffix from product names
    if "Product Name" in df.columns:
        SIZE_SUFFIXES_PS = {"XS","S","M","L","XL","2XL","3XL","4XL","5XL",
                            "ONE SIZE","FREE SIZE","36","37","38","39","40","41","42","43","44"}
        _dash_re_ps = re.compile(r'\s[-–]\s([A-Z0-9]{1,4})$')
        def _fix_ps_name(n):
            if "/" in n:
                parts = n.rsplit("/", 1)
                if parts[1].strip().upper() in SIZE_SUFFIXES_PS:
                    return parts[0].strip()
            m = _dash_re_ps.search(n)
            if m and m.group(1).upper() in SIZE_SUFFIXES_PS:
                return n[:m.start()].strip()
            return n
        df["Product Name"] = df["Product Name"].apply(_fix_ps_name)
    return df


@st.cache_resource(show_spinner=False)
def load_location_stock():
    """Load per-store stock from location_stock.xlsx — Store x Category sheet."""
    buf = _gdrive(GDRIVE_LOCSTK_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Store x Category", engine="openpyxl")
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export\exports"
        files = sorted(Path(base).glob("location_stock_*.xlsx"), reverse=True) if Path(base).exists() else []
        if files:
            try: df = pd.read_excel(files[0], sheet_name="Store x Category", engine="openpyxl")
            except: pass
    if df is None or df.empty: return None
    df.columns = [str(c).strip() for c in df.columns]
    cat_col   = df.columns[0]
    store_cols = [c for c in df.columns if c != cat_col]
    rows = []
    for _, row in df.iterrows():
        cat = str(row[cat_col]).strip()
        if not cat or cat.lower() in ("nan",""):
            continue
        for store in store_cols:
            qty = row[store]
            rows.append({"Category": cat, "Store": store.strip(),
                         "Stock": max(0.0, float(qty) if not pd.isna(qty) else 0.0)})
    return pd.DataFrame(rows) if rows else None
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
    # Fix: strip size suffix from product names
    if "Product Name" in df.columns:
        SIZE_SUFFIXES_PS = {"XS","S","M","L","XL","2XL","3XL","4XL","5XL",
                            "ONE SIZE","FREE SIZE","36","37","38","39","40","41","42","43","44"}
        _dash_re_ps = re.compile(r'\s[-–]\s([A-Z0-9]{1,4})$')
        def _fix_ps_name(n):
            if "/" in n:
                parts = n.rsplit("/", 1)
                if parts[1].strip().upper() in SIZE_SUFFIXES_PS:
                    return parts[0].strip()
            m = _dash_re_ps.search(n)
            if m and m.group(1).upper() in SIZE_SUFFIXES_PS:
                return n[:m.start()].strip()
            return n
        df["Product Name"] = df["Product Name"].apply(_fix_ps_name)
    return df


# ── Load ──────────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod           = load_products()
    size_df, color_df = load_variants()
    df_prodstore      = load_product_store()
    df_locstk         = load_location_stock()

if df_prod is None:
    st.error("Could not load product data."); st.stop()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🛒 Bulk Reorder Tool")
    st.markdown("---")

    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("","nan","True","False")])
    sel_brands = st.multiselect("Brand", brands, default=[brands[0]] if brands else [],
        help="Select one or more brands")

    # Categories cascade from selected brands
    _brand_df = df_prod[df_prod["Brand"].isin(sel_brands)] if sel_brands else df_prod
    cats = sorted([c for c in _brand_df["Category"].unique()
                   if c.strip().lower() not in JUNK_CATS])
    sel_cats = st.multiselect("Category", cats, default=[],
        help="Select one or more categories — leave empty to include all")

    # Sub-categories cascade from selected categories
    sel_subs = []
    if sel_cats and "Sub Category" in df_prod.columns:
        _cat_df = _brand_df[_brand_df["Category"].isin(sel_cats)]
        subs = sorted([s for s in _cat_df["Sub Category"].unique()
                       if s and s not in ("","nan")])
        if subs:
            sel_subs = st.multiselect("Sub Category", subs, default=[],
                help="Select one or more sub-categories — leave empty to include all")

    # Backwards-compat single values for header display
    sel_brand = ", ".join(sel_brands) if sel_brands else "All"
    sel_cat   = ", ".join(sel_cats)   if sel_cats   else "All"
    sel_sub   = ", ".join(sel_subs)   if sel_subs   else "All"

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
    show_zero = st.checkbox("Show products with no sales", value=False,
        help="When unchecked, hides products with zero all-time sales. Products with stock but low recent velocity are always shown.")

    st.markdown("---")
    st.markdown("**📈 Velocity Settings**")

    # Available windows depend on what columns the export has
    _has_30d  = "Recent Sold 30d"  in df_prod.columns if df_prod is not None else False
    _has_180d = "Recent Sold 180d" in df_prod.columns if df_prod is not None else False

    if _has_30d and _has_180d:
        _vel_options = [30, 60, 90, 180]
    elif _has_30d:
        _vel_options = [30, 60, 90]
    elif _has_180d:
        _vel_options = [60, 90, 180]
    else:
        _vel_options = [60, 90]

    velocity_days = st.select_slider(
        "Sales lookback window (days)", options=_vel_options, value=60,
        help=(
            "Sets the window for calculating daily sell rate.\n\n"
            "**30d** = recent momentum — aggressive, good mid-season\n"
            "**60d** = balanced (recommended)\n"
            "**90d** = smoothed average — conservative\n"
            "**180d** = half-year trend — very conservative\n\n"
            "Only windows available in the export are shown. "
            "Run `patch_odoo_windows.py` then re-export to unlock 30d and 180d."
        )
    )
    cover_days = st.slider(
        "Days of cover to reorder for", 30, 120, 60,
        step=15,
        help="How many days of stock to reorder. velocity × cover_days − stock = reorder qty"
    )

    st.markdown("---")
    st.markdown("**🏪 Store Filter**")
    all_stores = [s for s in LOCATION_ORDER]
    excluded_stores = st.multiselect(
        "Exclude stores",
        all_stores,
        default=list(DEFAULT_EXCLUDED_STORES),
        help="Lush stores are cosmetics-focused and skew clothing/accessories reorder data"
    )
    active_stores = [s for s in LOCATION_ORDER if s not in excluded_stores]
    date_opts = ["All time","Last 30 days","Last 60 days","Last 90 days",
                 "Older than 30 days","Older than 60 days","Older than 90 days"]
    sel_date = st.selectbox("Date filter", date_opts, index=0,
        help="'Older than X' excludes new arrivals that haven't had time to sell")

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_resource.clear(); st.rerun()

# ── Filter products ───────────────────────────────────────────────────────────
bdf = df_prod[df_prod["Brand"].isin(sel_brands)].copy() if sel_brands else df_prod.copy()
bdf = bdf[~bdf["Category"].str.strip().str.lower().isin(JUNK_CATS)]
if sel_cats:
    bdf = bdf[bdf["Category"].isin(sel_cats)]
if sel_subs and "Sub Category" in bdf.columns:
    bdf = bdf[bdf["Sub Category"].isin(sel_subs)]
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
grp_cols = ["Product Name","Brand","Category"]
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

# ── Velocity-based reorder — three-tier logic ─────────────────────────────────
#
# Tier 1 — No recent sales (Recent_60 = 0):
#   → Reorder = 0. No current demand signal. Show product but don't order.
#
# Tier 2 — New product (days_live < 90):
#   → Use recent velocity only (limited history, recent is most reliable)
#   → velocity = Recent_60 / 60
#
# Tier 3 — Established product (days_live ≥ 90, has recent sales):
#   → Use min(recent, lifetime) — conservative, avoids over-ordering
#   → If trend is dying: recent < lifetime → use recent (smaller order)
#   → If trending up: recent > lifetime → still limited by lifetime (safer)
#   → velocity = min(Recent_60/60, Total_Sold/days_live)
#
# Display velocity (Rate/wk) always shows recent when available,
# lifetime fallback when Recent_60 = 0.
# ─────────────────────────────────────────────────────────────────────────────

if "Create Date" in bdf.columns:
    dates = bdf.groupby("Product Name")["Create Date"].min().reset_index()
    dates["Create Date"] = pd.to_datetime(dates["Create Date"], errors="coerce")
    prod_sum = prod_sum.merge(dates, on="Product Name", how="left")
    prod_sum["days_live"] = ((today - prod_sum["Create Date"]).dt.days).fillna(365).clip(lower=7)
else:
    prod_sum["days_live"] = 365

has_recent = "Recent Sold 60d" in bdf.columns

NEW_PRODUCT_DAYS = 90  # threshold for Tier 2 vs Tier 3

if has_recent:
    recent_60 = bdf.groupby(grp_cols).agg(
        Recent_60 = ("Recent Sold 60d", "sum"),
    ).reset_index()
    prod_sum = prod_sum.merge(recent_60, on=grp_cols, how="left")
    prod_sum["Recent_60"] = prod_sum["Recent_60"].fillna(0)

    # Also load 90d window if available
    has_recent_90 = "Recent Sold 90d" in bdf.columns
    if has_recent_90:
        recent_90 = bdf.groupby(grp_cols).agg(
            Recent_90 = ("Recent Sold 90d", "sum"),
        ).reset_index()
        prod_sum = prod_sum.merge(recent_90, on=grp_cols, how="left")
        prod_sum["Recent_90"] = prod_sum["Recent_90"].fillna(0)
    else:
        prod_sum["Recent_90"] = prod_sum["Recent_60"]

    # Pick the best available window for velocity calculation
    # Use the exact window if available, else fall back to nearest
    _window_map = {}
    if "Recent Sold 30d"  in bdf.columns: _window_map[30]  = "Recent Sold 30d"
    if "Recent Sold 60d"  in bdf.columns: _window_map[60]  = "Recent Sold 60d"
    if "Recent Sold 90d"  in bdf.columns: _window_map[90]  = "Recent Sold 90d"
    if "Recent Sold 180d" in bdf.columns: _window_map[180] = "Recent Sold 180d"

    # Find the closest available window to what was selected
    _available = sorted(_window_map.keys())
    if velocity_days in _window_map:
        _vel_window = velocity_days
    else:
        # Use nearest available window
        _vel_window = min(_available, key=lambda x: abs(x - velocity_days))

    _vel_col = _window_map.get(_vel_window, "Recent Sold 60d")

    # Aggregate the chosen window
    vel_sales_agg = bdf.groupby(grp_cols).agg(
        _vel_sales = (_vel_col, "sum"),
    ).reset_index()
    prod_sum = prod_sum.merge(vel_sales_agg, on=grp_cols, how="left")
    prod_sum["_vel_sales"] = prod_sum["_vel_sales"].fillna(0)

    prod_sum["_recent_vel"]   = (prod_sum["_vel_sales"] / _vel_window).round(4)
    prod_sum["_lifetime_vel"] = (prod_sum["Total_Sold"] /
        prod_sum["days_live"].clip(upper=365).clip(lower=7)).round(4)

    def _calc_velocity(r):
        if r["_vel_sales"] == 0:
            return r["_lifetime_vel"]
        elif r["days_live"] < NEW_PRODUCT_DAYS:
            return r["_recent_vel"]
        else:
            return min(r["_recent_vel"], r["_lifetime_vel"])

    def _calc_reorder_vel(r):
        if r["_vel_sales"] == 0:
            return 0.0
        elif r["days_live"] < NEW_PRODUCT_DAYS:
            return r["_recent_vel"]
        else:
            return min(r["_recent_vel"], r["_lifetime_vel"])

    prod_sum["Daily_Velocity"]    = prod_sum.apply(_calc_velocity,     axis=1).round(4)
    prod_sum["_reorder_vel_daily"]= prod_sum.apply(_calc_reorder_vel,  axis=1).round(4)
    prod_sum["Weekly_Rate"]       = (prod_sum["Daily_Velocity"] * 7).round(2)
    prod_sum["effective_days"]    = velocity_days

    prod_sum["Reorder_Velocity"] = (
        prod_sum["_reorder_vel_daily"] * cover_days - prod_sum["Total_Stock"]
    ).clip(lower=0).round().astype(int)

    # Tier counts — use the selected velocity window
    t1 = (prod_sum["_vel_sales"] == 0).sum()
    t2 = ((prod_sum["_vel_sales"] > 0) & (prod_sum["days_live"] < NEW_PRODUCT_DAYS)).sum()
    t3 = ((prod_sum["_vel_sales"] > 0) & (prod_sum["days_live"] >= NEW_PRODUCT_DAYS)).sum()
    _window_note = f" (using {_vel_window}d window)" if _vel_window != velocity_days else ""
    st.sidebar.success(
        f"✅ Velocity tiers{_window_note}:\n"
        f"- {t3} established (min recent/lifetime)\n"
        f"- {t2} new <{NEW_PRODUCT_DAYS}d (recent only)\n"
        f"- {t1} no recent sales (reorder=0)")
else:
    _vel_window = velocity_days
    prod_sum["_vel_sales"]   = prod_sum["Total_Sold"]
    prod_sum["effective_days"]   = prod_sum["days_live"].clip(upper=velocity_days).clip(lower=7)
    prod_sum["Daily_Velocity"]   = (prod_sum["Total_Sold"] / prod_sum["effective_days"]).round(4)
    prod_sum["Weekly_Rate"]      = (prod_sum["Daily_Velocity"] * 7).round(2)
    prod_sum["Reorder_Velocity"] = (
        prod_sum["Daily_Velocity"] * cover_days - prod_sum["Total_Stock"]
    ).clip(lower=0).round().astype(int)
    st.sidebar.warning(
        "⚠️ Using all-time sales for velocity — re-export products to get Recent Sold 60d")


# Velocity tier label — always based on selected window vs lifetime
if has_recent:
    def _tier_label(r):
        if r["_vel_sales"] == 0:               return "🔴 No demand"
        if r["days_live"] < NEW_PRODUCT_DAYS:  return "🆕 New (<90d)"
        rv = r["_recent_vel"]; lv = r["_lifetime_vel"]
        if rv < lv * 0.8:  return "📉 Slowing"
        if rv > lv * 1.2:  return "📈 Trending"
        return "✅ Stable"
    prod_sum["Vel_Tier"] = prod_sum.apply(_tier_label, axis=1)
else:
    prod_sum["Vel_Tier"] = "—"

# Net_Sales for KPI display — use the window that matches velocity_days
prod_sum["Net_Sales"] = prod_sum["_vel_sales"] if has_recent else prod_sum["Total_Sold"]
prod_sum["weeks_live"]  = (prod_sum["days_live"] / 7).clip(lower=1)
prod_sum["Est_Value"]   = prod_sum["Reorder_Velocity"] * prod_sum["Avg_Price"]

# ── Last sold date — exact from export or estimated from windows ──────────────
has_last_sold = "Last Sold Date" in bdf.columns and "Days Not Sold" in bdf.columns

if has_last_sold:
    last_sold_agg = bdf.groupby(grp_cols).agg(
        _Last_Sold_Date = ("Last Sold Date", lambda x: x.dropna().max()),
        _Days_Not_Sold  = ("Days Not Sold",  lambda x: pd.to_numeric(x, errors="coerce").min()),
    ).reset_index()
    prod_sum = prod_sum.merge(last_sold_agg, on=grp_cols, how="left")

    def _last_sold_label(r):
        if r["Total_Sold"] == 0: return "Never sold"
        d = r.get("_Days_Not_Sold")
        if pd.isna(d) or d is None: return "> 90d"
        d = int(d)
        if d == 0:  return "Today"
        if d == 1:  return "1d ago"
        return f"{d}d ago"

    prod_sum["Last Sold"]      = prod_sum.apply(_last_sold_label, axis=1)
    prod_sum["Days Not Sold"]  = prod_sum["_Days_Not_Sold"].fillna(9999).astype(int)
    prod_sum["_days_not_sold"] = prod_sum["Days Not Sold"]

else:
    def _last_sold_signal(r):
        if r["Total_Sold"] == 0: return "Never sold", 9999
        vs = r.get("_vel_sales", 0)
        if vs > 0:    return f"< {_vel_window}d", 30
        else:         return f"> {_vel_window}d", _vel_window + 30

    signals = prod_sum.apply(_last_sold_signal, axis=1, result_type="expand")
    prod_sum["Last Sold"]      = signals[0]
    prod_sum["Days Not Sold"]  = signals[1]
    prod_sum["_days_not_sold"] = signals[1]

# ── Launch Date and Days Live (from Launch) ───────────────────────────────────
has_launch = "Launch Date" in bdf.columns
if has_launch:
    launch_agg = bdf.groupby(grp_cols).agg(
        Launch_Date = ("Launch Date", "min")
    ).reset_index()
    launch_agg["Launch_Date"] = pd.to_datetime(launch_agg["Launch_Date"], errors="coerce")
    prod_sum = prod_sum.merge(launch_agg, on=grp_cols, how="left")
    prod_sum["Days Live"] = ((today - prod_sum["Launch_Date"]).dt.days).fillna(
        prod_sum["days_live"]).clip(lower=1).astype(int)
    prod_sum["Launch Date"] = prod_sum["Launch_Date"].dt.strftime("%Y-%m-%d").fillna("")
else:
    prod_sum["Days Live"]   = prod_sum["days_live"].astype(int)
    prod_sum["Launch Date"] = ""

# Apply STR filter
prod_sum = prod_sum[prod_sum["STR_Pct"] >= min_str_pct]
if not show_zero:
    # Only hide products with zero sales entirely — keep well-stocked products
    # (a product with stock > cover target has Reorder=0 but is still relevant)
    prod_sum = prod_sum[prod_sum["Total_Sold"] > 0]

prod_sum = prod_sum.sort_values("Total_Sold", ascending=False)

total_units_vel   = int(prod_sum["Reorder_Velocity"].sum())
total_net_sales   = int(prod_sum["Net_Sales"].sum())
avg_velocity      = prod_sum["Daily_Velocity"].mean()
total_value       = prod_sum["Est_Value"].sum()
n_products        = len(prod_sum)
fast_count        = prod_sum["STR_Status"].isin(["Super Fast","Fast"]).sum()

# ── Page header ───────────────────────────────────────────────────────────────
st.title("🛒 Bulk Reorder Tool")
filter_parts = [sel_brand, sel_cat]
if sel_sub != "All": filter_parts.append(sel_sub)
if search.strip(): filter_parts.append(f'"{search}"')
filter_parts.append(sel_season_raw)
if sel_date != "All time": filter_parts.append(sel_date)
st.markdown(
    f"**{'  ·  '.join(filter_parts)}** · STR ≥ {min_str_pct}% · "
    f"Velocity {velocity_days}d lookback · {cover_days}d cover · {today.strftime('%b %d, %Y')}")

# ── KPIs ──────────────────────────────────────────────────────────────────────
c1,c2,c3,c4,c5 = st.columns(5)
net_lbl = f"Net Sales ({_vel_window}d)" if has_recent else "Total Sold (all-time)"
for col, val, lbl, clr in [
    (c1, f"{n_products:,}",        "Products",                         "#374151"),
    (c2, f"{fast_count:,}",        "Fast / Super Fast",                "#16a34a"),
    (c3, f"{total_net_sales:,}",   net_lbl,                            "#0f766e"),
    (c4, f"{total_units_vel:,}",   f"Order Qty ({cover_days}d cover)", "#1d4ed8"),
    (c5, fmt_npr(total_value),     "Est. Value",                       "#374151"),
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
    Products      = ("Product Name",    "count"),
    Units_Sold    = ("Total_Sold",      "sum"),
    Net_Sales     = ("Net_Sales",       "sum"),
    In_Stock      = ("Total_Stock",     "sum"),
    Avg_STR       = ("STR_Pct",         "mean"),
    Order_Vel     = ("Reorder_Velocity","sum"),
    Est_Value     = ("Est_Value",       "sum"),
).reset_index().sort_values(["Category","Order_Vel"], ascending=[True,False])

# Velocity at category level = total net sales / lookback days (not sum of individual velocities)
cat_sum["Velocity_Day"] = (cat_sum["Net_Sales"] / velocity_days).round(2)
cat_sum["Weekly_Rate"]  = (cat_sum["Velocity_Day"] * 7).round(1)
cat_sum["Avg_STR"]      = cat_sum["Avg_STR"].round(1)
cat_sum["Est_Value"]    = cat_sum["Est_Value"].apply(fmt_npr)
cat_sum = cat_sum.rename(columns={
    "Products":"# Products","Units_Sold":"Units Sold",
    "Net_Sales":net_lbl,
    "In_Stock":"In Stock","Avg_STR":"Avg STR %",
    "Velocity_Day":"Velocity (u/day)","Weekly_Rate":"Rate/wk",
    "Order_Vel":f"Order ({cover_days}d)","Est_Value":"Est. Value"
})

def _cat_style(val):
    if isinstance(val,(int,float)) and val > 0:
        return "background-color:#dbeafe;color:#1e40af;font-weight:700"
    return ""

def _vel_style(val):
    if isinstance(val,(int,float)) and val > 0:
        return "color:#0f766e;font-weight:600"
    return ""

disp_cat_cols = ["Category"] + \
    (["Sub Category"] if has_sub else []) + \
    ["# Products","Units Sold", net_lbl, "In Stock","Avg STR %",
     "Velocity (u/day)","Rate/wk", f"Order ({cover_days}d)","Est. Value"]
disp_cat_cols = [c for c in disp_cat_cols if c in cat_sum.columns]

st.dataframe(
    cat_sum[disp_cat_cols].style
        .map(_cat_style, subset=[f"Order ({cover_days}d)"])
        .map(_vel_style,  subset=["Velocity (u/day)"])
        .format({"Avg STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                 net_lbl:"{:,.0f}","Velocity (u/day)":"{:.2f}","Rate/wk":"{:.1f}",
                 f"Order ({cover_days}d)":"{:,.0f}"}),
    width='stretch', hide_index=True)
st.caption(
    f"**{net_lbl}** = units sold in the lookback window · "
    f"**Velocity (u/day)** = {net_lbl} ÷ {velocity_days}d · "
    f"**Order ({cover_days}d)** = velocity × {cover_days}d − stock"
)

# ── Size Breakdown by Category ────────────────────────────────────────────────
st.markdown('<div class="sec">📐 Size Breakdown by Category</div>', unsafe_allow_html=True)

if size_df is None:
    st.info("Size data not available — run `variant_export.py` first.")
else:
    # Filter size_df to current brand + category + sub filters
    _sz_cat = size_df[size_df["Brand"].str.strip().isin(sel_brands)].copy() if sel_brands else size_df.copy()
    _sz_cat = _sz_cat[~_sz_cat["Category"].str.strip().str.lower().isin(JUNK_CATS)]
    if sel_cats and "Category" in _sz_cat.columns:
        _sz_cat = _sz_cat[_sz_cat["Category"].str.strip().isin(sel_cats)]
    if sel_subs and "Sub Category" in _sz_cat.columns:
        _sz_cat = _sz_cat[_sz_cat["Sub Category"].str.strip().isin(sel_subs)]
    if search.strip():
        _sz_cat = _sz_cat[_sz_cat["Product Name"].str.contains(search.strip(), case=False, na=False)]

    if _sz_cat.empty:
        st.info("No size data for current filters.")
    else:
        # Style functions needed for dataframe styling
        def _sz_reorder_style(val):
            if isinstance(val,(int,float)) and val > 0:
                return "background-color:#dbeafe;color:#1e40af;font-weight:700"
            return ""
        def _sz_stock_style(val):
            if isinstance(val,(int,float)) and val == 0:
                return "background-color:#fee2e2;color:#991b1b"
            return ""
        def _sz_str_style(val):
            if not isinstance(val,(int,float)): return ""
            if val >= 70: return "background-color:#dcfce7;color:#166534"
            if val >= 30: return "background-color:#fef9c3;color:#854d0e"
            return "background-color:#fee2e2;color:#991b1b"

        # Aggregate by Category + Sub Category + Size
        sz_grp = ["Category"]
        if "Sub Category" in _sz_cat.columns and _sz_cat["Sub Category"].str.strip().ne("").any():
            sz_grp.append("Sub Category")
        sz_grp.append("Size")

        sz_cat_agg = _sz_cat.groupby(sz_grp, as_index=False).agg(
            Units_Sold = ("Units Sold", "sum"),
            In_Stock   = ("In Stock",   "sum"),
        )
        sz_cat_agg["STR %"] = (sz_cat_agg["Units_Sold"] /
            (sz_cat_agg["Units_Sold"] + sz_cat_agg["In_Stock"]).replace(0, float("nan")) * 100
        ).fillna(0).round(1)

        # Rate/wk (lifetime) — size_df has no recent data, use lifetime avg days
        avg_days_live = prod_sum["days_live"].mean() if "days_live" in prod_sum.columns else 365
        sz_cat_agg["Rate/wk"] = (sz_cat_agg["Units_Sold"] / max(avg_days_live, 1) * 7).round(2)

        # Reorder — use same velocity logic as category summary:
        # Distribute each category's total Order(60d) proportionally by size's share of Units Sold
        # This keeps size reorder CONSISTENT with category reorder total
        #
        # Build cat_key → Order(60d) lookup from cat_sum
        cat_order_lookup = {}
        for _, r in cat_sum.iterrows():
            key = tuple(str(r.get(c,"")).strip() for c in cat_grp)
            cat_order_lookup[key] = float(r.get(f"Order ({cover_days}d)", 0) or 0)

        def _size_reorder(row):
            # Find this size's category group
            key = tuple(str(row.get(c,"")).strip() for c in sz_grp[:-1])  # exclude Size
            cat_total_order = cat_order_lookup.get(key, 0)
            if cat_total_order == 0:
                return 0
            # This category's total units sold across all sizes
            cat_mask = pd.Series(True, index=sz_cat_agg.index)
            for ci, c in enumerate(sz_grp[:-1]):
                cat_mask = cat_mask & (sz_cat_agg[sz_grp[ci]] == row[sz_grp[ci]])
            cat_total_sold = sz_cat_agg.loc[cat_mask, "Units_Sold"].sum()
            if cat_total_sold == 0:
                return 0
            size_share = row["Units_Sold"] / cat_total_sold
            return round(cat_total_order * size_share)

        sz_cat_agg["Reorder"] = sz_cat_agg.apply(_size_reorder, axis=1).astype(int)

        # Sort sizes correctly
        sz_cat_agg["_sk"] = sz_cat_agg["Size"].apply(
            lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        sz_cat_agg = sz_cat_agg.sort_values(
            sz_grp[:-1] + ["_sk"], ascending=True
        ).drop(columns=["_sk"])

        sz_cat_agg = sz_cat_agg.rename(columns={"Units_Sold":"Units Sold","In_Stock":"In Stock"})

        disp_sz_cat = [c for c in sz_grp + ["Units Sold","In Stock","STR %","Rate/wk","Reorder"]
                       if c in sz_cat_agg.columns]

        st.dataframe(
            sz_cat_agg[disp_sz_cat].style
                .map(_sz_reorder_style, subset=["Reorder"])
                .map(_sz_stock_style,   subset=["In Stock"])
                .map(_sz_str_style,     subset=["STR %"])
                .format({"Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                         "STR %":"{:.1f}%","Rate/wk":"{:.2f}","Reorder":"{:,.0f}"}),
            width='stretch', hide_index=True)
        st.caption(
            f"{len(sz_cat_agg):,} size rows · "
            f"🔵 Reorder = category Order({cover_days}d) split proportionally by size · "
            f"🔴 Red stock = sold out · Rate/wk = lifetime average"
        )


st.markdown('<div class="sec">📋 Product-Level Reorder Plan</div>', unsafe_allow_html=True)

show_cols = ["Product Name","Brand","Category"] + \
    (["Sub Category"] if has_sub else []) + \
    ["STR_Status","STR_Pct","Total_Sold","Net_Sales",
     "Daily_Velocity","Weekly_Rate","Vel_Tier",
     "Launch Date","Days Live","Last Sold",
     "Total_Stock","Reorder_Velocity","Avg_Price","Est_Value"]
show_cols = [c for c in show_cols if c in prod_sum.columns]

disp = prod_sum[show_cols].copy().rename(columns={
    "STR_Status":"Status","STR_Pct":"STR %","Total_Sold":"Units Sold",
    "Net_Sales":net_lbl,"Daily_Velocity":"Velocity (u/day)",
    "Weekly_Rate":"Rate/wk","Vel_Tier":"Trend",
    "Launch Date":"Launch Date","Days Live":"Days Live",
    "Last Sold":"Last Sold",
    "Total_Stock":"In Stock",
    "Reorder_Velocity":f"Order ({cover_days}d)",
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

def _style_last_sold(val):
    if not isinstance(val, str): return ""
    if "< 60d"   in val: return "color:#16a34a;font-weight:600"   # green — recent
    if "60–90d"  in val: return "color:#d97706;font-weight:600"   # amber — getting old
    if "> 90d"   in val: return "color:#dc2626;font-weight:600"   # red — stale
    if "Never"   in val: return "color:#9ca3af"                   # grey — never sold
    return ""

def _style_last_sold(val):
    """Color the Last Sold column based on Days Not Sold."""
    if not isinstance(val, str): return ""
    if val in ("Today", "1d ago"): return "color:#16a34a;font-weight:700"
    try:
        # Extract number from "Nd ago"
        d = int(val.replace("d ago","").strip())
        if d <= 30:  return "color:#16a34a;font-weight:600"
        if d <= 90:  return "color:#d97706;font-weight:600"
        return "color:#dc2626;font-weight:600"
    except:
        if "Never" in val: return "color:#9ca3af"
        if "< 60d" in val: return "color:#16a34a;font-weight:600"
        if "60–90d" in val: return "color:#d97706;font-weight:600"
        if "> 90d" in val:  return "color:#dc2626;font-weight:600"
    return ""

def _style_days_live(val):
    if not isinstance(val, (int, float)): return ""
    if val <= 30:  return "color:#7c3aed"   # purple — very new
    if val <= 90:  return "color:#0369a1"   # blue — new
    return "color:#6b7280"                  # grey — established

fmt_d = {"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
         net_lbl:"{:,.0f}","Velocity (u/day)":"{:.3f}",
         "Rate/wk":"{:.2f}","Avg Price":"NPR {:,.0f}","Est. Value":"{:,.0f}",
         f"Order ({cover_days}d)":"{:,.0f}","Days Live":"{:,.0f}"}
_st = disp.style.map(_style_status, subset=["Status"])
if f"Order ({cover_days}d)" in disp.columns: _st = _st.map(_style_order,     subset=[f"Order ({cover_days}d)"])
if "Velocity (u/day)"       in disp.columns: _st = _st.map(_vel_style,       subset=["Velocity (u/day)"])
if "Last Sold"              in disp.columns: _st = _st.map(_style_last_sold, subset=["Last Sold"])
if "Days Live"              in disp.columns: _st = _st.map(_style_days_live, subset=["Days Live"])
st.dataframe(_st.format(fmt_d), width='stretch', hide_index=True)
st.caption(f"{len(disp):,} products · 🔵 Order ({cover_days}d) = velocity × {cover_days}d − stock · velocity = {net_lbl} ÷ {velocity_days}d")

# ── Pre-compute size / color / store data ─────────────────────────────────────
sz = pd.DataFrame()
if size_df is not None:
    _sz = size_df[size_df["Brand"].str.strip().isin(sel_brands)].copy() if sel_brands else size_df.copy()
    if sel_cats and "Category" in _sz.columns:
        _sz = _sz[_sz["Category"].str.strip().isin(sel_cats)]
    filtered_products_set = set(prod_sum["Product Name"].str.strip())
    _sz = _sz[_sz["Product Name"].str.strip().isin(filtered_products_set)]
    if not _sz.empty:
        _sz["_sk"] = _sz["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        _sz = _sz.sort_values(["Product Name","_sk"]).drop(columns=["_sk"])
        rate_map    = prod_sum.set_index("Product Name")["Weekly_Rate"].to_dict()
        reorder_map = prod_sum.set_index("Product Name")["Reorder_Velocity"].to_dict()
        _sz["_prod_rate"]    = _sz["Product Name"].map(rate_map).fillna(0)
        _sz["_prod_reorder"] = _sz["Product Name"].map(reorder_map).fillna(0)
        prod_total_sold = _sz.groupby("Product Name")["Units Sold"].transform("sum")
        _sz["_size_share"] = (_sz["Units Sold"] / prod_total_sold.replace(0, float("nan"))).fillna(
            1.0 / _sz.groupby("Product Name")["Units Sold"].transform("count"))
        _sz["Weekly Rate"] = (_sz["_prod_rate"]    * _sz["_size_share"]).round(2)
        # Order (Vel) = product's Reorder_Velocity (velocity-based, respects tiers)
        # split by this size's share of product sales — consistent with category breakdown
        _sz["Order (Vel)"] = (_sz["_prod_reorder"] * _sz["_size_share"]).round().astype(int)
        sz = _sz

cl = pd.DataFrame()
if color_df is not None:
    _cl = color_df[color_df["Brand"].str.strip().isin(sel_brands)].copy() if sel_brands else color_df.copy()
    if sel_cats and "Category" in _cl.columns:
        _cl = _cl[_cl["Category"].str.strip().isin(sel_cats)]
    filtered_products_set = set(prod_sum["Product Name"].str.strip())
    _cl = _cl[_cl["Product Name"].str.strip().isin(filtered_products_set)]
    if not _cl.empty:
        _cl = _cl[_cl["Status"].isin(["Super Fast","Fast"])]
        _cl = _cl.sort_values(["Product Name","Units Sold"], ascending=[True,False])

        # Add reorder qty per color — product reorder × color share of sales
        reorder_map = prod_sum.set_index("Product Name")["Reorder_Velocity"].to_dict()
        _cl["_prod_reorder"] = _cl["Product Name"].map(reorder_map).fillna(0)
        _prod_total = _cl.groupby("Product Name")["Units Sold"].transform("sum")
        _cl["Color Share %"] = (_cl["Units Sold"] / _prod_total.replace(0, float("nan")) * 100).fillna(0).round(1)
        _cl[f"Order ({cover_days}d)"] = (
            _cl["_prod_reorder"] * _cl["Color Share %"] / 100
        ).round().astype(int)
        cl = _cl

_ps_all = None
if df_prodstore is not None:
    _ps_all = df_prodstore[df_prodstore["Brand"].str.strip().isin(sel_brands)].copy() if sel_brands else df_prodstore.copy()
    if sel_cats and "Category" in _ps_all.columns:
        _ps_all = _ps_all[_ps_all["Category"].str.strip().isin(sel_cats)]

def _style_sz_status(val):
    return {"Super Fast":"background-color:#1B5E20;color:white","Fast":"background-color:#43A047;color:white",
            "Medium":"background-color:#F9A825;color:black","Slow":"background-color:#E53935;color:white",
            "Dead":"background-color:#424242;color:white"}.get(val,"")
def _style_sz_order(val):
    return "background-color:#dbeafe;color:#1e40af;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""
def _style_sz_str(val):
    return "background-color:#ede9fe;color:#5b21b6;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""

# ── Size Breakdown by Product (flat table) ────────────────────────────────────
st.markdown('<div class="sec">📏 Size Breakdown by Product</div>', unsafe_allow_html=True)
if sz.empty:
    st.info("No size data for the current filters." if size_df is not None else "Variant data not available.")
else:
    disp_sz = sz[["Product Name","Size","Units Sold","In Stock","STR %","Status","Weekly Rate","Order (Vel)"]].copy()
    _sst = (disp_sz.style
        .map(_style_sz_status, subset=["Status"])
        .map(_style_sz_order,  subset=["Order (Vel)"])
        .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                 "Weekly Rate":"{:.2f}","Order (Vel)":"{:,.0f}"}))
    st.dataframe(_sst, width='stretch', hide_index=True)
    st.caption(f"{len(sz):,} size rows · 🔵 Order (Vel) = velocity × {cover_days}d − stock")

# ── Color Breakdown by Product (flat table) ───────────────────────────────────
st.markdown('<div class="sec">🎨 Color Breakdown by Product</div>', unsafe_allow_html=True)
if cl.empty:
    st.info("No Fast/Super Fast colors for current filters.")
else:
    disp_cl_cols = ["Product Name","Color","Units Sold","In Stock","STR %",
                    "Color Share %","Status",f"Order ({cover_days}d)"]
    disp_cl_cols = [c for c in disp_cl_cols if c in cl.columns]
    disp_cl = cl[disp_cl_cols].copy()

    def _style_color_order(val):
        if isinstance(val,(int,float)) and val > 0:
            return "background-color:#dbeafe;color:#1e40af;font-weight:700"
        return ""

    _cst = (disp_cl.style
        .map(_style_sz_status, subset=["Status"])
        .map(_style_color_order, subset=[f"Order ({cover_days}d)"] if f"Order ({cover_days}d)" in disp_cl_cols else [])
        .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                 "Color Share %":"{:.1f}%",f"Order ({cover_days}d)":"{:,.0f}"}))
    st.dataframe(_cst, width='stretch', hide_index=True)
    st.caption(f"{len(cl):,} color rows · Fast/Super Fast only · 🔵 Order = product reorder × color share")

# ── Overall Store Distribution (category-level summary) ──────────────────────
st.markdown('<div class="sec">🏪 Overall Store Distribution</div>', unsafe_allow_html=True)
st.caption("Total reorder split across stores. Click a product in the popup above for per-product store breakdown.")

if df_prodstore is None:
    st.info("Store sales data not available. Run `fetch_product_store_sales.py` and set GDRIVE_PRODSTORE_ID.")
else:
    ps = df_prodstore[df_prodstore["Brand"].str.strip().isin(sel_brands)].copy() if sel_brands else df_prodstore.copy()
    if sel_cats and "Category" in ps.columns:
        ps = ps[ps["Category"].str.strip().isin(sel_cats)]
    if search.strip() and "Product Name" in ps.columns:
        ps = ps[ps["Product Name"].str.contains(search.strip(), case=False, na=False)]
    # Note: excluded stores are shown in table but get Order=0 (see store_totals logic)

    if ps.empty:
        st.info(f"No store sales data for **{sel_brand}**.")
    else:
        tab_store, tab_catstore = st.tabs(["📍 By Store", "📊 Category × Store"])

        store_totals = ps.groupby("Store").agg(
            Units_Sold = ("Units Sold","sum"),
            Revenue    = ("Revenue (NPR)","sum"),
        ).reset_index()
        store_totals["_order"] = store_totals["Store"].apply(
            lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
        store_totals = store_totals.sort_values("_order").drop(columns=["_order"])
        store_totals = store_totals[store_totals["Units_Sold"] > 0]

        # Add current stock per store from location_stock
        if df_locstk is not None and sel_cats:
            # Filter location stock to selected categories
            _lsk = df_locstk[df_locstk["Category"].isin(sel_cats)] if sel_cats else df_locstk
            _store_stock = _lsk.groupby("Store")["Stock"].sum().reset_index()
            _store_stock.columns = ["Store","On Hand"]
            store_totals = store_totals.merge(_store_stock, on="Store", how="left")
            store_totals["On Hand"] = store_totals["On Hand"].fillna(0).astype(int)
        elif df_locstk is not None and not sel_cats:
            _store_stock = df_locstk.groupby("Store")["Stock"].sum().reset_index()
            _store_stock.columns = ["Store","On Hand"]
            store_totals = store_totals.merge(_store_stock, on="Store", how="left")
            store_totals["On Hand"] = store_totals["On Hand"].fillna(0).astype(int)
        else:
            store_totals["On Hand"] = None

        # Share % and Order only from active (non-excluded) stores
        # Lush stores shown for visibility but don't receive reorder allocation
        active_mask   = store_totals["Store"].isin(active_stores)
        grand_sold    = store_totals.loc[active_mask, "Units_Sold"].sum()
        store_totals["Share_%"]   = store_totals.apply(
            lambda r: round(r["Units_Sold"] / grand_sold * 100, 1)
            if r["Store"] in active_stores and grand_sold > 0 else 0, axis=1)
        store_totals["Order_Vel"] = store_totals.apply(
            lambda r: round(r["Share_%"] / 100 * total_units_vel)
            if r["Store"] in active_stores else 0, axis=1).astype(int)
        store_totals["_excluded"] = ~store_totals["Store"].isin(active_stores)

        def _style_ord(val):
            return "background-color:#dbeafe;color:#1e40af;font-weight:700" if isinstance(val,(int,float)) and val > 0 else ""
        def _style_excluded_row(row):
            if row.get("Store","") in excluded_stores:
                return ["color:#94a3b8;font-style:italic"] * len(row)
            return [""] * len(row)

        with tab_store:
            col_tbl, col_bar = st.columns([2, 3])
            with col_tbl:
                _st_disp_cols = ["Store","Units_Sold","On Hand","Share_%","Order_Vel"]
                _st_disp_cols = [c for c in _st_disp_cols if c in store_totals.columns]
                disp_st = store_totals[_st_disp_cols].rename(columns={
                    "Units_Sold":"Units Sold","Share_%":"Share %",
                    "Order_Vel":f"Order ({cover_days}d)"})
                _fmt_st = {"Units Sold":"{:,.0f}","Share %":"{:.1f}%",
                           f"Order ({cover_days}d)":"{:,.0f}"}
                if "On Hand" in disp_st.columns:
                    _fmt_st["On Hand"] = "{:,.0f}"
                st.dataframe(
                    disp_st.style
                        .apply(_style_excluded_row, axis=1)
                        .map(_style_ord, subset=[f"Order ({cover_days}d)"])
                        .format(_fmt_st),
                    width='stretch', hide_index=True)
                active_order_total = store_totals.loc[active_mask, "Order_Vel"].sum()
                excl_note = f" · *{', '.join(excluded_stores)} shown but excluded from order*" if excluded_stores else ""
                st.caption(f"Total: {active_order_total:,} units · {active_mask.sum()} active stores{excl_note}")
            with col_bar:
                max_u = store_totals["Units_Sold"].max() or 1
                for _, row in store_totals.iterrows():
                    pct        = row["Units_Sold"] / max_u * 100
                    is_excl    = row["Store"] in excluded_stores
                    bar_color  = "#94a3b8" if is_excl else "#1d4ed8"
                    name_style = "color:#94a3b8;font-style:italic" if is_excl else ""
                    excl_tag   = ' <span style="font-size:10px;color:#94a3b8">(excluded from order)</span>' if is_excl else ""
                    st.markdown(
                        f'<div style="margin-bottom:6px">'
                        f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">'
                        f'<span style="{name_style}"><strong>{row["Store"]}</strong>{excl_tag}</span>'
                        f'<span style="color:#6b7280">{int(row["Units_Sold"]):,} units · '
                        f'{row["Share_%"]:.0f}% · <span style="color:{bar_color}">Order {int(row["Order_Vel"])}</span></span>'
                        f'</div>'
                        f'<div style="background:#e2e8f0;border-radius:4px;height:8px">'
                        f'<div style="background:{bar_color};width:{pct:.0f}%;height:8px;border-radius:4px"></div>'
                        f'</div></div>', unsafe_allow_html=True)

        with tab_catstore:
            # Apply sub-category filter to store data (ps is only filtered by category)
            ps_cat = ps.copy()
            if sel_subs and "Sub Category" in ps_cat.columns:
                ps_cat = ps_cat[ps_cat["Sub Category"].isin(sel_subs)]

            grp_key = ["Category","Sub Category","Store"] if "Sub Category" in ps_cat.columns else ["Category","Store"]
            cat_store = ps_cat.groupby(grp_key).agg(Units_Sold=("Units Sold","sum")).reset_index()
            # Show all stores in units sold table but only active stores in order table
            all_stores_present   = [s for s in LOCATION_ORDER if s in cat_store["Store"].unique()]
            stores_present       = [s for s in active_stores   if s in cat_store["Store"].unique()]
            pivot_cols = ["Category","Sub Category"] if "Sub Category" in cat_store.columns else ["Category"]

            # ── Units Sold pivot — shows ALL stores including Lush ──────────────
            pivot = cat_store.pivot_table(
                index=pivot_cols, columns="Store", values="Units_Sold",
                aggfunc="sum", fill_value=0
            ).reset_index()
            pivot.columns.name = None
            all_store_cols = [c for c in all_stores_present if c in pivot.columns]
            store_cols_present = [c for c in stores_present if c in pivot.columns]
            pivot["Total"] = pivot[all_store_cols].sum(axis=1)
            pivot = pivot.sort_values("Total", ascending=False)

            st.markdown("**Units Sold per Category per Store**")
            st.dataframe(
                pivot[[c for c in pivot_cols + all_store_cols + ["Total"] if c in pivot.columns]]
                .style.format({c: "{:,.0f}" for c in all_store_cols + ["Total"]}),
                width='stretch', hide_index=True)

            # ── Total Stock pivot per Category per Store ─────────────────────
            # Note: location_stock is at parent Category level (no Sub Category)
            # so we show one row per Category (not sub-category)
            stk_pivot = None
            if df_locstk is not None:
                _lsk_cat = df_locstk.copy()
                if sel_cats:
                    _lsk_cat = _lsk_cat[_lsk_cat["Category"].isin(sel_cats)]
                if not _lsk_cat.empty:
                    stk_pivot = _lsk_cat.pivot_table(
                        index="Category", columns="Store", values="Stock",
                        aggfunc="sum", fill_value=0
                    ).reset_index()
                    stk_pivot.columns.name = None
                    _stk_store_cols = [c for c in all_store_cols if c in stk_pivot.columns]
                    stk_pivot["Total Stock"] = stk_pivot[_stk_store_cols].sum(axis=1)
                    stk_pivot = stk_pivot.sort_values("Total Stock", ascending=False)
                    if stk_pivot["Total Stock"].sum() > 0:
                        st.markdown("**Total Stock per Category per Store**")
                        st.caption("Stock from location_stock.xlsx — parent category level only")
                        st.dataframe(
                            stk_pivot[[c for c in ["Category"] + _stk_store_cols + ["Total Stock"]
                                       if c in stk_pivot.columns]]
                            .style.format({c: "{:,.0f}" for c in _stk_store_cols + ["Total Stock"]}),
                            width='stretch', hide_index=True)

            # ── Reorder Qty pivot — distribute category Order(60d) by store share ──
            st.markdown(f"**Order ({cover_days}d) per Category per Store**")
            st.caption("Each category's total reorder qty split by that store's share of sales")

            # Build category total reorder lookup from cat_sum
            cat_reorder_map = {}
            for _, r in cat_sum.iterrows():
                key = tuple(r[c] for c in cat_grp)
                cat_reorder_map[key] = r.get(f"Order ({cover_days}d)", 0)

            reorder_rows = []
            for _, row in pivot.iterrows():
                key = tuple(row[c] for c in pivot_cols)
                # Match against cat_grp keys (may differ if has_sub mismatches)
                total_reorder = cat_reorder_map.get(key, 0)
                if not isinstance(total_reorder,(int,float)): total_reorder = 0
                row_total_sold = row["Total"]
                new_row = {c: row[c] for c in pivot_cols}
                for store in store_cols_present:
                    share = row[store] / row_total_sold if row_total_sold > 0 else 0
                    new_row[store] = round(total_reorder * share)
                new_row["Total Order"] = total_reorder
                reorder_rows.append(new_row)

            reorder_pivot = pd.DataFrame(reorder_rows)
            if not reorder_pivot.empty:
                reorder_pivot = reorder_pivot.sort_values("Total Order", ascending=False)

                def _ro_style(val):
                    if isinstance(val,(int,float)) and val > 0:
                        return "background-color:#dbeafe;color:#1e40af;font-weight:700"
                    return ""

                st.dataframe(
                    reorder_pivot.style
                        .map(_ro_style, subset=store_cols_present + ["Total Order"])
                        .format({c: "{:,.0f}" for c in store_cols_present + ["Total Order"]}),
                    width='stretch', hide_index=True)
            else:
                st.info("No reorder data to distribute for current filters.")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    cat_sum.to_excel(writer, sheet_name="Category Summary", index=False)

    full = prod_sum[["Product Name","Brand","Category"] +
                   (["Sub Category"] if "Sub Category" in prod_sum.columns else []) +
                   ["STR_Status","STR_Pct","Total_Sold","Net_Sales",
                    "Daily_Velocity","Weekly_Rate","Vel_Tier",
                    "Launch Date","Days Live","Last Sold",
                    "Total_Stock","Reorder_Velocity","Avg_Price","Est_Value"]].copy()
    full = full.rename(columns={"STR_Status":"Status","STR_Pct":"STR %",
                                "Total_Sold":"Units Sold","Total_Stock":"In Stock",
                                "Weekly_Rate":"Rate/wk","Vel_Tier":"Trend",
                                "Reorder_Velocity":f"Order ({cover_days}d)",
                                "Net_Sales":net_lbl,"Daily_Velocity":"Velocity (u/day)",
                                "Avg_Price":"Avg Price NPR",
                                "Est_Value":"Est. Value NPR"})
    full.to_excel(writer, sheet_name="Product Reorder Plan", index=False)

    if size_df is not None and "sz" in dir() and not sz.empty:
        # By Size — same detail level as By Color sheet
        sz_exp = sz.copy()
        # Add product-level columns (Trend, Launch Date, Last Sold, velocity)
        prod_merge_cols = ["Product Name","Brand","Category"] + \
            (["Sub Category"] if "Sub Category" in prod_sum.columns else []) + \
            ["STR_Status","STR_Pct","Total_Sold","Net_Sales",
             "Daily_Velocity","Weekly_Rate","Vel_Tier","Last Sold",
             "Launch Date","Days Live","Total_Stock","Reorder_Velocity","Avg_Price"]
        prod_merge_cols = [c for c in prod_merge_cols if c in prod_sum.columns]
        sz_exp = sz_exp.merge(
            prod_sum[prod_merge_cols],
            on="Product Name", how="left", suffixes=("","_prod")
        )
        sz_exp = sz_exp.rename(columns={
            "STR_Status":"Product Status","STR_Pct":"STR % (product)",
            "Total_Sold":"Total Units Sold","Net_Sales":net_lbl,
            "Daily_Velocity":"Velocity (u/day)","Weekly_Rate":"Product Rate/wk",
            "Vel_Tier":"Trend","Total_Stock":"Total Stock",
            "Reorder_Velocity":"Product Order","Avg_Price":"Avg Price",
            "STR %":"STR % (size)","Units Sold":"Units Sold (size)",
            "In Stock":"In Stock (size)",
        })
        export_sz_cols = [
            "Product Name","Brand","Category","Sub Category",
            "Size","Units Sold (size)","In Stock (size)",
            "STR % (size)","Status","Size Share %",
            "Total Units Sold","STR % (product)",net_lbl,
            "Velocity (u/day)","Product Rate/wk","Trend",
            "Launch Date","Days Live","Last Sold",
            "Total Stock","Product Order","Order (Vel)","Avg Price"
        ]
        export_sz_cols = [c for c in export_sz_cols if c in sz_exp.columns]

        # Add Size Share % if not already there
        if "Size Share %" not in sz_exp.columns and "_size_share" in sz_exp.columns:
            sz_exp["Size Share %"] = (sz_exp["_size_share"] * 100).round(1)

        sz_exp[export_sz_cols].to_excel(writer, sheet_name="By Size (Product)", index=False)

    # Size × Category breakdown
    if size_df is not None and "sz_cat_agg" in dir() and not sz_cat_agg.empty:
        sz_cat_exp = sz_cat_agg.rename(columns={"Units_Sold":"Units Sold","In_Stock":"In Stock"}) \
            if "Units_Sold" in sz_cat_agg.columns else sz_cat_agg.copy()
        sz_cat_exp.to_excel(writer, sheet_name="Size × Category", index=False)

    if color_df is not None and "cl" in dir() and not cl.empty:
        # Build color reorder sheet — matches product level data exactly
        # Get all colors (not just Fast/Super Fast) for the export
        _cl_all = color_df[color_df["Brand"].str.strip().isin(sel_brands)].copy() if sel_brands else color_df.copy()
        if sel_cats:
            _cl_all = _cl_all[_cl_all["Category"].str.strip().isin(sel_cats)]
        if sel_subs:
            _cl_all = _cl_all[_cl_all["Sub Category"].str.strip().isin(sel_subs)]
        # Only keep products that are in our filtered prod_sum
        _cl_all = _cl_all[_cl_all["Product Name"].str.strip().isin(
            set(prod_sum["Product Name"].str.strip()))]

        if not _cl_all.empty:
            # Merge product-level data (velocity, trend, order qty) onto color rows
            prod_export_cols = ["Product Name","Brand","Category","Sub Category",
                                "STR_Status","STR_Pct","Total_Sold","Net_Sales",
                                "Daily_Velocity","Weekly_Rate","Vel_Tier",
                                "Total_Stock","Reorder_Velocity","Avg_Price"]
            prod_export_cols = [c for c in prod_export_cols if c in prod_sum.columns]
            _prod_for_color  = prod_sum[prod_export_cols].copy()

            cl_exp = _cl_all.merge(_prod_for_color, on="Product Name", how="left",
                                   suffixes=("","_prod"))

            # Share of this color within the product
            _color_totals = _cl_all.groupby("Product Name")["Units Sold"].sum().rename("_color_total")
            cl_exp = cl_exp.merge(_color_totals, on="Product Name", how="left")
            cl_exp["Color Share %"] = (
                cl_exp["Units Sold"] / cl_exp["_color_total"].replace(0, float("nan")) * 100
            ).fillna(0).round(1)

            # Distribute product reorder by color share
            cl_exp[f"Order ({cover_days}d)"] = (
                cl_exp["Reorder_Velocity"].fillna(0) *
                cl_exp["Color Share %"] / 100
            ).round().astype(int)

            # Clean and rename for export
            cl_exp = cl_exp.rename(columns={
                "STR_Status":"Status","STR_Pct":"STR % (product)",
                "Total_Sold":"Total Units Sold","Net_Sales":net_lbl,
                "Daily_Velocity":"Velocity (u/day)","Weekly_Rate":"Rate/wk",
                "Vel_Tier":"Trend","Total_Stock":"Total Stock",
                "Reorder_Velocity":"Product Order","Avg_Price":"Avg Price",
                "STR %":"STR % (color)","Units Sold":"Units Sold (color)",
                "In Stock":"In Stock (color)",
            })

            # Fix: drop duplicate Status column from merge
            if "Status.1" in cl_exp.columns:
                cl_exp = cl_exp.drop(columns=["Status.1"])

            export_cols = [
                "Product Name","Brand","Category","Sub Category",
                "Color","Units Sold (color)","In Stock (color)",
                "STR % (color)","Status","Color Share %",
                "Total Units Sold","STR % (product)",net_lbl,
                "Velocity (u/day)","Rate/wk","Trend",
                "Total Stock","Product Order",f"Order ({cover_days}d)","Avg Price"
            ]
            export_cols = [c for c in export_cols if c in cl_exp.columns]
            cl_exp[export_cols].sort_values(
                ["Product Name","Units Sold (color)"], ascending=[True, False]
            ).to_excel(writer, sheet_name="By Color + Reorder", index=False)
        else:
            cl[["Product Name","Color","Units Sold","In Stock","STR %","Status"]].to_excel(
                writer, sheet_name="By Color + Reorder", index=False)

    if df_prodstore is not None and "store_totals" in dir() and not store_totals.empty:
        store_totals.rename(columns={"Units_Sold":"Units Sold","Share_%":"Share %",
                                     "Order_Vel":f"Order ({cover_days}d)"})\
            .to_excel(writer, sheet_name="By Store", index=False)

        # Category × Store — units sold pivot
        if "ps" in dir() and not ps.empty:
            try:
                ps_dl = ps.copy()
                if sel_subs and "Sub Category" in ps_dl.columns:
                    ps_dl = ps_dl[ps_dl["Sub Category"].isin(sel_subs)]
                grp_key_dl = ["Category","Sub Category","Store"] if "Sub Category" in ps_dl.columns else ["Category","Store"]
                cat_store_dl = ps_dl.groupby(grp_key_dl).agg(Units_Sold=("Units Sold","sum")).reset_index()
                stores_dl = [s for s in active_stores if s in cat_store_dl["Store"].unique()]
                pivot_cols_dl = ["Category","Sub Category"] if "Sub Category" in cat_store_dl.columns else ["Category"]
                pivot_dl = cat_store_dl.pivot_table(
                    index=pivot_cols_dl, columns="Store", values="Units_Sold",
                    aggfunc="sum", fill_value=0
                ).reset_index()
                pivot_dl.columns.name = None
                store_cols_dl = [c for c in stores_dl if c in pivot_dl.columns]
                pivot_dl["Total"] = pivot_dl[store_cols_dl].sum(axis=1)
                pivot_dl = pivot_dl.sort_values("Total", ascending=False)
                pivot_dl.to_excel(writer, sheet_name="Category × Store (Sold)", index=False)

                # Category × Store — reorder pivot
                reorder_rows_dl = []
                for _, row in pivot_dl.iterrows():
                    key = tuple(row[c] for c in pivot_cols_dl)
                    total_reorder = cat_reorder_map.get(key, 0)
                    if not isinstance(total_reorder,(int,float)): total_reorder = 0
                    row_total = row["Total"]
                    new_row = {c: row[c] for c in pivot_cols_dl}
                    for store in store_cols_dl:
                        share = row[store] / row_total if row_total > 0 else 0
                        new_row[store] = round(total_reorder * share)
                    new_row["Total Order"] = total_reorder
                    reorder_rows_dl.append(new_row)
                pd.DataFrame(reorder_rows_dl).sort_values(
                    "Total Order", ascending=False
                ).to_excel(writer, sheet_name="Category × Store (Order)", index=False)

                # Category × Store Stock
                if df_locstk is not None:
                    _lsk_dl = df_locstk.copy()
                    if sel_cats:
                        _lsk_dl = _lsk_dl[_lsk_dl["Category"].isin(sel_cats)]
                    if not _lsk_dl.empty:
                        stk_dl = _lsk_dl.pivot_table(
                            index="Category", columns="Store", values="Stock",
                            aggfunc="sum", fill_value=0
                        ).reset_index()
                        stk_dl.columns.name = None
                        _stk_cols_dl = [c for c in stores_dl if c in stk_dl.columns]
                        stk_dl["Total Stock"] = stk_dl[[c for c in stk_dl.columns
                                                          if c != "Category"]].sum(axis=1)
                        stk_dl.sort_values("Total Stock", ascending=False)\
                              .to_excel(writer, sheet_name="Category × Store (Stock)", index=False)
            except Exception as e:
                pass  # Skip if pivot fails

out.seek(0)
fname = f"reorder_{'-'.join(sel_brands) if sel_brands else 'All'}_{('-'.join(sel_cats) if sel_cats else 'AllCats').replace(' ','_')[:30]}.xlsx"
st.download_button(
    f"⬇️ Download Full Reorder Plan — {sel_brand} / {sel_cat if sel_cats else 'All'}",
    data=out, file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("Download includes: Category Summary · Product Plan · Size × Category · By Size (Product) · Color + Reorder · By Store · Category × Store (Sold) · Category × Store (Order) · Category × Store (Stock)")