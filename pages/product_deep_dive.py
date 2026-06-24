import streamlit as st
import pandas as pd
import re
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
.kpi{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;text-align:center}
.kpi-val{font-size:26px;font-weight:700;margin:0;line-height:1.1}
.kpi-lbl{font-size:11px;color:#6b7280;margin:4px 0 0}
.verdict{border-radius:10px;padding:14px 18px;margin:14px 0;font-size:14px;font-weight:500}
.verdict-reorder{background:#dcfce7;border-left:5px solid #16a34a;color:#166534}
.verdict-watch{background:#fef9c3;border-left:5px solid #d97706;color:#92400e}
.verdict-pause{background:#fee2e2;border-left:5px solid #dc2626;color:#991b1b}
.sec{font-size:13px;font-weight:700;color:#1F3864;text-transform:uppercase;
     letter-spacing:.08em;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin:18px 0 10px}
.insight{background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;
         padding:10px 14px;font-size:13px;color:#1e40af;margin-top:8px}
</style>
""", unsafe_allow_html=True)

# ── Drive IDs ─────────────────────────────────────────────────────────────────
GDRIVE_MAIN_ID      = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_VARIANT_ID   = "1LPeoGXDDd3ZAppTiuLskzY4q-71CJWfJ"
GDRIVE_STORE_ID     = "1B8_Ml_tAL59MSPrEDwKUR93ruFEC1m23"
GDRIVE_PRODSTORE_ID = "10ZvRKu4icGDw_g95PplVVdKmj_m-Zpo4"

SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL",
              "36","37","38","39","40","41","42","43","44","ONE SIZE","FREE SIZE"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def clean_name(name):
    name = str(name).strip()
    name = re.sub(r"^\[[^\]]+\]\s*", "", name)
    name = name.replace("\n", " ").replace("\t", " ")
    name = re.sub(r'^"+|"+$', "", name)
    return re.sub(r"\s+", " ", name).strip()

def strip_variant_suffix(name):
    name = clean_name(name)
    name = re.sub(r"\s*/\s*(XS|S|M|L|XL|2XL|3XL|4XL|XXL|36|37|38|39|40|41|42|43|44)$", "", name, flags=re.I)
    name = re.sub(r"/\w+$", "", name).strip()
    return name.strip()

def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

def str_color(status):
    return {"Super Fast":"#1B5E20","Fast":"#43A047","Medium":"#F9A825",
            "Slow":"#E53935","Dead":"#424242"}.get(status,"#9E9E9E")

# ── Style functions ───────────────────────────────────────────────────────────
def style_status(val):
    return {"Super Fast":"background-color:#1B5E20;color:white",
            "Fast":       "background-color:#43A047;color:white",
            "Medium":     "background-color:#F9A825;color:black",
            "Slow":       "background-color:#E53935;color:white",
            "Dead":       "background-color:#424242;color:white"}.get(val,"")

def style_reorder(val):
    if isinstance(val,(int,float)) and val > 0:
        return "background-color:#dcfce7;color:#166534;font-weight:700"
    return ""

def style_order_w(val):
    if isinstance(val,(int,float)) and val > 0:
        return "background-color:#dbeafe;color:#1e40af;font-weight:700"
    return ""

def style_doc(val):
    return {"Reorder Now":"background-color:#B71C1C;color:white",
            "Watch":"background-color:#F57F17;color:white",
            "OK":"background-color:#2E7D32;color:white"}.get(val,"")

# ── Google Drive loader ───────────────────────────────────────────────────────
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

# ── Data loaders (cache_resource = persists across all sessions) ──────────────
@st.cache_resource(show_spinner=False)
def load_products():
    skip = lambda c: c not in ("Image_Base64","Image")
    buf  = _gdrive(GDRIVE_MAIN_ID)
    df   = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl", usecols=skip)
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base+r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"), reverse=True) if Path(d).exists() else []
            if files:
                df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl", usecols=skip)
                break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    df["Product Name"] = df["Product Name"].fillna("").astype(str).apply(clean_name)
    for col in ["On Hand Qty","Total Units Sold","Revenue","Sell-Through %","Sales Price","Cost Price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Sell-Through %" in df.columns and df["Sell-Through %"].max() <= 1.0:
        df["Sell-Through %"] *= 100
    for col in ["Brand","Category","Sub Category","STR Status","ABC Class","DOC Status",
                "Color","Size","SKU / Variant","Sale Status"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    if "Template_ID" in df.columns:
        df["Template_ID"] = pd.to_numeric(df["Template_ID"], errors="coerce")
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
        df["Product Name"] = df["Product Name"].fillna("").astype(str).apply(clean_name)
        for col in ["Units Sold","In Stock","STR %"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in ["Size","Color","Brand","Category","Sub Category","Status"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str)\
                    .str.replace(r"^(Size|Color|Brand):\s*","",regex=True).str.strip()
        return df

    size_df  = _prep(size_df)
    color_df = _prep(color_df)

    # ── Size Breakdown: strip /Color suffix and aggregate by (product, size) ──
    # variant_analysis stores "Dress/Brown" and "Dress/Pink" as separate rows
    # We need to collapse them into one "Dress" row per size, summing sold+stock
    def _get_status(s):
        if s >= 95: return "Super Fast"
        if s >= 70: return "Fast"
        if s >= 30: return "Medium"
        if s > 0:   return "Slow"
        return "Dead"

    # Strip /Color from product name in size sheet
    size_df["Product Name"] = size_df["Product Name"].apply(
        lambda n: re.sub(r"/[^/]+$", "", n).strip())

    # Aggregate by (Product Name, Brand, Category, Sub Category, Size)
    grp_cols = [c for c in ["Product Name","Brand","Category","Sub Category","Size"]
                if c in size_df.columns]
    size_agg = size_df.groupby(grp_cols, as_index=False).agg(
        **{"Units Sold": ("Units Sold","sum"),
           "In Stock":   ("In Stock",  "sum")}
    )
    total = size_agg["Units Sold"] + size_agg["In Stock"]
    size_agg["STR %"]  = (size_agg["Units Sold"] / total.replace(0, float("nan")) * 100).fillna(0).round(1)
    size_agg["Status"] = size_agg["STR %"].apply(_get_status)
    size_df = size_agg

    # Color Breakdown: strip /Size suffix if present (mirror clean)
    color_df["Product Name"] = color_df["Product Name"].apply(
        lambda n: re.sub(r"/[^/]+$", "", n).strip())

    return size_df, color_df

@st.cache_resource(show_spinner=False)
def load_store():
    buf = _gdrive(GDRIVE_STORE_ID)
    df  = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="🏆 Top Products by Store", engine="openpyxl")
        except: pass
    if df is None:
        files = sorted(Path(r"C:\Users\Legion\Desktop\odoo_export\exports")
                       .glob("store_analysis*.xlsx"), reverse=True)
        if files:
            try: df = pd.read_excel(files[0], sheet_name="🏆 Top Products by Store", engine="openpyxl")
            except: pass
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    df["Product"] = df["Product"].fillna("").astype(str).apply(clean_name)
    for col in ["Revenue (NPR)","Units Sold"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Store" in df.columns:
        df["Store"] = df["Store"].replace("", pd.NA).ffill()
    return df

@st.cache_resource(show_spinner=False)
def load_product_store():
    if not GDRIVE_PRODSTORE_ID: return None
    buf = _gdrive(GDRIVE_PRODSTORE_ID)
    df  = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Product × Store", engine="openpyxl")
        except: pass
    if df is None:
        files = sorted(Path(r"C:\Users\Legion\Desktop\odoo_export\exports")
                       .glob("product_store_sales*.xlsx"), reverse=True)
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
            df[col] = df[col].fillna("").astype(str).str.strip().apply(clean_name)
    return df

@st.cache_resource(show_spinner=False)
def build_templates(_df):
    df = _df
    if "Template_ID" in df.columns and df["Template_ID"].notna().any():
        def canon(names):
            cleaned = [strip_variant_suffix(n) for n in names]
            from collections import Counter
            return Counter(cleaned).most_common(1)[0][0]
        result = df.groupby("Template_ID").agg(
            Product_Name =("Product Name",   lambda x: canon(x)),
            Brand        =("Brand",          lambda x: x.mode()[0] if len(x) else ""),
            Category     =("Category",       lambda x: x.mode()[0] if len(x) else ""),
            Sub_Category =("Sub Category",   lambda x: x.mode()[0] if len(x) else ""),
            Total_Sold   =("Total Units Sold","sum"),
            Total_Stock  =("On Hand Qty",    "sum"),
            Total_Revenue=("Revenue",        "sum"),
            Avg_Price    =("Sales Price",    "mean"),
            STR_Pct      =("Sell-Through %", "mean"),
            STR_Status   =("STR Status",     lambda x: x.mode()[0] if len(x) else ""),
            ABC_Class    =("ABC Class",      lambda x: x.mode()[0] if len(x) else ""),
            DOC_Status   =("DOC Status",     lambda x: x.mode()[0] if len(x) else ""),
            Variants     =("Variant_ID",     "count"),
            Create_Date  =("Create Date",    "min"),
        ).reset_index()
    else:
        df = df.copy()
        df["_base"] = df["Product Name"].apply(strip_variant_suffix)
        result = df.groupby("_base").agg(
            Product_Name =("_base",          "first"),
            Brand        =("Brand",          lambda x: x.mode()[0] if len(x) else ""),
            Category     =("Category",       lambda x: x.mode()[0] if len(x) else ""),
            Sub_Category =("Sub Category",   lambda x: x.mode()[0] if len(x) else ""),
            Total_Sold   =("Total Units Sold","sum"),
            Total_Stock  =("On Hand Qty",    "sum"),
            Total_Revenue=("Revenue",        "sum"),
            Avg_Price    =("Sales Price",    "mean"),
            STR_Pct      =("Sell-Through %", "mean"),
            STR_Status   =("STR Status",     lambda x: x.mode()[0] if len(x) else ""),
            ABC_Class    =("ABC Class",      lambda x: x.mode()[0] if len(x) else ""),
            DOC_Status   =("DOC Status",     lambda x: x.mode()[0] if len(x) else ""),
            Variants     =("Product Name",   "count"),
            Create_Date  =("Create Date",    "min") if "Create Date" in df.columns else ("Product Name","first"),
        ).reset_index(drop=True)
    result = result[result["Product_Name"].str.len() > 5]
    result = result[result["Product_Name"].str.contains(r" ", na=False)]
    result = result[result["Product_Name"].str.contains(r"[a-zA-Z]{3}", na=False)]
    result = result[~result["Product_Name"].str.match(r"^[\d\-\.\s/]+$", na=False)]
    return result

# ── Load data ─────────────────────────────────────────────────────────────────
_c = st.empty()
with _c.container():
    with st.spinner("Loading product catalog (first load ~15s, then instant)…"):
        df_raw = load_products()
with _c.container():
    with st.spinner("Loading variant analysis…"):
        size_df, color_df = load_variants()
with _c.container():
    with st.spinner("Loading store data…"):
        df_store     = load_store()
        df_prodstore = load_product_store()
_c.empty()

if df_raw is None:
    st.error("Could not load product data."); st.stop()

with st.spinner("Building product catalog…"):
    df_templates = build_templates(df_raw)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Product Deep Dive")
    st.markdown("---")
    brands = sorted([b for b in df_templates["Brand"].unique()
                     if b and b not in ("","nan","True","False")])
    sel_brand = st.selectbox("Brand", brands)
    brand_df  = df_templates[df_templates["Brand"] == sel_brand]
    cats      = ["All"] + sorted([c for c in brand_df["Category"].unique()
                                   if c and c not in ("","nan")])
    sel_cat   = st.selectbox("Category", cats)
    filtered_df = brand_df if sel_cat == "All" else brand_df[brand_df["Category"] == sel_cat]
    products    = sorted(filtered_df["Product_Name"].unique())
    search      = st.text_input("Search product", placeholder="Type to filter…")
    if search.strip():
        products = [p for p in products if search.lower() in p.lower()]
    if not products:
        st.warning("No products found."); st.stop()
    sel_product  = st.selectbox("Select Product", products)
    st.markdown("---")
    target_weeks = st.slider("Target weeks of stock", 2, 12, 4,
                              help="Weeks of cover to maintain per size.")
    st.markdown("---")
    st.caption(f"{len(products):,} products · {len(filtered_df):,} in category")
    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_resource.clear(); st.rerun()

# ── Selected product data ─────────────────────────────────────────────────────
prod_row = filtered_df[filtered_df["Product_Name"] == sel_product]
if prod_row.empty:
    st.warning(f"No data for {sel_product}"); st.stop()
prod = prod_row.iloc[0]

total_sold    = float(prod["Total_Sold"])
total_stock   = float(prod["Total_Stock"])
total_rev     = float(prod["Total_Revenue"])
avg_price     = float(prod["Avg_Price"])
str_pct       = float(prod["STR_Pct"])
str_status    = str(prod["STR_Status"])
category      = str(prod["Category"])
sub_cat       = str(prod["Sub_Category"])
num_variants  = int(prod["Variants"])
create_date   = prod.get("Create_Date", pd.NaT)

# ── Variant matching ──────────────────────────────────────────────────────────
def find_variant_rows(df, product_name):
    if df is None or df.empty: return pd.DataFrame()
    rows = df[df["Product Name"] == product_name]
    if not rows.empty: return rows
    pn_lower = product_name.lower()
    pn_words = set(pn_lower.split())
    mask = df["Product Name"].str.lower().apply(
        lambda n: n == pn_lower
        or (len(pn_lower) > 6 and (pn_lower in n or n in pn_lower))
        or len(set(n.split()) & pn_words) >= max(2, len(pn_words) - 1)
    )
    return df[mask].copy()

p_sizes  = find_variant_rows(size_df,  sel_product)
p_colors = find_variant_rows(color_df, sel_product)

if not p_sizes.empty and "Size" in p_sizes.columns:
    p_sizes["_sk"] = p_sizes["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
    p_sizes = p_sizes.sort_values("_sk").drop(columns=["_sk"])

# ── Store matching ────────────────────────────────────────────────────────────
def find_store_rows_full(df_ps, product_name):
    if df_ps is None or "Product Name" not in df_ps.columns: return pd.DataFrame()
    rows = df_ps[df_ps["Product Name"] == product_name]
    if rows.empty:
        pn_words = set(product_name.lower().split())
        mask = df_ps["Product Name"].str.lower().apply(
            lambda n: set(n.split()).issuperset(pn_words) and len(pn_words) >= 3)
        rows = df_ps[mask]
    if rows.empty: return pd.DataFrame()
    agg = rows.groupby("Store").agg(
        **{"Units Sold":    ("Units Sold",    "sum"),
           "Revenue (NPR)": ("Revenue (NPR)", "sum")}
    ).reset_index()
    return agg[agg["Units Sold"] > 0].sort_values("Units Sold", ascending=False)

def find_store_rows_top20(df_st, product_name):
    if df_st is None or "Product" not in df_st.columns: return pd.DataFrame()
    rows = df_st[df_st["Product"] == product_name]
    if rows.empty:
        pn_lower = product_name.lower()
        mask = df_st["Product"].str.lower().apply(
            lambda n: pn_lower in n or n in pn_lower
            or len(set(n.split()) & set(pn_lower.split())) >= max(2, len(pn_lower.split()) - 1))
        rows = df_st[mask]
    return rows[rows["Units Sold"] > 0].copy() if not rows.empty else pd.DataFrame()

p_stores = find_store_rows_full(df_prodstore, sel_product)
_store_source = "full"
if p_stores.empty:
    p_stores      = find_store_rows_top20(df_store, sel_product)
    _store_source = "top20"

# ── Reorder calculation ───────────────────────────────────────────────────────
# Compute product-level weekly rate from Create Date
today_ts = pd.Timestamp.today()
if pd.notna(create_date):
    try:
        weeks_live = max(4, (today_ts - pd.to_datetime(create_date)).days / 7)
    except:
        weeks_live = 52
else:
    weeks_live = 52
prod_weekly_rate = total_sold / weeks_live if weeks_live > 0 else 0

# Size-level: both methods
if not p_sizes.empty and "Units Sold" in p_sizes.columns:
    _tsz = p_sizes["Units Sold"].sum()

    def _size_rate(s):
        return prod_weekly_rate * (s / _tsz) if _tsz > 0 and prod_weekly_rate > 0 else 0

    p_sizes["Weekly Rate"]     = p_sizes["Units Sold"].apply(_size_rate).round(2)
    p_sizes["Weeks Cover"]     = p_sizes.apply(
        lambda r: round(r["In Stock"] / r["Weekly Rate"], 1)
                  if r["Weekly Rate"] > 0 else (999 if r["In Stock"] > 0 else 0), axis=1)
    p_sizes["Weeks Cover Fmt"] = p_sizes["Weeks Cover"].apply(
        lambda x: "—" if x >= 99 else f"{x:.1f} wks")
    p_sizes["Suggest (STR)"]   = p_sizes.apply(
        lambda r: max(0, round(r["Units Sold"] - r["In Stock"]))
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)
    p_sizes["Suggest (Weeks)"] = p_sizes.apply(
        lambda r: max(0, round(r["Weekly Rate"] * target_weeks - r["In Stock"]))
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)
    total_suggest_str   = int(p_sizes["Suggest (STR)"].sum())
    total_suggest_weeks = int(p_sizes["Suggest (Weeks)"].sum())
else:
    total_suggest_weeks = max(0, round(prod_weekly_rate * target_weeks - total_stock))
    total_suggest_str   = max(0, round(total_sold - total_stock))

total_suggest = total_suggest_weeks

# Color-level
if not p_colors.empty and "Units Sold" in p_colors.columns:
    _tcl = p_colors["Units Sold"].sum()

    def _color_rate(s):
        return prod_weekly_rate * (s / _tcl) if _tcl > 0 and prod_weekly_rate > 0 else 0

    p_colors["Weekly Rate"]     = p_colors["Units Sold"].apply(_color_rate).round(2)
    p_colors["Suggest (STR)"]   = p_colors.apply(
        lambda r: max(0, round(r["Units Sold"] - r["In Stock"]))
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)
    p_colors["Suggest (Weeks)"] = p_colors.apply(
        lambda r: max(0, round(r["Weekly Rate"] * target_weeks - r["In Stock"]))
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)

# ── Verdict ───────────────────────────────────────────────────────────────────
if str_status in ("Super Fast","Fast") and total_suggest_weeks > 0:
    vc, vi = "verdict-reorder", "✅"
    vt = (f"<strong>Reorder recommended — {total_suggest_weeks} units (weeks-based) "
          f"/ {total_suggest_str} units (STR-based).</strong> "
          f"Fast seller (STR {str_pct:.0f}%). See size & store tables below.")
elif str_status in ("Super Fast","Fast"):
    vc, vi = "verdict-watch", "📦"
    vt = f"Stock OK for now. Strong seller (STR {str_pct:.0f}%) — watch closely."
elif str_status == "Medium":
    vc, vi = "verdict-watch", "⚠️"
    vt = f"Medium performer (STR {str_pct:.0f}%). Reorder only if specific sizes are running out."
else:
    vc, vi = "verdict-pause", "🛑"
    vt = f"Slow/Dead seller (STR {str_pct:.0f}%). Do <strong>not</strong> reorder — clear existing stock first."

# ── Page header ───────────────────────────────────────────────────────────────
st.title("🔍 Product Deep Dive")
st.markdown(
    f"**{sel_product}** &nbsp;·&nbsp; {category}"
    + (f" › {sub_cat}" if sub_cat else "")
    + f" &nbsp;·&nbsp; {sel_brand} &nbsp;·&nbsp; {num_variants} variants",
    unsafe_allow_html=True)
st.markdown(f'<div class="verdict {vc}">{vi} {vt}</div>', unsafe_allow_html=True)

# ── KPI strip ─────────────────────────────────────────────────────────────────
c1,c2,c3,c4,c5,c6 = st.columns(6)
for col, val, lbl, clr in [
    (c1, f"{int(total_sold):,}",   "Total Units Sold",        "#1d4ed8"),
    (c2, f"{int(total_stock):,}",  "In Stock Now",            "#374151"),
    (c3, f"{str_pct:.0f}%",        "Sell-Through Rate",       str_color(str_status)),
    (c4, fmt_npr(total_rev),       "Total Revenue",           "#374151"),
    (c5, fmt_npr(avg_price),       "Avg Selling Price",       "#374151"),
    (c6, f"{total_suggest_weeks:,} u / {total_suggest_str:,} u",
         f"Reorder: Weeks/STR ({target_weeks}wk)",
         "#16a34a" if total_suggest_weeks > 0 else "#6b7280"),
]:
    with col:
        st.markdown(f'<div class="kpi"><p class="kpi-val" style="color:{clr}">{val}</p>'
                    f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Size performance ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">📏 Size Performance</div>', unsafe_allow_html=True)

if not p_sizes.empty and "Units Sold" in p_sizes.columns:
    _av  = ["Size","Weekly Rate","In Stock","Weeks Cover Fmt","STR %","Status",
            "Suggest (Weeks)","Suggest (STR)","Units Sold"]
    disp = p_sizes[[c for c in _av if c in p_sizes.columns]].copy()
    disp["STR %"] = disp["STR %"].round(1)
    disp = disp.rename(columns={"Weekly Rate":"Rate/wk","Weeks Cover Fmt":"Wks Cover",
                                 "Suggest (Weeks)":"Order (Wk)","Suggest (STR)":"Order (STR)"})

    def _sw(v):
        try:
            f = float(str(v).replace(" wks",""))
            if f <= 1:            return "color:#dc2626;font-weight:700"
            if f <= target_weeks: return "color:#d97706;font-weight:600"
            return "color:#16a34a"
        except: return ""

    _fd = {"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}"}
    if "Rate/wk"    in disp.columns: _fd["Rate/wk"]    = "{:.2f}"
    if "Order (Wk)" in disp.columns: _fd["Order (Wk)"] = "{:,.0f}"
    if "Order (STR)"in disp.columns: _fd["Order (STR)"]= "{:,.0f}"

    _st = disp.style.map(style_status, subset=["Status"])
    if "Wks Cover"  in disp.columns: _st = _st.map(_sw,           subset=["Wks Cover"])
    if "Order (Wk)" in disp.columns: _st = _st.map(style_order_w, subset=["Order (Wk)"])
    if "Order (STR)"in disp.columns: _st = _st.map(style_reorder, subset=["Order (STR)"])
    st.dataframe(_st.format(_fd), width='stretch', hide_index=True)
    st.caption(f"🔵 Order (Wk) = {target_weeks}-week buffer at current rate  ·  "
               f"🟢 Order (STR) = restore original stock level")

    fast = p_sizes[p_sizes["Status"].isin(["Super Fast","Fast"])]["Size"].tolist()
    dead = p_sizes[p_sizes["Status"].isin(["Dead","Slow"])]["Size"].tolist()
    parts = []
    if fast: parts.append(f"🟢 <strong>Fast sizes: {', '.join(fast)}</strong> — reorder these")
    if dead: parts.append(f"🔴 <strong>Stuck sizes: {', '.join(dead)}</strong> — hold off")
    if total_suggest_weeks > 0:
        parts.append(f"📦 <strong>Total suggested: {total_suggest_weeks} units (Wk) / {total_suggest_str} units (STR)</strong>")
    if parts:
        st.markdown(f'<div class="insight">{"  ·  ".join(parts)}</div>', unsafe_allow_html=True)
else:
    st.info(
        f"No size data for **{sel_product}** in variant_analysis. "
        + ("Product may have no size variants in Odoo, or each size is a separate Odoo template — "
           "check Full SKU Breakdown below." if size_df is not None
           else "Run `python variant_export.py` first.")
    )

# ── Color performance ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">🎨 Color Performance</div>', unsafe_allow_html=True)

if not p_colors.empty and "Units Sold" in p_colors.columns:
    p_colors = p_colors.sort_values("Units Sold", ascending=False)
    _cav  = ["Color","Weekly Rate","Units Sold","In Stock","STR %","Status",
             "Suggest (Weeks)","Suggest (STR)"]
    disp_c = p_colors[[c for c in _cav if c in p_colors.columns]].copy()
    disp_c["STR %"] = disp_c["STR %"].round(1)
    disp_c = disp_c.rename(columns={"Weekly Rate":"Rate/wk",
                                     "Suggest (Weeks)":"Order (Wk)",
                                     "Suggest (STR)":"Order (STR)"})
    _fc = {"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}"}
    if "Rate/wk"    in disp_c.columns: _fc["Rate/wk"]    = "{:.2f}"
    if "Order (Wk)" in disp_c.columns: _fc["Order (Wk)"] = "{:,.0f}"
    if "Order (STR)"in disp_c.columns: _fc["Order (STR)"]= "{:,.0f}"
    _sc = disp_c.style.map(style_status, subset=["Status"])
    if "Order (Wk)" in disp_c.columns: _sc = _sc.map(style_order_w, subset=["Order (Wk)"])
    if "Order (STR)"in disp_c.columns: _sc = _sc.map(style_reorder, subset=["Order (STR)"])
    st.dataframe(_sc.format(_fc), width='stretch', hide_index=True)

    fast_c = p_colors[p_colors["Status"].isin(["Super Fast","Fast"])]["Color"].tolist()
    dead_c = p_colors[p_colors["Status"].isin(["Dead","Slow"])]["Color"].tolist()
    parts_c = []
    if fast_c: parts_c.append(f"🟢 <strong>Top colors: {', '.join(fast_c[:4])}</strong>")
    if dead_c: parts_c.append(f"🔴 <strong>Not moving: {', '.join(dead_c[:4])}</strong>")
    if parts_c:
        st.markdown(f'<div class="insight">{"  ·  ".join(parts_c)}</div>', unsafe_allow_html=True)
else:
    st.info(
        f"No color variants for **{sel_product}** — "
        + ("product may be sizes-only." if color_df is not None
           else "Run `python variant_export.py` first.")
    )

# ── Store performance ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">🏪 Store Performance</div>', unsafe_allow_html=True)

if not p_stores.empty:
    col_s, col_b = st.columns([2,3])
    with col_s:
        p_stores_d = p_stores[["Store","Units Sold","Revenue (NPR)"]].copy()
        p_stores_d["Revenue (NPR)"] = p_stores_d["Revenue (NPR)"].apply(fmt_npr)
        st.dataframe(p_stores_d, width='stretch', hide_index=True)
    with col_b:
        max_u = p_stores["Units Sold"].max() or 1
        for _, row in p_stores.sort_values("Units Sold", ascending=False).iterrows():
            pct = row["Units Sold"] / max_u * 100
            st.markdown(
                f'<div style="margin-bottom:5px">'
                f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">'
                f'<span><strong>{row["Store"]}</strong></span>'
                f'<span style="color:#6b7280">{int(row["Units Sold"]):,} units</span></div>'
                f'<div style="background:#e2e8f0;border-radius:4px;height:8px">'
                f'<div style="background:#1d4ed8;width:{pct:.0f}%;height:8px;border-radius:4px"></div>'
                f'</div></div>', unsafe_allow_html=True)
    src_label = "product_store_sales.xlsx (all-time)" if _store_source == "full" else "store_analysis top-20"
    st.caption(f"Source: {src_label}")
else:
    if df_prodstore is not None:
        st.info(f"**{sel_product}** has no POS sales recorded in product_store_sales.xlsx.")
    elif df_store is not None:
        st.info(f"**{sel_product}** doesn't rank in the top 20 at any store. "
                "Run `fetch_product_store_sales.py` for full store breakdown.")
    else:
        st.info("Store data not available — check Google Drive.")

# ── Store reorder distribution ────────────────────────────────────────────────
if not p_stores.empty and total_suggest_weeks > 0:
    st.markdown('<div class="sec">📦 Reorder Distribution by Store</div>', unsafe_allow_html=True)
    total_su = p_stores["Units Sold"].sum()
    if total_su > 0:
        dist = p_stores.copy()
        dist["Share %"]     = (dist["Units Sold"] / total_su * 100).round(1)
        dist["Order (Wk)"]  = (dist["Units Sold"] / total_su * total_suggest_weeks).round().astype(int)
        dist["Order (STR)"] = (dist["Units Sold"] / total_su * total_suggest_str).round().astype(int)
        dist = dist[dist["Order (Wk)"] > 0].sort_values("Order (Wk)", ascending=False)
        dist_d = dist[["Store","Units Sold","Share %","Order (Wk)","Order (STR)"]]\
            .rename(columns={"Units Sold":"Sold"})

        def _sd(v):
            return "background-color:#dbeafe;color:#1e40af;font-weight:700" \
                   if isinstance(v,(int,float)) and v > 0 else ""

        st.dataframe(
            dist_d.style.map(_sd, subset=["Order (Wk)"])
                        .map(style_reorder, subset=["Order (STR)"])
                        .format({"Sold":"{:,.0f}","Share %":"{:.1f}%",
                                 "Order (Wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}),
            width='stretch', hide_index=True)
        st.caption(
            f"Total: {dist['Order (Wk)'].sum()} units (Wk) / "
            f"{dist['Order (STR)'].sum()} units (STR).  "
            f"Based on each store's share of recorded sales.")

# ── Full SKU breakdown ────────────────────────────────────────────────────────
st.markdown('<div class="sec">📋 Full SKU Breakdown — every size × color from Odoo</div>',
            unsafe_allow_html=True)

if "Template_ID" in df_raw.columns and "Template_ID" in prod_row.columns:
    tmpl_id = prod_row.iloc[0]["Template_ID"]
    sku_rows = df_raw[df_raw["Template_ID"] == tmpl_id].copy() \
               if pd.notna(tmpl_id) else \
               df_raw[df_raw["Product Name"].apply(strip_variant_suffix) == sel_product].copy()
else:
    sku_rows = df_raw[df_raw["Product Name"].apply(strip_variant_suffix) == sel_product].copy()

if not sku_rows.empty:
    sku_cols = [c for c in ["Color","Size","SKU / Variant","On Hand Qty","Total Units Sold",
                             "Sell-Through %","STR Status","Sales Price","DOC Status"]
                if c in sku_rows.columns]
    sku_d = sku_rows[sku_cols].copy()
    if "Size" in sku_d.columns:
        sku_d["_sk"] = sku_d["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        sku_d = sku_d.sort_values(["Color","_sk"]).drop(columns=["_sk"])
    if "Sell-Through %" in sku_d.columns: sku_d["Sell-Through %"] = sku_d["Sell-Through %"].round(1)
    _fmt = {}
    if "Sell-Through %" in sku_d.columns:   _fmt["Sell-Through %"]   = "{:.1f}%"
    if "On Hand Qty"    in sku_d.columns:   _fmt["On Hand Qty"]      = "{:,.0f}"
    if "Total Units Sold" in sku_d.columns: _fmt["Total Units Sold"] = "{:,.0f}"
    if "Sales Price"    in sku_d.columns:   _fmt["Sales Price"]      = "NPR {:,.0f}"
    _sty = sku_d.style.format(_fmt)
    if "STR Status" in sku_d.columns: _sty = _sty.map(style_status, subset=["STR Status"])
    if "DOC Status" in sku_d.columns: _sty = _sty.map(style_doc,    subset=["DOC Status"])
    st.dataframe(_sty, width='stretch', hide_index=True)
    st.caption(f"{len(sku_rows)} variants total")

# ── Category comparison ───────────────────────────────────────────────────────
st.markdown('<div class="sec">📊 How this product ranks in its category</div>',
            unsafe_allow_html=True)

cat_peers = df_templates[
    (df_templates["Brand"]        == sel_brand) &
    (df_templates["Category"]     == category)  &
    (df_templates["Product_Name"] != sel_product)
].sort_values("Total_Sold", ascending=False).head(15)

if not cat_peers.empty:
    current  = pd.DataFrame([{"Product_Name":f"➡️ {sel_product}",
                               "Total_Sold":total_sold,"Total_Stock":total_stock,
                               "Total_Revenue":total_rev,"STR_Pct":str_pct,"STR_Status":str_status}])
    combined = pd.concat([current, cat_peers], ignore_index=True)
    combined["Total_Revenue"] = combined["Total_Revenue"].apply(fmt_npr)
    combined["STR_Pct"]       = combined["STR_Pct"].round(1)
    combined = combined.rename(columns={"Product_Name":"Product","Total_Sold":"Units Sold",
                                        "Total_Stock":"In Stock","Total_Revenue":"Revenue",
                                        "STR_Pct":"STR %","STR_Status":"Status"})

    def _hi(row):
        return ["background-color:#eff6ff;font-weight:600"]*len(row) \
               if str(row["Product"]).startswith("➡️") else [""]*len(row)

    st.dataframe(
        combined.style.apply(_hi, axis=1)
                      .map(style_status, subset=["Status"])
                      .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}"}),
        width='stretch', hide_index=True)
    rank = (cat_peers["Total_Sold"] > total_sold).sum() + 1
    st.caption(f"Ranked #{rank} by units sold within {category} ({sel_brand})")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    pd.DataFrame([{"Product":sel_product,"Category":category,"Sub Category":sub_cat,
                   "Brand":sel_brand,"Total Sold":total_sold,"In Stock":total_stock,
                   "STR %":round(str_pct,1),"STR Status":str_status,
                   "Revenue":total_rev,"Avg Price":round(avg_price),
                   "Suggest (Weeks)":total_suggest_weeks,"Suggest (STR)":total_suggest_str}])\
        .to_excel(writer, sheet_name="Summary", index=False)
    if not p_sizes.empty:
        s_cols = [c for c in ["Size","Units Sold","In Stock","STR %","Status",
                               "Suggest (Weeks)","Suggest (STR)","Weekly Rate"] if c in p_sizes.columns]
        p_sizes[s_cols].to_excel(writer, sheet_name="By Size", index=False)
    if not p_colors.empty:
        c_cols = [c for c in ["Color","Units Sold","In Stock","STR %","Status",
                               "Suggest (Weeks)","Suggest (STR)","Weekly Rate"] if c in p_colors.columns]
        p_colors[c_cols].to_excel(writer, sheet_name="By Color", index=False)
    if not p_stores.empty:
        p_stores[["Store","Units Sold","Revenue (NPR)"]].to_excel(writer, sheet_name="By Store", index=False)
    if not sku_rows.empty:
        sku_rows[sku_cols].to_excel(writer, sheet_name="All SKUs", index=False)
out.seek(0)
st.download_button(
    f"⬇️ Download {sel_product[:40]} — full report", data=out,
    file_name=f"deep_dive_{re.sub(r'[^a-zA-Z0-9]','_',sel_product[:40])}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("💡 Need bulk analysis?"):
    st.markdown("""
- **Buying Brief** — category-level recommendations (Increase / Watch / Reduce)
- **Reorder Plan** — store-level reorder quantities across all categories
- **Variant Dashboard** — filter by Status = Super Fast to see all fast-moving sizes

This page is best for: *"Should we reorder this specific product?"*
For bulk supplier orders: use **Reorder Plan → Overall Summary → Download**.
    """)