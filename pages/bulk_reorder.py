import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pathlib import Path

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
.badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:11px;font-weight:700}
</style>
""", unsafe_allow_html=True)

# ── Google Drive IDs ──────────────────────────────────────────────────────────
GDRIVE_MAIN_ID    = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_VARIANT_ID = "1LPeoGXDDd3ZAppTiuLskzY4q-71CJWfJ"

SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL","5XL","ONE SIZE","FREE SIZE",
              "36","37","38","39","40","41","42","43","44"]

STR_COLORS = {
    "Super Fast": "#1B5E20", "Fast": "#43A047",
    "Medium": "#F9A825", "Slow": "#E53935", "Dead": "#424242",
}

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
    # Split category if needed
    SKIP = {"All","Saleable","PoS",""}
    if "Category" in df.columns and df["Category"].str.contains("/", na=False).any():
        def split_cat(raw):
            parts = [p.strip() for p in str(raw).split("/") if p.strip() not in SKIP]
            if not parts: return "", ""
            if len(parts) == 1: return parts[0], ""
            return parts[0], parts[1]
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
        df["Product Name"] = df["Product Name"].fillna("").astype(str)
        for col in ["Units Sold","In Stock","STR %"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in ["Brand","Category","Sub Category","Size","Color","Status","Product Name"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.replace(
                    r"^(Size|Color|Brand):\s*","",regex=True).str.strip()
        return df

    size_df  = _prep(size_df)
    color_df = _prep(color_df)

    # ── Strip /Color suffix from size sheet product names and aggregate ──────
    SIZE_SET = set(s.upper() for s in SIZE_ORDER)

    def parse_name_color(name):
        name = re.sub(r'^\[[^\]]+\]\s*', '', str(name)).strip()
        name = name.replace('\n',' ').replace('\t',' ')
        name = re.sub(r'\s+',' ', name).strip()
        if '/' in name:
            parts = name.rsplit('/', 1)
            suffix = parts[1].strip()
            if suffix.upper() in SIZE_SET:
                return parts[0].strip(), None
            else:
                return parts[0].strip(), suffix.strip()
        return name, None

    parsed = size_df["Product Name"].apply(parse_name_color)
    size_df["_base"]  = parsed.apply(lambda x: x[0])
    size_df["_color"] = parsed.apply(lambda x: x[1])

    # Aggregate sizes by (base_name, brand, category, sub_category, size)
    grp = [c for c in ["_base","Brand","Category","Sub Category","Size"] if c in size_df.columns]
    size_agg = size_df.groupby(grp, as_index=False).agg(
        **{"Units Sold":("Units Sold","sum"), "In Stock":("In Stock","sum")}
    ).rename(columns={"_base":"Product Name"})
    total = size_agg["Units Sold"] + size_agg["In Stock"]
    size_agg["STR %"]  = (size_agg["Units Sold"] / total.replace(0,float("nan")) * 100).fillna(0).round(1)

    def get_status(p):
        if p >= 95: return "Super Fast"
        if p >= 70: return "Fast"
        if p >= 30: return "Medium"
        if p > 0:   return "Slow"
        return "Dead"
    size_agg["Status"] = size_agg["STR %"].apply(get_status)
    size_df = size_agg

    # ── Build synthetic color rows for products missing from color sheet ──────
    color_df["Product Name"] = color_df["Product Name"].apply(
        lambda n: re.sub(r"/[^/]+$", "", n).strip())
    existing_colors = set(color_df["Product Name"].str.strip().str.lower())

    syn_rows = []
    for _, row in size_df[size_df.get("_color", pd.Series()).notna()].iterrows() if "_color" in size_df.columns else []:
        pass
    # Re-extract from original size_df before aggregation
    orig = _prep(pd.read_excel(_gdrive(GDRIVE_VARIANT_ID) or
                               Path(r"C:\Users\Legion\Desktop\odoo_export\variant_analysis.xlsx"),
                               sheet_name="Size Breakdown", engine="openpyxl"))
    orig["Product Name"] = orig["Product Name"].fillna("").astype(str)
    parsed2 = orig["Product Name"].apply(parse_name_color)
    orig["_base"]  = parsed2.apply(lambda x: x[0])
    orig["_color"] = parsed2.apply(lambda x: x[1])

    syn_src = orig[orig["_color"].notna() & ~orig["_base"].str.lower().isin(existing_colors)]
    if len(syn_src) > 0:
        grp_c = [c for c in ["_base","Brand","Category","Sub Category","_color"] if c in syn_src.columns]
        syn_agg = syn_src.groupby(grp_c, as_index=False).agg(
            **{"Units Sold":("Units Sold","sum"), "In Stock":("In Stock","sum")}
        ).rename(columns={"_base":"Product Name","_color":"Color"})
        total_c = syn_agg["Units Sold"] + syn_agg["In Stock"]
        syn_agg["STR %"]  = (syn_agg["Units Sold"] / total_c.replace(0,float("nan")) * 100).fillna(0).round(1)
        syn_agg["Status"] = syn_agg["STR %"].apply(get_status)
        color_df = pd.concat([color_df, syn_agg], ignore_index=True)

    return size_df, color_df


# ── Load ──────────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod     = load_products()
    size_df, color_df = load_variants()

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
                              if c and c not in ("","nan")])
    sel_cat = st.selectbox("Category", cats)

    st.markdown("---")
    st.markdown("**Reorder Filter**")
    min_str = st.slider("Min STR % to include", 0, 100, 50,
        help="Only show products at or above this sell-through rate")
    target_weeks = st.slider("Target weeks of cover", 2, 12, 4,
        help="How many weeks of stock buffer you want")
    show_zero = st.checkbox("Include products needing 0 units", value=False)

    st.markdown("---")
    data_source = st.radio("Data source for reorder", ["Variant-level (size/color breakdown)", "Product-level (totals only)"],
        help="Variant-level shows per-size and per-color breakdown. Product-level uses total stock/sales.")

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_resource.clear(); st.rerun()

# ── Filter ────────────────────────────────────────────────────────────────────
bdf = df_prod[df_prod["Brand"] == sel_brand].copy()
if sel_cat != "All":
    bdf = bdf[bdf["Category"] == sel_cat]

# ── Build product-level summary ───────────────────────────────────────────────
grp_cols = ["Product Name","Category","Sub Category"] if "Sub Category" in bdf.columns else ["Product Name","Category"]
prod_sum = bdf.groupby(grp_cols).agg(
    Total_Sold  = ("Total Units Sold","sum"),
    Total_Stock = ("On Hand Qty",     "sum"),
    Avg_Price   = ("Sales Price",     "mean"),
    Total_Rev   = ("Revenue",         "sum"),
).reset_index()

prod_sum["STR_Pct"] = (prod_sum["Total_Sold"] /
    (prod_sum["Total_Sold"] + prod_sum["Total_Stock"]).replace(0, float("nan")) * 100
).fillna(0).round(1)

def str_status(p):
    if p >= 95: return "Super Fast"
    if p >= 70: return "Fast"
    if p >= 30: return "Medium"
    if p > 0:   return "Slow"
    return "Dead"

prod_sum["STR_Status"] = prod_sum["STR_Pct"].apply(str_status)

# Weekly rate: use Create Date if available, else 52 weeks default
today = pd.Timestamp.today()
if "Create Date" in bdf.columns:
    dates = bdf.groupby("Product Name")["Create Date"].min().reset_index()
    dates["Create Date"] = pd.to_datetime(dates["Create Date"], errors="coerce")
    prod_sum = prod_sum.merge(dates, on="Product Name", how="left")
    prod_sum["weeks_live"] = (today - prod_sum["Create Date"]).dt.days / 7
    prod_sum["weeks_live"] = prod_sum["weeks_live"].fillna(52).clip(lower=4)
else:
    prod_sum["weeks_live"] = 52

prod_sum["Weekly_Rate"]  = (prod_sum["Total_Sold"] / prod_sum["weeks_live"]).round(2)
prod_sum["Target_Stock"] = (prod_sum["Weekly_Rate"] * target_weeks).round(0)
prod_sum["Reorder_Wk"]   = (prod_sum["Target_Stock"] - prod_sum["Total_Stock"]).clip(lower=0).round().astype(int)
prod_sum["Reorder_STR"]  = (prod_sum["Total_Sold"] - prod_sum["Total_Stock"]).clip(lower=0).round().astype(int)
prod_sum["Est_Value"]    = prod_sum["Reorder_Wk"] * prod_sum["Avg_Price"]

# Apply STR filter
prod_sum = prod_sum[prod_sum["STR_Pct"] >= min_str]
if not show_zero:
    prod_sum = prod_sum[(prod_sum["Reorder_Wk"] > 0) | (prod_sum["Reorder_STR"] > 0)]

prod_sum = prod_sum.sort_values("Total_Sold", ascending=False)

# ── KPIs ──────────────────────────────────────────────────────────────────────
st.title("🛒 Bulk Reorder Tool")
st.markdown(f"**{sel_brand}** · {sel_cat} · STR ≥ {min_str}% · {target_weeks}-week target")

total_units_wk  = int(prod_sum["Reorder_Wk"].sum())
total_units_str = int(prod_sum["Reorder_STR"].sum())
total_value     = prod_sum["Est_Value"].sum()
n_products      = len(prod_sum)
fast_count      = (prod_sum["STR_Status"].isin(["Super Fast","Fast"])).sum()

c1,c2,c3,c4,c5 = st.columns(5)
for col, val, lbl, clr in [
    (c1, f"{n_products:,}",            "Products shown",     "#374151"),
    (c2, f"{fast_count:,}",            "Fast / Super Fast",  "#16a34a"),
    (c3, f"{total_units_wk:,} units",  f"Order (Weeks/{target_weeks}wk)", "#1d4ed8"),
    (c4, f"{total_units_str:,} units", "Order (STR restore)","#7c3aed"),
    (c5, fmt_npr(total_value),         "Est. Value (Wk)",    "#374151"),
]:
    with col:
        st.markdown(f'<div class="kpi"><p class="kpi-val" style="color:{clr}">{val}</p>'
                    f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Explanation ───────────────────────────────────────────────────────────────
with st.expander("💡 How reorder quantities are calculated"):
    st.markdown("""
**Order (Weeks)** = weeks needed × weekly sell rate − current stock  
→ *"How many to buy to reach your target buffer"*  
→ Increase the **Target weeks of cover** slider for more buffer (e.g. before peak season)

**Order (STR)** = total units ever sold − units currently in stock  
→ *"How many to buy to restore stock back to original level"*  
→ Use this when you want to fully restock a proven seller

**Which to use?**  
- Ongoing replenishment → **Order (Weeks)**  
- Seasonal restock of a bestseller → **Order (STR)**
    """)

# ── Category Summary ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">📊 Category Summary</div>', unsafe_allow_html=True)

# Group by Category + Sub Category if sub cat exists and is populated
has_sub = "Sub Category" in prod_sum.columns and prod_sum["Sub Category"].str.strip().ne("").any()
cat_grp_cols = ["Category","Sub Category"] if has_sub else ["Category"]

cat_sum = prod_sum.groupby(cat_grp_cols).agg(
    Products    = ("Product Name","count"),
    Total_Sold  = ("Total_Sold",  "sum"),
    Total_Stock = ("Total_Stock", "sum"),
    Avg_STR     = ("STR_Pct",     "mean"),
    Order_Wk    = ("Reorder_Wk",  "sum"),
    Order_STR   = ("Reorder_STR", "sum"),
    Est_Value   = ("Est_Value",   "sum"),
).reset_index().sort_values(["Category","Order_Wk"], ascending=[True,False])

cat_sum["Avg_STR"]   = cat_sum["Avg_STR"].round(1)
cat_sum["Est_Value"] = cat_sum["Est_Value"].apply(fmt_npr)
cat_sum = cat_sum.rename(columns={
    "Products":"# Products","Total_Sold":"Units Sold","Total_Stock":"In Stock",
    "Avg_STR":"Avg STR %","Order_Wk":f"Order (Wk/{target_weeks}wk)",
    "Order_STR":"Order (STR)","Est_Value":"Est. Value"
})

def _cat_style(val):
    if isinstance(val,(int,float)) and val > 0:
        return "background-color:#dbeafe;color:#1e40af;font-weight:700"
    return ""

# Column order: Category, Sub Category (if present), then rest
display_cols = (["Category","Sub Category"] if has_sub else ["Category"]) + \
               ["# Products","Units Sold","In Stock","Avg STR %",
                f"Order (Wk/{target_weeks}wk)","Order (STR)","Est. Value"]
display_cols = [c for c in display_cols if c in cat_sum.columns]

st.dataframe(
    cat_sum[display_cols].style.map(_cat_style, subset=[f"Order (Wk/{target_weeks}wk)","Order (STR)"])
                 .format({"Avg STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
                          f"Order (Wk/{target_weeks}wk)":"{:,.0f}","Order (STR)":"{:,.0f}"}),
    width='stretch', hide_index=True)
st.caption("Sorted by Category A–Z, then by Order Qty descending within each category")

# ── Product Table ─────────────────────────────────────────────────────────────
st.markdown('<div class="sec">📋 Product-Level Reorder Plan</div>', unsafe_allow_html=True)

show_cols = ["Product Name","Category","Sub Category","STR_Status","STR_Pct",
             "Total_Sold","Total_Stock","Weekly_Rate","Reorder_Wk","Reorder_STR","Avg_Price","Est_Value"]
show_cols = [c for c in show_cols if c in prod_sum.columns]

disp = prod_sum[show_cols].copy()
disp = disp.rename(columns={
    "STR_Status":"Status","STR_Pct":"STR %","Total_Sold":"Units Sold",
    "Total_Stock":"In Stock","Weekly_Rate":"Rate/wk",
    "Reorder_Wk":f"Order (Wk)","Reorder_STR":"Order (STR)",
    "Avg_Price":"Avg Price","Est_Value":"Est. Value"
})

def _style_status(val):
    return {
        "Super Fast":"background-color:#1B5E20;color:white",
        "Fast":      "background-color:#43A047;color:white",
        "Medium":    "background-color:#F9A825;color:black",
        "Slow":      "background-color:#E53935;color:white",
        "Dead":      "background-color:#424242;color:white",
    }.get(val,"")

def _style_order(val):
    if isinstance(val,(int,float)) and val > 0:
        return "background-color:#dbeafe;color:#1e40af;font-weight:700"
    return ""

fmt = {"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}",
       "Rate/wk":"{:.2f}","Avg Price":"NPR {:,.0f}","Est. Value":"{:,.0f}"}
if "Order (Wk)" in disp.columns: fmt["Order (Wk)"] = "{:,.0f}"
if "Order (STR)" in disp.columns: fmt["Order (STR)"] = "{:,.0f}"

_st = disp.style.map(_style_status, subset=["Status"])
if "Order (Wk)"  in disp.columns: _st = _st.map(_style_order, subset=["Order (Wk)"])
if "Order (STR)" in disp.columns: _st = _st.map(_style_order, subset=["Order (STR)"])

st.dataframe(_st.format(fmt), width='stretch', hide_index=True)
st.caption(f"{len(disp):,} products · 🔵 Order (Wk) = {target_weeks}-week buffer · 🟣 Order (STR) = restore to original stock")

# ── Size Breakdown (if variant data available) ────────────────────────────────
if data_source.startswith("Variant") and size_df is not None and sel_cat != "All":
    st.markdown('<div class="sec">📏 Size Breakdown — Fast Products in This Category</div>',
                unsafe_allow_html=True)

    cat_sizes = size_df[
        (size_df["Brand"].str.strip() == sel_brand) &
        (size_df["Category"].str.strip() == sel_cat) &
        (size_df["Status"].isin(["Super Fast","Fast"]))
    ].copy()

    if not cat_sizes.empty:
        cat_sizes["_sk"] = cat_sizes["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        cat_sizes = cat_sizes.sort_values(["Product Name","_sk"]).drop(columns=["_sk"])
        cat_sizes["Suggest"] = (cat_sizes["Units Sold"] - cat_sizes["In Stock"]).clip(lower=0).round().astype(int)

        disp_s = cat_sizes[["Product Name","Size","Units Sold","In Stock","STR %","Status","Suggest"]].copy()
        _ss = disp_s.style.map(_style_status, subset=["Status"]).map(_style_order, subset=["Suggest"])
        st.dataframe(_ss.format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}","Suggest":"{:,.0f}"}),
                     width='stretch', hide_index=True)
        st.caption(f"Suggest = units sold − in stock per size (for Fast/Super Fast only)")
    else:
        st.info(f"No Super Fast or Fast sizes found for {sel_brand} / {sel_cat}.")

# ── Color Breakdown ───────────────────────────────────────────────────────────
if data_source.startswith("Variant") and color_df is not None and sel_cat != "All":
    st.markdown('<div class="sec">🎨 Color Breakdown — Fast Colors in This Category</div>',
                unsafe_allow_html=True)

    cat_colors = color_df[
        (color_df["Brand"].str.strip() == sel_brand) &
        (color_df["Category"].str.strip() == sel_cat) &
        (color_df["Status"].isin(["Super Fast","Fast"]))
    ].copy() if color_df is not None else pd.DataFrame()

    if not cat_colors.empty:
        cat_colors = cat_colors.sort_values("Units Sold", ascending=False)
        cat_colors["Suggest"] = (cat_colors["Units Sold"] - cat_colors["In Stock"]).clip(lower=0).round().astype(int)

        disp_c = cat_colors[["Product Name","Color","Units Sold","In Stock","STR %","Status","Suggest"]].copy()
        _sc = disp_c.style.map(_style_status, subset=["Status"]).map(_style_order, subset=["Suggest"])
        st.dataframe(_sc.format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}","Suggest":"{:,.0f}"}),
                     width='stretch', hide_index=True)
    else:
        st.info(f"No Super Fast or Fast colors found for {sel_brand} / {sel_cat}.")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    # Summary sheet
    cat_sum.to_excel(writer, sheet_name="Category Summary", index=False)

    # Full product list
    full = prod_sum[["Product Name","Category"] +
                   (["Sub Category"] if "Sub Category" in prod_sum.columns else []) +
                   ["STR_Status","STR_Pct","Total_Sold","Total_Stock",
                    "Weekly_Rate","Reorder_Wk","Reorder_STR","Avg_Price","Est_Value"]].copy()
    full = full.rename(columns={"STR_Status":"Status","STR_Pct":"STR %",
                                "Total_Sold":"Units Sold","Total_Stock":"In Stock",
                                "Weekly_Rate":"Rate/wk","Reorder_Wk":f"Order (Wk {target_weeks}wk)",
                                "Reorder_STR":"Order (STR)","Avg_Price":"Avg Price NPR",
                                "Est_Value":"Est. Value NPR"})
    full.to_excel(writer, sheet_name="Product Reorder Plan", index=False)

    # Size breakdown if available
    if size_df is not None:
        sz_export = size_df[size_df["Brand"].str.strip() == sel_brand].copy()
        if sel_cat != "All":
            sz_export = sz_export[sz_export["Category"].str.strip() == sel_cat]
        sz_export["Suggest"] = (sz_export["Units Sold"] - sz_export["In Stock"]).clip(lower=0).round().astype(int)
        sz_cols = [c for c in ["Product Name","Category","Sub Category","Size","Units Sold","In Stock","STR %","Status","Suggest"]
                   if c in sz_export.columns]
        sz_export[sz_cols].sort_values(["Product Name"]).to_excel(writer, sheet_name="By Size", index=False)

    # Color breakdown if available
    if color_df is not None:
        cl_export = color_df[color_df["Brand"].str.strip() == sel_brand].copy()
        if sel_cat != "All":
            cl_export = cl_export[cl_export["Category"].str.strip() == sel_cat]
        cl_export["Suggest"] = (cl_export["Units Sold"] - cl_export["In Stock"]).clip(lower=0).round().astype(int)
        cl_cols = [c for c in ["Product Name","Category","Sub Category","Color","Units Sold","In Stock","STR %","Status","Suggest"]
                   if c in cl_export.columns]
        cl_export[cl_cols].sort_values(["Product Name"]).to_excel(writer, sheet_name="By Color", index=False)

out.seek(0)
fname = f"reorder_{sel_brand.replace(' ','_')}_{sel_cat.replace(' ','_') if sel_cat!='All' else 'AllCats'}.xlsx"
st.download_button(
    f"⬇️ Download Full Reorder Plan — {sel_brand} / {sel_cat}",
    data=out, file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Download includes: Category Summary · Product Plan · Size Breakdown · Color Breakdown")