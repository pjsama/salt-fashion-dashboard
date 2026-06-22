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
GDRIVE_MAIN_ID    = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_VARIANT_ID = "1LPeoGXDDd3ZAppTiuLskzY4q-71CJWfJ"
GDRIVE_STORE_ID   = "1B8_Ml_tAL59MSPrEDwKUR93ruFEC1m23"

SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL",
              "36","37","38","39","40","41","42","43","44","ONE SIZE","FREE SIZE"]

# ── Name cleaners ─────────────────────────────────────────────────────────────
def clean_name(name):
    """Strip [SKU] prefix, quotes, newlines, tabs, extra spaces."""
    name = str(name).strip()
    name = re.sub(r"^\[[^\]]+\]\s*", "", name)
    name = name.replace("\n", " ").replace("\t", " ")
    name = re.sub(r'^"+|"+$', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name

def clean_store_name(name):
    """Clean store_analysis product names — strip [SKU] prefix AND color/size suffixes.
    Store names look like: '[SA-11112-XL] Salt Trench Coat - Black - XL'
    We want: 'Salt Trench Coat - Black'  to match template names.
    """
    name = clean_name(name)
    # Strip " - Size" suffix at end (e.g. " - L", " - XL")
    name = re.sub(r"\s*-\s*(XS|S|M|L|XL|2XL|3XL|4XL|XXL)\s*$", "", name, flags=re.I)
    # Strip "/Color (Size)" suffix (e.g. "/Dark Brown (S)")
    name = re.sub(r"/[\w\s]+\s*\((XS|S|M|L|XL|2XL|3XL|4XL|XXL)\)\s*$", "", name, flags=re.I)
    # Strip " (Size)" suffix
    name = re.sub(r"\s*\((XS|S|M|L|XL|2XL|3XL|4XL|XXL)\)\s*$", "", name, flags=re.I)
    # Strip "-Size" glued suffix (e.g. "Jacket-M")
    name = re.sub(r"-\s*(XS|S|M|L|XL|2XL|3XL|4XL|XXL)\s*$", "", name, flags=re.I)
    return name.strip()

def strip_variant_suffix(name):
    """Remove color/size suffixes: 'Dress/Black' → 'Dress', 'Dress Green / L' → 'Dress Green'."""
    name = clean_name(name)
    # Remove "/ Size" at end
    name = re.sub(r"\s*/\s*(XS|S|M|L|XL|2XL|3XL|4XL|XXL|36|37|38|39|40|41|42|43|44)$", "", name, flags=re.I)
    # Remove "/Color" at end (single word after slash)
    name = re.sub(r"/\w+$", "", name).strip()
    return name.strip()

# ── Style functions (module-level) ────────────────────────────────────────────
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

def style_doc(val):
    return {"Reorder Now":"background-color:#B71C1C;color:white",
            "Watch":"background-color:#F57F17;color:white",
            "OK":"background-color:#2E7D32;color:white"}.get(val,"")

def str_color(status):
    return {"Super Fast":"#1B5E20","Fast":"#43A047","Medium":"#F9A825",
            "Slow":"#E53935","Dead":"#424242"}.get(status,"#9E9E9E")

def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

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

@st.cache_resource(show_spinner=False)
def load_products():
    buf = _gdrive(GDRIVE_MAIN_ID)
    df = None
    # Skip Image_Base64 and Image columns — they add 85MB and 17s to load time
    _skip_cols = lambda c: c not in ("Image_Base64", "Image")
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl", usecols=_skip_cols)
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base+r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"), reverse=True) if Path(d).exists() else []
            if files:
                df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl", usecols=_skip_cols)
                break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    df["Product Name"] = df["Product Name"].fillna("").astype(str).apply(clean_name)
    for col in ["On Hand Qty","Total Units Sold","Revenue","Sell-Through %","Sales Price","Cost Price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Sell-Through %" in df.columns and df["Sell-Through %"].max() <= 1.0:
        df["Sell-Through %"] *= 100
    for col in ["Brand","Category","Sub Category","STR Status","ABC Class","DOC Status","Color","Size","SKU / Variant"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    # ── Group variants into products using Template_ID ─────────────────────
    # odoo_products has one row per variant (size/color)
    # Template_ID groups all variants of the same product
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
    for df in [size_df, color_df]:
        df.columns = [c.strip() for c in df.columns]
        # Clean product names: strip [SKU] prefix, \n chars
        df["Product Name"] = df["Product Name"].fillna("").astype(str).apply(clean_name)
        for col in ["Units Sold","In Stock","STR %"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in ["Size","Color","Brand","Category","Sub Category","Status"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.replace(r"^(Size|Color|Brand):\s*","",regex=True).str.strip()
    return size_df, color_df

@st.cache_resource(show_spinner=False)
def load_store():
    buf = _gdrive(GDRIVE_STORE_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="🏆 Top Products by Store", engine="openpyxl")
        except: pass
    if df is None:
        files = sorted(Path(r"C:\Users\Legion\Desktop\odoo_export\exports").glob("store_analysis*.xlsx"), reverse=True)
        if files:
            try: df = pd.read_excel(files[0], sheet_name="🏆 Top Products by Store", engine="openpyxl")
            except: pass
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    # Strip [SKU] prefix from product names in store_analysis
    df["Product"] = df["Product"].fillna("").astype(str).apply(clean_store_name)
    for col in ["Revenue (NPR)","Units Sold"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    # Forward-fill Store column
    if "Store" in df.columns:
        df["Store"] = df["Store"].replace("", pd.NA).ffill()
    return df

# ── Data loading ──────────────────────────────────────────────────────────────
# Show loading messages — products file takes ~5s even optimised
_load_container = st.empty()
with _load_container.container():
    with st.spinner("Loading product catalog (first load may take ~10s)…"):
        df_raw = load_products()
with _load_container.container():
    with st.spinner("Loading variant analysis…"):
        size_df, color_df = load_variants()
with _load_container.container():
    with st.spinner("Loading store data…"):
        df_store = load_store()
_load_container.empty()

if df_raw is None:
    st.error("Could not load product data."); st.stop()

# ── Build template-level product catalog ─────────────────────────────────────
# Group variant rows by Template_ID → one row per real product
if "Template_ID" in df_raw.columns and df_raw["Template_ID"].notna().any():
    # Get the most common (canonical) name per template
    def canonical_name(names):
        clean = [strip_variant_suffix(n) for n in names]
        from collections import Counter
        return Counter(clean).most_common(1)[0][0]

    df_templates = df_raw.groupby("Template_ID").agg(
        Product_Name   =("Product Name",  lambda x: canonical_name(x)),
        Brand          =("Brand",         lambda x: x.mode()[0] if len(x) else ""),
        Category       =("Category",      lambda x: x.mode()[0] if len(x) else ""),
        Sub_Category   =("Sub Category",  lambda x: x.mode()[0] if len(x) else ""),
        Total_Sold     =("Total Units Sold","sum"),
        Total_Stock    =("On Hand Qty",   "sum"),
        Total_Revenue  =("Revenue",       "sum"),
        Avg_Price      =("Sales Price",   "mean"),
        STR_Pct        =("Sell-Through %","mean"),
        STR_Status     =("STR Status",    lambda x: x.mode()[0] if len(x) else ""),
        ABC_Class      =("ABC Class",     lambda x: x.mode()[0] if len(x) else ""),
        DOC_Status     =("DOC Status",    lambda x: x.mode()[0] if len(x) else ""),
        Variants       =("Variant_ID",    "count"),
    ).reset_index()
else:
    # Fallback: group by stripped product name
    df_raw["_base"] = df_raw["Product Name"].apply(strip_variant_suffix)
    df_templates = df_raw.groupby("_base").agg(
        Product_Name   =("_base",         "first"),
        Brand          =("Brand",         lambda x: x.mode()[0] if len(x) else ""),
        Category       =("Category",      lambda x: x.mode()[0] if len(x) else ""),
        Sub_Category   =("Sub Category",  lambda x: x.mode()[0] if len(x) else ""),
        Total_Sold     =("Total Units Sold","sum"),
        Total_Stock    =("On Hand Qty",   "sum"),
        Total_Revenue  =("Revenue",       "sum"),
        Avg_Price      =("Sales Price",   "mean"),
        STR_Pct        =("Sell-Through %","mean"),
        STR_Status     =("STR Status",    lambda x: x.mode()[0] if len(x) else ""),
        ABC_Class      =("ABC Class",     lambda x: x.mode()[0] if len(x) else ""),
        DOC_Status     =("DOC Status",    lambda x: x.mode()[0] if len(x) else ""),
        Variants       =("Variant_ID",    "count") if "Variant_ID" in df_raw.columns else ("Product Name","count"),
    ).reset_index(drop=True)

# Clean product catalog
df_templates = df_templates[df_templates["Product_Name"].str.len() > 5]
df_templates = df_templates[df_templates["Product_Name"].str.contains(r" ", na=False)]
df_templates = df_templates[df_templates["Product_Name"].str.contains(r"[a-zA-Z]{3}", na=False)]
df_templates = df_templates[~df_templates["Product_Name"].str.match(r"^[\d\-\.\s/]+$", na=False)]

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔍 Product Deep Dive")
    st.markdown("---")

    brands = sorted([b for b in df_templates["Brand"].unique()
                     if b and b not in ("","nan","True","False")])
    sel_brand = st.selectbox("Brand", brands)

    brand_df = df_templates[df_templates["Brand"] == sel_brand]

    cats = ["All"] + sorted([c for c in brand_df["Category"].unique() if c and c not in ("","nan")])
    sel_cat = st.selectbox("Category", cats)

    filtered_df = brand_df if sel_cat == "All" else brand_df[brand_df["Category"] == sel_cat]

    products = sorted(filtered_df["Product_Name"].unique())
    search = st.text_input("Search product", placeholder="Type to filter…")
    if search.strip():
        products = [p for p in products if search.lower() in p.lower()]

    if not products:
        st.warning("No products found."); st.stop()

    sel_product = st.selectbox("Select Product", products)

    st.markdown("---")
    target_weeks = st.slider("Target weeks of stock", 2, 12, 4,
        help="How many weeks of selling you want in stock.")
    st.markdown("---")
    st.caption(f"{len(products):,} products · {len(filtered_df):,} total in category")
    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_resource.clear(); st.rerun()

# ── Get template row for selected product ─────────────────────────────────────
prod_row = filtered_df[filtered_df["Product_Name"] == sel_product]
if prod_row.empty:
    st.warning(f"No data for {sel_product}"); st.stop()
prod = prod_row.iloc[0]

total_sold   = prod["Total_Sold"]
total_stock  = prod["Total_Stock"]
total_rev    = prod["Total_Revenue"]
avg_price    = prod["Avg_Price"]
str_pct      = prod["STR_Pct"]
str_status   = prod["STR_Status"]
category     = prod["Category"]
sub_cat      = prod["Sub_Category"]
num_variants = prod["Variants"]

# ── Match variant_analysis data ───────────────────────────────────────────────
# variant_analysis Product Name is already cleaned (strip_[SKU]_prefix done in loader)
# Now match against the canonical template name
def find_variant_rows(df, product_name):
    """Match product_name against variant_analysis Product Name.
    Uses exact match first. Fuzzy match requires high word overlap to avoid
    wrong matches (e.g. 'Black Long Dress' matching 'Black Dress', 'Black Top' etc.)
    """
    if df is None or df.empty: return pd.DataFrame()
    # Exact match first
    rows = df[df["Product Name"] == product_name]
    if not rows.empty: return rows
    # Strict fuzzy: require ALL words of the shorter name to appear in the longer name
    pn_words = set(product_name.lower().split())
    pn_lower  = product_name.lower()
    def strict_match(n):
        n_lower = n.lower()
        n_words  = set(n_lower.split())
        if n_lower == pn_lower: return True
        # Must share ALL words of shorter name (not just some)
        shorter = pn_words if len(pn_words) <= len(n_words) else n_words
        longer  = n_words  if len(pn_words) <= len(n_words) else pn_words
        return shorter.issubset(longer) and len(shorter) >= 3
    mask = df["Product Name"].str.lower().apply(strict_match)
    return df[mask].copy()

p_sizes  = find_variant_rows(size_df,  sel_product)
p_colors = find_variant_rows(color_df, sel_product)

# Sort sizes
if not p_sizes.empty and "Size" in p_sizes.columns:
    p_sizes["_sk"] = p_sizes["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
    p_sizes = p_sizes.sort_values("_sk").drop(columns=["_sk"])

# ── Match store data ──────────────────────────────────────────────────────────
def find_store_rows(df_store, product_name):
    if df_store is None or "Product" not in df_store.columns: return pd.DataFrame()
    rows = df_store[df_store["Product"] == product_name]
    if not rows.empty: return rows
    pn_lower = product_name.lower()
    mask = df_store["Product"].str.lower().apply(
        lambda n: pn_lower in n or n in pn_lower or
        len(set(pn_lower.split()) & set(n.split())) >= max(2, len(pn_lower.split()) - 1)
    )
    return df_store[mask].copy()

p_stores = find_store_rows(df_store, sel_product)

# ── Reorder calculation ───────────────────────────────────────────────────────
# Per-size: suggest = max(0, Units_Sold - In_Stock) for Fast/Super Fast sizes
# Product-level: all-time STR-based signal
if not p_sizes.empty and "Units Sold" in p_sizes.columns:
    p_sizes["Suggest Reorder"] = p_sizes.apply(
        lambda r: max(0, round(r["Units Sold"] - r["In Stock"]))
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)
    total_suggest = int(p_sizes["Suggest Reorder"].sum())
else:
    # Fallback: all-time weekly rate × target - stock
    weekly_rate = total_sold / 52 if total_sold > 0 else 0
    total_suggest = max(0, round(weekly_rate * target_weeks - total_stock))

if not p_colors.empty and "Units Sold" in p_colors.columns:
    p_colors["Suggest Reorder"] = p_colors.apply(
        lambda r: max(0, round(r["Units Sold"] - r["In Stock"]))
        if r.get("Status","") in ("Super Fast","Fast") else 0, axis=1)

# ── Verdict ───────────────────────────────────────────────────────────────────
if str_status in ("Super Fast","Fast") and total_suggest > 0:
    vc, vi = "verdict-reorder", "✅"
    vt = f"<strong>Reorder recommended — {total_suggest} units.</strong> Fast seller (STR {str_pct:.0f}%). Suggest quantity = units sold minus current stock per fast-moving size."
elif str_status in ("Super Fast","Fast"):
    vc, vi = "verdict-watch", "📦"
    vt = f"Stock level OK. Strong seller (STR {str_pct:.0f}%) — watch closely, may need reorder soon."
elif str_status == "Medium":
    vc, vi = "verdict-watch", "⚠️"
    vt = f"Medium performer (STR {str_pct:.0f}%). Monitor — reorder only if specific sizes are running out."
else:
    vc, vi = "verdict-pause", "🛑"
    vt = f"Slow/Dead seller (STR {str_pct:.0f}%). Do <strong>not</strong> reorder — focus on clearing existing stock first."

# ── Header ────────────────────────────────────────────────────────────────────
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
    (c1, f"{int(total_sold):,}",    "Total Units Sold",        "#1d4ed8"),
    (c2, f"{int(total_stock):,}",   "In Stock Now",            "#374151"),
    (c3, f"{str_pct:.0f}%",         "Sell-Through Rate",       str_color(str_status)),
    (c4, fmt_npr(total_rev),        "Total Revenue",           "#374151"),
    (c5, fmt_npr(avg_price),        "Avg Selling Price",       "#374151"),
    (c6, f"{total_suggest:,} units",f"Suggest Reorder ({target_weeks}wk)",
         "#16a34a" if total_suggest > 0 else "#6b7280"),
]:
    with col:
        st.markdown(f'<div class="kpi"><p class="kpi-val" style="color:{clr}">{val}</p>'
                    f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Size performance ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">📏 Size Performance — which sizes sell vs which are stuck</div>', unsafe_allow_html=True)

if not p_sizes.empty and "Units Sold" in p_sizes.columns:
    disp = p_sizes[["Size","Units Sold","In Stock","STR %","Status","Suggest Reorder"]].copy()
    disp["STR %"] = disp["STR %"].round(1)
    styled = (disp.style
              .map(style_status,  subset=["Status"])
              .map(style_reorder, subset=["Suggest Reorder"])
              .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}","Suggest Reorder":"{:,.0f}"}))
    st.dataframe(styled, width='stretch', hide_index=True)

    fast   = p_sizes[p_sizes["Status"].isin(["Super Fast","Fast"])]["Size"].tolist()
    dead   = p_sizes[p_sizes["Status"].isin(["Dead","Slow"])]["Size"].tolist()
    parts  = []
    if fast:  parts.append(f"🟢 <strong>Fast sizes: {', '.join(fast)}</strong> — reorder these")
    if dead:  parts.append(f"🔴 <strong>Stuck sizes: {', '.join(dead)}</strong> — don't reorder")
    if total_suggest > 0: parts.append(f"📦 <strong>Total suggested: {total_suggest} units</strong>")
    if parts:
        st.markdown(f'<div class="insight">{"  ·  ".join(parts)}</div>', unsafe_allow_html=True)
else:
    if size_df is not None:
        st.info(
            f"No size breakdown found for **{sel_product}** in variant_analysis. "
            "Two common reasons: ① The product has no size attribute set in Odoo (e.g. single-size items), "
            "or ② Each size was created as a **separate Odoo product** instead of variants — "
            "in that case, see the Full SKU Breakdown table below which shows each size separately."
        )
    else:
        st.info("Size data requires variant_analysis.xlsx — run `python variant_export.py`.")

# ── Color performance ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">🎨 Color Performance — which colors customers want</div>', unsafe_allow_html=True)

if not p_colors.empty and "Units Sold" in p_colors.columns:
    p_colors = p_colors.sort_values("Units Sold", ascending=False)
    disp_c = p_colors[["Color","Units Sold","In Stock","STR %","Status","Suggest Reorder"]].copy()
    disp_c["STR %"] = disp_c["STR %"].round(1)
    styled_c = (disp_c.style
                .map(style_status,  subset=["Status"])
                .map(style_reorder, subset=["Suggest Reorder"])
                .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}","Suggest Reorder":"{:,.0f}"}))
    st.dataframe(styled_c, width='stretch', hide_index=True)

    fast_c = p_colors[p_colors["Status"].isin(["Super Fast","Fast"])]["Color"].tolist()
    dead_c = p_colors[p_colors["Status"].isin(["Dead","Slow"])]["Color"].tolist()
    parts_c = []
    if fast_c: parts_c.append(f"🟢 <strong>Top colors: {', '.join(fast_c[:4])}</strong>")
    if dead_c: parts_c.append(f"🔴 <strong>Not moving: {', '.join(dead_c[:4])}</strong>")
    if parts_c:
        st.markdown(f'<div class="insight">{"  ·  ".join(parts_c)}</div>', unsafe_allow_html=True)
else:
    if color_df is not None:
        st.info(
            f"No color breakdown for **{sel_product}**. "
            "This is normal if: ① the product has no color attribute in Odoo (size-only variants), "
            "or ② the product name differs between the product catalog and variant_analysis file."
        )
    else:
        st.info("Color data requires variant_analysis.xlsx — run `python variant_export.py`.")

# ── Store performance ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">🏪 Store Performance — where this product sells most</div>', unsafe_allow_html=True)

if not p_stores.empty:
    p_stores_d = p_stores[p_stores["Units Sold"] > 0][["Store","Units Sold","Revenue (NPR)"]].copy()
    p_stores_d = p_stores_d.sort_values("Units Sold", ascending=False)
    p_stores_d["Revenue (NPR)"] = p_stores_d["Revenue (NPR)"].apply(fmt_npr)
    col_s, col_b = st.columns([2,3])
    with col_s:
        st.dataframe(p_stores_d, width='stretch', hide_index=True)
    with col_b:
        max_u = p_stores[["Units Sold"]].max().iloc[0] or 1
        for _, row in p_stores.sort_values("Units Sold", ascending=False).iterrows():
            pct = row["Units Sold"] / max_u * 100
            st.markdown(
                f'<div style="margin-bottom:5px">'
                f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">'
                f'<span><strong>{row["Store"]}</strong></span><span style="color:#6b7280">{int(row["Units Sold"]):,} units</span></div>'
                f'<div style="background:#e2e8f0;border-radius:4px;height:8px">'
                f'<div style="background:#1d4ed8;width:{pct:.0f}%;height:8px;border-radius:4px"></div></div></div>',
                unsafe_allow_html=True)
else:
    if df_store is not None:
        st.info(
            f"**{sel_product}** sold {int(total_sold)} units total but doesn't rank in the "
            f"top 20 at any store. The store_analysis file tracks only the top 20 products per store. "
            f"To appear here, a SALT product needs ~7+ units at Lazimpat/Kumaripati or ~4+ at Pokhara "
            f"in the current export period. Summer dresses won't appear during winter season — "
            f"that's expected, not a data error."
        )
    else:
        st.info("Store data requires store_analysis.xlsx — check Google Drive.")

# ── Full SKU breakdown ────────────────────────────────────────────────────────
st.markdown('<div class="sec">📋 Full SKU Breakdown — every size × color from Odoo</div>', unsafe_allow_html=True)

# Get all variant rows for this template
if "Template_ID" in df_raw.columns and not prod_row.empty and "Template_ID" in prod_row.columns:
    tmpl_id = prod_row.iloc[0]["Template_ID"] if "Template_ID" in prod_row.columns else None
    if pd.notna(tmpl_id):
        sku_rows = df_raw[df_raw["Template_ID"] == tmpl_id].copy()
    else:
        sku_rows = df_raw[df_raw["Product Name"].apply(strip_variant_suffix) == sel_product].copy()
else:
    sku_rows = df_raw[df_raw["Product Name"].apply(strip_variant_suffix) == sel_product].copy()

if not sku_rows.empty:
    sku_cols = [c for c in ["Color","Size","SKU / Variant","On Hand Qty","Total Units Sold",
                             "Sell-Through %","STR Status","Sales Price","DOC Status"] if c in sku_rows.columns]
    sku_d = sku_rows[sku_cols].copy()
    if "Size" in sku_d.columns:
        sku_d["_sk"] = sku_d["Size"].apply(lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99)
        sku_d = sku_d.sort_values(["Color","_sk"]).drop(columns=["_sk"])
    if "Sell-Through %" in sku_d.columns: sku_d["Sell-Through %"] = sku_d["Sell-Through %"].round(1)
    fmt = {}
    if "Sell-Through %" in sku_d.columns:   fmt["Sell-Through %"]   = "{:.1f}%"
    if "On Hand Qty" in sku_d.columns:      fmt["On Hand Qty"]      = "{:,.0f}"
    if "Total Units Sold" in sku_d.columns: fmt["Total Units Sold"] = "{:,.0f}"
    if "Sales Price" in sku_d.columns:      fmt["Sales Price"]      = "NPR {:,.0f}"
    apply_s = [c for c in ["STR Status"] if c in sku_d.columns]
    apply_d = [c for c in ["DOC Status"] if c in sku_d.columns]
    styled_sku = sku_d.style.format(fmt)
    if apply_s: styled_sku = styled_sku.map(style_status, subset=apply_s)
    if apply_d: styled_sku = styled_sku.map(style_doc,    subset=apply_d)
    st.dataframe(styled_sku, width='stretch', hide_index=True)
    st.caption(f"{len(sku_rows)} variants total")

# ── Category comparison ───────────────────────────────────────────────────────
st.markdown('<div class="sec">📊 How this product ranks in its category</div>', unsafe_allow_html=True)

cat_peers = df_templates[
    (df_templates["Brand"] == sel_brand) &
    (df_templates["Category"] == category) &
    (df_templates["Product_Name"] != sel_product)
].sort_values("Total_Sold", ascending=False).head(15)

if not cat_peers.empty:
    current = pd.DataFrame([{"Product_Name": f"➡️ {sel_product}",
                              "Total_Sold": total_sold, "Total_Stock": total_stock,
                              "Total_Revenue": total_rev, "STR_Pct": str_pct,
                              "STR_Status": str_status}])
    combined = pd.concat([current, cat_peers], ignore_index=True)
    combined["Total_Revenue"] = combined["Total_Revenue"].apply(fmt_npr)
    combined["STR_Pct"] = combined["STR_Pct"].round(1)
    combined = combined.rename(columns={"Product_Name":"Product","Total_Sold":"Units Sold",
                                        "Total_Stock":"In Stock","Total_Revenue":"Revenue",
                                        "STR_Pct":"STR %","STR_Status":"Status"})

    def highlight_current(row):
        if str(row["Product"]).startswith("➡️"):
            return ["background-color:#eff6ff;font-weight:600"] * len(row)
        return [""] * len(row)

    styled_p = (combined.style
                .apply(highlight_current, axis=1)
                .map(style_status, subset=["Status"])
                .format({"STR %":"{:.1f}%","Units Sold":"{:,.0f}","In Stock":"{:,.0f}"}))
    st.dataframe(styled_p, width='stretch', hide_index=True)
    rank = (cat_peers["Total_Sold"] > total_sold).sum() + 1
    st.caption(f"Ranked #{rank} by units sold within {category} ({sel_brand})")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    pd.DataFrame([{"Product": sel_product, "Category": category, "Sub Category": sub_cat,
                   "Brand": sel_brand, "Total Sold": total_sold, "In Stock": total_stock,
                   "STR %": round(str_pct,1), "STR Status": str_status,
                   "Revenue": total_rev, "Avg Price": round(avg_price),
                   "Suggested Reorder": total_suggest}])\
      .to_excel(writer, sheet_name="Summary", index=False)
    if not p_sizes.empty:
        p_sizes[["Size","Units Sold","In Stock","STR %","Status","Suggest Reorder"]]\
            .to_excel(writer, sheet_name="By Size", index=False)
    if not p_colors.empty:
        p_colors[["Color","Units Sold","In Stock","STR %","Status","Suggest Reorder"]]\
            .to_excel(writer, sheet_name="By Color", index=False)
    if not p_stores.empty:
        p_stores[["Store","Units Sold","Revenue (NPR)"]].to_excel(writer, sheet_name="By Store", index=False)
    if not sku_rows.empty:
        sku_rows[sku_cols].to_excel(writer, sheet_name="All SKUs", index=False)
out.seek(0)
st.download_button(
    f"⬇️ Download {sel_product[:40]} — full report",
    data=out,
    file_name=f"deep_dive_{re.sub(r'[^a-zA-Z0-9]','_',sel_product[:40])}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── Bulk analysis note ────────────────────────────────────────────────────────
with st.expander("💡 Need bulk analysis for multiple products?"):
    st.markdown("""
**For bulk reorder decisions across all products:**

1. **Buying Brief page** — already shows top performing categories with recommendations (Increase/Maintain/Reduce/Watch) and Top 10 Winners
2. **Reorder Plan page** — shows category-level reorder quantities across all stores
3. **Variant Dashboard** — filter by Category + Status = "Super Fast" to see all fast-moving size/color combinations at once

**This page is best for:** When your supervisor asks *"should we reorder this specific product?"* — you pick it, she sees the full picture in one view.

**For bulk supplier orders:** Use the Reorder Plan → Overall Summary tab → Download Supplier Order as Excel.
    """)