import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
from datetime import datetime, timedelta

st.set_page_config(
    page_title="Salt Fashion — Reorder Plan",
    page_icon="📦", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container{padding:1.5rem 2rem}
.reorder-card{background:#ffffff;border-radius:10px;border:1px solid #e2e8f0;padding:16px 18px;margin-bottom:8px}
.urgent{border-left:4px solid #dc2626}
.warning{border-left:4px solid #f59e0b}
.ok{border-left:4px solid #16a34a}
.kpi-box{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;text-align:center}
.kpi-val{font-size:26px;font-weight:700;margin:0}
.kpi-lbl{font-size:11px;color:#6b7280;margin:4px 0 0}
.src-badge{display:inline-block;padding:1px 8px;border-radius:8px;font-size:10px;
           font-weight:600;margin-left:6px}
</style>
""", unsafe_allow_html=True)

# ── Google Drive IDs ──────────────────────────────────────────────────────────
GDRIVE_MAIN_ID      = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_POS_ID       = "1YcW30p_dUfeeaQj-XXmGhMHP0ldAM32X"
GDRIVE_LOCSTK_ID    = "1zgTBhh7vOTjxEIz-LO3YSM-TXJeDUrBT"
GDRIVE_RECENTCAT_ID = "1EMEw10v7zEwsMzrocJWCjkyRfy14LaIM"
# Variant file (Odoo product.product export with Barcode, SKU, Name, Qty On Hand)
# Upload your Product_Variant__product_product_.xlsx to Google Drive and paste the ID here
GDRIVE_VARIANT_STOCK_ID = "1zNSEJReRHXPBNpU0mjmvH-ddu29QsVeM"   # ← paste file ID after uploading

# ── Planning constants ────────────────────────────────────────────────────────
MIN_REORDER_QTY = 5   # suppress "Reorder Soon" when the gap is trivially small

# ── Display stock configuration ───────────────────────────────────────────────
# Store floor areas in sq ft — used to scale display stock proportionally.
# Bigger store = more floor displays = more stock tied up on the floor.
STORE_AREA = {
    "Baneshwor":  2800,
    "Pokhara":    2800,
    "Kumaripati": 2200,
    "Lazimpat":   2000,
    "Chitwan":     500,
    "Online":        0,   # no physical display
    "Baneshwor Lush":  0,
    "Chitwan Lush":    0,
    "Pokhara Lush":    0,
}
MAX_AREA = max(v for v in STORE_AREA.values() if v > 0)  # 2800

# Base display units per category at the largest store (2800 sq ft).
# Smaller stores get a proportionally smaller display requirement.
# These are editable in the sidebar — these are just sensible starting defaults.
DEFAULT_DISPLAY_BASE = {
    "Tops":         30,
    "Dress":        25,
    "Denim Pant":   20,
    "Shorts":       20,
    "Skirt":        20,
    "Skort":        15,
    "Basic Top":    20,
    "Co-Ord Set":   15,
    "Jacket":       15,
    "Coat":         12,
    "Sweater":      12,
    "Cardigan":     12,
    "Sweatshirt":   10,
    "Hoodie":       10,
    "Leggings":     15,
    "Jeans":        20,
    "T-Shirts":     20,
    "Fashion Accessories": 10,
}
DEFAULT_DISPLAY_FALLBACK = 10  # for categories not listed above

LOCATION_ORDER = ["Baneshwor","Lazimpat","Kumaripati","Chitwan","Pokhara","Online",
                  "Baneshwor Lush","Chitwan Lush","Pokhara Lush"]

STORE_NAME_FIX = {
    "lazimpat":       "Lazimpat",
    "baneshwor":      "Baneshwor",
    "chitwan":        "Chitwan",
    "kumaripati":     "Kumaripati",
    "pokhara":        "Pokhara",
    "online":         "Online",
    "main warehouse": "Main Warehouse",
}

SKIP_PARTS = {"All","Saleable","PoS",""}

def split_cat(raw):
    """Returns (category, sub_category) from Odoo path like 'Jacket / Fur Regular'."""
    parts = [p.strip() for p in str(raw).split("/") if p.strip() not in SKIP_PARTS]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    # parts[0] = parent category, parts[1] = sub-category
    return parts[0], parts[1]

def norm_store(name):
    return STORE_NAME_FIX.get(str(name).strip().lower(), str(name).strip())

# ── Seasonal category classification ──────────────────────────────────────────
WINTER_CATEGORIES = {
    "Coat","Jacket","Sweater","Cardigan","Sweatshirt","Hoodie","Waistcoat",
    "Pajamas Set","Vest","Knitted","Fur Regular","Wool",
    # Accessories that are winter-only
    "Beanie","Boots","Scarves & Mufflers","Mufflers","Scarves",
    "Fashion Accessories",  # catch-all for winter accessories
    "Gloves","Earmuffs",
}
SUMMER_CATEGORIES = {
    "T-Shirts","Shorts","Tops","Dress","Co-Ord Set","Tank Top","Swim Wear",
    "Skirt","Skort","Sundress","Basic Top",
}
# Categories explicitly NOT winter accessories (override Fashion Accessories catch-all)
SUMMER_ACCESSORIES = {"Sunglasses","Handbags","Bags","Sandals"}

def season_for_month(month):
    if month in (11,12,1,2): return "Winter"
    if month in (5,6,7,8,9):  return "Summer"
    return "Transition"

CURRENT_SEASON = season_for_month(pd.Timestamp.today().month)

def category_season(cat):
    if cat in WINTER_CATEGORIES: return "Winter"
    if cat in SUMMER_CATEGORIES: return "Summer"
    return "All-Season"

# ── Loaders ───────────────────────────────────────────────────────────────────
def gdrive_bytes(file_id):
    if not file_id:
        return None, None
    try:
        from google.oauth2.service_account import Credentials
        import googleapiclient.discovery
        from googleapiclient.http import MediaIoBaseDownload
        import json as _j
        raw   = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            _j.loads(_j.dumps(dict(raw))),
            scopes=["https://www.googleapis.com/auth/drive"])
        svc  = googleapiclient.discovery.build("drive","v3",credentials=creds)
        req  = svc.files().get_media(fileId=file_id)
        buf  = BytesIO()
        dl   = MediaIoBaseDownload(buf, req)
        done = False
        while not done: _, done = dl.next_chunk()
        buf.seek(0); return buf, None
    except Exception as e:
        return None, str(e)

@st.cache_data(ttl=600, show_spinner=False)
def load_products():
    buf, _ = gdrive_bytes(GDRIVE_MAIN_ID)
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl")
        except: df = None
    else: df = None
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base+r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"),reverse=True) if Path(d).exists() else []
            if files: df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl"); break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    for col in ["Sales Price","On Hand Qty","Total Units Sold"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── Split Category into parent + sub ──────────────────────────────────────
    # The export may already have a "Sub Category" column (from odoo_export_products.py)
    # OR Category may still contain slashes. Handle both cases.
    if "Category" in df.columns:
        has_sub_col = "Sub Category" in df.columns
        has_slashes = df["Category"].str.contains("/", na=False).any()

        if not has_sub_col or has_slashes:
            split = df["Category"].apply(split_cat)
            df["Category"]     = split.apply(lambda x: x[0])
            df["Sub Category"] = split.apply(lambda x: x[1])
        # If Sub Category already exists and no slashes, leave both columns as-is
        if "Sub Category" not in df.columns:
            df["Sub Category"] = ""

    for col in ["Brand","Category","Sub Category","Product Name"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df

@st.cache_data(ttl=600, show_spinner=False)
def load_pos():
    buf, _ = gdrive_bytes(GDRIVE_POS_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Point of Sale Analysis", engine="openpyxl")
        except: pass
    if df is None:
        files = sorted(Path(r"C:\Users\Legion\Desktop\odoo_export\exports").glob("pos_analysis_*.xlsx"),reverse=True) \
                if Path(r"C:\Users\Legion\Desktop\odoo_export\exports").exists() else []
        if files: df = pd.read_excel(files[0], sheet_name="Point of Sale Analysis", engine="openpyxl")
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Location"] != "TOTAL"].dropna(subset=["Location"])
    date_col = "Total" if "Total" in df.columns else "Date"
    df["Date"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["Date"])
    for col in ["Ticket Sold","QTY","Sales Amount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df

@st.cache_data(ttl=600, show_spinner=False)
def load_location_stock():
    buf, err = gdrive_bytes(GDRIVE_LOCSTK_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Store x Category", engine="openpyxl")
        except Exception as e: err = f"Drive file found but sheet read failed: {e}"
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export\exports"
        files = sorted(Path(base).glob("location_stock_*.xlsx"), reverse=True) if Path(base).exists() else []
        if files:
            try:
                df = pd.read_excel(files[0], sheet_name="Store x Category", engine="openpyxl")
                err = None
            except Exception as e:
                err = f"Local file found but sheet read failed: {e}"
    if df is None or df.empty:
        return None, set(), err

    df.columns = [str(c).strip() for c in df.columns]
    cat_col = df.columns[0]
    store_cols = [c for c in df.columns if c != cat_col]

    long_rows = []
    covered_stores = set()
    for _, row in df.iterrows():
        cat = str(row[cat_col]).strip()
        if not cat or cat.lower() in ("nan",""):
            continue
        for store in store_cols:
            covered_stores.add(norm_store(store))
            qty = row[store]
            qty_val = 0.0 if pd.isna(qty) else float(qty)
            long_rows.append({
                "Location": norm_store(store),
                "Category": cat,
                "On_Hand_Real": max(0.0, qty_val),
            })
    if not long_rows:
        return None, set(), err
    return pd.DataFrame(long_rows), covered_stores, None

@st.cache_data(ttl=600, show_spinner=False)
def load_recent_category_sales():
    buf, err = gdrive_bytes(GDRIVE_RECENTCAT_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Recent Category Sales", engine="openpyxl")
        except Exception as e: err = f"Drive file found but sheet read failed: {e}"
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export\exports"
        files = sorted(Path(base).glob("category_sales_recent_*.xlsx"), reverse=True) \
                if Path(base).exists() else []
        if files:
            try:
                df = pd.read_excel(files[0], sheet_name="Recent Category Sales", engine="openpyxl")
                err = None
            except Exception as e:
                err = f"Local file found but sheet read failed: {e}"
    if df is None or df.empty:
        return None, err
    df.columns = [str(c).strip() for c in df.columns]
    df["Location"] = df["Location"].apply(norm_store)
    return df, None

@st.cache_data(ttl=600, show_spinner=False)
def load_variant_stock():
    """Odoo product.product export: Barcode, Internal Reference, Name, Qty On Hand."""
    buf, _ = gdrive_bytes(GDRIVE_VARIANT_STOCK_ID) if GDRIVE_VARIANT_STOCK_ID else (None, None)
    df = None
    if buf:
        try: df = pd.read_excel(buf, engine="openpyxl")
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for fname in ["Product_Variant__product_product_.xlsx","product_variants.xlsx"]:
            p = Path(base) / fname
            if p.exists():
                try: df = pd.read_excel(p, engine="openpyxl"); break
                except: pass
    if df is None or df.empty: return None
    df.columns = [str(c).strip() for c in df.columns]
    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if "qty" in cl or "quantity" in cl: col_map[c] = "Qty"
        elif "barcode" in cl:               col_map[c] = "Barcode"
        elif "internal" in cl:              col_map[c] = "SKU"
        elif c.strip().lower() == "name":   col_map[c] = "Name"
    df = df.rename(columns=col_map)
    if "Name" not in df.columns or "Qty" not in df.columns: return None
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Base Name"] = df["Name"].str.split("/").str[0].str.strip().str.strip("\n").str.strip()
    return df


def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

# ── Main ──────────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod   = load_products()
    df_pos    = load_pos()
    df_locstk, covered_stores, locstk_err = load_location_stock()
    df_recent_cat, recentcat_err = load_recent_category_sales()
    df_variants = load_variant_stock()

if df_prod is None or df_pos is None:
    st.error("Could not load data. Make sure both product and POS files are on Google Drive.")
    st.stop()

USING_REAL_STOCK     = df_locstk is not None and not df_locstk.empty
USING_SEASONAL_RATES = df_recent_cat is not None and not df_recent_cat.empty

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📦 Reorder Planner")
    st.markdown("---")

    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("nan","True","False","None","")])
    sel_brand = st.selectbox("Brand", brands, index=0)

    st.markdown("---")
    st.markdown("**Planning settings**")
    target_weeks = st.slider("Target weeks of cover", 2, 12, 4,
        help="How many weeks of stock you want to always have.")
    lookback_weeks = st.slider("Sales lookback (weeks)", 2, 12, 4,
        help="How many recent weeks to use for calculating weekly sell rate.")
    min_weekly_rate = st.number_input("Min weekly rate to show (units)", 0, 50, 1,
        help="Hide categories selling fewer than this per week.")

    st.markdown("---")
    locations = ["All"] + [l for l in LOCATION_ORDER if l in df_pos["Location"].unique()]
    sel_loc = st.selectbox("Filter by location", locations)

    # ── Category filter (parent only) ─────────────────────────────────────────
    prod_brand = df_prod[df_prod["Brand"] == sel_brand]
    parent_cats = sorted([c for c in prod_brand["Category"].unique()
                          if c and c not in ("nan","")])
    sel_cat = st.selectbox("Filter by category", ["All"] + parent_cats)

    # ── Sub-category filter (cascades from parent selection) ──────────────────
    sel_sub_cat = "All"
    if sel_cat != "All" and "Sub Category" in prod_brand.columns:
        sub_cats = sorted([s for s in
                           prod_brand[prod_brand["Category"] == sel_cat]["Sub Category"].unique()
                           if s and s not in ("nan", "")])
        if sub_cats:
            sel_sub_cat = st.selectbox(
                "Filter by sub-category",
                ["All"] + sub_cats,
                help=f"Sub-types within {sel_cat}")

    SEASON_OPTIONS = ["All", "Summer (+ All-Season)", "Winter (+ All-Season)", "All-Season only"]
    sel_season_raw = st.selectbox(
        "Season",
        SEASON_OPTIONS,
        index=1,  # default to Summer
        help="Summer and Winter both include All-Season items (Denim, Leggings etc.)"
    )
    sel_season = {
        "All":                      "All",
        "Summer (+ All-Season)":    "Summer",
        "Winter (+ All-Season)":    "Winter",
        "All-Season only":          "All-Season",
    }[sel_season_raw]

    # ── Display stock settings ─────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**🪟 Display Stock Settings**",
                help="Display stock = units permanently on the shop floor that can't be sold "
                     "from the back. Deducting this gives the true free/buffer stock.")

    show_display = st.toggle("Enable display stock deduction", value=True)

    display_base = {}
    display_overrides = {}   # (store, category) -> override units

    if show_display:
        # Let user edit the base units per category
        with st.expander("Base display units (at largest store)", expanded=False):
            st.caption("Units of each category on display at a 2800 sq ft store. "
                       "Smaller stores are scaled down automatically by floor area.")
            # Show only categories present in the current brand's data
            visible_cats = sorted([c for c in prod_brand["Category"].unique()
                                   if c and c not in ("nan","")])
            for cat in visible_cats:
                default_val = DEFAULT_DISPLAY_BASE.get(cat, DEFAULT_DISPLAY_FALLBACK)
                display_base[cat] = st.number_input(
                    cat, min_value=0, max_value=200,
                    value=default_val, step=1, key=f"disp_{cat}")

        # Per-store overrides
        with st.expander("Store overrides (optional)", expanded=False):
            st.caption("Override the area-formula for a specific store + category. "
                       "Leave 0 to use the formula.")
            override_store = st.selectbox("Store", [s for s in LOCATION_ORDER
                                                     if STORE_AREA.get(s,0) > 0],
                                          key="ovr_store")
            override_cat   = st.selectbox("Category", ["(none)"] + sorted(display_base.keys()),
                                          key="ovr_cat")
            override_units = st.number_input("Override units", min_value=0, max_value=500,
                                             value=0, step=1, key="ovr_units")
            if st.button("➕ Add override") and override_cat != "(none)" and override_units > 0:
                st.session_state[f"override_{override_store}_{override_cat}"] = override_units

            # Show active overrides stored in session state
            active = {k: v for k, v in st.session_state.items()
                      if k.startswith("override_") and v > 0}
            if active:
                st.markdown("**Active overrides:**")
                for k, v in active.items():
                    parts = k.replace("override_","").rsplit("_",1)
                    st.markdown(f"- {parts[0]} · {parts[1] if len(parts)>1 else ''} = **{v} units**")
                    display_overrides[tuple(k.replace("override_","").rsplit("_",1))] = v
                if st.button("🗑 Clear all overrides"):
                    for k in list(st.session_state.keys()):
                        if k.startswith("override_"):
                            del st.session_state[k]
                    st.rerun()

    st.markdown("---")
    if USING_REAL_STOCK:
        st.success("✅ Using real per-location stock")
    else:
        msg = "⚠️ Real location stock not found — using estimated split. "
        if locstk_err:
            msg += f"Error: {locstk_err}. "
        msg += "Run `python fetch_location_stock.py` and set GDRIVE_LOCSTK_ID."
        st.warning(msg)

    if USING_SEASONAL_RATES:
        st.success("✅ Using current-season sell rates")
    else:
        msg = "⚠️ Seasonal sales data not found — winter items may show as 'urgent'. "
        if recentcat_err:
            msg += f"Error: {recentcat_err}. "
        msg += "Run `python fetch_recent_category_sales.py --brand SALT` and set GDRIVE_RECENTCAT_ID."
        st.warning(msg)

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ── Calculate weekly sell rates from POS data ─────────────────────────────────
today          = pd.Timestamp.today().normalize()
lookback_start = today - pd.Timedelta(weeks=lookback_weeks)
pos_recent     = df_pos[df_pos["Date"] >= lookback_start].copy()

if "Brand" in pos_recent.columns:
    b_key = "Lush" if "Lush" in sel_brand else "Salt"
    pos_recent = pos_recent[pos_recent["Brand"].str.contains(b_key, case=False, na=False)]

qty_col = "QTY" if "QTY" in pos_recent.columns else "Units"
rev_col = "Sales Amount" if "Sales Amount" in pos_recent.columns else "Revenue"

pos_agg = pos_recent.groupby("Location").agg(
    Total_Units=(qty_col, "sum"),
    Total_Revenue=(rev_col, "sum"),
    Days=("Date", "nunique"),
).reset_index()
pos_agg["Weekly_Rate"]    = pos_agg["Total_Units"] / lookback_weeks
pos_agg["Daily_Rate"]     = pos_agg["Total_Units"] / (lookback_weeks * 7)
pos_agg["Weekly_Revenue"] = pos_agg["Total_Revenue"] / lookback_weeks

# ── Stock per (Category, Sub Category) from product data ─────────────────────
prod_brand = df_prod[df_prod["Brand"] == sel_brand].copy()

# Aggregate at (Category, Sub Category) level so sub-category is preserved
group_cols = ["Category", "Sub Category"] if "Sub Category" in prod_brand.columns else ["Category"]

cat_stock = prod_brand.groupby(group_cols).agg(
    On_Hand=("On Hand Qty","sum"),
    Products=("Product Name","nunique"),
    Avg_Price=("Sales Price","mean"),
).reset_index()

# Avg price lookup keyed by (Category, Sub Category)
avg_price_map = {
    (row["Category"], row.get("Sub Category","")): row["Avg_Price"]
    for _, row in cat_stock.iterrows()
}

# ── Display stock calculator ──────────────────────────────────────────────────
def calc_display_stock(store, cat, display_base, display_overrides, show_display):
    """
    Returns how many units of `cat` are permanently on display at `store`.
    Priority: override > area-formula > 0 (if disabled).
    """
    if not show_display:
        return 0
    # Check override (stored as session state key override_{store}_{cat})
    ovr_key = (store, cat)
    if ovr_key in display_overrides:
        return display_overrides[ovr_key]
    # Area-based formula
    area      = STORE_AREA.get(store, 0)
    if area == 0:
        return 0
    base      = display_base.get(cat, DEFAULT_DISPLAY_FALLBACK)
    area_frac = area / MAX_AREA          # e.g. Chitwan: 500/2800 = 0.179
    return round(base * area_frac)

# ── Build reorder plan ────────────────────────────────────────────────────────
total_units = pos_agg["Total_Units"].sum()
pos_agg["Location_Share"] = pos_agg["Total_Units"] / total_units if total_units > 0 else 0

cat_sold = prod_brand.groupby(group_cols)["Total Units Sold"].sum().reset_index()
cat_sold.columns = group_cols + ["Total_Sold"]
cat_data = cat_stock.merge(cat_sold, on=group_cols, how="left").fillna(0)
total_sold_all = cat_data["Total_Sold"].sum()

# Real stock lookup: (location, category) -> on-hand
# Note: location_stock only has parent category, so we look up by parent cat
real_stock_map = {}
if USING_REAL_STOCK:
    for _, r in df_locstk.iterrows():
        real_stock_map[(r["Location"], r["Category"])] = r["On_Hand_Real"]

# Recent seasonal rate lookup: (location, category) -> weekly rate
recent_rate_map = {}
if USING_SEASONAL_RATES:
    for _, r in df_recent_cat.iterrows():
        recent_rate_map[(r["Location"], r["Category"])] = float(r.get("Weekly Rate", 0) or 0)

rows = []
for _, loc_row in pos_agg.iterrows():
    loc             = loc_row["Location"]
    share           = loc_row["Location_Share"]
    loc_weekly_rate = loc_row["Weekly_Rate"]
    loc_weekly_rev  = loc_row["Weekly_Revenue"]

    for _, cat_row in cat_data.iterrows():
        cat     = cat_row["Category"]
        sub_cat = cat_row.get("Sub Category", "") if "Sub Category" in cat_row.index else ""
        if not cat or cat in ("nan","","All"): continue

        # ── Stock: real per parent-category if available, else proportional ──
        real_val = real_stock_map.get((loc, cat))
        if real_val is not None:
            # When there are multiple sub-cats under one parent, split the
            # parent-level real stock proportionally by their historical sales.
            # This is the best we can do — location_stock tracks at category level.
            sub_sold_total = cat_data[cat_data["Category"] == cat]["Total_Sold"].sum()
            if sub_sold_total > 0:
                sub_fraction = cat_row["Total_Sold"] / sub_sold_total
            else:
                sub_count    = len(cat_data[cat_data["Category"] == cat])
                sub_fraction = 1.0 / sub_count if sub_count > 0 else 1.0
            est_stock    = max(0, real_val * sub_fraction)
            stock_source = "real"
        elif loc in covered_stores:
            est_stock    = 0
            stock_source = "real"
        else:
            est_stock    = cat_row["On_Hand"] * share
            stock_source = "est"

        # ── Weekly rate: seasonal if available, else all-time share ──────────
        cat_share_of_total = cat_row["Total_Sold"] / total_sold_all if total_sold_all > 0 else 0
        if (loc, cat) in recent_rate_map:
            # Seasonal rate is also at parent-category level — split same way
            sub_sold_total = cat_data[cat_data["Category"] == cat]["Total_Sold"].sum()
            sub_fraction_rate = (cat_row["Total_Sold"] / sub_sold_total
                                 if sub_sold_total > 0 else
                                 1.0 / max(1, len(cat_data[cat_data["Category"] == cat])))
            weekly_rate = recent_rate_map[(loc, cat)] * sub_fraction_rate
            rate_source = "seasonal"
        else:
            weekly_rate = loc_weekly_rate * cat_share_of_total
            rate_source = "alltime"

        if weekly_rate < min_weekly_rate: continue

        daily_rate   = weekly_rate / 7
        days_cover   = est_stock / daily_rate if daily_rate > 0 else 999
        weeks_cover  = days_cover / 7
        target_stock = target_weeks * weekly_rate
        reorder_qty  = max(0, round(target_stock - est_stock))

        # ── Display stock deduction ───────────────────────────────────────────
        display_units  = calc_display_stock(loc, cat, display_base, display_overrides, show_display)
        free_stock     = max(0, est_stock - display_units)   # stock not locked on floor

        # Adjusted metrics using free stock instead of total stock
        days_cover_adj  = free_stock / daily_rate if daily_rate > 0 else 999
        weeks_cover_adj = days_cover_adj / 7
        reorder_qty_adj = max(0, round(target_stock - free_stock))

        # ── Urgency (based on raw stock — shown alongside adjusted) ──────────
        if weeks_cover <= 1:
            urgency     = "🔴 Urgent"
            urgency_key = 0
        elif weeks_cover < target_weeks and reorder_qty >= MIN_REORDER_QTY:
            urgency     = "🟡 Reorder Soon"
            urgency_key = 1
        else:
            urgency     = "🟢 OK"
            urgency_key = 2

        # Adjusted urgency (using free stock)
        if weeks_cover_adj <= 1:
            urgency_adj     = "🔴 Urgent"
            urgency_key_adj = 0
        elif weeks_cover_adj < target_weeks and reorder_qty_adj >= MIN_REORDER_QTY:
            urgency_adj     = "🟡 Reorder Soon"
            urgency_key_adj = 1
        else:
            urgency_adj     = "🟢 OK"
            urgency_key_adj = 2

        avg_price = avg_price_map.get((cat, sub_cat), 0)

        rows.append({
            "Location":          loc,
            "Category":          cat,
            "Sub Category":      sub_cat,
            "Season":            category_season(cat),
            # Raw (no display deduction)
            "Est. Stock":        round(est_stock),
            "Stock Source":      stock_source,
            "Rate Source":       rate_source,
            "Weekly Rate":       round(weekly_rate, 1),
            "Weeks Cover":       round(weeks_cover, 1),
            "Target Stock":      round(target_stock),
            "Reorder Qty":       reorder_qty,
            "Est. Value":        round(reorder_qty * avg_price),
            "Urgency":           urgency,
            "_urgency_key":      urgency_key,
            # Adjusted (display stock deducted)
            "Display Stock":     display_units,
            "Free Stock":        round(free_stock),
            "Weeks Cover (Adj)": round(weeks_cover_adj, 1),
            "Reorder Qty (Adj)": reorder_qty_adj,
            "Est. Value (Adj)":  round(reorder_qty_adj * avg_price),
            "Urgency (Adj)":     urgency_adj,
            "_urgency_key_adj":  urgency_key_adj,
            "_weekly_rev":       loc_weekly_rev * cat_share_of_total if rate_source == "alltime" else 0,
        })

df_plan = pd.DataFrame(rows)

if df_plan.empty:
    st.warning("No reorder data — check that POS and product data are loaded correctly.")
    st.stop()

# ── Apply filters ─────────────────────────────────────────────────────────────
if sel_loc != "All":
    df_plan = df_plan[df_plan["Location"] == sel_loc]
if sel_cat != "All":
    df_plan = df_plan[df_plan["Category"] == sel_cat]
if sel_sub_cat != "All":
    df_plan = df_plan[df_plan["Sub Category"] == sel_sub_cat]
# Season filter: Summer/Winter includes All-Season items (Denim, Leggings — year-round)
if sel_season != "All":
    df_plan = df_plan[
        (df_plan["Season"] == sel_season) |
        (df_plan["Season"] == "All-Season")
    ]

df_plan = df_plan.sort_values(["_urgency_key","Reorder Qty"], ascending=[True,False])

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📦 Reorder Planner")
src_badge = ('<span class="src-badge" style="background:#dcfce7;color:#166534">Real stock</span>'
              if USING_REAL_STOCK else
              '<span class="src-badge" style="background:#fef3c7;color:#92400e">Estimated stock</span>')
filter_desc = sel_cat if sel_cat != "All" else "All Categories"
if sel_sub_cat != "All":
    filter_desc += f" › {sel_sub_cat}"
st.markdown(
    f"{sel_brand} · {filter_desc} · {target_weeks}-week target · "
    f"Last {lookback_weeks} weeks · {today.strftime('%B %d, %Y')} {src_badge}",
    unsafe_allow_html=True)

# ── KPI strip ─────────────────────────────────────────────────────────────────
urgent_count       = len(df_plan[df_plan["_urgency_key"]==0])
reorder_count      = len(df_plan[df_plan["_urgency_key"]<=1])
total_units_needed = df_plan["Reorder Qty"].sum()
total_value_needed = df_plan["Est. Value"].sum()

# Adjusted (display stock deducted)
urgent_count_adj       = len(df_plan[df_plan["_urgency_key_adj"]==0])
reorder_count_adj      = len(df_plan[df_plan["_urgency_key_adj"]<=1])
total_units_needed_adj = df_plan["Reorder Qty (Adj)"].sum()
total_value_needed_adj = df_plan["Est. Value (Adj)"].sum()

if show_display:
    # Show 2 rows of KPIs: raw on top, adjusted below
    st.markdown("**Without display stock deduction** *(raw)*")
    c1,c2,c3,c4 = st.columns(4)
    for col, val, lbl, clr in [
        (c1, f"🔴 {urgent_count}",              "Urgent — under 1 week",   "#dc2626"),
        (c2, f"🟡 {reorder_count-urgent_count}", "Reorder Soon",            "#d97706"),
        (c3, f"{int(total_units_needed):,}",    "Units to Reorder",         "#1d4ed8"),
        (c4, fmt_npr(total_value_needed),       "Est. Value",               "#374151"),
    ]:
        with col:
            st.markdown(f'<div class="kpi-box"><p class="kpi-val" style="color:{clr}">{val}</p>'
                        f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**After display stock deduction** *(adjusted — supervisor view)*")
    d1,d2,d3,d4 = st.columns(4)
    for col, val, lbl, clr in [
        (d1, f"🔴 {urgent_count_adj}",                  "Urgent — under 1 week",   "#dc2626"),
        (d2, f"🟡 {reorder_count_adj-urgent_count_adj}", "Reorder Soon",            "#d97706"),
        (d3, f"{int(total_units_needed_adj):,}",         "Units to Reorder (Adj)",  "#1d4ed8"),
        (d4, fmt_npr(total_value_needed_adj),            "Est. Value (Adj)",        "#374151"),
    ]:
        with col:
            st.markdown(f'<div class="kpi-box" style="border-color:#6366f1">'
                        f'<p class="kpi-val" style="color:{clr}">{val}</p>'
                        f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)
else:
    c1,c2,c3,c4 = st.columns(4)
    for col, val, lbl, clr in [
        (c1, f"🔴 {urgent_count}",              "Urgent — under 1 week stock", "#dc2626"),
        (c2, f"🟡 {reorder_count-urgent_count}", "Reorder Soon — under target", "#d97706"),
        (c3, f"{int(total_units_needed):,}",    "Total Units to Reorder",       "#1d4ed8"),
        (c4, fmt_npr(total_value_needed),       "Est. Reorder Value",           "#374151"),
    ]:
        with col:
            st.markdown(f'<div class="kpi-box"><p class="kpi-val" style="color:{clr}">{val}</p>'
                        f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📦 Overall Summary", "🔴 Urgent & Reorder", "📊 Full Plan", "📍 By Location"])

# ── Tab 1: Overall Summary ────────────────────────────────────────────────────
with tab1:
    qty_col_o = "Reorder Qty (Adj)" if show_display else "Reorder Qty"

    # Pool all stores — treat as one shared warehouse
    # Order Qty = max(0, total_target − total_free_stock)
    pool = df_plan.groupby(["Category", "Sub Category"]).agg(
        Stores           =("Location",    "nunique"),
        Total_Stock      =("Est. Stock",  "sum"),
        Total_Free       =("Free Stock",  "sum"),
        Total_Rate       =("Weekly Rate", "sum"),
    ).reset_index()

    pool["Target"]     = (target_weeks * pool["Total_Rate"]).round()
    pool["Order_Qty"]  = (pool["Target"] - pool["Total_Free"]).clip(lower=0).round().astype(int)
    pool["Wks_Cover"]  = (
        pool["Total_Free"] /
        (pool["Total_Rate"] / 7).replace(0, float("nan")) / 7
    ).round(1)

    def pool_status(row):
        wk = row["Wks_Cover"]
        oq = row["Order_Qty"]
        import math
        if math.isnan(wk) or wk > target_weeks: return ("🟢 OK", 3)
        if wk <= 1:                              return ("🔴 Urgent", 0)
        if oq >= 5:                              return ("🟡 Reorder Soon", 1)
        if oq >= 1:                              return ("⚠️ Watch", 2)
        return ("🟢 OK", 3)

    pool[["Status","_sort"]] = pd.DataFrame(
        pool.apply(pool_status, axis=1).tolist(), index=pool.index
    )
    pool["Wks_Cover_fmt"] = pool["Wks_Cover"].apply(
        lambda x: "—" if pd.isna(x) or x > 50 else f"{x:.1f} wks"
    )
    pool = pool.sort_values(["_sort","Order_Qty"], ascending=[True,False])

    # KPIs
    p_urgent  = (pool["_sort"] == 0).sum()
    p_reorder = (pool["_sort"] == 1).sum()
    p_watch   = (pool["_sort"] == 2).sum()
    p_units   = int(pool["Order_Qty"].sum())

    st.markdown("### 🏭 Supplier Order Summary")
    st.caption(
        "All stores treated as one shared pool. "
        "**Order Qty = units to buy from supplier.** "
        "For store-level gaps, see 📍 By Location."
    )
    st.markdown("<br>", unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    for col, val, lbl, clr in [
        (k1, f"🔴 {p_urgent}",   "Urgent (pooled)",       "#dc2626"),
        (k2, f"🟡 {p_reorder}",  "Reorder Soon (pooled)", "#d97706"),
        (k3, f"⚠️ {p_watch}",    "Watch (pooled)",        "#f97316"),
        (k4, f"{p_units:,}",     "Total Units to Buy",    "#1d4ed8"),
    ]:
        with col:
            st.markdown(
                f'<div class="kpi"><p class="kpi-val" style="color:{clr}">{val}</p>'
                f'<p class="kpi-lbl">{lbl}</p></div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # Category rollup
    st.markdown("**By Category — total to buy from supplier**")
    cat_pool = pool.groupby("Category").agg(
        Sub_cats   =("Sub Category", "nunique"),
        Stores     =("Stores",       "max"),
        Total_Stock=("Total_Stock",  "sum"),
        Free_Stock =("Total_Free",   "sum"),
        Weekly_Rate=("Total_Rate",   "sum"),
        Order_Qty  =("Order_Qty",    "sum"),
    ).reset_index()
    cat_pool["Weeks Cover"] = (
        cat_pool["Free_Stock"] /
        (cat_pool["Weekly_Rate"] / 7).replace(0, float("nan")) / 7
    ).round(1).apply(lambda x: "—" if pd.isna(x) or x > 50 else f"{x:.1f} wks")
    cat_pool["Status"] = cat_pool["Order_Qty"].apply(
        lambda q: "🟡 Reorder" if q >= 5 else ("⚠️ Watch" if q >= 1 else "🟢 OK"))
    cat_pool["_s"] = cat_pool["Order_Qty"].apply(lambda q: 0 if q >= 5 else (1 if q >= 1 else 2))
    cat_pool = cat_pool.sort_values(["_s","Order_Qty"], ascending=[True,False]).drop(columns=["_s"])
    cat_pool["Weekly_Rate"] = cat_pool["Weekly_Rate"].round(1)
    cat_pool = cat_pool.rename(columns={
        "Total_Stock":"Total Stock", "Free_Stock":"Free Stock",
        "Weekly_Rate":"Weekly Rate", "Order_Qty":"Order Qty",
    })
    st.dataframe(
        cat_pool[["Status","Category","Sub_cats","Stores","Total Stock",
                  "Free Stock","Weekly Rate","Weeks Cover","Order Qty"]],
        use_container_width=True, hide_index=True
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # Sub-category detail
    st.markdown("**By Sub-Category — detail (pooled)**")
    sub_disp = pool[[
        "Status","Category","Sub Category","Stores",
        "Total_Stock","Total_Free","Total_Rate",
        "Wks_Cover_fmt","Target","Order_Qty"
    ]].copy().rename(columns={
        "Sub Category":"Sub-cat",
        "Total_Stock":"Total Stock",
        "Total_Free":"Free Stock",
        "Total_Rate":"Weekly Rate",
        "Wks_Cover_fmt":"Weeks Cover",
        "Target":"Target Stock",
        "Order_Qty":"Order Qty",
    })
    sub_disp["Weekly Rate"] = sub_disp["Weekly Rate"].round(1)
    st.dataframe(
        sub_disp[["Status","Category","Sub-cat","Stores","Total Stock",
                  "Free Stock","Weekly Rate","Weeks Cover","Target Stock","Order Qty"]],
        use_container_width=True, hide_index=True
    )

    st.info(
        "💡 **Order Qty = 0** for a category means total chain stock already meets "
        "the target — but individual stores may still need stock moved. "
        "See **📍 By Location** for distribution gaps."
    )

    # ── Sold-Out Products with Recent Sales ───────────────────────────────────
    # Shows specific products that are completely out of stock (all variants = 0)
    # but were selling recently — these are the real reorder candidates the
    # category pool misses because old dead stock inflates category totals.
    st.markdown("---")
    st.markdown("### 🚨 Sold-Out Products with Recent Sales")
    st.caption(
        "These specific products are completely sold out (every size/color = 0 stock) "
        "but had recent POS sales — meaning customers wanted them. "
        "The category pool shows Order Qty = 0 because other old styles still have stock, "
        "but **these specific styles need to be reordered.**"
    )

    if df_variants is not None:
        # Filter to selected brand products from the product export
        # Use df_prod (which has Brand column) to get product names for this brand
        brand_products = df_prod[df_prod["Brand"] == sel_brand]["Product Name"].str.strip().unique()

        # Aggregate variant stock by base product name
        v = df_variants.copy()
        prod_stock = v.groupby("Base Name").agg(
            Total_Qty   =("Qty", "sum"),
            Variants    =("Qty", "count"),
            Zero_Vars   =("Qty", lambda x: (x <= 0).sum()),
        ).reset_index()

        # Completely sold out = all variants are zero or negative
        sold_out_all = prod_stock[
            (prod_stock["Total_Qty"] <= 0) &
            (prod_stock["Zero_Vars"] == prod_stock["Variants"])
        ]["Base Name"].tolist()

        # Cross with brand products to filter to selected brand
        # Match by checking if the product name is close to something in brand_products
        # Use simple contains matching since names may differ slightly
        sold_out_brand = []
        brand_prod_set = set(p.lower() for p in brand_products)
        for name in sold_out_all:
            # Direct match
            if name.lower() in brand_prod_set:
                sold_out_brand.append(name)
            # Fuzzy: check if any brand product contains this name or vice versa
            elif any(name.lower() in bp or bp in name.lower()
                     for bp in brand_prod_set if len(name) > 5):
                sold_out_brand.append(name)

        if not sold_out_brand:
            # Fall back to all sold-out products if brand match yields nothing
            sold_out_brand = sold_out_all[:50]

        if sold_out_brand:
            # Get the variant detail for sold-out products
            sold_detail = v[v["Base Name"].isin(sold_out_brand)].copy()

            # Build size breakdown per product
            def get_sizes(name, df):
                subs = df[df["Base Name"] == name]
                # Try to extract size from SKU (format: XXXXX-Color-Size)
                sizes = []
                for _, row in subs.iterrows():
                    sku = str(row.get("SKU", ""))
                    parts = sku.split("-")
                    last = parts[-1].upper() if parts else ""
                    SIZE_SET = {"XS","S","M","L","XL","XXL","2XL","3XL",
                                "36","37","38","39","40","41","42","43","44"}
                    if last in SIZE_SET:
                        sizes.append(last)
                return ", ".join(sorted(set(sizes))) if sizes else "—"

            # Deduplicate and build display table
            seen = set()
            rows_out = []
            for name in sold_out_brand:
                if name in seen: continue
                seen.add(name)
                sub = prod_stock[prod_stock["Base Name"] == name].iloc[0]
                sizes = get_sizes(name, v)
                rows_out.append({
                    "Product Name":  name,
                    "Variants":      int(sub["Variants"]),
                    "Sizes":         sizes,
                    "Total Qty":     int(sub["Total_Qty"]),
                })

            # Apply category filter if set
            if sel_cat != "All" and df_prod is not None:
                cat_prods = df_prod[
                    (df_prod["Brand"] == sel_brand) &
                    (df_prod["Category"] == sel_cat)
                ]["Product Name"].str.strip().str.lower().tolist()
                rows_out = [r for r in rows_out
                            if any(cp in r["Product Name"].lower() or
                                   r["Product Name"].lower() in cp
                                   for cp in cat_prods)]

            if rows_out:
                df_out = pd.DataFrame(rows_out).sort_values("Product Name")

                # KPI
                col_kpi1, col_kpi2 = st.columns(2)
                with col_kpi1:
                    st.markdown(
                        f'<div class="kpi-box">'                        f'<p class="kpi-val" style="color:#dc2626">{len(df_out)}</p>'                        f'<p class="kpi-lbl">Styles completely sold out</p></div>',
                        unsafe_allow_html=True
                    )
                with col_kpi2:
                    st.markdown(
                        f'<div class="kpi-box">'                        f'<p class="kpi-val" style="color:#d97706">{df_out["Variants"].sum()}</p>'                        f'<p class="kpi-lbl">SKUs (size/color combinations) at zero</p></div>',
                        unsafe_allow_html=True
                    )
                st.markdown("<br>", unsafe_allow_html=True)
                st.dataframe(df_out, use_container_width=True, hide_index=True)

                # Download
                out_so = BytesIO()
                df_out.to_excel(out_so, index=False, engine="openpyxl")
                out_so.seek(0)
                st.download_button(
                    "⬇️ Download Sold-Out List",
                    data=out_so,
                    file_name=f"sold_out_{sel_brand}_{today.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.success(f"✅ No completely sold-out products found for {sel_brand} / {sel_cat}.")
        else:
            st.success(f"✅ No completely sold-out products found for {sel_brand}.")
    else:
        st.warning(
            "⚠️ Variant stock file not loaded. Upload **Product_Variant__product_product_.xlsx** "
            "from Odoo (Products → Export → select Barcode, Internal Reference, Name, Qty On Hand) "
            "to Google Drive and set **GDRIVE_VARIANT_STOCK_ID** at the top of this file."
        )

    out_pool = BytesIO()
    sub_disp.to_excel(out_pool, index=False, engine="openpyxl")
    out_pool.seek(0)
    st.download_button(
        "⬇️ Download Supplier Order as Excel",
        data=out_pool,
        file_name=f"supplier_order_{sel_brand}_{today.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ── Tab 2: Urgent & Reorder ───────────────────────────────────────────────────
with tab2:
    # When display stock is on, show adjusted urgent list as primary
    if show_display:
        needs_action = df_plan[df_plan["_urgency_key_adj"] <= 1].copy()
    else:
        needs_action = df_plan[df_plan["_urgency_key"] <= 1].copy()

    if needs_action.empty:
        st.success("✅ All categories have sufficient stock cover. No urgent reorders needed.")
    else:
        label = "adjusted (display-deducted)" if show_display else "raw"
        st.markdown(f"**{len(needs_action)} category-location combinations need action** *({label})*")
        for _, r in needs_action.iterrows():
            # Use adjusted urgency for card colour when display is on
            uk  = r["_urgency_key_adj"] if show_display else r["_urgency_key"]
            css = "urgent" if uk == 0 else "warning"

            raw_wks  = f"{r['Weeks Cover']:.1f} wks"
            adj_wks  = f"{r['Weeks Cover (Adj)']:.1f} wks"
            sub_cat  = r.get("Sub Category","")
            cat_label = r["Category"]
            if sub_cat and sub_cat not in ("","nan"):
                cat_label = f"{r['Category']} <span style='color:#94a3b8;font-weight:400'>› {sub_cat}</span>"

            src_tag = ('<span class="src-badge" style="background:#dcfce7;color:#166534">real stock</span>'
                       if r["Stock Source"]=="real" else
                       '<span class="src-badge" style="background:#fef3c7;color:#92400e">est stock</span>')
            season_tag = ""
            if r["Season"] != "All-Season" and r["Season"] != CURRENT_SEASON:
                season_tag = (f'<span class="src-badge" style="background:#f1f5f9;color:#475569">'
                               f'{r["Season"]} item — off-season</span>')

            # Display stock badge
            disp_tag = ""
            if show_display and r["Display Stock"] > 0:
                disp_tag = (f'<span class="src-badge" style="background:#ede9fe;color:#5b21b6">'
                             f'🪟 {int(r["Display Stock"])} on display</span>')

            if show_display:
                # 6-cell grid: stock, display, free, rate, weeks(adj), reorder(adj)
                grid = (
                    f'<div style="display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-top:10px">'
                    f'<div><div style="font-size:10px;color:#94a3b8">Est. Stock {src_tag}</div>'
                    f'<div style="font-size:14px;font-weight:600">{int(r["Est. Stock"]):,}</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">🪟 Display</div>'
                    f'<div style="font-size:14px;font-weight:600;color:#7c3aed">{int(r["Display Stock"]):,}</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Free Stock</div>'
                    f'<div style="font-size:14px;font-weight:600;color:#0f766e">{int(r["Free Stock"]):,}</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Weekly Rate</div>'
                    f'<div style="font-size:14px;font-weight:600">{r["Weekly Rate"]:.1f} u/wk</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Weeks Cover <span style="color:#7c3aed">(adj)</span></div>'
                    f'<div style="font-size:14px;font-weight:600;color:{"#dc2626" if uk==0 else "#d97706"}">{adj_wks} <span style="font-size:10px;color:#94a3b8">raw:{raw_wks}</span></div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Reorder Qty <span style="color:#7c3aed">(adj)</span></div>'
                    f'<div style="font-size:14px;font-weight:600;color:#1d4ed8">{int(r["Reorder Qty (Adj)"]):,} <span style="font-size:10px;color:#94a3b8">raw:{int(r["Reorder Qty"]):,}</span></div></div>'
                    f'</div>'
                )
            else:
                weeks_str = f"{r['Weeks Cover']:.1f} weeks" if r['Weeks Cover'] < 99 else "No sales"
                grid = (
                    f'<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-top:10px">'
                    f'<div><div style="font-size:10px;color:#94a3b8">Est. Stock {src_tag}</div>'
                    f'<div style="font-size:15px;font-weight:600">{int(r["Est. Stock"]):,}</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Weekly Rate</div>'
                    f'<div style="font-size:15px;font-weight:600">{r["Weekly Rate"]:.1f} u/wk</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Weeks Cover</div>'
                    f'<div style="font-size:15px;font-weight:600;color:{"#dc2626" if uk==0 else "#d97706"}">{weeks_str}</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Reorder Qty</div>'
                    f'<div style="font-size:15px;font-weight:600;color:#1d4ed8">{int(r["Reorder Qty"]):,} units</div></div>'
                    f'<div><div style="font-size:10px;color:#94a3b8">Est. Value</div>'
                    f'<div style="font-size:15px;font-weight:600">{fmt_npr(r["Est. Value"])}</div></div>'
                    f'</div>'
                )

            urgency_show = r["Urgency (Adj)"] if show_display else r["Urgency"]
            card_html = (
                f'<div class="reorder-card {css}">'
                f'<div style="display:flex;justify-content:space-between;align-items:center">'
                f'<div>'
                f'<span style="font-size:14px;font-weight:600;color:#0f172a">{cat_label}</span>'
                f'<span style="font-size:12px;color:#64748b;margin-left:8px">📍 {r["Location"]}</span>'
                f'{season_tag}{disp_tag}'
                f'</div>'
                f'<span style="font-size:13px;font-weight:600">{urgency_show}</span>'
                f'</div>'
                f'{grid}'
                f'</div>'
            )
            st.markdown(card_html, unsafe_allow_html=True)

with tab3:
    st.markdown("**Full reorder plan — all categories and locations**")
    if show_display:
        st.info("🪟 Display stock is enabled. Each row shows both raw and adjusted (display-deducted) figures side by side.")
        display_cols = [
            "Urgency","Location","Category","Sub Category","Season",
            "Est. Stock","Display Stock","Free Stock","Stock Source",
            "Weekly Rate","Weeks Cover","Weeks Cover (Adj)",
            "Target Stock","Reorder Qty","Reorder Qty (Adj)",
            "Est. Value","Est. Value (Adj)","Urgency (Adj)"
        ]
    else:
        display_cols = [
            "Urgency","Location","Category","Sub Category","Season",
            "Est. Stock","Stock Source","Weekly Rate",
            "Weeks Cover","Target Stock","Reorder Qty","Est. Value"
        ]
    display = df_plan[display_cols].copy()
    display["Stock Source"] = display["Stock Source"].map({"real":"✅ Real","est":"≈ Estimated"})
    for vcol in ["Est. Value","Est. Value (Adj)"]:
        if vcol in display.columns:
            display[vcol] = display[vcol].apply(fmt_npr)
    st.dataframe(display, use_container_width=True, hide_index=True)

    if st.button("⬇️ Download reorder plan as Excel"):
        out = BytesIO()
        export_df = df_plan[display_cols].copy()
        export_df.to_excel(out, index=False, engine="openpyxl")
        out.seek(0)
        st.download_button(
            "📥 Download Excel",
            data=out,
            file_name=f"reorder_plan_{sel_brand}_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab4:
    st.markdown("**Reorder need by location — total units and value**")
    loc_summary = df_plan.groupby("Location").agg(
        Categories=("Category","nunique"),
        Urgent=("_urgency_key", lambda x: (x==0).sum()),
        Reorder_Soon=("_urgency_key", lambda x: (x==1).sum()),
        Total_Units=("Reorder Qty","sum"),
        Total_Value=("Est. Value","sum"),
    ).reset_index()
    loc_summary["_o"] = loc_summary["Location"].apply(
        lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
    loc_summary = loc_summary.sort_values("_o").drop(columns=["_o"])
    loc_summary["Total_Value"] = loc_summary["Total_Value"].apply(fmt_npr)
    loc_summary = loc_summary.rename(columns={
        "Total_Units":"Units to Reorder",
        "Total_Value":"Est. Value",
        "Reorder_Soon":"Reorder Soon",
    })
    st.dataframe(loc_summary, use_container_width=True, hide_index=True)

    # Sub-category breakdown when a parent is selected
    if sel_cat != "All":
        st.markdown(f"**Sub-category breakdown — {sel_cat}**")
        sub_summary = df_plan.groupby("Sub Category").agg(
            Locations=("Location","nunique"),
            Urgent=("_urgency_key", lambda x: (x==0).sum()),
            Reorder_Soon=("_urgency_key", lambda x: (x==1).sum()),
            Total_Units=("Reorder Qty","sum"),
            Total_Value=("Est. Value","sum"),
            Avg_Weeks_Cover=("Weeks Cover","mean"),
        ).reset_index()
        sub_summary = sub_summary.sort_values("Total_Units", ascending=False)
        sub_summary["Total_Value"]      = sub_summary["Total_Value"].apply(fmt_npr)
        sub_summary["Avg_Weeks_Cover"]  = sub_summary["Avg_Weeks_Cover"].round(1)
        sub_summary = sub_summary.rename(columns={
            "Total_Units":"Units to Reorder",
            "Total_Value":"Est. Value",
            "Reorder_Soon":"Reorder Soon",
            "Avg_Weeks_Cover":"Avg Weeks Cover",
        })
        st.dataframe(sub_summary, use_container_width=True, hide_index=True)