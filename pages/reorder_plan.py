import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Salt Fashion — Reorder Plan",
    page_icon="📦", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container { padding: 1.5rem 2rem }

/* Cards */
.card {
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 18px 20px;
    margin-bottom: 10px;
}
.card-urgent  { border-left: 5px solid #dc2626 }
.card-warning { border-left: 5px solid #f59e0b }
.card-ok      { border-left: 5px solid #16a34a }

/* KPI boxes */
.kpi {
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 16px 12px;
    text-align: center;
}
.kpi-adj { border-color: #7c3aed; border-width: 2px }
.kpi-val { font-size: 28px; font-weight: 700; margin: 0; line-height: 1.2 }
.kpi-lbl { font-size: 11px; color: #6b7280; margin: 5px 0 0 }

/* Section divider */
.section-label {
    font-size: 12px;
    font-weight: 600;
    color: #64748b;
    letter-spacing: .06em;
    text-transform: uppercase;
    margin: 20px 0 8px;
    padding-bottom: 4px;
    border-bottom: 1px solid #e2e8f0;
}

/* Inline badges */
.badge {
    display: inline-block;
    padding: 2px 9px;
    border-radius: 9px;
    font-size: 10px;
    font-weight: 600;
    margin-left: 6px;
    vertical-align: middle;
}
.badge-real    { background: #dcfce7; color: #166534 }
.badge-est     { background: #fef3c7; color: #92400e }
.badge-display { background: #ede9fe; color: #5b21b6 }
.badge-free    { background: #ecfeff; color: #155e75 }
.badge-offseason { background: #f1f5f9; color: #475569 }
</style>
""", unsafe_allow_html=True)

# ── Google Drive IDs ──────────────────────────────────────────────────────────
GDRIVE_MAIN_ID      = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_POS_ID       = "1YcW30p_dUfeeaQj-XXmGhMHP0ldAM32X"
GDRIVE_LOCSTK_ID    = "1zgTBhh7vOTjxEIz-LO3YSM-TXJeDUrBT"
GDRIVE_RECENTCAT_ID = "1EMEw10v7zEwsMzrocJWCjkyRfy14LaIM"

# ── Constants ─────────────────────────────────────────────────────────────────
MIN_REORDER_QTY = 5

# ── Real furniture-based display capacity per store per category ──────────────
# Source: Salt_furniture.xlsx — floor display furniture only.
# Excludes "4 Step rack" (backstore storage) and "Hanging fix rail on backstore".
# Methodology: count × capacity_each, distributed across furniture categories.
# These reflect Summer capacity (maximum floor display).
# Winter note from file: "quantities will drastically be impacted by season."
FURNITURE_DISPLAY = {
    "Lazimpat": {
        "Basic Top": 116, "Denim Pant": 126, "Dress": 66,
        "Formal Pant": 137, "Jeans": 126, "Leggings": 74,
        "Shorts": 23, "Skirts": 87, "Skort": 87,
        "T-Shirts": 8, "Tops": 116,
    },
    "Kumaripati": {
        "Basic Top": 105, "Denim Pant": 154, "Dress": 88,
        "Formal Pant": 154, "Jeans": 154, "Leggings": 154,
        "Shorts": 88, "Skirts": 171, "Skort": 171,
        "T-Shirts": 66, "Tops": 105,
    },
    "Baneshwor": {
        "Basic Top": 258, "Denim Pant": 163, "Dress": 107,
        "Formal Pant": 179, "Jeans": 163, "Leggings": 163,
        "Shorts": 71, "Skirts": 173, "Skort": 173,
        "T-Shirts": 50, "Tops": 258,
    },
    "Chitwan": {
        # Recalculated using floor display rails only (30 standard Long Rails × 50).
        # The 90 "Long Rail large" rows in the Excel are backstore overflow, not floor display
        # (120 rails in a single store is physically impossible as floor display).
        # Floor: 30 rails×50 + 8 T-hangers×12 = 1,596 (hanging) + 20 tables×30 + 15 tables×20 = 900 (folded)
        "Basic Top": 266, "Tops": 266,
        "Dress": 266,
        "Denim Pant": 416, "Jeans": 416, "Leggings": 416, "Formal Pant": 416,
        "Shorts": 150, "Skirts": 150, "Skort": 150,
        "T-Shirts": 100,
    },
    "Pokhara": {
        "Basic Top": 153, "Denim Pant": 111, "Dress": 102,
        "Formal Pant": 112, "Jeans": 111, "Leggings": 111,
        "Shorts": 63, "Skirts": 155, "Skort": 155,
        "T-Shirts": 7, "Tops": 153,
    },
}
FURNITURE_DISPLAY_FALLBACK = 0  # no display data = no deduction

LOCATION_ORDER = [
    "Baneshwor", "Lazimpat", "Kumaripati", "Chitwan", "Pokhara",
    "Online", "Baneshwor Lush", "Chitwan Lush", "Pokhara Lush",
]

STORE_NAME_FIX = {
    "lazimpat": "Lazimpat", "baneshwor": "Baneshwor",
    "chitwan": "Chitwan", "kumaripati": "Kumaripati",
    "pokhara": "Pokhara", "online": "Online",
    "main warehouse": "Main Warehouse",
}

SKIP_PARTS = {"All", "Saleable", "PoS", ""}

WINTER_CATS = {"Coat","Jacket","Sweater","Cardigan","Sweatshirt","Hoodie",
               "Waistcoat","Pajamas Set","Vest","Knitted","Fur Regular","Wool"}
SUMMER_CATS = {"T-Shirts","Shorts","Tops","Dress","Co-Ord Set","Tank Top",
               "Swim Wear","Skirt","Skort","Sundress","Basic Top"}


def split_cat(raw):
    parts = [p.strip() for p in str(raw).split("/") if p.strip() not in SKIP_PARTS]
    if not parts: return "", ""
    if len(parts) == 1: return parts[0], ""
    return parts[0], parts[1]

def norm_store(name):
    return STORE_NAME_FIX.get(str(name).strip().lower(), str(name).strip())

def cat_season(cat):
    if cat in WINTER_CATS: return "Winter"
    if cat in SUMMER_CATS: return "Summer"
    return "All-Season"

def current_season():
    m = pd.Timestamp.today().month
    if m in (11,12,1,2): return "Winter"
    if m in (5,6,7,8,9): return "Summer"
    return "Transition"

CURRENT_SEASON = current_season()

def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

# ── Google Drive loader ───────────────────────────────────────────────────────
def gdrive_bytes(file_id):
    if not file_id: return None, None
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
        buf.seek(0)
        return buf, None
    except Exception as e:
        return None, str(e)

# ── Data loaders ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=600, show_spinner=False)
def load_products():
    buf, _ = gdrive_bytes(GDRIVE_MAIN_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl")
        except: pass
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base + r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"), reverse=True) if Path(d).exists() else []
            if files:
                df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl")
                break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    for col in ["Sales Price", "On Hand Qty", "Total Units Sold"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Category" in df.columns:
        has_sub   = "Sub Category" in df.columns
        has_slash = df["Category"].str.contains("/", na=False).any()
        if not has_sub or has_slash:
            split = df["Category"].apply(split_cat)
            df["Category"]     = split.apply(lambda x: x[0])
            df["Sub Category"] = split.apply(lambda x: x[1])
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
        files = sorted(
            Path(r"C:\Users\Legion\Desktop\odoo_export\exports").glob("pos_analysis_*.xlsx"),
            reverse=True
        ) if Path(r"C:\Users\Legion\Desktop\odoo_export\exports").exists() else []
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
        except Exception as e: err = str(e)
    if df is None:
        files = sorted(
            Path(r"C:\Users\Legion\Desktop\odoo_export\exports").glob("location_stock_*.xlsx"),
            reverse=True
        ) if Path(r"C:\Users\Legion\Desktop\odoo_export\exports").exists() else []
        if files:
            try:
                df = pd.read_excel(files[0], sheet_name="Store x Category", engine="openpyxl")
                err = None
            except Exception as e: err = str(e)
    if df is None or df.empty:
        return None, set(), err
    df.columns = [str(c).strip() for c in df.columns]
    cat_col    = df.columns[0]
    store_cols = [c for c in df.columns if c != cat_col]
    long_rows      = []
    covered_stores = set()
    for _, row in df.iterrows():
        cat = str(row[cat_col]).strip()
        if not cat or cat.lower() in ("nan", ""): continue
        for store in store_cols:
            covered_stores.add(norm_store(store))
            qty = row[store]
            qty_val = 0.0 if pd.isna(qty) else float(qty)
            long_rows.append({
                "Location":    norm_store(store),
                "Category":    cat,
                "On_Hand_Real": max(0.0, qty_val),
            })
    if not long_rows: return None, set(), err
    return pd.DataFrame(long_rows), covered_stores, None

@st.cache_data(ttl=600, show_spinner=False)
def load_recent_cat_sales():
    buf, err = gdrive_bytes(GDRIVE_RECENTCAT_ID)
    df = None
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Recent Category Sales", engine="openpyxl")
        except Exception as e: err = str(e)
    if df is None:
        files = sorted(
            Path(r"C:\Users\Legion\Desktop\odoo_export\exports").glob("category_sales_recent_*.xlsx"),
            reverse=True
        ) if Path(r"C:\Users\Legion\Desktop\odoo_export\exports").exists() else []
        if files:
            try:
                df = pd.read_excel(files[0], sheet_name="Recent Category Sales", engine="openpyxl")
                err = None
            except Exception as e: err = str(e)
    if df is None or df.empty: return None, err
    df.columns = [str(c).strip() for c in df.columns]
    df["Location"] = df["Location"].apply(norm_store)
    return df, None

# ── Display stock calculator ──────────────────────────────────────────────────
def calc_display(store, cat, user_overrides=None):
    """
    Returns floor display units for (store, category) from real furniture data.
    user_overrides: dict of {(store, cat): int} for manual adjustments.
    """
    if user_overrides and (store, cat) in user_overrides:
        return user_overrides[(store, cat)]
    return FURNITURE_DISPLAY.get(store, {}).get(cat, FURNITURE_DISPLAY_FALLBACK)

# ── Load data ─────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod                       = load_products()
    df_pos                        = load_pos()
    df_locstk, covered_stores, _  = load_location_stock()
    df_recent_cat, _              = load_recent_cat_sales()

if df_prod is None or df_pos is None:
    st.error("Could not load data. Make sure product and POS files are on Google Drive.")
    st.stop()

USING_REAL_STOCK     = df_locstk is not None and not df_locstk.empty
USING_SEASONAL_RATES = df_recent_cat is not None and not df_recent_cat.empty

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📦 Reorder Planner")
    st.markdown("---")

    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("nan","True","False","None","")])
    sel_brand = st.selectbox("Brand", brands)

    st.markdown("---")
    st.markdown("**Planning settings**")
    target_weeks   = st.slider("Target weeks of cover", 2, 12, 4)
    lookback_weeks = st.slider("Sales lookback (weeks)", 2, 12, 4)
    min_weekly_rate = st.number_input("Min weekly rate to show", 0, 50, 1)

    st.markdown("---")
    locations = ["All"] + [l for l in LOCATION_ORDER if l in df_pos["Location"].unique()]
    sel_loc   = st.selectbox("Location", locations)

    prod_brand = df_prod[df_prod["Brand"] == sel_brand]
    parent_cats = sorted([c for c in prod_brand["Category"].unique()
                          if c and c not in ("nan","")])
    sel_cat = st.selectbox("Category", ["All"] + parent_cats)

    sel_sub_cat = "All"
    if sel_cat != "All" and "Sub Category" in prod_brand.columns:
        sub_cats = sorted([s for s in
                           prod_brand[prod_brand["Category"] == sel_cat]["Sub Category"].unique()
                           if s and s not in ("nan","")])
        if sub_cats:
            sel_sub_cat = st.selectbox("Sub-category", ["All"] + sub_cats)

    sel_season = st.selectbox("Season", ["All","Summer","Winter","All-Season"])

    # ── Display stock ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**🪟 Display Stock**")
    st.caption(
        "Units permanently on the shop floor (from real furniture counts). "
        "Deducting these shows the true free buffer stock."
    )
    show_display = st.toggle("Enable display stock deduction", value=True)

    user_overrides = {}
    if show_display:
        with st.expander("📋 View furniture display capacity", expanded=False):
            st.caption("Calculated from Salt_furniture.xlsx — floor display furniture only.")
            if sel_loc != "All":
                store_show = [sel_loc]
            else:
                store_show = [s for s in ["Lazimpat","Kumaripati","Baneshwor","Chitwan","Pokhara"]
                              if s in FURNITURE_DISPLAY]
            for s in store_show:
                st.markdown(f"**{s}**")
                fd = FURNITURE_DISPLAY.get(s, {})
                for cat in sorted(fd.keys()):
                    if fd[cat] > 0:
                        st.caption(f"  {cat}: {fd[cat]} units")

        with st.expander("✏️ Override a value (optional)", expanded=False):
            st.caption("Use this to correct a specific store/category if the furniture data is outdated.")
            ov_store = st.selectbox("Store", list(FURNITURE_DISPLAY.keys()), key="ov_store")
            ov_cats  = sorted(FURNITURE_DISPLAY.get(ov_store, {}).keys())
            if ov_cats:
                ov_cat   = st.selectbox("Category", ov_cats, key="ov_cat")
                current  = FURNITURE_DISPLAY.get(ov_store, {}).get(ov_cat, 0)
                ov_val   = st.number_input(
                    f"Display units (current: {current})",
                    min_value=0, max_value=2000, value=current, step=1, key="ov_val"
                )
                if ov_val != current:
                    user_overrides[(ov_store, ov_cat)] = ov_val

    st.markdown("---")
    st.success("✅ Real stock" if USING_REAL_STOCK else "⚠️ Estimated stock")
    st.success("✅ Seasonal rates" if USING_SEASONAL_RATES else "⚠️ All-time rates")

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ── Build weekly rates from POS ───────────────────────────────────────────────
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
).reset_index()
pos_agg["Weekly_Rate"]    = pos_agg["Total_Units"] / lookback_weeks
pos_agg["Weekly_Revenue"] = pos_agg["Total_Revenue"] / lookback_weeks

total_units_all = pos_agg["Total_Units"].sum()
pos_agg["Share"] = pos_agg["Total_Units"] / total_units_all if total_units_all > 0 else 0

# ── Stock per category ────────────────────────────────────────────────────────
prod_brand = df_prod[df_prod["Brand"] == sel_brand].copy()
group_cols = ["Category", "Sub Category"] if "Sub Category" in prod_brand.columns else ["Category"]

cat_stock = prod_brand.groupby(group_cols).agg(
    On_Hand=("On Hand Qty", "sum"),
    Total_Sold=("Total Units Sold", "sum"),
    Avg_Price=("Sales Price", "mean"),
).reset_index()

total_sold_all = cat_stock["Total_Sold"].sum()

real_stock_map = {}
if USING_REAL_STOCK:
    for _, r in df_locstk.iterrows():
        real_stock_map[(r["Location"], r["Category"])] = r["On_Hand_Real"]

recent_rate_map = {}
if USING_SEASONAL_RATES:
    for _, r in df_recent_cat.iterrows():
        recent_rate_map[(r["Location"], r["Category"])] = float(r.get("Weekly Rate", 0) or 0)

# ── Build plan rows ───────────────────────────────────────────────────────────
rows = []
for _, loc_row in pos_agg.iterrows():
    loc   = loc_row["Location"]
    share = loc_row["Share"]
    loc_wr = loc_row["Weekly_Rate"]

    for _, cat_row in cat_stock.iterrows():
        cat     = cat_row["Category"]
        sub_cat = cat_row.get("Sub Category", "") if "Sub Category" in cat_row.index else ""
        if not cat or cat in ("nan","","All"): continue

        # Stock
        real_val = real_stock_map.get((loc, cat))
        if real_val is not None:
            sub_sold_total = cat_stock[cat_stock["Category"] == cat]["Total_Sold"].sum()
            if sub_sold_total > 0:
                sub_frac = cat_row["Total_Sold"] / sub_sold_total
            else:
                sub_count = len(cat_stock[cat_stock["Category"] == cat])
                sub_frac  = 1.0 / sub_count if sub_count > 0 else 1.0
            est_stock    = max(0, real_val * sub_frac)
            stock_source = "real"
        elif loc in covered_stores:
            est_stock    = 0
            stock_source = "real"
        else:
            est_stock    = cat_row["On_Hand"] * share
            stock_source = "est"

        # Rate
        cat_share = cat_row["Total_Sold"] / total_sold_all if total_sold_all > 0 else 0
        if (loc, cat) in recent_rate_map:
            sub_sold_total = cat_stock[cat_stock["Category"] == cat]["Total_Sold"].sum()
            sub_frac_r = (cat_row["Total_Sold"] / sub_sold_total
                         if sub_sold_total > 0
                         else 1.0 / max(1, len(cat_stock[cat_stock["Category"] == cat])))
            weekly_rate = recent_rate_map[(loc, cat)] * sub_frac_r
        else:
            weekly_rate = loc_wr * cat_share

        if weekly_rate < min_weekly_rate: continue

        daily_rate   = weekly_rate / 7
        weeks_cover  = (est_stock / daily_rate / 7) if daily_rate > 0 else 999
        target_stock = target_weeks * weekly_rate
        reorder_qty  = max(0, round(target_stock - est_stock))

        # Display stock
        # The furniture data holds the TOTAL display capacity for the parent category
        # (e.g. Lazimpat Tops = 116 total units across ALL Tops sub-categories).
        # We must split this by the same sub_frac used for stock, so sub-categories
        # share the parent display budget rather than each claiming the full total.
        if show_display:
            parent_display = calc_display(loc, cat, user_overrides)
            # Determine the sub-fraction: same logic as the stock split above
            _sub_sold_total = cat_stock[cat_stock["Category"] == cat]["Total_Sold"].sum()
            if _sub_sold_total > 0:
                _display_frac = cat_row["Total_Sold"] / _sub_sold_total
            else:
                _n_subs = len(cat_stock[cat_stock["Category"] == cat])
                _display_frac = 1.0 / _n_subs if _n_subs > 0 else 1.0
            display_units = round(parent_display * _display_frac)
        else:
            display_units = 0
        # Cap display at est_stock — can't display more than you own.
        # When display >= est_stock it means all stock is on the floor,
        # nothing is in the back room. Free stock = 0 is correct in that case.
        display_units  = min(display_units, round(est_stock))
        free_stock     = max(0, est_stock - display_units)
        weeks_cover_adj = (free_stock / daily_rate / 7) if daily_rate > 0 else 999
        reorder_qty_adj = max(0, round(target_stock - free_stock))

        # Urgency (raw)
        if weeks_cover <= 1:
            urgency, uk = "🔴 Urgent", 0
        elif weeks_cover < target_weeks and reorder_qty >= MIN_REORDER_QTY:
            urgency, uk = "🟡 Reorder Soon", 1
        else:
            urgency, uk = "🟢 OK", 2

        # Urgency (adjusted)
        if weeks_cover_adj <= 1:
            urgency_adj, uk_adj = "🔴 Urgent", 0
        elif weeks_cover_adj < target_weeks and reorder_qty_adj >= MIN_REORDER_QTY:
            urgency_adj, uk_adj = "🟡 Reorder Soon", 1
        else:
            urgency_adj, uk_adj = "🟢 OK", 2

        rows.append({
            "Location":           loc,
            "Category":           cat,
            "Sub Category":       sub_cat,
            "Season":             cat_season(cat),
            # Raw
            "Est. Stock":         round(est_stock),
            "Stock Source":       stock_source,
            "Weekly Rate":        round(weekly_rate, 1),
            "Weeks Cover":        round(weeks_cover, 1),
            "Target Stock":       round(target_stock),
            "Reorder Qty":        reorder_qty,
            "Est. Value":         round(reorder_qty * cat_row["Avg_Price"]),
            "Urgency":            urgency,
            "_uk":                uk,
            # Adjusted
            "Display Stock":      display_units,
            "Free Stock":         round(free_stock),
            "Weeks Cover (Adj)":  round(weeks_cover_adj, 1),
            "Reorder Qty (Adj)":  reorder_qty_adj,
            "Est. Value (Adj)":   round(reorder_qty_adj * cat_row["Avg_Price"]),
            "Urgency (Adj)":      urgency_adj,
            "_uk_adj":            uk_adj,
        })

df_plan = pd.DataFrame(rows)

if df_plan.empty:
    st.warning("No reorder data. Check that POS and product data are loaded correctly.")
    st.stop()

# Filters
if sel_loc     != "All": df_plan = df_plan[df_plan["Location"]     == sel_loc]
if sel_cat     != "All": df_plan = df_plan[df_plan["Category"]     == sel_cat]
if sel_sub_cat != "All": df_plan = df_plan[df_plan["Sub Category"] == sel_sub_cat]
if sel_season  != "All": df_plan = df_plan[df_plan["Season"]       == sel_season]

# Sort by adjusted urgency when display is on, raw urgency otherwise
sort_key = "_uk_adj" if show_display else "_uk"
df_plan  = df_plan.sort_values([sort_key, "Reorder Qty (Adj)" if show_display else "Reorder Qty"],
                                ascending=[True, False])

# ── Page header ───────────────────────────────────────────────────────────────
st.title("📦 Reorder Planner")
stock_badge = (
    '<span class="badge badge-real">✅ Real stock</span>'
    if USING_REAL_STOCK else
    '<span class="badge badge-est">≈ Estimated stock</span>'
)
filter_label = sel_cat if sel_cat != "All" else "All Categories"
if sel_sub_cat != "All": filter_label += f" › {sel_sub_cat}"

st.markdown(
    f"{sel_brand} &nbsp;·&nbsp; {filter_label} &nbsp;·&nbsp; "
    f"{target_weeks}-week target &nbsp;·&nbsp; {lookback_weeks}-week lookback &nbsp;·&nbsp; "
    f"{today.strftime('%B %d, %Y')} {stock_badge}",
    unsafe_allow_html=True,
)
st.markdown("<br>", unsafe_allow_html=True)

# ── KPI section ───────────────────────────────────────────────────────────────
# Raw KPIs
urgent_n     = (df_plan["_uk"] == 0).sum()
reorder_n    = (df_plan["_uk"] <= 1).sum()
units_raw    = int(df_plan["Reorder Qty"].sum())
value_raw    = df_plan["Est. Value"].sum()

# Adjusted KPIs
urgent_n_adj  = (df_plan["_uk_adj"] == 0).sum()
reorder_n_adj = (df_plan["_uk_adj"] <= 1).sum()
units_adj     = int(df_plan["Reorder Qty (Adj)"].sum())
value_adj     = df_plan["Est. Value (Adj)"].sum()

if show_display:
    # Two rows: raw on top, adjusted (purple) below
    st.markdown('<p class="section-label">Without display stock (raw)</p>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    for col, val, lbl, clr in [
        (c1, f"🔴 {urgent_n}",          "Urgent (under 1 week)",    "#dc2626"),
        (c2, f"🟡 {reorder_n-urgent_n}", "Reorder Soon",             "#d97706"),
        (c3, f"{units_raw:,}",          "Units to Order",            "#1d4ed8"),
        (c4, fmt_npr(value_raw),        "Estimated Value",           "#374151"),
    ]:
        with col:
            st.markdown(
                f'<div class="kpi">'
                f'<p class="kpi-val" style="color:{clr}">{val}</p>'
                f'<p class="kpi-lbl">{lbl}</p></div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<p class="section-label" style="color:#7c3aed;border-color:#c4b5fd">'
        '🪟 After display stock deduction (adjusted — supervisor view)</p>',
        unsafe_allow_html=True,
    )
    d1, d2, d3, d4 = st.columns(4)
    for col, val, lbl, clr in [
        (d1, f"🔴 {urgent_n_adj}",              "Urgent (under 1 week)",    "#dc2626"),
        (d2, f"🟡 {reorder_n_adj-urgent_n_adj}", "Reorder Soon",             "#d97706"),
        (d3, f"{units_adj:,}",                  "Units to Order (Adj)",     "#1d4ed8"),
        (d4, fmt_npr(value_adj),                "Estimated Value (Adj)",    "#374151"),
    ]:
        with col:
            st.markdown(
                f'<div class="kpi kpi-adj">'
                f'<p class="kpi-val" style="color:{clr}">{val}</p>'
                f'<p class="kpi-lbl">{lbl}</p></div>',
                unsafe_allow_html=True,
            )
else:
    c1, c2, c3, c4 = st.columns(4)
    for col, val, lbl, clr in [
        (c1, f"🔴 {urgent_n}",          "Urgent — under 1 week",    "#dc2626"),
        (c2, f"🟡 {reorder_n-urgent_n}", "Reorder Soon",             "#d97706"),
        (c3, f"{units_raw:,}",          "Total Units to Order",      "#1d4ed8"),
        (c4, fmt_npr(value_raw),        "Estimated Value",           "#374151"),
    ]:
        with col:
            st.markdown(
                f'<div class="kpi">'
                f'<p class="kpi-val" style="color:{clr}">{val}</p>'
                f'<p class="kpi-lbl">{lbl}</p></div>',
                unsafe_allow_html=True,
            )

st.markdown("<br>", unsafe_allow_html=True)

# ── Display stock explainer (shown when enabled) ──────────────────────────────
if show_display:
    with st.expander("ℹ️ How display stock works", expanded=False):
        st.markdown("""
**Display stock** is the number of units permanently on the shop floor as display pieces.
They cannot be pulled to replenish — they must stay on the floor to keep shelves looking full.

| Term | Meaning |
|------|---------|
| **Est. Stock** | Total on-hand quantity from Odoo warehouse |
| **Display Stock** | Units locked on the floor (from real furniture counts) |
| **Free Stock** | Est. Stock − Display Stock → the true available buffer |
| **Order Qty** | Units needed to bring Free Stock up to the target buffer |

Display figures are calculated from **Salt_furniture.xlsx** — each store's actual
furniture pieces × their capacity, for floor display furniture only.
Backstore racks and rails are excluded (those are already counted as buffer stock).

> *Source: Long Rails, T-Hangers, Tables, Square Rails — all floor display furniture.*
""")
        # Show display capacity table for current filter
        import pandas as _pd
        rows_ex = []
        for s in ["Lazimpat","Kumaripati","Baneshwor","Chitwan","Pokhara"]:
            fd = FURNITURE_DISPLAY.get(s, {})
            show_cats = [sel_cat] if sel_cat != "All" else sorted(fd.keys())
            for cat in show_cats:
                disp = fd.get(cat, 0)
                rows_ex.append({"Store": s, "Category": cat, "Display Units": disp,
                                 "Source": "Furniture data"})
        if rows_ex:
            st.dataframe(_pd.DataFrame(rows_ex), use_container_width=True, hide_index=True)

st.markdown("---")

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["🔴 Urgent & Reorder", "📊 Full Plan", "📍 By Location"])

# ── Tab 1: Action list ────────────────────────────────────────────────────────
with tab1:
    action_key = "_uk_adj" if show_display else "_uk"
    needs_action = df_plan[df_plan[action_key] <= 1].copy()

    if needs_action.empty:
        st.success("✅ All categories have sufficient cover — no urgent reorders.")
    else:
        view_label = "adjusted (after display deduction)" if show_display else "raw"
        st.markdown(f"**{len(needs_action)} items need action** *({view_label})*")
        st.markdown("")

        for _, r in needs_action.iterrows():
            uk   = r["_uk_adj"] if show_display else r["_uk"]
            css  = "card-urgent" if uk == 0 else "card-warning"

            # Category label with sub-category
            cat_label = r["Category"]
            sub       = r.get("Sub Category","")
            if sub and sub not in ("","nan"):
                cat_label = f'{r["Category"]} <span style="color:#94a3b8;font-weight:400">› {sub}</span>'

            # Badges
            src_badge = (
                '<span class="badge badge-real">real stock</span>'
                if r["Stock Source"] == "real" else
                '<span class="badge badge-est">est. stock</span>'
            )
            season_badge = ""
            if r["Season"] != "All-Season" and r["Season"] != CURRENT_SEASON:
                season_badge = f'<span class="badge badge-offseason">{r["Season"]} — off-season</span>'

            urgency_show = r["Urgency (Adj)"] if show_display else r["Urgency"]

            if show_display:
                # 6-column grid showing the full deduction story
                grid = "".join([
                    '<div style="display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-top:12px">',
                    # Est Stock
                    f'<div><div style="font-size:10px;color:#94a3b8">Est. Stock {src_badge}</div>'
                    f'<div style="font-size:15px;font-weight:700">{int(r["Est. Stock"]):,}</div></div>',
                    # Display stock
                    f'<div><div style="font-size:10px;color:#7c3aed">🪟 Display Stock</div>'
                    f'<div style="font-size:15px;font-weight:700;color:#7c3aed">−{int(r["Display Stock"]):,}</div></div>',
                    # Free stock
                    f'<div><div style="font-size:10px;color:#0f766e;font-weight:600">= Free Stock</div>'
                    f'<div style="font-size:15px;font-weight:700;color:#0f766e">{int(r["Free Stock"]):,}</div></div>',
                    # Weekly rate
                    f'<div><div style="font-size:10px;color:#94a3b8">Weekly Rate</div>'
                    f'<div style="font-size:15px;font-weight:700">{r["Weekly Rate"]:.1f} u/wk</div></div>',
                    # Weeks cover adj
                    f'<div><div style="font-size:10px;color:#7c3aed">Weeks Cover (adj)</div>'
                    f'<div style="font-size:15px;font-weight:700;color:{"#dc2626" if uk==0 else "#d97706"}">'
                    + ("Out of stock" if r["Weeks Cover (Adj)"] == 0 else f"{r['Weeks Cover (Adj)']:.1f} wks") +
                    f'</div></div>',
                    # Reorder adj
                    f'<div><div style="font-size:10px;color:#7c3aed">Order Qty</div>'
                    f'<div style="font-size:15px;font-weight:700;color:#1d4ed8">'
                    f'{int(r["Reorder Qty (Adj)"]):,} units'
                    f'</div></div>',
                    '</div>',
                ])
            else:
                wks_str = f'{r["Weeks Cover"]:.1f} weeks' if r["Weeks Cover"] < 99 else "No sales"
                grid = "".join([
                    '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-top:12px">',
                    f'<div><div style="font-size:10px;color:#94a3b8">Est. Stock {src_badge}</div>'
                    f'<div style="font-size:15px;font-weight:700">{int(r["Est. Stock"]):,}</div></div>',
                    f'<div><div style="font-size:10px;color:#94a3b8">Weekly Rate</div>'
                    f'<div style="font-size:15px;font-weight:700">{r["Weekly Rate"]:.1f} u/wk</div></div>',
                    f'<div><div style="font-size:10px;color:#94a3b8">Weeks Cover</div>'
                    f'<div style="font-size:15px;font-weight:700;color:{"#dc2626" if uk==0 else "#d97706"}">{wks_str}</div></div>',
                    f'<div><div style="font-size:10px;color:#94a3b8">Reorder Qty</div>'
                    f'<div style="font-size:15px;font-weight:700;color:#1d4ed8">{int(r["Reorder Qty"]):,}</div></div>',
                    f'<div><div style="font-size:10px;color:#94a3b8">Est. Value</div>'
                    f'<div style="font-size:15px;font-weight:700">{fmt_npr(r["Est. Value"])}</div></div>',
                    '</div>',
                ])

            card_html = "".join([
                f'<div class="card {css}">',
                '<div style="display:flex;justify-content:space-between;align-items:center">',
                f'<div><span style="font-size:15px;font-weight:700;color:#0f172a">{cat_label}</span>',
                f'<span style="font-size:12px;color:#64748b;margin-left:10px">📍 {r["Location"]}</span>',
                season_badge, '</div>',
                f'<span style="font-size:13px;font-weight:700">{urgency_show}</span>',
                '</div>',
                grid,
                '</div>',
            ])
            st.markdown(card_html, unsafe_allow_html=True)

# ── Tab 2: Full table ─────────────────────────────────────────────────────────
with tab2:
    # Always show the clean supervisor view — adjusted figures only when display is on.
    # "Est. Stock" is kept as context; raw Reorder Qty / Weeks Cover are hidden to avoid confusion.
    if show_display:
        # Build a clean renamed copy so column headers read naturally
        tbl = df_plan[[
            "Urgency (Adj)", "Location", "Category", "Sub Category", "Season",
            "Est. Stock", "Display Stock", "Free Stock", "Stock Source",
            "Weekly Rate", "Weeks Cover (Adj)", "Target Stock",
            "Reorder Qty (Adj)", "Est. Value (Adj)",
        ]].copy().rename(columns={
            "Urgency (Adj)":    "Status",
            "Free Stock":       "Free Stock ✓",
            "Weeks Cover (Adj)":"Weeks Cover",
            "Reorder Qty (Adj)":"Order Qty",
            "Est. Value (Adj)": "Est. Value",
        })
        tbl["Stock Source"]  = tbl["Stock Source"].map({"real":"✅ Real","est":"≈ Est."})
        tbl["Est. Value"]    = tbl["Est. Value"].apply(fmt_npr)
        tbl["Weeks Cover"]   = tbl["Weeks Cover"].apply(
            lambda x: "Out of stock" if x == 0 else f"{x:.1f} wks"
        )
        st.caption(
            "🪟 Display stock deducted — **Order Qty** and **Weeks Cover** reflect "
            "only the free (non-display) stock. Est. Stock shown for reference."
        )
    else:
        tbl = df_plan[[
            "Urgency", "Location", "Category", "Sub Category", "Season",
            "Est. Stock", "Stock Source", "Weekly Rate",
            "Weeks Cover", "Target Stock", "Reorder Qty", "Est. Value",
        ]].copy().rename(columns={"Urgency": "Status", "Reorder Qty": "Order Qty"})
        tbl["Stock Source"] = tbl["Stock Source"].map({"real":"✅ Real","est":"≈ Est."})
        tbl["Est. Value"]   = tbl["Est. Value"].apply(fmt_npr)
        tbl["Weeks Cover"]  = tbl["Weeks Cover"].apply(
            lambda x: "Out of stock" if x == 0 else f"{x:.1f} wks"
        )

    st.dataframe(tbl, use_container_width=True, hide_index=True)

    out = BytesIO()
    tbl.to_excel(out, index=False, engine="openpyxl")
    out.seek(0)
    st.download_button(
        "⬇️ Download as Excel",
        data=out,
        file_name=f"reorder_plan_{sel_brand}_{today.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ── Tab 3: By location ────────────────────────────────────────────────────────
with tab3:
    uk_col     = "_uk_adj" if show_display else "_uk"
    qty_col_s  = "Reorder Qty (Adj)" if show_display else "Reorder Qty"
    val_col_s  = "Est. Value (Adj)"  if show_display else "Est. Value"

    loc_summ = df_plan.groupby("Location").agg(
        Categories=("Category", "nunique"),
        Urgent=(uk_col, lambda x: (x == 0).sum()),
        Reorder_Soon=(uk_col, lambda x: (x == 1).sum()),
        Units=(qty_col_s, "sum"),
        Value=(val_col_s, "sum"),
    ).reset_index()
    loc_summ["_o"] = loc_summ["Location"].apply(
        lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
    loc_summ = loc_summ.sort_values("_o").drop(columns=["_o"])
    loc_summ["Value"] = loc_summ["Value"].apply(fmt_npr)
    loc_summ = loc_summ.rename(columns={
        "Units": "Units to Order",
        "Value": "Est. Value",
        "Reorder_Soon": "Reorder Soon",
    })
    st.dataframe(loc_summ, use_container_width=True, hide_index=True)

    if sel_cat != "All":
        st.markdown(f"**Sub-category breakdown — {sel_cat}**")
        sub_summ = df_plan.groupby("Sub Category").agg(
            Locations=("Location", "nunique"),
            Urgent=(uk_col, lambda x: (x == 0).sum()),
            Reorder_Soon=(uk_col, lambda x: (x == 1).sum()),
            Units=(qty_col_s, "sum"),
            Value=(val_col_s, "sum"),
            Avg_Wks_Cover=("Weeks Cover (Adj)" if show_display else "Weeks Cover", "mean"),
        ).reset_index()
        sub_summ = sub_summ.sort_values("Units", ascending=False)
        sub_summ["Value"]         = sub_summ["Value"].apply(fmt_npr)
        sub_summ["Avg_Wks_Cover"] = sub_summ["Avg_Wks_Cover"].round(1)
        sub_summ = sub_summ.rename(columns={
            "Units": "Units to Order", "Value": "Est. Value",
            "Reorder_Soon": "Reorder Soon", "Avg_Wks_Cover": "Avg Weeks Cover",
        })
        st.dataframe(sub_summ, use_container_width=True, hide_index=True)