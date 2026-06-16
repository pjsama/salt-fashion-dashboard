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
GDRIVE_LOCSTK_ID    = "1zgTBhh7vOTjxEIz-LO3YSM-TXJeDUrBT"   # location_stock_*.xlsx
GDRIVE_RECENTCAT_ID = "1EMEw10v7zEwsMzrocJWCjkyRfy14LaIM"   # ← fill in once you upload category_sales_recent_*.xlsx to Drive

LOCATION_ORDER = ["Baneshwor","Lazimpat","Kumaripati","Chitwan","Pokhara","Online",
                  "Baneshwor Lush","Chitwan Lush","Pokhara Lush"]

# Normalise store names from the location-stock export to match POS location names
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
    parts = [p.strip() for p in str(raw).split("/") if p.strip() not in SKIP_PARTS]
    if not parts: return ""
    return parts[-2] if len(parts) >= 2 else parts[0]

def norm_store(name):
    return STORE_NAME_FIX.get(str(name).strip().lower(), str(name).strip())

# ── Seasonal category classification ──────────────────────────────────────────
# Helps flag winter-heavy categories during summer months (and vice versa) so
# buying decisions aren't driven purely by short-term stockout math on
# off-season leftovers.
WINTER_CATEGORIES = {
    "Coat","Jacket","Sweater","Cardigan","Sweatshirt","Hoodie","Waistcoat",
    "Pajamas Set","Vest","Knitted","Fur Regular","Wool",
}
SUMMER_CATEGORIES = {
    "T-Shirts","Shorts","Tops","Dress","Co-Ord Set","Tank Top","Swim Wear",
    "Skirt","Skort","Sundress","Basic Top",
}

def season_for_month(month):
    # Nepal: roughly Nov-Feb = winter, Apr/Oct = transition, May-Sep = summer
    if month in (11,12,1,2): return "Winter"
    if month in (5,6,7,8,9):  return "Summer"
    return "Transition"  # Mar, Apr, Oct — shoulder months

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
    if "Category" in df.columns:
        df["Category"] = df["Category"].apply(split_cat)
    for col in ["Brand","Category","Product Name"]:
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
    """
    Loads exports/location_stock_*.xlsx -> 'Store x Category' sheet.
    Returns (dataframe_or_None, covered_stores_set, error_or_None)
    """
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
    cat_col = df.columns[0]  # "Category"
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
            # NaN means this category has zero stock at this (real, covered)
            # location — not "data missing". Treat as real 0, not estimated.
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
    """
    Loads category_sales_recent_*.xlsx -> 'Recent Category Sales' sheet.
    Tries Google Drive first (GDRIVE_RECENTCAT_ID), then local exports/ folder
    (only works when running locally, not on Streamlit Cloud).
    Returns (dataframe_or_None, error_or_None)
    """
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

if df_prod is None or df_pos is None:
    st.error("Could not load data. Make sure both product and POS files are on Google Drive.")
    st.stop()

USING_REAL_STOCK = df_locstk is not None and not df_locstk.empty
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
        help="How many weeks of stock you want to always have. 4 = monthly buying cycle.")
    lookback_weeks = st.slider("Sales lookback (weeks)", 2, 12, 4,
        help="How many recent weeks to use for calculating weekly sell rate.")
    min_weekly_rate = st.number_input("Min weekly rate to show (units)", 0, 50, 1,
        help="Hide categories selling fewer than this per week — filters out very slow movers.")

    st.markdown("---")
    locations = ["All"] + [l for l in LOCATION_ORDER if l in df_pos["Location"].unique()]
    sel_loc = st.selectbox("Filter by location", locations)

    # Category filter
    prod_brand = df_prod[df_prod["Brand"] == sel_brand]
    cats = ["All"] + sorted([c for c in prod_brand["Category"].unique()
                              if c and c not in ("nan","")])
    sel_cat = st.selectbox("Filter by category", cats)

    season_options = ["All", "Summer", "Winter", "All-Season"]
    sel_season = st.selectbox("Filter by season", season_options, index=0,
        help=f"Current season: {CURRENT_SEASON}. Winter items (coats, jackets, "
             f"sweaters) often show as 'Urgent' in summer due to leftover stock "
             f"selling slowly — filter to 'Summer' or 'All-Season' to focus on "
             f"what's actually relevant to buy right now.")

    st.markdown("---")
    if USING_REAL_STOCK:
        st.success("✅ Using real per-location stock")
    else:
        msg = "⚠️ Real location stock not found — using estimated split. "
        if locstk_err:
            msg += f"Error: {locstk_err}. "
        msg += ("Run `python fetch_location_stock.py`, upload the file to "
                "Drive, and set GDRIVE_LOCSTK_ID.")
        st.warning(msg)

    if USING_SEASONAL_RATES:
        st.success("✅ Using current-season sell rates")
    else:
        msg = "⚠️ Seasonal sales data not found — winter items may show as 'urgent' based on all-time sales. "
        if recentcat_err:
            msg += f"Error: {recentcat_err}. "
        msg += ("Run `python fetch_recent_category_sales.py --brand SALT`, "
                "upload the output to Drive, and set GDRIVE_RECENTCAT_ID.")
        st.warning(msg)

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ── Calculate weekly sell rates from POS data ─────────────────────────────────
today      = pd.Timestamp.today().normalize()
lookback_start = today - pd.Timedelta(weeks=lookback_weeks)

pos_recent = df_pos[df_pos["Date"] >= lookback_start].copy()

# Filter by brand (Salt locations vs Lush locations)
if "Brand" in pos_recent.columns:
    b_key = "Lush" if "Lush" in sel_brand else "Salt"
    pos_recent = pos_recent[pos_recent["Brand"].str.contains(b_key, case=False, na=False)]

# Weekly units sold per location
qty_col = "QTY" if "QTY" in pos_recent.columns else "Units"
rev_col = "Sales Amount" if "Sales Amount" in pos_recent.columns else "Revenue"

# Aggregate: units per week per location
pos_agg = pos_recent.groupby("Location").agg(
    Total_Units=(qty_col, "sum"),
    Total_Revenue=(rev_col, "sum"),
    Days=("Date", "nunique"),
).reset_index()
pos_agg["Weekly_Rate"]   = pos_agg["Total_Units"] / lookback_weeks
pos_agg["Daily_Rate"]    = pos_agg["Total_Units"] / (lookback_weeks * 7)
pos_agg["Weekly_Revenue"]= pos_agg["Total_Revenue"] / lookback_weeks

# ── Stock per category from product data (fallback / brand total) ─────────────
prod_brand = df_prod[df_prod["Brand"] == sel_brand].copy()

# Category-level stock aggregation (used as fallback when real stock missing)
cat_stock = prod_brand.groupby("Category").agg(
    On_Hand=("On Hand Qty","sum"),
    Products=("Product Name","nunique"),
    Avg_Price=("Sales Price","mean"),
).reset_index()

# Avg price lookup per category (used regardless of stock source)
avg_price_map = cat_stock.set_index("Category")["Avg_Price"].to_dict()

# ── Build reorder plan ────────────────────────────────────────────────────────
total_units = pos_agg["Total_Units"].sum()
pos_agg["Location_Share"] = pos_agg["Total_Units"] / total_units if total_units > 0 else 0

# Per category, estimate units sold (annualised -> weekly), used to split a
# location's overall weekly rate across categories
cat_sold = prod_brand.groupby("Category")["Total Units Sold"].sum().reset_index()
cat_sold.columns = ["Category","Total_Sold"]
cat_data = cat_stock.merge(cat_sold, on="Category", how="left").fillna(0)
total_sold_all = cat_data["Total_Sold"].sum()

# Real stock lookup: (location, category) -> on-hand
real_stock_map = {}
if USING_REAL_STOCK:
    for _, r in df_locstk.iterrows():
        real_stock_map[(r["Location"], r["Category"])] = r["On_Hand_Real"]

# Recent seasonal weekly-rate lookup: (location, category) -> weekly rate
# This is the FIX for the winter-coat-in-summer problem: instead of splitting
# a location's current overall weekly rate by each category's ALL-TIME sales
# share (which keeps winter items looking "hot" in summer), we use each
# category's ACTUAL recent (last N weeks) sales rate at that location.
recent_rate_map = {}
if USING_SEASONAL_RATES:
    for _, r in df_recent_cat.iterrows():
        recent_rate_map[(r["Location"], r["Category"])] = float(r.get("Weekly Rate", 0) or 0)

rows = []
for _, loc_row in pos_agg.iterrows():
    loc   = loc_row["Location"]
    share = loc_row["Location_Share"]
    loc_weekly_rate = loc_row["Weekly_Rate"]
    loc_weekly_rev  = loc_row["Weekly_Revenue"]

    for _, cat_row in cat_data.iterrows():
        cat = cat_row["Category"]
        if not cat or cat in ("nan","","All"): continue

        # ── Stock: real per-location if available, else proportional estimate ──
        real_val = real_stock_map.get((loc, cat))
        if real_val is not None:
            est_stock = max(0, real_val)
            stock_source = "real"
        elif loc in covered_stores:
            # Location has real stock coverage but this category had no
            # quant records there at all -> genuinely zero stock, not unknown.
            est_stock = 0
            stock_source = "real"
        else:
            est_stock = cat_row["On_Hand"] * share
            stock_source = "est"

        # ── Weekly rate: recent seasonal data if available, else all-time share ──
        cat_share_of_total = cat_row["Total_Sold"] / total_sold_all if total_sold_all > 0 else 0
        if (loc, cat) in recent_rate_map:
            weekly_rate = recent_rate_map[(loc, cat)]
            rate_source = "seasonal"
        else:
            weekly_rate = loc_weekly_rate * cat_share_of_total
            rate_source = "alltime"

        if weekly_rate < min_weekly_rate: continue

        # Days / weeks of cover
        daily_rate  = weekly_rate / 7
        days_cover  = est_stock / daily_rate if daily_rate > 0 else 999
        weeks_cover = days_cover / 7

        # Reorder quantity
        target_stock = target_weeks * weekly_rate
        reorder_qty  = max(0, round(target_stock - est_stock))

        # Urgency
        if weeks_cover <= 1:
            urgency = "🔴 Urgent"
            urgency_key = 0
        elif weeks_cover <= target_weeks:
            urgency = "🟡 Reorder Soon"
            urgency_key = 1
        else:
            urgency = "🟢 OK"
            urgency_key = 2

        rows.append({
            "Location":      loc,
            "Category":      cat,
            "Season":        category_season(cat),
            "Est. Stock":    round(est_stock),
            "Stock Source":  stock_source,
            "Rate Source":   rate_source,
            "Weekly Rate":   round(weekly_rate, 1),
            "Weeks Cover":   round(weeks_cover, 1),
            "Target Stock":  round(target_stock),
            "Reorder Qty":   reorder_qty,
            "Est. Value":    round(reorder_qty * avg_price_map.get(cat, 0)),
            "Urgency":       urgency,
            "_urgency_key":  urgency_key,
            "_weekly_rev":   loc_weekly_rev * cat_share_of_total if rate_source == "alltime" else 0,
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
if sel_season != "All":
    df_plan = df_plan[df_plan["Season"] == sel_season]

df_plan = df_plan.sort_values(["_urgency_key","Reorder Qty"], ascending=[True,False])

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📦 Reorder Planner")
src_badge = ('<span class="src-badge" style="background:#dcfce7;color:#166534">Real stock</span>'
              if USING_REAL_STOCK else
              '<span class="src-badge" style="background:#fef3c7;color:#92400e">Estimated stock</span>')
st.markdown(
    f"{sel_brand} · {target_weeks}-week target cover · Based on last {lookback_weeks} weeks of sales · "
    f"{today.strftime('%B %d, %Y')} {src_badge}",
    unsafe_allow_html=True)

# ── KPI strip ─────────────────────────────────────────────────────────────────
urgent_count  = len(df_plan[df_plan["_urgency_key"]==0])
reorder_count = len(df_plan[df_plan["_urgency_key"]<=1])
total_units_needed = df_plan["Reorder Qty"].sum()
total_value_needed = df_plan["Est. Value"].sum()

c1,c2,c3,c4 = st.columns(4)
for col, val, lbl, clr in [
    (c1, f"🔴 {urgent_count}",             "Urgent — under 1 week stock",  "#dc2626"),
    (c2, f"🟡 {reorder_count-urgent_count}","Reorder Soon — under target",  "#d97706"),
    (c3, f"{int(total_units_needed):,}",   "Total Units to Reorder",        "#1d4ed8"),
    (c4, fmt_npr(total_value_needed),      "Est. Reorder Value",            "#374151"),
]:
    with col:
        st.markdown(f'<div class="kpi-box"><p class="kpi-val" style="color:{clr}">{val}</p>'
                    f'<p class="kpi-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["🔴 Urgent & Reorder", "📊 Full Plan Table", "📍 By Location"])

with tab1:
    needs_action = df_plan[df_plan["_urgency_key"] <= 1].copy()
    if needs_action.empty:
        st.success("✅ All categories have sufficient stock cover. No urgent reorders needed.")
    else:
        st.markdown(f"**{len(needs_action)} category-location combinations need action**")
        for _, r in needs_action.iterrows():
            css = "urgent" if r["_urgency_key"]==0 else "warning"
            weeks_str = f"{r['Weeks Cover']:.1f} weeks" if r['Weeks Cover'] < 99 else "No sales"

            src_tag = ('<span class="src-badge" style="background:#dcfce7;color:#166534">real stock</span>'
                       if r["Stock Source"]=="real" else
                       '<span class="src-badge" style="background:#fef3c7;color:#92400e">est stock</span>')

            season_tag = ""
            if r["Season"] != "All-Season" and r["Season"] != CURRENT_SEASON:
                season_tag = (f'<span class="src-badge" style="background:#f1f5f9;color:#475569">'
                               f'{r["Season"]} item — off-season</span>')

            # Build the card as a single-line HTML string (no leading
            # whitespace on any line) — Streamlit's markdown renderer treats
            # indented lines inside a triple-quoted f-string as a code block,
            # which is what caused raw HTML tags to show up as text.
            card_html = (
                f'<div class="reorder-card {css}">'
                f'<div style="display:flex;justify-content:space-between;align-items:center">'
                f'<div>'
                f'<span style="font-size:14px;font-weight:600;color:#0f172a">{r["Category"]}</span>'
                f'<span style="font-size:12px;color:#64748b;margin-left:8px">📍 {r["Location"]}</span>'
                f'{season_tag}'
                f'</div>'
                f'<span style="font-size:13px;font-weight:600">{r["Urgency"]}</span>'
                f'</div>'
                f'<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-top:10px">'
                f'<div><div style="font-size:10px;color:#94a3b8">Est. Stock {src_tag}</div>'
                f'<div style="font-size:15px;font-weight:600">{int(r["Est. Stock"]):,}</div></div>'
                f'<div><div style="font-size:10px;color:#94a3b8">Weekly Rate</div>'
                f'<div style="font-size:15px;font-weight:600">{r["Weekly Rate"]:.1f} u/wk</div></div>'
                f'<div><div style="font-size:10px;color:#94a3b8">Weeks Cover</div>'
                f'<div style="font-size:15px;font-weight:600;color:{"#dc2626" if r["_urgency_key"]==0 else "#d97706"}">{weeks_str}</div></div>'
                f'<div><div style="font-size:10px;color:#94a3b8">Reorder Qty</div>'
                f'<div style="font-size:15px;font-weight:600;color:#1d4ed8">{int(r["Reorder Qty"]):,} units</div></div>'
                f'<div><div style="font-size:10px;color:#94a3b8">Est. Value</div>'
                f'<div style="font-size:15px;font-weight:600">{fmt_npr(r["Est. Value"])}</div></div>'
                f'</div>'
                f'</div>'
            )
            st.markdown(card_html, unsafe_allow_html=True)

with tab2:
    st.markdown("**Full reorder plan — all categories and locations**")
    display = df_plan[[
        "Urgency","Location","Category","Season","Est. Stock","Stock Source","Weekly Rate",
        "Weeks Cover","Target Stock","Reorder Qty","Est. Value"
    ]].copy()
    display["Stock Source"] = display["Stock Source"].map({"real":"✅ Real","est":"≈ Estimated"})
    display["Est. Value"] = display["Est. Value"].apply(lambda x: fmt_npr(x))
    st.dataframe(display, use_container_width=True, hide_index=True)

    # Excel download
    if st.button("⬇️ Download reorder plan as Excel"):
        out = BytesIO()
        export_df = df_plan[[
            "Urgency","Location","Category","Season","Est. Stock","Stock Source","Weekly Rate",
            "Weeks Cover","Target Stock","Reorder Qty","Est. Value"
        ]].copy()
        export_df.to_excel(out, index=False, engine="openpyxl")
        out.seek(0)
        st.download_button(
            "📥 Download Excel",
            data=out,
            file_name=f"reorder_plan_{sel_brand}_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab3:
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