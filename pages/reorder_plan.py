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
</style>
""", unsafe_allow_html=True)

# ── Google Drive IDs ──────────────────────────────────────────────────────────
GDRIVE_MAIN_ID  = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_POS_ID   = "1YcW30p_dUfeeaQj-XXmGhMHP0ldAM32X"

LOCATION_ORDER = ["Baneshwor","Lazimpat","Kumaripati","Chitwan","Pokhara","Online",
                  "Baneshwor Lush","Chitwan Lush","Pokhara Lush"]

SKIP_PARTS = {"All","Saleable","PoS",""}

def split_cat(raw):
    parts = [p.strip() for p in str(raw).split("/") if p.strip() not in SKIP_PARTS]
    if not parts: return ""
    return parts[-2] if len(parts) >= 2 else parts[0]

# ── Loaders ───────────────────────────────────────────────────────────────────
def gdrive_bytes(file_id):
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
        buf.seek(0); return buf
    except: return None

@st.cache_data(ttl=600, show_spinner=False)
def load_products():
    buf = gdrive_bytes(GDRIVE_MAIN_ID)
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
    buf = gdrive_bytes(GDRIVE_POS_ID)
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

def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if v >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if v >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

# ── Main ──────────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod = load_products()
    df_pos  = load_pos()

if df_prod is None or df_pos is None:
    st.error("Could not load data. Make sure both product and POS files are on Google Drive.")
    st.stop()

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

    st.markdown("---")
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

# ── Stock per category from product data ─────────────────────────────────────
prod_brand = df_prod[df_prod["Brand"] == sel_brand].copy()

# Category-level stock aggregation
cat_stock = prod_brand.groupby("Category").agg(
    On_Hand=("On Hand Qty","sum"),
    Products=("Product Name","nunique"),
    Avg_Price=("Sales Price","mean"),
).reset_index()

# ── Build reorder plan ────────────────────────────────────────────────────────
# We need: category stock × location weekly rate
# Since POS data doesn't split by category per location, we:
# 1. Get each location's % of total sales
# 2. Distribute category stock proportionally across locations
# 3. Calculate reorder need per location

total_units = pos_agg["Total_Units"].sum()
pos_agg["Location_Share"] = pos_agg["Total_Units"] / total_units if total_units > 0 else 0

# Per category, estimate units sold at each location
# Use variant-level or template-level sold data
cat_sold = prod_brand.groupby("Category")["Total Units Sold"].sum().reset_index()
cat_sold.columns = ["Category","Total_Sold"]

# Merge
cat_data = cat_stock.merge(cat_sold, on="Category", how="left").fillna(0)
cat_data["Avg_Weekly_Sold"] = cat_data["Total_Sold"] / 52  # annualised to weekly

# Build location × category reorder table
rows = []
for _, loc_row in pos_agg.iterrows():
    loc   = loc_row["Location"]
    share = loc_row["Location_Share"]
    loc_weekly_rate = loc_row["Weekly_Rate"]
    loc_weekly_rev  = loc_row["Weekly_Revenue"]

    for _, cat_row in cat_data.iterrows():
        cat        = cat_row["Category"]
        if not cat or cat in ("nan","","All"): continue

        # Estimated stock at this location (proportional share)
        est_stock      = cat_row["On_Hand"] * share
        # Weekly sell rate for this category at this location
        cat_share_of_total = cat_row["Total_Sold"] / cat_data["Total_Sold"].sum() \
                             if cat_data["Total_Sold"].sum() > 0 else 0
        weekly_rate    = loc_weekly_rate * cat_share_of_total
        if weekly_rate < min_weekly_rate: continue

        # Days of cover
        daily_rate     = weekly_rate / 7
        days_cover     = est_stock / daily_rate if daily_rate > 0 else 999
        weeks_cover    = days_cover / 7

        # Reorder quantity
        target_stock   = target_weeks * weekly_rate
        reorder_qty    = max(0, round(target_stock - est_stock))

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
            "Est. Stock":    round(est_stock),
            "Weekly Rate":   round(weekly_rate, 1),
            "Weeks Cover":   round(weeks_cover, 1),
            "Target Stock":  round(target_stock),
            "Reorder Qty":   reorder_qty,
            "Est. Value":    round(reorder_qty * cat_row["Avg_Price"]),
            "Urgency":       urgency,
            "_urgency_key":  urgency_key,
            "_weekly_rev":   loc_weekly_rev * cat_share_of_total,
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

df_plan = df_plan.sort_values(["_urgency_key","Reorder Qty"], ascending=[True,False])

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📦 Reorder Planner")
st.caption(f"{sel_brand} · {target_weeks}-week target cover · Based on last {lookback_weeks} weeks of sales · {today.strftime('%B %d, %Y')}")

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
            st.markdown(f"""
            <div class="reorder-card {css}">
              <div style="display:flex;justify-content:space-between;align-items:center">
                <div>
                  <span style="font-size:14px;font-weight:600;color:#0f172a">{r['Category']}</span>
                  <span style="font-size:12px;color:#64748b;margin-left:8px">📍 {r['Location']}</span>
                </div>
                <span style="font-size:13px;font-weight:600">{r['Urgency']}</span>
              </div>
              <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-top:10px">
                <div><div style="font-size:10px;color:#94a3b8">Est. Stock</div>
                     <div style="font-size:15px;font-weight:600">{int(r['Est. Stock']):,}</div></div>
                <div><div style="font-size:10px;color:#94a3b8">Weekly Rate</div>
                     <div style="font-size:15px;font-weight:600">{r['Weekly Rate']:.1f} u/wk</div></div>
                <div><div style="font-size:10px;color:#94a3b8">Weeks Cover</div>
                     <div style="font-size:15px;font-weight:600;color:{'#dc2626' if r['_urgency_key']==0 else '#d97706'}">{weeks_str}</div></div>
                <div><div style="font-size:10px;color:#94a3b8">Reorder Qty</div>
                     <div style="font-size:15px;font-weight:600;color:#1d4ed8">{int(r['Reorder Qty']):,} units</div></div>
                <div><div style="font-size:10px;color:#94a3b8">Est. Value</div>
                     <div style="font-size:15px;font-weight:600">{fmt_npr(r['Est. Value'])}</div></div>
              </div>
            </div>
            """, unsafe_allow_html=True)

with tab2:
    st.markdown("**Full reorder plan — all categories and locations**")
    display = df_plan[[
        "Urgency","Location","Category","Est. Stock","Weekly Rate",
        "Weeks Cover","Target Stock","Reorder Qty","Est. Value"
    ]].copy()
    display["Est. Value"] = display["Est. Value"].apply(lambda x: fmt_npr(x))
    st.dataframe(display, use_container_width=True, hide_index=True)

    # Excel download
    if st.button("⬇️ Download reorder plan as Excel"):
        out = BytesIO()
        export_df = df_plan[[
            "Urgency","Location","Category","Est. Stock","Weekly Rate",
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
        "Total_Units":"Units to Reorder","Total_Value":"Est. Value",
        "Reorder_Soon":"Reorder Soon"})
    st.dataframe(loc_summary, use_container_width=True, hide_index=True)

    # Horizontal bar chart — reorder units by location
    st.markdown("**Units needed by location**")
    chart_df = df_plan.groupby("Location")["Reorder Qty"].sum().reset_index()
    chart_df["_o"] = chart_df["Location"].apply(
        lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
    chart_df = chart_df.sort_values("_o").drop(columns=["_o"])
    st.bar_chart(chart_df.set_index("Location")["Reorder Qty"])

# ── Footer note ───────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    f"⚠️ Stock estimates per location are calculated proportionally from total stock × location sales share. "
    f"For exact per-location stock counts, a warehouse integration would be needed. "
    f"Reorder Qty = ({target_weeks} weeks × weekly rate) − estimated current stock. "
    f"All figures based on last {lookback_weeks} weeks of POS data."
)
