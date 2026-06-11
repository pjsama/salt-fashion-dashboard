import streamlit as st
import pandas as pd
import json
from io import BytesIO
from pathlib import Path
from datetime import datetime, date
import base64

st.set_page_config(
    page_title="Salt Fashion — Buying Brief",
    page_icon="📋", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:wght@700&display=swap');

* { box-sizing: border-box; }
.block-container { padding: 1.5rem 2rem; font-family: 'Inter', sans-serif; }

/* ── Section headers ── */
.brief-title {
    font-family: 'Playfair Display', serif;
    font-size: 32px; font-weight: 700;
    color: #0f172a; letter-spacing: -0.5px;
    margin-bottom: 2px;
}
.brief-subtitle {
    font-size: 13px; color: #64748b;
    font-weight: 400; margin-bottom: 20px;
    text-transform: uppercase; letter-spacing: 1.5px;
}
.section-label {
    font-size: 10px; font-weight: 700; color: #94a3b8;
    text-transform: uppercase; letter-spacing: 2px;
    margin-bottom: 6px;
}
.section-title {
    font-size: 18px; font-weight: 600; color: #0f172a;
    margin-bottom: 14px; padding-bottom: 8px;
    border-bottom: 2px solid #f1f5f9;
}

/* ── KPI strip ── */
.kpi-strip {
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 12px; margin-bottom: 24px;
}
.kpi-box {
    background: #ffffff; border: 1px solid #e2e8f0;
    border-radius: 10px; padding: 16px 18px;
    border-left: 4px solid var(--accent);
}
.kpi-val  { font-size: 26px; font-weight: 700; color: #0f172a; margin: 0; line-height:1; }
.kpi-lbl  { font-size: 11px; color: #64748b; margin: 4px 0 0 0; }
.kpi-sub  { font-size: 11px; color: #94a3b8; margin: 2px 0 0 0; }

/* ── Recommendation cards ── */
.rec-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px; }
.rec-card {
    background: #fff; border: 1px solid #e2e8f0;
    border-radius: 10px; padding: 16px;
}
.rec-card.increase { border-left: 4px solid #16a34a; }
.rec-card.maintain { border-left: 4px solid #2563eb; }
.rec-card.reduce   { border-left: 4px solid #dc2626; }
.rec-card.watch    { border-left: 4px solid #d97706; }
.rec-action { font-size: 10px; font-weight: 700; text-transform: uppercase;
              letter-spacing: 1.5px; margin-bottom: 4px; }
.rec-action.increase { color: #16a34a; }
.rec-action.maintain { color: #2563eb; }
.rec-action.reduce   { color: #dc2626; }
.rec-action.watch    { color: #d97706; }
.rec-name  { font-size: 14px; font-weight: 600; color: #0f172a; margin-bottom: 2px; }
.rec-stats { font-size: 12px; color: #64748b; }

/* ── Season badge ── */
.season-badge {
    display: inline-block; padding: 3px 12px;
    border-radius: 20px; font-size: 11px; font-weight: 600;
    background: #f8fafc; border: 1px solid #e2e8f0; color: #475569;
    margin-right: 6px;
}
.season-badge.active {
    background: #0f172a; color: #f8fafc; border-color: #0f172a;
}

/* ── Heatmap table ── */
.heat-table { width:100%; border-collapse:collapse; font-size:12px; }
.heat-table th { background:#f8fafc; font-weight:600; padding:8px 10px;
                 text-align:left; border-bottom:2px solid #e2e8f0; color:#374151; }
.heat-table td { padding:7px 10px; border-bottom:1px solid #f1f5f9; }
.heat-table tr:hover td { background:#fafafa; }

/* ── Divider ── */
.brief-divider { border: none; border-top: 1px solid #e2e8f0; margin: 24px 0; }

/* ── Signal pill ── */
.signal { display:inline-block; padding:2px 9px; border-radius:10px;
          font-size:11px; font-weight:600; margin-right:4px; }
.signal-up    { background:#dcfce7; color:#15803d; }
.signal-down  { background:#fee2e2; color:#b91c1c; }
.signal-flat  { background:#f1f5f9; color:#475569; }
.signal-warn  { background:#fef3c7; color:#92400e; }

/* ── Print area ── */
@media print {
    .stSidebar, .stButton, [data-testid="stToolbar"] { display: none !important; }
    .block-container { padding: 0 !important; }
}
</style>
""", unsafe_allow_html=True)

# ── Google Drive IDs ────────────────────────────────────────────────────────────
GDRIVE_MAIN_ID    = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_VARIANT_ID = "1LPeoGXDDd3ZAppTiuLskzY4q-71CJWfJ"
GDRIVE_POS_ID     = "1YcW30p_dUfeeaQj-XXmGhMHP0ldAM32X"

# ── Season definitions ─────────────────────────────────────────────────────────
SEASONS = {
    "Summer 2026": (date(2026, 3, 1), date(2026, 8, 31)),
    "Winter 2025": (date(2025, 9, 1), date(2026, 2, 28)),
    "Summer 2025": (date(2025, 3, 1), date(2025, 8, 31)),
    "Winter 2024": (date(2024, 9, 1), date(2025, 2, 28)),
}

# ── Loaders ────────────────────────────────────────────────────────────────────
def gdrive_load_excel(file_id, sheet=0):
    try:
        from google.oauth2.service_account import Credentials
        import googleapiclient.discovery
        from googleapiclient.http import MediaIoBaseDownload
        import json as _j
        raw = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(
            _j.loads(_j.dumps(dict(raw))),
            scopes=["https://www.googleapis.com/auth/drive"])
        svc = googleapiclient.discovery.build("drive","v3",credentials=creds)
        req = svc.files().get_media(fileId=file_id)
        buf = BytesIO()
        dl  = MediaIoBaseDownload(buf, req)
        done = False
        while not done: _, done = dl.next_chunk()
        buf.seek(0)
        return pd.read_excel(buf, sheet_name=sheet, engine="openpyxl")
    except Exception as e:
        return None

@st.cache_data(ttl=600, show_spinner=False)
def load_products():
    df = gdrive_load_excel(GDRIVE_MAIN_ID, sheet="Products")
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base + r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"), reverse=True) if Path(d).exists() else []
            if files:
                df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl")
                break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    for col in ["Sales Price","Cost Price","On Hand Qty","Total Units Sold","Revenue","Sell-Through %"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Sell-Through %" in df.columns and df["Sell-Through %"].max() <= 1.0:
        df["Sell-Through %"] = df["Sell-Through %"] * 100
    if "Create Date" in df.columns:
        df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce")
    if "Launch Date" in df.columns:
        df["Launch Date"] = pd.to_datetime(df["Launch Date"], errors="coerce")
    for col in ["Brand","Category","STR Status","ABC Class","DOC Status","Product Name","Color","Size"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df

@st.cache_data(ttl=600, show_spinner=False)
def load_variants():
    size_df  = gdrive_load_excel(GDRIVE_VARIANT_ID, sheet="Size Breakdown")
    color_df = gdrive_load_excel(GDRIVE_VARIANT_ID, sheet="Color Breakdown")
    if size_df is None or color_df is None:
        local = r"C:\Users\Legion\Desktop\odoo_export\variant_analysis.xlsx"
        if Path(local).exists():
            size_df  = pd.read_excel(local, sheet_name="Size Breakdown",  engine="openpyxl")
            color_df = pd.read_excel(local, sheet_name="Color Breakdown", engine="openpyxl")
    if size_df is not None:
        size_df.columns  = [c.strip() for c in size_df.columns]
        color_df.columns = [c.strip() for c in color_df.columns]
        for df in [size_df, color_df]:
            for col in ["Units Sold","In Stock","STR %"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            for col in ["Brand","Category","Color","Size","Status","Product Name"]:
                if col in df.columns:
                    df[col] = df[col].fillna("").astype(str).str.replace(r"^(Color|Size|Brand):\s*","",regex=True).str.strip()
    return size_df, color_df

@st.cache_data(ttl=600, show_spinner=False)
def load_pos():
    df = gdrive_load_excel(GDRIVE_POS_ID, sheet="Point of Sale Analysis")
    if df is None:
        files = sorted(Path(r"C:\Users\Legion\Desktop\odoo_export\exports").glob("pos_analysis_*.xlsx"), reverse=True)
        if files: df = pd.read_excel(files[0], sheet_name="Point of Sale Analysis", engine="openpyxl")
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Location"] != "TOTAL"].dropna(subset=["Location"])
    df["Date"] = pd.to_datetime(df.get("Total", df.get("Date","")), errors="coerce")
    df = df.dropna(subset=["Date"])
    for col in ["Ticket Sold","QTY","Sales Amount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df

# ── Helpers ────────────────────────────────────────────────────────────────────
def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if abs(v) >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if abs(v) >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

def signal_html(val, good_direction="up"):
    if val > 5:
        return f'<span class="signal signal-up">▲ {val:.0f}%</span>'
    elif val < -5:
        return f'<span class="signal signal-down">▼ {abs(val):.0f}%</span>'
    return f'<span class="signal signal-flat">→ {val:.0f}%</span>'

def str_color(pct):
    if pct >= 70: return "#16a34a"
    if pct >= 40: return "#d97706"
    return "#dc2626"

def season_window(season_name):
    return SEASONS.get(season_name, (date(2026,3,1), date(2026,8,31)))

# ── Main ───────────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod    = load_products()
    sz_df, cl_df = load_variants()
    df_pos     = load_pos()

if df_prod is None:
    st.error("No product data found. Run scheduler.py first.")
    st.stop()

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Buying Brief")
    st.markdown("---")

    # Brand
    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("nan","True","False","None","")])
    sel_brand = st.selectbox("Brand", brands, index=0)

    st.markdown("---")

    # Review season (past)
    st.markdown("**Review Season** *(what happened)*")
    review_season = st.selectbox("Review", list(SEASONS.keys()), index=1)

    # Plan season (future)
    st.markdown("**Plan Season** *(what to buy)*")
    plan_season = st.selectbox("Plan", list(SEASONS.keys()), index=0)

    st.markdown("---")

    # Filters
    st.markdown("**Filters**")
    bdf_cats = df_prod[df_prod["Brand"] == sel_brand] if sel_brand else df_prod
    cats = ["All Categories"] + sorted([c for c in bdf_cats["Category"].unique()
                                        if c and c not in ("nan","")])
    sel_cat = st.selectbox("Category", cats, index=0)

    min_str = st.slider("Min STR % to show", 0, 100, 0, 5,
                        help="Hide categories below this STR — useful to focus on movers")

    st.markdown("---")
    gen_pdf = st.button("⬇️ Download PDF Brief", use_container_width=True)

    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ── Filter to brand ────────────────────────────────────────────────────────────
bdf = df_prod[df_prod["Brand"] == sel_brand].copy()
if sel_cat != "All Categories":
    bdf = bdf[bdf["Category"] == sel_cat]

# Season windows
rev_start, rev_end = season_window(review_season)
plan_start, plan_end = season_window(plan_season)

# Filter by review season using Create Date as proxy
def season_filter(df, start, end):
    if "Create Date" not in df.columns: return df
    cd = df["Create Date"]
    return df[(cd >= pd.Timestamp(start)) & (cd <= pd.Timestamp(end))]

bdf_rev = season_filter(bdf, rev_start, rev_end)
# If season filter gives too few rows, fall back to all
if len(bdf_rev) < 10:
    bdf_rev = bdf

# ── Header ─────────────────────────────────────────────────────────────────────
col_h, col_date = st.columns([3,1])
with col_h:
    st.markdown(f'<div class="brief-title">{sel_brand} — Buying Brief</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="brief-subtitle">Prepared {datetime.today().strftime("%B %d, %Y")}</div>', unsafe_allow_html=True)
with col_date:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f'<span class="season-badge">Review: {review_season}</span>'
        f'<span class="season-badge active">Plan: {plan_season}</span>',
        unsafe_allow_html=True)

st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — SEASON SCORECARD
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Section 1</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📊 Season Scorecard — What Happened</div>', unsafe_allow_html=True)

total_products  = len(bdf_rev)
total_revenue   = bdf_rev["Revenue"].sum()
total_units     = bdf_rev["Total Units Sold"].sum()
overall_str     = bdf_rev["Sell-Through %"].mean()
dead_count      = len(bdf_rev[bdf_rev["STR Status"].isin(["Dead","Slow"])])
dead_stock_val  = (bdf_rev[bdf_rev["STR Status"].isin(["Dead","Slow"])]["On Hand Qty"] *
                   bdf_rev[bdf_rev["STR Status"].isin(["Dead","Slow"])]["Sales Price"]).sum()
abc_a_rev       = bdf_rev[bdf_rev["ABC Class"]=="A"]["Revenue"].sum() if "ABC Class" in bdf_rev.columns else 0
abc_a_pct       = abc_a_rev / total_revenue * 100 if total_revenue else 0

st.markdown(f"""
<div class="kpi-strip">
  <div class="kpi-box" style="--accent:#2563eb">
    <p class="kpi-val">{fmt_npr(total_revenue)}</p>
    <p class="kpi-lbl">Total Revenue</p>
    <p class="kpi-sub">{total_products:,} products · {review_season}</p>
  </div>
  <div class="kpi-box" style="--accent:#16a34a">
    <p class="kpi-val">{overall_str:.1f}%</p>
    <p class="kpi-lbl">Average Sell-Through</p>
    <p class="kpi-sub">Super Fast ≥95% · Fast ≥70%</p>
  </div>
  <div class="kpi-box" style="--accent:#7c3aed">
    <p class="kpi-val">{int(total_units):,}</p>
    <p class="kpi-lbl">Units Sold</p>
    <p class="kpi-sub">Across all categories</p>
  </div>
  <div class="kpi-box" style="--accent:#dc2626">
    <p class="kpi-val">{fmt_npr(dead_stock_val)}</p>
    <p class="kpi-lbl">Dead/Slow Stock Value</p>
    <p class="kpi-sub">{dead_count:,} products need action</p>
  </div>
</div>
""", unsafe_allow_html=True)

# STR distribution bar
str_counts = bdf_rev["STR Status"].value_counts()
str_order  = ["Super Fast","Fast","Medium","Slow","Dead"]
str_colors_map = {"Super Fast":"#15803d","Fast":"#22c55e","Medium":"#f59e0b","Slow":"#f97316","Dead":"#ef4444"}
total_c = len(bdf_rev)

if total_c > 0:
    st.markdown("**Sell-Through Distribution**")
    bar_html = '<div style="display:flex;height:28px;border-radius:6px;overflow:hidden;margin-bottom:6px">'
    legend_html = '<div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:20px">'
    for s in str_order:
        cnt = str_counts.get(s, 0)
        pct = cnt / total_c * 100
        if pct > 0:
            bar_html += f'<div style="width:{pct:.1f}%;background:{str_colors_map[s]};' \
                        f'display:flex;align-items:center;justify-content:center;' \
                        f'font-size:10px;font-weight:700;color:white" title="{s}: {cnt}">' \
                        f'{"" if pct<6 else f"{pct:.0f}%"}</div>'
        legend_html += f'<span style="font-size:11px;color:#374151">' \
                       f'<span style="display:inline-block;width:10px;height:10px;' \
                       f'background:{str_colors_map[s]};border-radius:2px;margin-right:4px"></span>' \
                       f'{s} <b>{cnt:,}</b></span>'
    bar_html += '</div>'
    legend_html += '</div>'
    st.markdown(bar_html + legend_html, unsafe_allow_html=True)

st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — CATEGORY PERFORMANCE
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Section 2</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">🗂️ Category Performance</div>', unsafe_allow_html=True)

if "Category" in bdf_rev.columns and "Revenue" in bdf_rev.columns:
    cat_agg = bdf_rev.groupby("Category").agg(
        Products      = ("Product Name","nunique"),
        Revenue       = ("Revenue","sum"),
        Units_Sold    = ("Total Units Sold","sum"),
        Avg_STR       = ("Sell-Through %","mean"),
        Dead_Count    = ("STR Status", lambda x: (x.isin(["Dead","Slow"])).sum()),
        On_Hand_Units = ("On Hand Qty","sum"),
    ).reset_index()
    cat_agg["Stock_Value"] = (bdf_rev.groupby("Category").apply(
        lambda g: (g["On Hand Qty"] * g["Sales Price"]).sum())).values
    cat_agg["Rev_Share_%"] = cat_agg["Revenue"] / cat_agg["Revenue"].sum() * 100
    cat_agg = cat_agg[cat_agg["Avg_STR"] >= min_str]
    cat_agg = cat_agg.sort_values("Revenue", ascending=False)

    # Color-coded table
    def color_str(val):
        c = str_color(val)
        return f'<td style="color:{c};font-weight:600">{val:.1f}%</td>'

    tbl = '<table class="heat-table"><thead><tr>'
    for h in ["Category","Products","Revenue","Units Sold","Avg STR %","Dead/Slow","Stock Value","Rev Share"]:
        tbl += f'<th>{h}</th>'
    tbl += '</tr></thead><tbody>'
    for _, row in cat_agg.iterrows():
        tbl += f'<tr><td><b>{row["Category"]}</b></td>'
        tbl += f'<td>{int(row["Products"]):,}</td>'
        tbl += f'<td>{fmt_npr(row["Revenue"])}</td>'
        tbl += f'<td>{int(row["Units_Sold"]):,}</td>'
        tbl += color_str(row["Avg_STR"])
        dead_pct = row["Dead_Count"] / row["Products"] * 100 if row["Products"] else 0
        dead_clr = "#dc2626" if dead_pct > 40 else ("#d97706" if dead_pct > 20 else "#374151")
        tbl += f'<td style="color:{dead_clr}">{int(row["Dead_Count"]):,} ({dead_pct:.0f}%)</td>'
        tbl += f'<td>{fmt_npr(row["Stock_Value"])}</td>'
        tbl += f'<td>{row["Rev_Share_%"]:.1f}%</td>'
        tbl += '</tr>'
    tbl += '</tbody></table>'
    st.markdown(tbl, unsafe_allow_html=True)

st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — SIZE & COLOR INTELLIGENCE
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Section 3</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📐 Size & Color Intelligence</div>', unsafe_allow_html=True)

col_sz, col_cl = st.columns(2)

SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL","5XL","Free Size","One Size",
              "26","27","28","29","30","31","32","33","34","36","38","40","42"]

with col_sz:
    st.markdown("**Top Sizes by Units Sold**")
    if sz_df is not None and "Brand" in sz_df.columns:
        sz_brand = sz_df[sz_df["Brand"] == sel_brand] if sel_brand else sz_df
        if sel_cat != "All Categories" and "Category" in sz_brand.columns:
            sz_brand = sz_brand[sz_brand["Category"] == sel_cat]
        if len(sz_brand) > 0:
            sz_agg = sz_brand.groupby("Size").agg(
                Units=("Units Sold","sum"), Stock=("In Stock","sum")).reset_index()
            sz_agg["STR_%"] = sz_agg.apply(
                lambda r: min(r["Units"]/(r["Units"]+max(0,r["Stock"]))*100,100) if (r["Units"]+max(0,r["Stock"]))>0 else 0, axis=1)
            ordered = [s for s in SIZE_ORDER if s in sz_agg["Size"].values]
            others  = [s for s in sz_agg["Size"].values if s not in SIZE_ORDER]
            sz_agg = sz_agg.set_index("Size").reindex(ordered+others).dropna().reset_index()
            sz_agg = sz_agg.sort_values("Units", ascending=False).head(12)

            tbl2 = '<table class="heat-table"><thead><tr><th>Size</th><th>Units Sold</th><th>In Stock</th><th>STR %</th></tr></thead><tbody>'
            for _, r in sz_agg.iterrows():
                clr = str_color(r["STR_%"])
                tbl2 += f'<tr><td><b>{r["Size"]}</b></td><td>{int(r["Units"]):,}</td><td>{int(r["Stock"]):,}</td>'
                tbl2 += f'<td style="color:{clr};font-weight:600">{r["STR_%"]:.0f}%</td></tr>'
            tbl2 += '</tbody></table>'
            st.markdown(tbl2, unsafe_allow_html=True)
        else:
            st.info("No size data for this brand/category")
    else:
        st.info("Run variant_export.py to see size breakdown")

with col_cl:
    st.markdown("**Top Colors by Units Sold**")
    if cl_df is not None and "Brand" in cl_df.columns:
        cl_brand = cl_df[cl_df["Brand"] == sel_brand] if sel_brand else cl_df
        if sel_cat != "All Categories" and "Category" in cl_brand.columns:
            cl_brand = cl_brand[cl_brand["Category"] == sel_cat]
        if len(cl_brand) > 0:
            cl_agg = cl_brand.groupby("Color").agg(
                Units=("Units Sold","sum"), Stock=("In Stock","sum")).reset_index()
            cl_agg["STR_%"] = cl_agg.apply(
                lambda r: min(r["Units"]/(r["Units"]+max(0,r["Stock"]))*100,100) if (r["Units"]+max(0,r["Stock"]))>0 else 0, axis=1)
            cl_agg = cl_agg.sort_values("Units", ascending=False).head(12)

            tbl3 = '<table class="heat-table"><thead><tr><th>Color</th><th>Units Sold</th><th>In Stock</th><th>STR %</th></tr></thead><tbody>'
            for _, r in cl_agg.iterrows():
                clr = str_color(r["STR_%"])
                tbl3 += f'<tr><td><b>{r["Color"]}</b></td><td>{int(r["Units"]):,}</td><td>{int(r["Stock"]):,}</td>'
                tbl3 += f'<td style="color:{clr};font-weight:600">{r["STR_%"]:.0f}%</td></tr>'
            tbl3 += '</tbody></table>'
            st.markdown(tbl3, unsafe_allow_html=True)
        else:
            st.info("No color data for this brand/category")
    else:
        st.info("Run variant_export.py to see color breakdown")

st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — STORE PERFORMANCE
# ═══════════════════════════════════════════════════════════════════════════════
if df_pos is not None:
    st.markdown('<div class="section-label">Section 4</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🏪 Store Performance — Revenue by Location</div>', unsafe_allow_html=True)

    pos_filtered = df_pos.copy()
    # Filter by brand if column exists
    if "Brand" in pos_filtered.columns:
        brand_map = {"SALT": "Salt", "Jeevee Lush": "Lush"}
        b_key = brand_map.get(sel_brand, sel_brand)
        pos_filtered = pos_filtered[pos_filtered["Brand"].str.contains(b_key, case=False, na=False)]

    # Filter by review season
    pos_filtered = pos_filtered[
        (pos_filtered["Date"].dt.date >= rev_start) &
        (pos_filtered["Date"].dt.date <= rev_end)
    ]

    if len(pos_filtered) > 0:
        rev_col = "Sales Amount" if "Sales Amount" in pos_filtered.columns else "Revenue"
        tkt_col = "Ticket Sold"  if "Ticket Sold"  in pos_filtered.columns else "Tickets"

        store_agg = pos_filtered.groupby("Location").agg(
            Revenue  = (rev_col, "sum"),
            Tickets  = (tkt_col, "sum"),
        ).reset_index()
        store_agg["ATV"]       = store_agg["Revenue"] / store_agg["Tickets"].replace(0, pd.NA)
        store_agg["Rev_Share"] = store_agg["Revenue"] / store_agg["Revenue"].sum() * 100
        store_agg = store_agg.sort_values("Revenue", ascending=False)

        STORE_ORDER = ["Baneshwor","Lazimpat","Kumaripati","Chitwan","Pokhara","Online",
                       "Baneshwor Lush","Chitwan Lush","Pokhara Lush"]
        store_agg["_o"] = store_agg["Location"].apply(lambda x: STORE_ORDER.index(x) if x in STORE_ORDER else 99)
        store_agg = store_agg.sort_values("_o").drop(columns=["_o"])

        tbl4 = '<table class="heat-table"><thead><tr><th>Location</th><th>Revenue</th><th>Tickets</th><th>ATV</th><th>Rev Share</th></tr></thead><tbody>'
        for _, r in store_agg.iterrows():
            tbl4 += f'<tr><td><b>{r["Location"]}</b></td>'
            tbl4 += f'<td>{fmt_npr(r["Revenue"])}</td>'
            tbl4 += f'<td>{int(r["Tickets"]):,}</td>'
            tbl4 += f'<td>{fmt_npr(r["ATV"])}</td>'
            tbl4 += f'<td>{r["Rev_Share"]:.1f}%</td></tr>'
        tbl4 += '</tbody></table>'
        st.markdown(tbl4, unsafe_allow_html=True)
    else:
        st.info(f"No POS data found for {sel_brand} in {review_season}")

    st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — NEXT SEASON BUYING RECOMMENDATIONS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Section 5</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">🛒 Next Season Recommendations — {}</div>'.format(plan_season), unsafe_allow_html=True)

st.markdown("""
<div class="insight" style="background:#f0fdf4;border-color:#bbf7d0;color:#14532d;margin-bottom:16px">
💡 Recommendations are generated from your actual sell-through data.
Increase = STR ≥ 70% and ABC-A/B · Maintain = STR 40–69% · Reduce = STR &lt; 30% · Watch = high dead stock value.
</div>
""", unsafe_allow_html=True)

if "Category" in bdf_rev.columns and len(cat_agg) > 0:
    def get_recommendation(row):
        str_val  = row["Avg_STR"]
        dead_pct = row["Dead_Count"] / row["Products"] * 100 if row["Products"] else 0
        rev_share = row["Rev_Share_%"]
        if str_val >= 70 and dead_pct < 20:
            return "increase", "Increase buying depth", \
                   f"STR {str_val:.0f}% · Revenue share {rev_share:.1f}% · Strong sell-through justifies deeper stock."
        elif str_val >= 40 and dead_pct < 35:
            return "maintain", "Maintain current volumes", \
                   f"STR {str_val:.0f}% · Performing adequately. Hold current buying quantities and monitor."
        elif dead_pct >= 40:
            return "reduce", "Reduce or pause buying", \
                   f"STR {str_val:.0f}% · {dead_pct:.0f}% of products are dead/slow. Clear existing stock before reordering."
        else:
            return "watch", "Buy cautiously — watch closely", \
                   f"STR {str_val:.0f}% · Mixed signals. Buy smaller quantities with reorder options if it moves."

    rec_html = '<div class="rec-grid">'
    for _, row in cat_agg.iterrows():
        action_key, action_label, rationale = get_recommendation(row)
        rec_html += f'''
        <div class="rec-card {action_key}">
            <div class="rec-action {action_key}">{action_label}</div>
            <div class="rec-name">{row["Category"]}</div>
            <div class="rec-stats">{rationale}</div>
        </div>'''
    rec_html += '</div>'
    st.markdown(rec_html, unsafe_allow_html=True)

st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — TOP 10 WINNERS & LOSERS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Section 6</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">🏆 Top 10 Winners & Losers to Know</div>', unsafe_allow_html=True)

col_w, col_l = st.columns(2)

with col_w:
    st.markdown("**🟢 Reorder These — Top Performers**")
    winners = bdf_rev[bdf_rev["STR Status"].isin(["Super Fast","Fast"])].nlargest(10,"Revenue")
    if len(winners) > 0:
        tw = '<table class="heat-table"><thead><tr><th>Product</th><th>Category</th><th>Revenue</th><th>STR %</th></tr></thead><tbody>'
        for _, r in winners.iterrows():
            name = str(r.get("Product Name",""))[:38]
            tw += f'<tr><td title="{r.get("Product Name","")}">{name}{"…" if len(str(r.get("Product Name","")))>38 else ""}</td>'
            tw += f'<td>{r.get("Category","")}</td>'
            tw += f'<td>{fmt_npr(r.get("Revenue",0))}</td>'
            tw += f'<td style="color:#16a34a;font-weight:600">{r.get("Sell-Through %",0):.0f}%</td></tr>'
        tw += '</tbody></table>'
        st.markdown(tw, unsafe_allow_html=True)
    else:
        st.info("No fast movers in current filter")

with col_l:
    st.markdown("**🔴 Clear These — Dead Stock with Value**")
    losers = bdf_rev[bdf_rev["STR Status"].isin(["Dead","Slow"])].copy()
    losers["Stock_Value"] = losers["On Hand Qty"] * losers["Sales Price"]
    losers = losers.nlargest(10, "Stock_Value")
    if len(losers) > 0:
        tl = '<table class="heat-table"><thead><tr><th>Product</th><th>Stock Qty</th><th>Stock Value</th><th>STR %</th></tr></thead><tbody>'
        for _, r in losers.iterrows():
            name = str(r.get("Product Name",""))[:38]
            tl += f'<tr><td title="{r.get("Product Name","")}">{name}{"…" if len(str(r.get("Product Name","")))>38 else ""}</td>'
            tl += f'<td>{int(r.get("On Hand Qty",0)):,}</td>'
            tl += f'<td>{fmt_npr(r.get("Stock_Value",0))}</td>'
            tl += f'<td style="color:#dc2626;font-weight:600">{r.get("Sell-Through %",0):.0f}%</td></tr>'
        tl += '</tbody></table>'
        st.markdown(tl, unsafe_allow_html=True)
    else:
        st.info("No slow/dead stock in current filter")

st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — PRICE POINT ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Section 7</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">💰 Price Point Analysis — Where Customers Buy</div>', unsafe_allow_html=True)

if "Sales Price" in bdf_rev.columns and len(bdf_rev) > 0:
    bdf_price = bdf_rev[bdf_rev["Sales Price"] > 0].copy()
    if len(bdf_price) > 0:
        bins   = [0, 500, 1000, 1500, 2000, 3000, 5000, 999999]
        labels = ["Under 500","500–1K","1K–1.5K","1.5K–2K","2K–3K","3K–5K","Over 5K"]
        bdf_price["Price Band"] = pd.cut(bdf_price["Sales Price"], bins=bins, labels=labels)
        price_agg = bdf_price.groupby("Price Band", observed=True).agg(
            Products   = ("Product Name","nunique"),
            Revenue    = ("Revenue","sum"),
            Units_Sold = ("Total Units Sold","sum"),
            Avg_STR    = ("Sell-Through %","mean"),
        ).reset_index()
        price_agg["Rev_Share"] = price_agg["Revenue"] / price_agg["Revenue"].sum() * 100

        tp = '<table class="heat-table"><thead><tr><th>Price Band (NPR)</th><th>Products</th><th>Revenue</th><th>Units Sold</th><th>Avg STR %</th><th>Rev Share</th></tr></thead><tbody>'
        for _, r in price_agg.iterrows():
            clr = str_color(r["Avg_STR"])
            tp += f'<tr><td><b>{r["Price Band"]}</b></td>'
            tp += f'<td>{int(r["Products"]):,}</td>'
            tp += f'<td>{fmt_npr(r["Revenue"])}</td>'
            tp += f'<td>{int(r["Units_Sold"]):,}</td>'
            tp += f'<td style="color:{clr};font-weight:600">{r["Avg_STR"]:.1f}%</td>'
            tp += f'<td>{r["Rev_Share"]:.1f}%</td></tr>'
        tp += '</tbody></table>'
        st.markdown(tp, unsafe_allow_html=True)

        # Sweet spot callout
        best_band = price_agg.loc[price_agg["Avg_STR"].idxmax()]
        top_rev_band = price_agg.loc[price_agg["Revenue"].idxmax()]
        st.markdown(f"""
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:14px">
          <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:12px">
            <div style="font-size:10px;font-weight:700;color:#15803d;text-transform:uppercase;letter-spacing:1px">Sweet Spot</div>
            <div style="font-size:16px;font-weight:700;color:#14532d">NPR {best_band["Price Band"]}</div>
            <div style="font-size:12px;color:#166534">Highest sell-through at {best_band["Avg_STR"]:.0f}% STR — buy more here</div>
          </div>
          <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:12px">
            <div style="font-size:10px;font-weight:700;color:#1d4ed8;text-transform:uppercase;letter-spacing:1px">Revenue Driver</div>
            <div style="font-size:16px;font-weight:700;color:#1e3a8a">NPR {top_rev_band["Price Band"]}</div>
            <div style="font-size:12px;color:#1d4ed8">Highest revenue at {fmt_npr(top_rev_band["Revenue"])} — protect this range</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown('<hr class="brief-divider">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PDF DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Export</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">⬇️ Download Buying Brief</div>', unsafe_allow_html=True)

st.info("""
**To save this brief as PDF:**
1. Press **Ctrl + P** (Windows) or **Cmd + P** (Mac)
2. Set destination to **Save as PDF**
3. Layout: **Landscape** · Margins: **Minimum**
4. Untick "Headers and footers" → **Save**

This gives you a clean printable version to bring to supplier meetings.
""")

# Excel export of key tables
if st.button("⬇️ Download as Excel (all tables)", use_container_width=False):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        if "Category" in bdf_rev.columns:
            cat_agg.to_excel(writer, sheet_name="Category Performance", index=False)
        if sz_df is not None and "Brand" in sz_df.columns:
            sz_brand2 = sz_df[sz_df["Brand"] == sel_brand]
            sz_brand2.to_excel(writer, sheet_name="Size Breakdown", index=False)
        if cl_df is not None and "Brand" in cl_df.columns:
            cl_brand2 = cl_df[cl_df["Brand"] == sel_brand]
            cl_brand2.to_excel(writer, sheet_name="Color Breakdown", index=False)
        winners.to_excel(writer, sheet_name="Top Winners", index=False)
        losers.to_excel(writer, sheet_name="Clear These", index=False)
    out.seek(0)
    st.download_button(
        label="📥 Download Excel",
        data=out,
        file_name=f"buying_brief_{sel_brand}_{plan_season.replace(' ','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="margin-top:32px;padding:16px;background:#f8fafc;border-radius:8px;
font-size:11px;color:#94a3b8;text-align:center">
Salt Fashion Intelligence Platform · Buying Brief · {sel_brand} · Generated {datetime.today().strftime("%B %d, %Y")} ·
Data source: Odoo {datetime.today().strftime("%Y")} · Review period: {review_season} · Planning for: {plan_season}
</div>
""", unsafe_allow_html=True)
