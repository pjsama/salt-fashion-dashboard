import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
from datetime import datetime, date

st.set_page_config(
    page_title="Salt Fashion — Buying Brief",
    page_icon="📋", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:wght@700&display=swap');

/* ── Force light background everywhere ── */
.stApp, .stApp > div, [data-testid="stAppViewContainer"],
[data-testid="stMain"], .block-container {
    background-color: #f8fafc !important;
    color: #0f172a !important;
}
.stSidebar, [data-testid="stSidebar"] {
    background-color: #ffffff !important;
}
.stSidebar *, [data-testid="stSidebar"] * {
    color: #0f172a !important;
}
.block-container {
    padding: 2rem 2.5rem !important;
    max-width: 1200px;
    font-family: 'Inter', sans-serif;
}

/* ── Typography ── */
.brief-title {
    font-family: 'Playfair Display', serif;
    font-size: 36px; font-weight: 700;
    color: #0f172a !important;
    letter-spacing: -0.5px; line-height: 1.1;
    margin-bottom: 4px;
}
.brief-meta {
    font-size: 12px; color: #64748b !important;
    text-transform: uppercase; letter-spacing: 2px;
    margin-bottom: 0;
}
.section-eyebrow {
    font-size: 10px; font-weight: 700; color: #94a3b8 !important;
    text-transform: uppercase; letter-spacing: 2.5px;
    margin-bottom: 4px; margin-top: 28px;
}
.section-heading {
    font-size: 20px; font-weight: 600; color: #0f172a !important;
    margin-bottom: 16px; padding-bottom: 10px;
    border-bottom: 2px solid #e2e8f0;
}

/* ── KPI cards — explicit light styling ── */
.kpi-row { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 24px; }
.kpi-card {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 12px; padding: 18px 20px;
    border-left: 4px solid var(--ac) !important;
}
.kpi-val  { font-size: 28px; font-weight: 700; color: #0f172a !important; margin: 0; line-height: 1.1; }
.kpi-lbl  { font-size: 12px; font-weight: 500; color: #374151 !important; margin: 5px 0 2px; }
.kpi-sub  { font-size: 11px; color: #6b7280 !important; margin: 0; }

/* ── STR bar ── */
.str-bar-wrap { margin-bottom: 6px; }
.str-bar { display: flex; height: 32px; border-radius: 8px; overflow: hidden; }
.str-seg { display: flex; align-items: center; justify-content: center;
           font-size: 11px; font-weight: 700; color: #fff; }
.str-legend { display: flex; gap: 16px; flex-wrap: wrap; margin-top: 8px; }
.str-dot { display: inline-block; width: 10px; height: 10px;
           border-radius: 3px; margin-right: 5px; vertical-align: middle; }

/* ── Data tables ── */
.brief-table { width: 100%; border-collapse: collapse; font-size: 13px;
               background: #ffffff !important; border-radius: 10px;
               overflow: hidden; border: 1px solid #e2e8f0; }
.brief-table thead tr { background: #f1f5f9 !important; }
.brief-table th { padding: 10px 14px; text-align: left; font-weight: 600;
                  color: #374151 !important; font-size: 12px;
                  border-bottom: 2px solid #e2e8f0; }
.brief-table td { padding: 9px 14px; border-bottom: 1px solid #f1f5f9;
                  color: #1e293b !important; }
.brief-table tbody tr:last-child td { border-bottom: none; }
.brief-table tbody tr:hover td { background: #f8fafc !important; }

/* ── Rec cards ── */
.rec-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 8px; }
.rec-card {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px; padding: 16px 18px;
}
.rec-card.increase { border-left: 4px solid #16a34a !important; }
.rec-card.maintain { border-left: 4px solid #2563eb !important; }
.rec-card.reduce   { border-left: 4px solid #dc2626 !important; }
.rec-card.watch    { border-left: 4px solid #d97706 !important; }
.rec-act  { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 5px; }
.rec-act.increase { color: #16a34a !important; }
.rec-act.maintain { color: #2563eb !important; }
.rec-act.reduce   { color: #dc2626 !important; }
.rec-act.watch    { color: #d97706 !important; }
.rec-name  { font-size: 15px; font-weight: 600; color: #0f172a !important; margin-bottom: 3px; }
.rec-why   { font-size: 12px; color: #64748b !important; line-height: 1.4; }

/* ── Callout boxes ── */
.callout {
    background: #f0fdf4 !important; border: 1px solid #bbf7d0 !important;
    border-radius: 10px; padding: 14px 18px; margin-bottom: 20px;
    font-size: 13px; color: #14532d !important;
}
.callout-tip {
    background: #eff6ff !important; border: 1px solid #bfdbfe !important;
    border-radius: 10px; padding: 14px 18px; margin-bottom: 20px;
    font-size: 13px; color: #1e3a8a !important;
}

/* ── Sweet spot row ── */
.sweetspot-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 14px; }
.sweetspot-box {
    border-radius: 10px; padding: 14px 16px;
    background: #f0fdf4 !important; border: 1px solid #bbf7d0 !important;
}
.sweetspot-box.rev {
    background: #eff6ff !important; border: 1px solid #bfdbfe !important;
}
.ss-eyebrow { font-size: 10px; font-weight: 700; text-transform: uppercase;
              letter-spacing: 1.5px; margin-bottom: 4px; }
.ss-eyebrow.grn { color: #15803d !important; }
.ss-eyebrow.blu { color: #1d4ed8 !important; }
.ss-val { font-size: 18px; font-weight: 700; color: #0f172a !important; margin-bottom: 3px; }
.ss-desc { font-size: 12px; color: #374151 !important; }

/* ── Season badges ── */
.badge-row { display: flex; gap: 8px; justify-content: flex-end; align-items: flex-start; }
.sbadge {
    padding: 4px 14px; border-radius: 20px; font-size: 12px; font-weight: 500;
    background: #f1f5f9 !important; border: 1px solid #e2e8f0 !important;
    color: #475569 !important;
}
.sbadge.active {
    background: #0f172a !important; color: #f8fafc !important;
    border-color: #0f172a !important;
}

/* ── Divider ── */
.bdiv { border: none; border-top: 1px solid #e2e8f0; margin: 28px 0; }

/* ── Footer ── */
.brief-footer {
    margin-top: 40px; padding: 16px 20px;
    background: #f1f5f9 !important; border-radius: 8px;
    font-size: 11px; color: #94a3b8 !important; text-align: center;
}

@media print {
    .stSidebar, [data-testid="stSidebar"],
    .stButton, [data-testid="stToolbar"],
    [data-testid="stDecoration"] { display: none !important; }
    .block-container { padding: 0 !important; }
    .stApp { background: white !important; }
}
</style>
""", unsafe_allow_html=True)

# ── Google Drive IDs ─────────────────────────────────────────────────────────
GDRIVE_MAIN_ID    = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
GDRIVE_VARIANT_ID = "1LPeoGXDDd3ZAppTiuLskzY4q-71CJWfJ"
GDRIVE_POS_ID     = "1YcW30p_dUfeeaQj-XXmGhMHP0ldAM32X"

SEASONS = {
    "Summer 2026": (date(2026, 3, 1),  date(2026, 8, 31)),
    "Winter 2025": (date(2025, 9, 1),  date(2026, 2, 28)),
    "Summer 2025": (date(2025, 3, 1),  date(2025, 8, 31)),
    "Winter 2024": (date(2024, 9, 1),  date(2025, 2, 28)),
}

SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL","5XL",
              "Free Size","One Size","26","27","28","29","30",
              "31","32","33","34","36","38","40","42"]

# ── Loaders ──────────────────────────────────────────────────────────────────
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
        buf.seek(0)
        return buf
    except:
        return None

@st.cache_data(ttl=600, show_spinner=False)
def load_products():
    buf = gdrive_bytes(GDRIVE_MAIN_ID)
    if buf:
        try: df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl")
        except: df = None
    else:
        df = None
    if df is None:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        for d in [base+r"\exports", base]:
            files = sorted(Path(d).glob("odoo_products*.xlsx"),reverse=True) if Path(d).exists() else []
            if files: df = pd.read_excel(files[0], sheet_name="Products", engine="openpyxl"); break
    if df is None: return None
    df.columns = [c.strip() for c in df.columns]
    for col in ["Sales Price","Cost Price","On Hand Qty","Total Units Sold","Revenue","Sell-Through %"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Sell-Through %" in df.columns and df["Sell-Through %"].max() <= 1.0:
        df["Sell-Through %"] = df["Sell-Through %"] * 100
    if "Create Date" in df.columns:
        df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce")
    for col in ["Brand","Category","STR Status","ABC Class","DOC Status","Product Name","Color","Size"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df

@st.cache_data(ttl=600, show_spinner=False)
def load_variants():
    buf = gdrive_bytes(GDRIVE_VARIANT_ID)
    size_df = color_df = None
    if buf:
        try:
            size_df  = pd.read_excel(buf, sheet_name="Size Breakdown",  engine="openpyxl")
            buf.seek(0)
            color_df = pd.read_excel(buf, sheet_name="Color Breakdown", engine="openpyxl")
        except: pass
    if size_df is None:
        local = r"C:\Users\Legion\Desktop\odoo_export\variant_analysis.xlsx"
        if Path(local).exists():
            size_df  = pd.read_excel(local, sheet_name="Size Breakdown",  engine="openpyxl")
            color_df = pd.read_excel(local, sheet_name="Color Breakdown", engine="openpyxl")
    if size_df is not None:
        for df in [size_df, color_df]:
            df.columns = [c.strip() for c in df.columns]
            for col in ["Units Sold","In Stock","STR %"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            for col in ["Brand","Category","Color","Size","Status","Product Name"]:
                if col in df.columns:
                    df[col] = df[col].fillna("").astype(str)\
                        .str.replace(r"^(Color|Size|Brand):\s*","",regex=True).str.strip()
    return size_df, color_df

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

# ── Helpers ──────────────────────────────────────────────────────────────────
def fmt_npr(v):
    if pd.isna(v) or v == 0: return "—"
    if abs(v) >= 1_000_000: return f"NPR {v/1_000_000:.1f}M"
    if abs(v) >= 1_000:     return f"NPR {v/1_000:.0f}K"
    return f"NPR {v:,.0f}"

def str_clr(pct):
    if pct >= 70: return "#16a34a"
    if pct >= 40: return "#d97706"
    return "#dc2626"

def clean_cat(cat):
    """'Denim Pant / Baggy' → 'Denim Pant · Baggy'  keeps it readable"""
    return str(cat).replace(" / ", " · ")

def get_rec(avg_str, dead_pct, rev_share):
    if avg_str >= 70 and dead_pct < 20:
        return "increase","Increase buying depth", \
               f"STR {avg_str:.0f}% · {rev_share:.1f}% revenue share · Strong sell-through."
    elif avg_str >= 40 and dead_pct < 35:
        return "maintain","Maintain current volumes", \
               f"STR {avg_str:.0f}% · Adequate performance. Hold quantities and monitor."
    elif dead_pct >= 40:
        return "reduce","Reduce or pause buying", \
               f"STR {avg_str:.0f}% · {dead_pct:.0f}% dead/slow. Clear existing stock first."
    else:
        return "watch","Buy cautiously — watch closely", \
               f"STR {avg_str:.0f}% · Mixed signals. Buy small with reorder options."

# ── Load data ────────────────────────────────────────────────────────────────
with st.spinner("Loading data…"):
    df_prod       = load_products()
    sz_df, cl_df  = load_variants()
    df_pos        = load_pos()

if df_prod is None:
    st.error("No product data. Run scheduler.py first."); st.stop()

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Buying Brief")
    st.markdown("---")

    brands = sorted([b for b in df_prod["Brand"].unique()
                     if b and b not in ("nan","True","False","None","")])
    sel_brand = st.selectbox("Brand", brands, index=0)

    st.markdown("---")
    st.markdown("**Review Season** *(what happened)*")
    review_season = st.selectbox("Review", list(SEASONS.keys()), index=1)
    st.markdown("**Plan Season** *(what to buy)*")
    plan_season   = st.selectbox("Plan",   list(SEASONS.keys()), index=0)

    st.markdown("---")
    st.markdown("**Filters**")
    bdf_all = df_prod[df_prod["Brand"] == sel_brand]
    cats    = ["All Categories"] + sorted([c for c in bdf_all["Category"].unique()
                                           if c and c not in ("nan","")])
    sel_cat   = st.selectbox("Category", cats, index=0)
    min_str   = st.slider("Min STR % to show", 0, 100, 0, 5)

    st.markdown("---")
    if st.button("⬇️ Download PDF  (Ctrl+P → Save as PDF)", use_container_width=True):
        st.info("Press Ctrl+P → Save as PDF → Landscape → Minimum margins")
    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear(); st.rerun()

# ── Filtered dataframe ───────────────────────────────────────────────────────
bdf = df_prod[df_prod["Brand"] == sel_brand].copy()
if sel_cat != "All Categories":
    bdf = bdf[bdf["Category"] == sel_cat]

rev_start, rev_end = SEASONS[review_season]

# Season filter on Create Date — fall back to all if too sparse
def season_df(df, s, e):
    if "Create Date" not in df.columns: return df
    filt = df[(df["Create Date"] >= pd.Timestamp(s)) & (df["Create Date"] <= pd.Timestamp(e))]
    return filt if len(filt) >= 10 else df

bdf_rev = season_df(bdf, rev_start, rev_end)

# ════════════════════════════════════════════════════════════════════════════
# HEADER
# ════════════════════════════════════════════════════════════════════════════
col_h, col_b = st.columns([3,1])
with col_h:
    st.markdown(f'<div class="brief-title">{sel_brand} — Buying Brief</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="brief-meta">Prepared {datetime.today().strftime("%B %d, %Y")}</div>', unsafe_allow_html=True)
with col_b:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f'<div class="badge-row">'
        f'<span class="sbadge">Review: {review_season}</span>'
        f'<span class="sbadge active">Plan: {plan_season}</span>'
        f'</div>', unsafe_allow_html=True)

st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 1 — SCORECARD
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-eyebrow">Section 1</div>', unsafe_allow_html=True)
st.markdown('<div class="section-heading">📊 Season Scorecard — What Happened</div>', unsafe_allow_html=True)

total_rev   = bdf_rev["Revenue"].sum()
total_units = bdf_rev["Total Units Sold"].sum()
avg_str     = bdf_rev["Sell-Through %"].mean()
n_prod      = len(bdf_rev)
dead_mask   = bdf_rev["STR Status"].isin(["Dead","Slow"])
dead_count  = dead_mask.sum()
dead_val    = (bdf_rev[dead_mask]["On Hand Qty"] * bdf_rev[dead_mask]["Sales Price"]).sum()

st.markdown(f"""
<div class="kpi-row">
  <div class="kpi-card" style="--ac:#2563eb">
    <p class="kpi-val">{fmt_npr(total_rev)}</p>
    <p class="kpi-lbl">Total Revenue</p>
    <p class="kpi-sub">{n_prod:,} products · {review_season}</p>
  </div>
  <div class="kpi-card" style="--ac:#16a34a">
    <p class="kpi-val">{avg_str:.1f}%</p>
    <p class="kpi-lbl">Average Sell-Through</p>
    <p class="kpi-sub">Super Fast ≥95% · Fast ≥70%</p>
  </div>
  <div class="kpi-card" style="--ac:#7c3aed">
    <p class="kpi-val">{int(total_units):,}</p>
    <p class="kpi-lbl">Units Sold</p>
    <p class="kpi-sub">Across all categories</p>
  </div>
  <div class="kpi-card" style="--ac:#dc2626">
    <p class="kpi-val">{fmt_npr(dead_val)}</p>
    <p class="kpi-lbl">Dead / Slow Stock Value</p>
    <p class="kpi-sub">{dead_count:,} products need action</p>
  </div>
</div>
""", unsafe_allow_html=True)

# STR distribution bar
str_cfg = [
    ("Super Fast","#15803d"), ("Fast","#22c55e"),
    ("Medium","#f59e0b"),     ("Slow","#f97316"), ("Dead","#ef4444"),
]
total_c = len(bdf_rev)
if total_c > 0:
    st.markdown("**Sell-Through Distribution**")
    counts = bdf_rev["STR Status"].value_counts()
    bar = '<div class="str-bar">'
    leg = '<div class="str-legend">'
    for label, color in str_cfg:
        cnt = counts.get(label, 0)
        pct = cnt / total_c * 100
        if pct > 0:
            bar += f'<div class="str-seg" style="width:{pct:.1f}%;background:{color}">' \
                   f'{"" if pct<5 else f"{pct:.0f}%"}</div>'
        leg += f'<span><span class="str-dot" style="background:{color}"></span>' \
               f'{label} <b>{cnt:,}</b></span>'
    bar += '</div>'; leg += '</div>'
    st.markdown(f'<div class="str-bar-wrap">{bar}{leg}</div>', unsafe_allow_html=True)

st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 2 — CATEGORY PERFORMANCE
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-eyebrow">Section 2</div>', unsafe_allow_html=True)
st.markdown('<div class="section-heading">🗂️ Category Performance</div>', unsafe_allow_html=True)

if "Category" in bdf_rev.columns:
    cat_agg = bdf_rev.groupby("Category").agg(
        Products   =("Product Name","nunique"),
        Revenue    =("Revenue","sum"),
        Units_Sold =("Total Units Sold","sum"),
        Avg_STR    =("Sell-Through %","mean"),
        Dead_Count =("STR Status", lambda x: x.isin(["Dead","Slow"]).sum()),
        On_Hand    =("On Hand Qty","sum"),
    ).reset_index()
    # Stock value per category
    sv = bdf_rev.groupby("Category").apply(
        lambda g: (g["On Hand Qty"] * g["Sales Price"]).sum()).reset_index(name="Stock_Value")
    cat_agg = cat_agg.merge(sv, on="Category", how="left")
    cat_agg["Rev_Share"] = cat_agg["Revenue"] / cat_agg["Revenue"].sum() * 100
    cat_agg = cat_agg[cat_agg["Avg_STR"] >= min_str].sort_values("Revenue", ascending=False)

    tbl = ('<table class="brief-table"><thead><tr>'
           '<th>Category</th><th>Products</th><th>Revenue</th>'
           '<th>Units Sold</th><th>Avg STR %</th><th>Dead/Slow</th>'
           '<th>Stock Value</th><th>Rev Share</th>'
           '</tr></thead><tbody>')
    for _, r in cat_agg.iterrows():
        dp  = r["Dead_Count"]/r["Products"]*100 if r["Products"] else 0
        dc  = "#dc2626" if dp>40 else ("#d97706" if dp>20 else "#374151")
        sc  = str_clr(r["Avg_STR"])
        tbl += (f'<tr>'
                f'<td><b>{clean_cat(r["Category"])}</b></td>'
                f'<td>{int(r["Products"]):,}</td>'
                f'<td>{fmt_npr(r["Revenue"])}</td>'
                f'<td>{int(r["Units_Sold"]):,}</td>'
                f'<td style="color:{sc};font-weight:600">{r["Avg_STR"]:.1f}%</td>'
                f'<td style="color:{dc}">{int(r["Dead_Count"]):,} ({dp:.0f}%)</td>'
                f'<td>{fmt_npr(r["Stock_Value"])}</td>'
                f'<td>{r["Rev_Share"]:.1f}%</td>'
                f'</tr>')
    tbl += '</tbody></table>'
    st.markdown(tbl, unsafe_allow_html=True)
else:
    st.info("No category data available.")
    cat_agg = pd.DataFrame()

st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 3 — SIZE & COLOR
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-eyebrow">Section 3</div>', unsafe_allow_html=True)
st.markdown('<div class="section-heading">📐 Size & Color Intelligence</div>', unsafe_allow_html=True)

col_sz, col_cl = st.columns(2)

def size_color_table(rows, col1, col1_label):
    t = (f'<table class="brief-table"><thead><tr>'
         f'<th>{col1_label}</th><th>Units Sold</th><th>In Stock</th><th>STR %</th>'
         f'</tr></thead><tbody>')
    for _, r in rows.iterrows():
        pct = r.get("STR_%", r.get("STR %", 0))
        clr = str_clr(pct)
        t += (f'<tr><td><b>{r[col1]}</b></td>'
              f'<td>{int(r.get("Units",r.get("Units Sold",0))):,}</td>'
              f'<td>{int(r.get("Stock",r.get("In Stock",0))):,}</td>'
              f'<td style="color:{clr};font-weight:600">{pct:.0f}%</td></tr>')
    t += '</tbody></table>'
    return t

with col_sz:
    st.markdown("**Top Sizes by Units Sold**")
    if sz_df is not None and "Brand" in sz_df.columns:
        sf2 = sz_df[sz_df["Brand"] == sel_brand].copy()
        if sel_cat != "All Categories" and "Category" in sf2.columns:
            sf2 = sf2[sf2["Category"] == sel_cat]
        if len(sf2) > 0:
            sa = sf2.groupby("Size").agg(Units=("Units Sold","sum"),Stock=("In Stock","sum")).reset_index()
            sa["STR_%"] = sa.apply(lambda r: min(r["Units"]/(r["Units"]+max(0,r["Stock"]))*100,100)
                                   if r["Units"]+max(0,r["Stock"])>0 else 0, axis=1)
            ordered = [s for s in SIZE_ORDER if s in sa["Size"].values]
            others  = [s for s in sa["Size"].values if s not in SIZE_ORDER]
            sa = sa.set_index("Size").reindex(ordered+others).dropna().reset_index()
            sa = sa.sort_values("Units", ascending=False).head(14)
            st.markdown(size_color_table(sa, "Size", "Size"), unsafe_allow_html=True)
        else:
            st.info("No size data for this filter")
    else:
        st.info("Run variant_export.py to enable size data")

with col_cl:
    st.markdown("**Top Colors by Units Sold**")
    if cl_df is not None and "Brand" in cl_df.columns:
        cf2 = cl_df[cl_df["Brand"] == sel_brand].copy()
        if sel_cat != "All Categories" and "Category" in cf2.columns:
            cf2 = cf2[cf2["Category"] == sel_cat]
        if len(cf2) > 0:
            ca = cf2.groupby("Color").agg(Units=("Units Sold","sum"),Stock=("In Stock","sum")).reset_index()
            ca["STR_%"] = ca.apply(lambda r: min(r["Units"]/(r["Units"]+max(0,r["Stock"]))*100,100)
                                   if r["Units"]+max(0,r["Stock"])>0 else 0, axis=1)
            ca = ca.sort_values("Units", ascending=False).head(14)
            st.markdown(size_color_table(ca, "Color", "Color"), unsafe_allow_html=True)
        else:
            st.info("No color data for this filter")
    else:
        st.info("Run variant_export.py to enable color data")

st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 4 — STORE PERFORMANCE
# ════════════════════════════════════════════════════════════════════════════
if df_pos is not None:
    st.markdown('<div class="section-eyebrow">Section 4</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-heading">🏪 Store Performance</div>', unsafe_allow_html=True)

    pf = df_pos.copy()
    if "Brand" in pf.columns:
        b_key = "Lush" if "Lush" in sel_brand else "Salt"
        pf = pf[pf["Brand"].str.contains(b_key, case=False, na=False)]
    pf = pf[(pf["Date"].dt.date >= rev_start) & (pf["Date"].dt.date <= rev_end)]

    if len(pf) > 0:
        rc = "Sales Amount" if "Sales Amount" in pf.columns else "Revenue"
        tc = "Ticket Sold"  if "Ticket Sold"  in pf.columns else "Tickets"
        sa = pf.groupby("Location").agg(Revenue=(rc,"sum"),Tickets=(tc,"sum")).reset_index()
        sa["ATV"]       = sa["Revenue"] / sa["Tickets"].replace(0, pd.NA)
        sa["Rev_Share"] = sa["Revenue"] / sa["Revenue"].sum() * 100
        SLOC = ["Baneshwor","Lazimpat","Kumaripati","Chitwan","Pokhara","Online",
                "Baneshwor Lush","Chitwan Lush","Pokhara Lush"]
        sa["_o"] = sa["Location"].apply(lambda x: SLOC.index(x) if x in SLOC else 99)
        sa = sa.sort_values("_o").drop(columns=["_o"])

        t = ('<table class="brief-table"><thead><tr>'
             '<th>Location</th><th>Revenue</th><th>Tickets</th>'
             '<th>ATV</th><th>Rev Share</th>'
             '</tr></thead><tbody>')
        for _, r in sa.iterrows():
            t += (f'<tr><td><b>{r["Location"]}</b></td>'
                  f'<td>{fmt_npr(r["Revenue"])}</td>'
                  f'<td>{int(r["Tickets"]):,}</td>'
                  f'<td>{fmt_npr(r["ATV"])}</td>'
                  f'<td>{r["Rev_Share"]:.1f}%</td></tr>')
        t += '</tbody></table>'
        st.markdown(t, unsafe_allow_html=True)
    else:
        st.info(f"No POS data for {sel_brand} in {review_season}. Check date range.")

    st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 5 — BUYING RECOMMENDATIONS
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-eyebrow">Section 5</div>', unsafe_allow_html=True)
st.markdown(f'<div class="section-heading">🛒 Buying Recommendations — {plan_season}</div>',
            unsafe_allow_html=True)

st.markdown(
    '<div class="callout-tip">💡 Recommendations are based on sell-through data from the review season. '
    '<b>Increase</b> = STR ≥ 70% and low dead stock · '
    '<b>Maintain</b> = STR 40–69% · '
    '<b>Reduce</b> = high dead stock (≥ 40%) · '
    '<b>Watch</b> = mixed signals</div>',
    unsafe_allow_html=True)

if len(cat_agg) > 0:
    rec_html = '<div class="rec-grid">'
    for _, row in cat_agg.iterrows():
        dp  = row["Dead_Count"] / row["Products"] * 100 if row["Products"] else 0
        key, label, why = get_rec(row["Avg_STR"], dp, row["Rev_Share"])
        rec_html += (f'<div class="rec-card {key}">'
                     f'<div class="rec-act {key}">{label}</div>'
                     f'<div class="rec-name">{clean_cat(row["Category"])}</div>'
                     f'<div class="rec-why">{why}</div>'
                     f'</div>')
    rec_html += '</div>'
    st.markdown(rec_html, unsafe_allow_html=True)
else:
    st.info("No category data to generate recommendations.")

st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 6 — WINNERS & LOSERS
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-eyebrow">Section 6</div>', unsafe_allow_html=True)
st.markdown('<div class="section-heading">🏆 Top 10 Winners & Losers</div>', unsafe_allow_html=True)

col_w, col_l = st.columns(2)

with col_w:
    st.markdown("**🟢 Reorder These — Top Performers**")
    winners = bdf_rev[bdf_rev["STR Status"].isin(["Super Fast","Fast"])].nlargest(10,"Revenue")
    if len(winners) > 0:
        t = ('<table class="brief-table"><thead><tr>'
             '<th>Product</th><th>Category</th><th>Revenue</th><th>STR %</th>'
             '</tr></thead><tbody>')
        for _, r in winners.iterrows():
            nm = str(r.get("Product Name",""))[:36]
            t += (f'<tr><td title="{r.get("Product Name","")}">{nm}{"…" if len(str(r.get("Product Name","")))>36 else ""}</td>'
                  f'<td>{clean_cat(r.get("Category",""))}</td>'
                  f'<td>{fmt_npr(r.get("Revenue",0))}</td>'
                  f'<td style="color:#16a34a;font-weight:600">{r.get("Sell-Through %",0):.0f}%</td></tr>')
        t += '</tbody></table>'
        st.markdown(t, unsafe_allow_html=True)
    else:
        st.info("No fast movers in current filter")

with col_l:
    st.markdown("**🔴 Clear These — Dead Stock with Value**")
    losers = bdf_rev[bdf_rev["STR Status"].isin(["Dead","Slow"])].copy()
    losers["_sv"] = losers["On Hand Qty"] * losers["Sales Price"]
    losers = losers.nlargest(10,"_sv")
    if len(losers) > 0:
        t = ('<table class="brief-table"><thead><tr>'
             '<th>Product</th><th>Stock Qty</th><th>Stock Value</th><th>STR %</th>'
             '</tr></thead><tbody>')
        for _, r in losers.iterrows():
            nm = str(r.get("Product Name",""))[:36]
            t += (f'<tr><td title="{r.get("Product Name","")}">{nm}{"…" if len(str(r.get("Product Name","")))>36 else ""}</td>'
                  f'<td>{int(r.get("On Hand Qty",0)):,}</td>'
                  f'<td>{fmt_npr(r.get("_sv",0))}</td>'
                  f'<td style="color:#dc2626;font-weight:600">{r.get("Sell-Through %",0):.0f}%</td></tr>')
        t += '</tbody></table>'
        st.markdown(t, unsafe_allow_html=True)
    else:
        st.info("No slow/dead stock in current filter")

st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 7 — PRICE POINT ANALYSIS
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-eyebrow">Section 7</div>', unsafe_allow_html=True)
st.markdown('<div class="section-heading">💰 Price Point Analysis</div>', unsafe_allow_html=True)

bp = bdf_rev[bdf_rev["Sales Price"] > 0].copy()
if len(bp) > 0:
    bins   = [0,500,1000,1500,2000,3000,5000,999999]
    labels = ["Under 500","500–1K","1K–1.5K","1.5K–2K","2K–3K","3K–5K","Over 5K"]
    bp["Band"] = pd.cut(bp["Sales Price"], bins=bins, labels=labels)
    pa = bp.groupby("Band", observed=True).agg(
        Products  =("Product Name","nunique"),
        Revenue   =("Revenue","sum"),
        Units     =("Total Units Sold","sum"),
        Avg_STR   =("Sell-Through %","mean"),
    ).reset_index()
    pa["Rev_Share"] = pa["Revenue"] / pa["Revenue"].sum() * 100

    t = ('<table class="brief-table"><thead><tr>'
         '<th>Price Band (NPR)</th><th>Products</th><th>Revenue</th>'
         '<th>Units Sold</th><th>Avg STR %</th><th>Rev Share</th>'
         '</tr></thead><tbody>')
    for _, r in pa.iterrows():
        sc = str_clr(r["Avg_STR"])
        t += (f'<tr><td><b>{r["Band"]}</b></td>'
              f'<td>{int(r["Products"]):,}</td>'
              f'<td>{fmt_npr(r["Revenue"])}</td>'
              f'<td>{int(r["Units"]):,}</td>'
              f'<td style="color:{sc};font-weight:600">{r["Avg_STR"]:.1f}%</td>'
              f'<td>{r["Rev_Share"]:.1f}%</td></tr>')
    t += '</tbody></table>'
    st.markdown(t, unsafe_allow_html=True)

    # Sweet spot callout
    best   = pa.loc[pa["Avg_STR"].idxmax()]
    top_rv = pa.loc[pa["Revenue"].idxmax()]
    st.markdown(f"""
    <div class="sweetspot-row">
      <div class="sweetspot-box">
        <div class="ss-eyebrow grn">Sweet Spot — Best STR</div>
        <div class="ss-val">NPR {best["Band"]}</div>
        <div class="ss-desc">Highest sell-through at {best["Avg_STR"]:.0f}% STR — buy more at this price point</div>
      </div>
      <div class="sweetspot-box rev">
        <div class="ss-eyebrow blu">Revenue Driver</div>
        <div class="ss-val">NPR {top_rv["Band"]}</div>
        <div class="ss-desc">Highest revenue at {fmt_npr(top_rv["Revenue"])} — protect this range</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<hr class="bdiv">', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# EXPORT
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-eyebrow">Export</div>', unsafe_allow_html=True)
st.markdown('<div class="section-heading">⬇️ Download</div>', unsafe_allow_html=True)

col_pdf, col_xl = st.columns(2)
with col_pdf:
    st.markdown("""
    **Save as PDF for supplier meetings:**
    1. Press **Ctrl + P** (Windows) or **Cmd + P** (Mac)
    2. Destination → **Save as PDF**
    3. Layout → **Landscape** · Margins → **Minimum**
    4. Untick *Headers and footers* → **Save**
    """)
with col_xl:
    if st.button("⬇️ Export all tables as Excel"):
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            if len(cat_agg) > 0:
                cat_agg.to_excel(w, sheet_name="Category Performance", index=False)
            if sz_df is not None and "Brand" in sz_df.columns:
                sz_df[sz_df["Brand"]==sel_brand].to_excel(w, sheet_name="Size Breakdown", index=False)
            if cl_df is not None and "Brand" in cl_df.columns:
                cl_df[cl_df["Brand"]==sel_brand].to_excel(w, sheet_name="Color Breakdown", index=False)
            if len(winners) > 0:
                winners.to_excel(w, sheet_name="Top Winners", index=False)
            if len(losers) > 0:
                losers.to_excel(w, sheet_name="Clear These", index=False)
        out.seek(0)
        st.download_button("📥 Download Excel",
            data=out,
            file_name=f"buying_brief_{sel_brand}_{plan_season.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Footer
st.markdown(
    f'<div class="brief-footer">Salt Fashion Intelligence Platform · '
    f'{sel_brand} · {review_season} → {plan_season} · '
    f'Generated {datetime.today().strftime("%B %d, %Y")}</div>',
    unsafe_allow_html=True)