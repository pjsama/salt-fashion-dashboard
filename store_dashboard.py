import streamlit as st
import pandas as pd
import json
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Salt Fashion — Store Intelligence",
    page_icon="🏪", layout="wide",
    initial_sidebar_state="expanded",
)

# ── Google Drive loader (reuse pattern from main dashboard) ───────────────────
GDRIVE_FILE_ID = "REPLACE_WITH_STORE_ANALYSIS_FILE_ID"

@st.cache_data(ttl=300)
def load_from_gdrive(file_id):
    try:
        import gdown
        url = f"https://drive.google.com/uc?id={file_id}"
        buf = BytesIO()
        gdown.download(url, buf, quiet=True)
        buf.seek(0)
        return buf
    except Exception:
        return None

@st.cache_data(ttl=300)
def load_data():
    # Try Google Drive first
    buf = load_from_gdrive(GDRIVE_FILE_ID)
    if buf:
        try:
            return pd.read_excel(buf, sheet_name=None)
        except Exception:
            pass

    # Try service account (same as main dashboard)
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        import json as _j; raw = st.secrets["gcp_service_account"]; creds_info = _j.loads(_j.dumps(dict(raw)))
        creds = Credentials.from_service_account_info(
            creds_info,
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        service = build("drive", "v3", credentials=creds)
        req = service.files().get_media(fileId=GDRIVE_FILE_ID)
        buf = BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)
        return pd.read_excel(buf, sheet_name=None)
    except Exception as e:
        st.error(f"Could not load store data from Google Drive: {e}")

    # Local fallback
    local = Path("exports")
    files = sorted(local.glob("store_analysis_*.xlsx"), reverse=True) if local.exists() else []
    if files:
        return pd.read_excel(files[0], sheet_name=None)

    st.error("No store analysis file found. Run store_export.py first.")
    return None


def fmt_npr(val):
    if pd.isna(val) or val == 0:
        return "—"
    if val >= 1_000_000:
        return f"NPR {val/1_000_000:.1f}M"
    return f"NPR {val/1_000:.0f}K"

def metric(label, value, delta=None):
    st.metric(label, value, delta)


# ── Main ───────────────────────────────────────────────────────────────────────
sheets = load_data()
if not sheets:
    st.stop()

df_overview  = sheets.get("📊 Store Overview",  pd.DataFrame())
df_brand_store = sheets.get("🏷️ Brand × Store", pd.DataFrame())
df_monthly   = sheets.get("📅 Monthly by Store", pd.DataFrame())
df_top       = sheets.get("🏆 Top Products by Store", pd.DataFrame())
df_categ     = sheets.get("🗂️ Category × Store", pd.DataFrame())
df_building  = sheets.get("🏢 Building Summary", pd.DataFrame())

# Clean overview
df_overview  = df_overview.dropna(subset=["Building"])
df_overview  = df_overview[df_overview["Building"] != "TOTAL"]

# ── Sidebar filters ────────────────────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/color/96/shop.png", width=60)
st.sidebar.title("🏪 Store Intelligence")
st.sidebar.markdown("---")

# Building filter
buildings = ["All"] + sorted(df_overview["Building"].dropna().unique().tolist())
sel_building = st.sidebar.selectbox("📍 Building / Location", buildings)

# Brand filter (from Brand × Store sheet)
if not df_brand_store.empty:
    brands_list = ["All"] + [b for b in df_brand_store.iloc[:, 0].dropna().tolist() if b != "Unknown"]
else:
    brands_list = ["All"]
sel_brand = st.sidebar.selectbox("🏷️ Brand", brands_list)

# Month range filter
if not df_monthly.empty:
    months = [m for m in df_monthly.iloc[:, 0].dropna().tolist() if m and str(m) != "nan"]
    months = sorted([str(m) for m in months])
    if months:
        col1, col2 = st.sidebar.columns(2)
        m_from = col1.selectbox("From", months, index=0)
        m_to   = col2.selectbox("To",   months, index=len(months)-1)
    else:
        m_from = m_to = None
else:
    m_from = m_to = None

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.rerun()

# ── Filter overview by building ────────────────────────────────────────────────
ov = df_overview.copy()
if sel_building != "All":
    ov = ov[ov["Building"] == sel_building]

# ── Page title ─────────────────────────────────────────────────────────────────
st.title("🏪 Store Intelligence Dashboard")
if sel_building != "All":
    st.caption(f"Filtered: **{sel_building}** · {len(ov)} store(s)")
else:
    st.caption("All locations · Sep 2024 → May 2026")

# ── Top KPI metrics ────────────────────────────────────────────────────────────
total_rev  = ov["Total Revenue (NPR)"].sum()
total_units = ov["Total Units"].sum()
total_orders = ov["Total Orders"].sum()
avg_aov = total_rev / total_orders if total_orders else 0
n_stores = len(ov)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("💰 Total Revenue", fmt_npr(total_rev))
c2.metric("📦 Units Sold", f"{int(total_units):,}")
c3.metric("🧾 Orders", f"{int(total_orders):,}")
c4.metric("🛒 Avg Order Value", f"NPR {avg_aov:,.0f}")
c5.metric("🏪 Stores", n_stores)

st.markdown("---")

# ── Store Overview Table ───────────────────────────────────────────────────────
st.subheader("📊 Store Performance")

ov_display = ov[["Building", "Floor", "POS Terminal", "Brands Present",
                  "Total Revenue (NPR)", "Total Units", "Total Orders",
                  "Avg Order Value", "Revenue Share %"]].copy()
ov_display["Total Revenue (NPR)"] = ov_display["Total Revenue (NPR)"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
ov_display["Avg Order Value"]     = ov_display["Avg Order Value"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
ov_display["Total Units"]         = ov_display["Total Units"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")
ov_display["Total Orders"]        = ov_display["Total Orders"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")
ov_display["Revenue Share %"]     = ov_display["Revenue Share %"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
ov_display["Floor"] = ov_display["Floor"].fillna("")

st.dataframe(ov_display, use_container_width=True, hide_index=True)

st.markdown("---")

# ── Brand × Store Revenue ──────────────────────────────────────────────────────
st.subheader("🏷️ Brand Revenue by Store")

if not df_brand_store.empty:
    bdf = df_brand_store.copy()
    bdf.columns = [str(c) for c in bdf.columns]
    brand_col = bdf.columns[0]

    # Filter brand
    if sel_brand != "All":
        bdf = bdf[bdf[brand_col] == sel_brand]

    # Filter stores (columns) by building
    if sel_building != "All":
        keep_cols = [brand_col, "TOTAL"]
        for col in bdf.columns:
            if sel_building.lower() in col.lower():
                keep_cols.append(col)
        bdf = bdf[[c for c in keep_cols if c in bdf.columns]]

    # Format numbers
    num_cols = [c for c in bdf.columns if c != brand_col]
    bdf_display = bdf.copy()
    for col in num_cols:
        bdf_display[col] = bdf_display[col].apply(
            lambda x: f"NPR {x:,.0f}" if pd.notna(x) and x != 0 else "—")

    st.dataframe(bdf_display, use_container_width=True, hide_index=True)

st.markdown("---")

# ── Monthly Trend ──────────────────────────────────────────────────────────────
st.subheader("📅 Monthly Revenue Trend")

if not df_monthly.empty and m_from and m_to:
    mdf = df_monthly.copy()
    mdf.columns = [str(c) for c in mdf.columns]
    month_col = mdf.columns[0]
    mdf = mdf[mdf[month_col].notna()]
    mdf[month_col] = mdf[month_col].astype(str)
    mdf = mdf[(mdf[month_col] >= m_from) & (mdf[month_col] <= m_to)]

    # Filter by building (columns)
    if sel_building != "All":
        keep = [month_col, "TOTAL"]
        for col in mdf.columns:
            if sel_building.lower() in col.lower():
                keep.append(col)
        mdf = mdf[[c for c in keep if c in mdf.columns]]

    # Chart — melt for line chart
    num_cols = [c for c in mdf.columns if c != month_col and c != "TOTAL"]
    if num_cols:
        chart_df = mdf[[month_col] + num_cols].set_index(month_col)
        chart_df = chart_df.apply(pd.to_numeric, errors="coerce").fillna(0)
        st.line_chart(chart_df, use_container_width=True)

    # Table
    mdf_display = mdf.copy()
    for col in [c for c in mdf.columns if c != month_col]:
        mdf_display[col] = mdf_display[col].apply(
            lambda x: f"NPR {x:,.0f}" if pd.notna(x) and x else "—")
    st.dataframe(mdf_display, use_container_width=True, hide_index=True)

st.markdown("---")

# ── Top Products ───────────────────────────────────────────────────────────────
st.subheader("🏆 Top Products by Store")

if not df_top.empty:
    tdf = df_top.copy()

    # Forward-fill Store column
    tdf["Store"] = tdf["Store"].ffill()

    # Filter by building
    if sel_building != "All":
        tdf = tdf[tdf["Store"].str.contains(sel_building, case=False, na=False)]

    # Filter by brand
    if sel_brand != "All":
        tdf = tdf[tdf["Brand"] == sel_brand]

    if not tdf.empty:
        tdf = tdf.dropna(subset=["Product"])
        tdf["Revenue (NPR)"] = tdf["Revenue (NPR)"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
        tdf["Units Sold"]    = tdf["Units Sold"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")
        st.dataframe(tdf[["Store", "Rank", "Product", "Brand", "Category",
                           "Revenue (NPR)", "Units Sold"]],
                     use_container_width=True, hide_index=True)
    else:
        st.info("No products match current filters.")

st.markdown("---")

# ── Category × Store ───────────────────────────────────────────────────────────
st.subheader("🗂️ Category Breakdown by Store")

if not df_categ.empty:
    cdf = df_categ.copy()
    cdf.columns = [str(c) for c in cdf.columns]
    categ_col = cdf.columns[0]

    if sel_building != "All":
        keep = [categ_col, "TOTAL"]
        for col in cdf.columns:
            if sel_building.lower() in col.lower():
                keep.append(col)
        cdf = cdf[[c for c in keep if c in cdf.columns]]

    cdf_display = cdf.copy()
    for col in [c for c in cdf.columns if c != categ_col]:
        cdf_display[col] = cdf_display[col].apply(
            lambda x: f"NPR {x:,.0f}" if pd.notna(x) and x else "—")

    st.dataframe(cdf_display, use_container_width=True, hide_index=True)

st.markdown("---")

# ── Building Summary ───────────────────────────────────────────────────────────
st.subheader("🏢 Brand Split per Building")

if not df_building.empty:
    bld = df_building.copy()

    if sel_building != "All":
        # Find rows for this building
        bld["_building"] = bld["Building"].ffill()
        bld = bld[bld["_building"].str.upper() == sel_building.upper()]
        bld = bld.drop(columns=["_building"])

    bld = bld.dropna(subset=["Brand"])
    bld["Revenue (NPR)"] = bld["Revenue (NPR)"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
    bld["Units Sold"]    = bld["Units Sold"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")
    bld["Orders"]        = bld["Orders"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")
    bld["% of Building"] = bld["% of Building"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
    bld["Building"] = bld["Building"].fillna("")
    bld["Floor"]    = bld["Floor"].fillna("")

    st.dataframe(bld[["Building", "Floor", "Brand", "Revenue (NPR)", "Units Sold", "Orders", "% of Building"]],
                 use_container_width=True, hide_index=True)