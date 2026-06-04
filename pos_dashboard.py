import streamlit as st
import pandas as pd
import json
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Salt Fashion — POS Analysis",
    page_icon="📊", layout="wide",
    initial_sidebar_state="expanded",
)

GDRIVE_FILE_ID = "1YcW30p_dUfeeaQj-XXmGhMHP0ldAM32X"

LOCATION_ORDER = [
    "Baneshwor", "Lazimpat", "Kumaripati", "Chitwan", "Pokhara", "Online",
    "Baneshwor Lush", "Chitwan Lush", "Pokhara Lush",
]

LOC_COLORS = {
    "Baneshwor":      "#3B82F6",
    "Lazimpat":       "#10B981",
    "Kumaripati":     "#F59E0B",
    "Chitwan":        "#EF4444",
    "Pokhara":        "#8B5CF6",
    "Online":         "#F97316",
    "Baneshwor Lush": "#06B6D4",
    "Chitwan Lush":   "#EC4899",
    "Pokhara Lush":   "#6366F1",
}


@st.cache_data(ttl=300)
def load_data():
    # Google Drive
    if "gcp_service_account" in st.secrets:
        try:
            from google.oauth2.service_account import Credentials
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaIoBaseDownload
            import json as _j

            raw = st.secrets["gcp_service_account"]
            creds_info = _j.loads(_j.dumps(dict(raw)))
            creds = Credentials.from_service_account_info(
                creds_info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
            service = build("drive", "v3", credentials=creds)
            req = service.files().get_media(fileId=GDRIVE_FILE_ID)
            buf = BytesIO()
            dl = MediaIoBaseDownload(buf, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            buf.seek(0)
            df = pd.read_excel(buf, sheet_name="Point of Sale Analysis", engine="openpyxl")
            return df
        except Exception as e:
            st.warning(f"Google Drive load failed: {e}")

    # Local fallback
    files = sorted(Path("exports").glob("pos_analysis_*.xlsx"), reverse=True) if Path("exports").exists() else []
    if files:
        return pd.read_excel(files[0], sheet_name="Point of Sale Analysis", engine="openpyxl")

    return None


def fmt_npr(val):
    if pd.isna(val) or val == 0: return "—"
    if val >= 1_000_000: return f"NPR {val/1_000_000:.1f}M"
    if val >= 1_000:     return f"NPR {val/1_000:.0f}K"
    return f"NPR {val:,.0f}"


# ── Load & clean ──────────────────────────────────────────────────────────────
df_raw = load_data()
if df_raw is None:
    st.error("No POS analysis file found. Run pos_analysis_export.py first.")
    st.stop()

df = df_raw.copy()
df.columns = [str(c).strip() for c in df.columns]

# Remove TOTAL row
df = df[df["Location"] != "TOTAL"]
df = df.dropna(subset=["Location"])

# Parse date
df["Date"] = pd.to_datetime(df["Total"], errors="coerce")
df = df.dropna(subset=["Date"])
df["Month"]   = df["Date"].dt.to_period("M").astype(str)
df["Year"]    = df["Date"].dt.year
df["Quarter"] = df["Date"].dt.to_period("Q").astype(str)

# Rename columns to standard names
df = df.rename(columns={
    "Ticket Sold": "Tickets",
    "QTY":         "Units",
    "Sales Amount":"Revenue",
    "Footfall":    "Footfall",
    "ATV":         "ATV",
    "UPT":         "UPT",
    "Brand":       "Brand",
})

# Ensure numeric
for col in ["Tickets", "Units", "Revenue", "Footfall", "ATV", "UPT"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.title("📊 POS Analysis")
st.sidebar.markdown("---")

# Brand filter
brands = ["All"] + sorted(df["Brand"].dropna().unique().tolist()) if "Brand" in df.columns else ["All"]
sel_brand = st.sidebar.selectbox("🏷️ Brand", brands)

# Location filter
all_locs = [l for l in LOCATION_ORDER if l in df["Location"].unique()]
sel_locs = st.sidebar.multiselect("📍 Location", all_locs, default=all_locs)

# Date range
min_date = df["Date"].min().date()
max_date = df["Date"].max().date()
d_from, d_to = st.sidebar.date_input("📅 Date Range",
    value=(min_date, max_date), min_value=min_date, max_value=max_date)

# Group by
group_by = st.sidebar.radio("📆 Group By", ["Daily", "Monthly", "Quarterly"])

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh"):
    st.cache_data.clear()
    st.rerun()

# ── Apply filters ─────────────────────────────────────────────────────────────
dff = df.copy()
if sel_brand != "All" and "Brand" in dff.columns:
    dff = dff[dff["Brand"] == sel_brand]
if sel_locs:
    dff = dff[dff["Location"].isin(sel_locs)]
dff = dff[(dff["Date"].dt.date >= d_from) & (dff["Date"].dt.date <= d_to)]

# ── Title ─────────────────────────────────────────────────────────────────────
st.title("📊 POS Analysis Dashboard")
brand_tag = f" · {sel_brand}" if sel_brand != "All" else ""
st.caption(f"{d_from} → {d_to}{brand_tag} · {len(sel_locs)} location(s)")

if dff.empty:
    st.warning("No data for selected filters.")
    st.stop()

# ── KPI metrics ───────────────────────────────────────────────────────────────
total_rev     = dff["Revenue"].sum()
total_tickets = dff["Tickets"].sum()
total_units   = dff["Units"].sum()
total_footfall= dff["Footfall"].sum() if dff["Footfall"].notna().any() else 0
avg_atv       = dff["Revenue"].sum() / dff["Tickets"].sum() if dff["Tickets"].sum() else 0
avg_upt       = dff["Units"].sum() / dff["Tickets"].sum() if dff["Tickets"].sum() else 0
conversion    = (total_tickets / total_footfall * 100) if total_footfall else None

cols = st.columns(7)
kpis = [
    ("💰 Revenue",     fmt_npr(total_rev)),
    ("🧾 Tickets Sold",f"{int(total_tickets):,}"),
    ("📦 Units Sold",  f"{int(total_units):,}"),
    ("👣 Footfall",    f"{int(total_footfall):,}" if total_footfall else "—"),
    ("🛒 ATV",         f"NPR {avg_atv:,.0f}"),
    ("📦 UPT",         f"{avg_upt:.2f}"),
    ("🎯 Conversion",  f"{conversion:.1f}%" if conversion else "—"),
]
for col, (label, val) in zip(cols, kpis):
    col.markdown(f"**{label}**")
    col.markdown(f"### {val}")

st.markdown("---")

# ── Period grouping ───────────────────────────────────────────────────────────
period_col = {"Daily": "Date", "Monthly": "Month", "Quarterly": "Quarter"}[group_by]

tab1, tab2, tab3, tab4 = st.tabs(["📈 Trends", "🏪 By Location", "📋 Data Table", "🔢 Location Summary"])

# ── Tab 1: Trends ─────────────────────────────────────────────────────────────
with tab1:
    grp = dff.groupby(period_col).agg(
        Revenue=("Revenue", "sum"),
        Tickets=("Tickets", "sum"),
        Units=("Units", "sum"),
        Footfall=("Footfall", "sum"),
    ).reset_index()
    grp["ATV"] = grp["Revenue"] / grp["Tickets"].replace(0, pd.NA)
    grp["UPT"] = grp["Units"]   / grp["Tickets"].replace(0, pd.NA)
    grp["Conversion"] = grp["Tickets"] / grp["Footfall"].replace(0, pd.NA) * 100

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("💰 Revenue Trend")
        st.line_chart(grp.set_index(period_col)[["Revenue"]], use_container_width=True)
    with c2:
        st.subheader("🧾 Tickets Sold Trend")
        st.line_chart(grp.set_index(period_col)[["Tickets"]], use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.subheader("🛒 ATV Trend")
        st.line_chart(grp.set_index(period_col)[["ATV"]], use_container_width=True)
    with c4:
        st.subheader("📦 UPT Trend")
        st.line_chart(grp.set_index(period_col)[["UPT"]], use_container_width=True)

    if grp["Footfall"].notna().any() and grp["Footfall"].sum() > 0:
        st.subheader("🎯 Conversion Rate Trend")
        st.line_chart(grp.set_index(period_col)[["Conversion"]], use_container_width=True)

# ── Tab 2: By Location ────────────────────────────────────────────────────────
with tab2:
    loc_grp = dff.groupby(["Location", period_col]).agg(
        Revenue=("Revenue", "sum"),
        Tickets=("Tickets", "sum"),
    ).reset_index()

    st.subheader("💰 Revenue by Location")
    pivot_rev = loc_grp.pivot_table(index=period_col, columns="Location",
                                     values="Revenue", aggfunc="sum").fillna(0)
    # Order columns by LOCATION_ORDER
    ordered_cols = [l for l in LOCATION_ORDER if l in pivot_rev.columns]
    st.line_chart(pivot_rev[ordered_cols], use_container_width=True)

    st.subheader("🧾 Tickets by Location")
    pivot_tkt = loc_grp.pivot_table(index=period_col, columns="Location",
                                     values="Tickets", aggfunc="sum").fillna(0)
    ordered_cols2 = [l for l in LOCATION_ORDER if l in pivot_tkt.columns]
    st.line_chart(pivot_tkt[ordered_cols2], use_container_width=True)

# ── Tab 3: Data Table ─────────────────────────────────────────────────────────
with tab3:
    st.subheader("📋 Raw Data")
    show_cols = [c for c in ["Date", "Location", "Brand", "Tickets", "Units",
                              "Revenue", "Footfall", "ATV", "UPT"] if c in dff.columns]
    display = dff[show_cols].copy().sort_values(["Date", "Location"], ascending=[False, True])
    display["Revenue"] = display["Revenue"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
    display["ATV"]     = display["ATV"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
    display["UPT"]     = display["UPT"].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "—")
    display["Footfall"]= display["Footfall"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")
    st.dataframe(display, use_container_width=True, hide_index=True)

# ── Tab 4: Location Summary ───────────────────────────────────────────────────
with tab4:
    st.subheader("🔢 Location Summary")
    summ = dff.groupby("Location").agg(
        Days=("Date", "nunique"),
        Revenue=("Revenue", "sum"),
        Tickets=("Tickets", "sum"),
        Units=("Units", "sum"),
        Footfall=("Footfall", "sum"),
    ).reset_index()
    summ["ATV"]        = summ["Revenue"] / summ["Tickets"].replace(0, pd.NA)
    summ["UPT"]        = summ["Units"]   / summ["Tickets"].replace(0, pd.NA)
    summ["Rev/Day"]    = summ["Revenue"] / summ["Days"].replace(0, pd.NA)
    summ["Conversion"] = summ["Tickets"] / summ["Footfall"].replace(0, pd.NA) * 100

    # Sort by LOCATION_ORDER
    summ["_order"] = summ["Location"].apply(
        lambda x: LOCATION_ORDER.index(x) if x in LOCATION_ORDER else 99)
    summ = summ.sort_values("_order").drop(columns=["_order"])

    summ_display = summ.copy()
    summ_display["Revenue"]    = summ_display["Revenue"].apply(lambda x: f"NPR {x:,.0f}")
    summ_display["ATV"]        = summ_display["ATV"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
    summ_display["UPT"]        = summ_display["UPT"].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "—")
    summ_display["Rev/Day"]    = summ_display["Rev/Day"].apply(lambda x: f"NPR {x:,.0f}" if pd.notna(x) else "—")
    summ_display["Tickets"]    = summ_display["Tickets"].apply(lambda x: f"{int(x):,}")
    summ_display["Units"]      = summ_display["Units"].apply(lambda x: f"{int(x):,}")
    summ_display["Footfall"]   = summ_display["Footfall"].apply(lambda x: f"{int(x):,}" if pd.notna(x) and x > 0 else "—")
    summ_display["Conversion"] = summ_display["Conversion"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")

    st.dataframe(summ_display[["Location", "Days", "Revenue", "Tickets", "Units",
                                "Footfall", "ATV", "UPT", "Rev/Day", "Conversion"]],
                 use_container_width=True, hide_index=True)
