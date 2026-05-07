import streamlit as st
import pandas as pd
import requests
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from io import BytesIO

st.set_page_config(
    page_title="Variant Intelligence",
    page_icon="👗", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container{padding:1.5rem 2rem}
.metric-card{background:white;border-radius:10px;padding:12px 16px;
             border:1px solid #e8edf3;text-align:center}
.metric-val{font-size:24px;font-weight:600;margin:0}
.metric-lbl{font-size:11px;color:#6b7280;margin:0;margin-top:2px}
.divider{border-top:1px solid #e5e7eb;margin:12px 0}
.insight{background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;
         padding:10px 14px;font-size:13px;color:#1e40af;margin-bottom:12px}
.sec-title{font-size:16px;font-weight:600;color:#111827;margin-bottom:8px}
</style>
""", unsafe_allow_html=True)

# ── FILE ID — your variant_analysis.xlsx on Google Drive ─────────────────────
VARIANT_FILE_ID = "1LMzZJpQZo2NHOqtn4d00RNwBtngA-L9C"
# ─────────────────────────────────────────────────────────────────────────────

STR_ORDER  = ["Super Fast","Fast","Medium","Slow","Dead"]
SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL","5XL",
              "Free Size","One Size",
              "26","27","28","29","30","31","32","33","34","36","38","40","42"]

def calc_str(sold, stock):
    sold  = float(sold  or 0)
    stock = max(0.0, float(stock or 0))
    total = sold + stock
    if total <= 0 or sold <= 0: return 0.0
    return min(round(sold / total * 100, 1), 100.0)

def str_status(pct):
    if pct >= 95: return "Super Fast"
    if pct >= 70: return "Fast"
    if pct >= 30: return "Medium"
    if pct >  0:  return "Slow"
    return "Dead"

# ── Download .xlsx from Google Drive (works for real xlsx, not just Sheets) ───
@st.cache_data(ttl=300, show_spinner=False)
def load_sheets():
    """
    Downloads variant_analysis.xlsx via the Drive files.get API
    using service account credentials from Streamlit secrets.
    Returns (size_df, color_df, error_string)
    """
    try:
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        # Refresh token
        creds.refresh(Request())

        url  = f"https://www.googleapis.com/drive/v3/files/{VARIANT_FILE_ID}?alt=media"
        resp = requests.get(
            url,
            headers={"Authorization": f"Bearer {creds.token}"},
            timeout=180,
        )
        if resp.status_code != 200:
            return None, None, (
                f"Google Drive returned HTTP {resp.status_code}.\n"
                f"Response: {resp.text[:400]}\n\n"
                f"Make sure the file is shared with:\n"
                f"salt-dashboard@salt-dashboard-494810.iam.gserviceaccount.com"
            )

        bio      = BytesIO(resp.content)
        size_df  = pd.read_excel(bio, sheet_name="Size Breakdown",  engine="openpyxl")
        bio.seek(0)
        color_df = pd.read_excel(bio, sheet_name="Color Breakdown", engine="openpyxl")
        return size_df, color_df, None

    except Exception as e:
        return None, None, str(e)

# ── Load ──────────────────────────────────────────────────────────────────────
st.markdown("## 👗 Salt Fashion — Variant Intelligence")
st.caption("Size × Color breakdown — which variants sell fastest")

with st.spinner("Downloading variant_analysis.xlsx from Google Drive..."):
    size_df, color_df, err = load_sheets()

if err or size_df is None:
    st.error(f"**Could not load data.**\n\n{err}")
    st.stop()

# Clean column names (strip whitespace)
size_df.columns  = [str(c).strip() for c in size_df.columns]
color_df.columns = [str(c).strip() for c in color_df.columns]

# ── Show what columns exist (debug — remove after confirmed working) ──────────
with st.expander("🔍 Debug: column names in file"):
    st.write("**Size Breakdown columns:**", list(size_df.columns))
    st.write("**Color Breakdown columns:**", list(color_df.columns))
    st.write(f"Size rows: {len(size_df):,}  |  Color rows: {len(color_df):,}")

# ── Detect columns flexibly ───────────────────────────────────────────────────
def find(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    return None

size_col    = find(size_df,  ["Size","size_value","Size Value"])
color_col   = find(color_df, ["Color","color_value","Color Value"])
brand_col_s = find(size_df,  ["Brand","brand"])
brand_col_c = find(color_df, ["Brand","brand"])
sold_s      = find(size_df,  ["Units Sold","units_sold","Total Units Sold","Sold"])
sold_c      = find(color_df, ["Units Sold","units_sold","Total Units Sold","Sold"])
stock_s     = find(size_df,  ["In Stock","on_hand_qty","In_Stock","Stock","On Hand"])
stock_c     = find(color_df, ["In Stock","on_hand_qty","In_Stock","Stock","On Hand"])
str_s       = find(size_df,  ["STR %","STR_%","Sell_Through_%","Sell Through %"])
str_c       = find(color_df, ["STR %","STR_%","Sell_Through_%","Sell Through %"])
stat_s      = find(size_df,  ["Status","STR Status","Velocity"])
stat_c      = find(color_df, ["Status","STR Status","Velocity"])
name_s      = find(size_df,  ["Product Name","product_name","Name"])
name_c      = find(color_df, ["Product Name","product_name","Name"])

# Numeric coerce
for df, cols in [(size_df, [sold_s, stock_s, str_s]),
                 (color_df,[sold_c, stock_c, str_c])]:
    for col in cols:
        if col:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace("%","").str.replace(",",""),
                errors="coerce").fillna(0)

# Derive STR / Status if columns missing
if not str_s and sold_s and stock_s:
    size_df["STR %"] = size_df.apply(
        lambda r: calc_str(r[sold_s], r[stock_s]), axis=1); str_s = "STR %"
if not str_c and sold_c and stock_c:
    color_df["STR %"] = color_df.apply(
        lambda r: calc_str(r[sold_c], r[stock_c]), axis=1); str_c = "STR %"
if not stat_s and str_s:
    size_df["Status"]  = size_df[str_s].apply(str_status);  stat_s = "Status"
if not stat_c and str_c:
    color_df["Status"] = color_df[str_c].apply(str_status); stat_c = "Status"

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📊 Filters")

    # Brand selector
    brands = ["All Brands"]
    for bc in [brand_col_s, brand_col_c]:
        if bc:
            blist = [str(x) for x in size_df[bc].dropna().unique()
                     if str(x) not in ("","nan","None","False")]
            if blist:
                brands = ["All Brands"] + sorted(set(blist))
                break
    sel_brand = st.selectbox("Brand", brands)

    view = st.radio("View", [
        "📏 Size Performance",
        "🎨 Color Performance",
        "⚡ Top Performers",
        "🚨 Dead Stock Alert",
    ])
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    sel_status = []
    st.markdown("**STR Status**")
    for s in STR_ORDER:
        if st.checkbox(s, value=True, key=f"chk_{s}"):
            sel_status.append(s)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    search = st.text_input("Search product", placeholder="e.g. Kurti, Jeans...")
    if st.button("🔄 Refresh"):
        st.cache_data.clear(); st.rerun()

# ── Filter ────────────────────────────────────────────────────────────────────
def apply_filters(df, bc, nc, sc):
    r = df.copy()
    if sel_brand != "All Brands" and bc and bc in r.columns:
        r = r[r[bc].astype(str) == sel_brand]
    if sc and sel_status:
        r = r[r[sc].isin(sel_status)]
    if search.strip() and nc and nc in r.columns:
        r = r[r[nc].str.contains(search.strip(), case=False, na=False)]
    return r

sf = apply_filters(size_df,  brand_col_s, name_s, stat_s)
cf = apply_filters(color_df, brand_col_c, name_c, stat_c)

# ── Top metrics ───────────────────────────────────────────────────────────────
total_sold  = int(sf[sold_s].sum())  if sold_s  else 0
total_stock = int(sf[stock_s].sum()) if stock_s else 0
overall_str = calc_str(total_sold, total_stock)
sf_cnt = int((sf[stat_s]  == "Super Fast").sum()) if stat_s  else 0
cf_cnt = int((cf[stat_c]  == "Super Fast").sum()) if stat_c  else 0

for col, val, lbl, clr in zip(
    st.columns(5),
    [f"{total_sold:,}", f"{total_stock:,}", f"{overall_str:.1f}%", sf_cnt, cf_cnt],
    ["Total Units Sold","Units In Stock","Overall STR","⚡ Super Fast Sizes","⚡ Super Fast Colors"],
    ["#1d4ed8","#374151","#1B5E20","#1B5E20","#1B5E20"],
):
    col.markdown(
        f'<div class="metric-card">'
        f'<p class="metric-val" style="color:{clr}">{val}</p>'
        f'<p class="metric-lbl">{lbl}</p></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
if view == "📏 Size Performance":
    st.markdown('<div class="sec-title">📏 Size Performance</div>', unsafe_allow_html=True)
    if not size_col or not sold_s:
        st.warning("Size or Units Sold column not found in file."); st.stop()

    agg = sf.groupby(size_col).agg(
        Units_Sold =(sold_s,  "sum"),
        In_Stock   =(stock_s, "sum")   if stock_s else (sold_s,"count"),
        Products   =(name_s,  "nunique") if name_s else (sold_s,"count"),
    ).reset_index()
    agg["STR_%"]  = agg.apply(lambda r: calc_str(r["Units_Sold"], r["In_Stock"]), axis=1)
    agg["Status"] = agg["STR_%"].apply(str_status)
    agg = agg.sort_values("Units_Sold", ascending=False)

    top  = agg.iloc[0][size_col] if len(agg) else "N/A"
    best = agg.loc[agg["STR_%"].idxmax(), size_col] if len(agg) else "N/A"
    st.markdown(
        f'<div class="insight">💡 Best selling size: <b>{top}</b> by units &nbsp;|&nbsp; '
        f'Best STR%: <b>{best}</b></div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Units Sold by Size**")
        chart = agg.set_index(size_col)["Units_Sold"]
        ordered = [s for s in SIZE_ORDER if s in chart.index] + \
                  [s for s in chart.index if s not in SIZE_ORDER]
        st.bar_chart(chart.reindex(ordered).dropna())
    with c2:
        st.markdown("**STR % by Size**")
        chart2 = agg.set_index(size_col)["STR_%"]
        ordered2 = [s for s in SIZE_ORDER if s in chart2.index] + \
                   [s for s in chart2.index if s not in SIZE_ORDER]
        st.bar_chart(chart2.reindex(ordered2).dropna())

    st.markdown("**Full Size Table**")
    disp = agg[[size_col,"Units_Sold","In_Stock","STR_%","Status","Products"]].copy()
    disp.columns = ["Size","Units Sold","In Stock","STR %","Status","Products"]
    st.dataframe(disp, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
elif view == "🎨 Color Performance":
    st.markdown('<div class="sec-title">🎨 Color Performance</div>', unsafe_allow_html=True)
    if not color_col or not sold_c:
        st.warning("Color or Units Sold column not found in file."); st.stop()

    agg = cf.groupby(color_col).agg(
        Units_Sold =(sold_c,  "sum"),
        In_Stock   =(stock_c, "sum")   if stock_c else (sold_c,"count"),
        Products   =(name_c,  "nunique") if name_c else (sold_c,"count"),
    ).reset_index()
    agg["STR_%"]  = agg.apply(lambda r: calc_str(r["Units_Sold"], r["In_Stock"]), axis=1)
    agg["Status"] = agg["STR_%"].apply(str_status)
    agg = agg.sort_values("Units_Sold", ascending=False)

    top  = agg.iloc[0][color_col] if len(agg) else "N/A"
    best = agg.loc[agg["STR_%"].idxmax(), color_col] if len(agg) else "N/A"
    st.markdown(
        f'<div class="insight">💡 Best selling color: <b>{top}</b> &nbsp;|&nbsp; '
        f'Best STR%: <b>{best}</b></div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Top 20 Colors by Units Sold**")
        st.bar_chart(agg.head(20).set_index(color_col)["Units_Sold"])
    with c2:
        st.markdown("**Top 20 Colors by STR %**")
        st.bar_chart(agg.nlargest(20,"STR_%").set_index(color_col)["STR_%"])

    st.markdown("**Full Color Table**")
    disp = agg[[color_col,"Units_Sold","In_Stock","STR_%","Status","Products"]].copy()
    disp.columns = ["Color","Units Sold","In Stock","STR %","Status","Products"]
    st.dataframe(disp, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
elif view == "⚡ Top Performers":
    st.markdown('<div class="sec-title">⚡ Super Fast Variants</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**🏆 Top Sizes — Super Fast**")
        if stat_s and size_col:
            top_sf = sf[sf[stat_s]=="Super Fast"]
            if sold_s: top_sf = top_sf.sort_values(sold_s, ascending=False)
            show = [c for c in [name_s, brand_col_s, size_col, sold_s, stock_s, str_s] if c]
            st.dataframe(top_sf[show].head(30), use_container_width=True, hide_index=True)
    with c2:
        st.markdown("**🏆 Top Colors — Super Fast**")
        if stat_c and color_col:
            top_cf = cf[cf[stat_c]=="Super Fast"]
            if sold_c: top_cf = top_cf.sort_values(sold_c, ascending=False)
            show = [c for c in [name_c, brand_col_c, color_col, sold_c, stock_c, str_c] if c]
            st.dataframe(top_cf[show].head(30), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════════════════════
elif view == "🚨 Dead Stock Alert":
    st.markdown('<div class="sec-title">🚨 Dead Stock — Needs Clearance</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Dead Sizes (never sold)**")
        if stat_s and size_col:
            dead = sf[sf[stat_s]=="Dead"]
            if stock_s: dead = dead.sort_values(stock_s, ascending=False)
            show = [c for c in [name_s, brand_col_s, size_col, sold_s, stock_s, str_s] if c]
            st.dataframe(dead[show].head(30), use_container_width=True, hide_index=True)
            if stock_s:
                st.warning(f"⚠️ {int(dead[stock_s].sum()):,} units stuck in dead sizes")
    with c2:
        st.markdown("**Dead Colors (never sold)**")
        if stat_c and color_col:
            dead = cf[cf[stat_c]=="Dead"]
            if stock_c: dead = dead.sort_values(stock_c, ascending=False)
            show = [c for c in [name_c, brand_col_c, color_col, sold_c, stock_c, str_c] if c]
            st.dataframe(dead[show].head(30), use_container_width=True, hide_index=True)
            if stock_c:
                st.warning(f"⚠️ {int(dead[stock_c].sum()):,} units stuck in dead colors")