import streamlit as st
import pandas as pd
import os
from io import BytesIO
from pathlib import Path
from collections import defaultdict

st.set_page_config(
    page_title="Salt Fashion — Variant Intelligence",
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
.badge{display:inline-block;padding:2px 10px;border-radius:12px;
       font-size:11px;font-weight:600}
.divider{border-top:1px solid #e5e7eb;margin:12px 0}
.insight{background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;
         padding:10px 14px;font-size:13px;color:#1e40af;margin-bottom:12px}
.sec-title{font-size:16px;font-weight:600;color:#111827;margin-bottom:8px}
</style>
""", unsafe_allow_html=True)

GDRIVE_FILE_ID = "1eUynsHj3U95qJAhuCtfQKDJ7pgrzDmsl"

STR_COLORS = {
    "Super Fast": ("#1B5E20","#FFFFFF"),
    "Fast":       ("#43A047","#FFFFFF"),
    "Medium":     ("#F9A825","#000000"),
    "Slow":       ("#E53935","#FFFFFF"),
    "Dead":       ("#424242","#FFFFFF"),
}
STR_ORDER = ["Super Fast","Fast","Medium","Slow","Dead"]

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

SIZE_ORDER = ["XS","S","M","L","XL","2XL","3XL","4XL","5XL",
              "Free Size","One Size",
              "26","27","28","29","30","31","32","33","34","36","38","40","42"]


# ── Load variant Excel from Google Drive ──────────────────────────────────────
@st.cache_data(ttl=300)
def load_variant_data():
    """Load variant_analysis.xlsx — tries Google Drive first, then local."""

    # ── 1. Google Drive (primary — works on Streamlit Cloud) ──────────────
    if "gcp_service_account" in st.secrets:
        try:
            from google.oauth2.service_account import Credentials
            import googleapiclient.discovery
            from googleapiclient.http import MediaIoBaseDownload

            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=["https://www.googleapis.com/auth/drive"]
            )
            svc     = googleapiclient.discovery.build("drive", "v3", credentials=creds)
            request = svc.files().get_media(fileId=GDRIVE_FILE_ID)
            buf     = BytesIO()
            dl      = MediaIoBaseDownload(buf, request)
            done    = False
            while not done:
                _, done = dl.next_chunk()
            buf.seek(0)
            size_df  = pd.read_excel(buf, sheet_name="Size Breakdown",  engine="openpyxl")
            buf.seek(0)
            color_df = pd.read_excel(buf, sheet_name="Color Breakdown", engine="openpyxl")
            return size_df, color_df, None
        except Exception as e:
            gdrive_err = str(e)
    else:
        gdrive_err = "No GCP secrets configured"

    # ── 2. Local fallback (works on your laptop) ──────────────────────────
    base   = r"C:\Users\Legion\Desktop\odoo_export"
    local  = os.path.join(base, "variant_analysis.xlsx")
    if os.path.exists(local):
        try:
            size_df  = pd.read_excel(local, sheet_name="Size Breakdown",  engine="openpyxl")
            color_df = pd.read_excel(local, sheet_name="Color Breakdown", engine="openpyxl")
            return size_df, color_df, None
        except Exception as e:
            return None, None, f"Local load failed: {e}"

    return None, None, (
        f"Could not load variant data.\n"
        f"Google Drive error: {gdrive_err}\n"
        f"Local file not found: {local}\n\n"
        f"Run: python variant_export.py"
    )


@st.cache_data(ttl=300)
def load_main_data():
    """Load main products Excel for brand/category filter — Google Drive first."""
    MAIN_FILE_ID = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"   # same as dashboard.py

    if "gcp_service_account" in st.secrets:
        try:
            from google.oauth2.service_account import Credentials
            import googleapiclient.discovery
            from googleapiclient.http import MediaIoBaseDownload

            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=["https://www.googleapis.com/auth/drive"]
            )
            svc     = googleapiclient.discovery.build("drive", "v3", credentials=creds)
            request = svc.files().get_media(fileId=MAIN_FILE_ID)
            buf     = BytesIO()
            dl      = MediaIoBaseDownload(buf, request)
            done    = False
            while not done:
                _, done = dl.next_chunk()
            buf.seek(0)
            df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl")
            df.columns = [c.strip() for c in df.columns]
            return df
        except Exception:
            pass

    # Local fallback
    base = r"C:\Users\Legion\Desktop\odoo_export"
    dirs = [os.path.join(base, "exports"), base]
    candidates = []
    for d in dirs:
        if os.path.exists(d):
            candidates += list(Path(d).glob("odoo_products*.xlsx"))
    if not candidates:
        return None
    latest = str(max(candidates, key=lambda f: f.stat().st_mtime))
    try:
        df = pd.read_excel(latest, sheet_name="Products", engine="openpyxl")
        df.columns = [c.strip() for c in df.columns]
        return df
    except:
        return None


def badge_html(status):
    bg, fg = STR_COLORS.get(status, ("#9E9E9E","#FFFFFF"))
    return f'<span class="badge" style="background:{bg};color:{fg}">{status}</span>'


def main():
    st.markdown("## 👗 Salt Fashion — Variant Intelligence")
    st.markdown("Size × Color × Type breakdown — which variants sell fastest")

    size_df, color_df, err = load_variant_data()
    main_df = load_main_data()

    if err or size_df is None:
        st.error(f"Could not load variant data")
        st.code(err or "Unknown error")
        st.info("""
        **To generate variant data:**
        1. Run: `python variant_export.py` on your laptop
        2. Upload `variant_analysis.xlsx` to Google Drive (Salt Dashboard Data folder)
        3. Share it with: `salt-dashboard@salt-dashboard-494810.iam.gserviceaccount.com`
        4. Refresh this page
        """)
        st.stop()

    # Clean columns
    size_df.columns  = [c.strip() for c in size_df.columns]
    color_df.columns = [c.strip() for c in color_df.columns]

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 📊 Variant Filters")
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        brands = ["All Brands"]
        if main_df is not None and "Brand" in main_df.columns:
            brands += sorted([b for b in main_df["Brand"].dropna().unique()
                              if str(b).strip() not in ("","nan","None","True","False")])
        sel_brand = st.selectbox("Brand", brands, index=0)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        view = st.radio("View", ["Size Performance","Color Performance",
                                 "Size × Color Matrix","Top Performers"])

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown("**Filter by STR Status**")
        sel_status = []
        for s in STR_ORDER:
            if st.checkbox(s, value=True, key=f"s_{s}"):
                sel_status.append(s)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        cats = ["All Categories"]
        if main_df is not None and "Category" in main_df.columns:
            bdf = main_df[main_df["Brand"] == sel_brand] \
                  if sel_brand != "All Brands" else main_df
            cats += sorted([str(c) for c in bdf["Category"].dropna().unique()
                           if str(c).strip() not in ("","nan","None")])
        sel_cat = st.selectbox("Category", cats, index=0)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        search = st.text_input("Search product", placeholder="e.g. Kurti, Jeans...")

        if st.button("🔄 Refresh"):
            st.cache_data.clear(); st.rerun()

    # ── Filter ────────────────────────────────────────────────────────────────
    def filter_df(df):
        result = df.copy()
        if "Status" in result.columns and sel_status:
            result = result[result["Status"].isin(sel_status)]
        if search.strip() and "Product Name" in result.columns:
            result = result[result["Product Name"].str.contains(
                search.strip(), case=False, na=False)]
        return result

    sf = filter_df(size_df)
    cf = filter_df(color_df)

    # ── Overall metrics ───────────────────────────────────────────────────────
    if "Units Sold" in size_df.columns and "In Stock" in size_df.columns:
        total_sold  = size_df["Units Sold"].sum()
        total_stock = size_df["In Stock"].sum()
        overall_str = calc_str(total_sold, total_stock)

        size_status_counts  = size_df["Status"].value_counts() \
                              if "Status" in size_df.columns else {}
        color_status_counts = color_df["Status"].value_counts() \
                              if "Status" in color_df.columns else {}

        c1,c2,c3,c4,c5 = st.columns(5)
        for col, val, lbl, clr in [
            (c1, f"{int(total_sold):,}",    "Total Units Sold",   "#1d4ed8"),
            (c2, f"{int(total_stock):,}",   "Units In Stock",     "#374151"),
            (c3, f"{overall_str:.1f}%",     "Overall STR",        "#1B5E20"),
            (c4, size_status_counts.get("Super Fast",0),  "Super Fast Sizes",  "#1B5E20"),
            (c5, color_status_counts.get("Super Fast",0), "Super Fast Colors", "#1B5E20"),
        ]:
            with col:
                st.markdown(
                    f'<div class="metric-card">'
                    f'<p class="metric-val" style="color:{clr}">{val}</p>'
                    f'<p class="metric-lbl">{lbl}</p></div>',
                    unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

    # ── Views ─────────────────────────────────────────────────────────────────
    if view == "Size Performance":
        st.markdown('<div class="sec-title">📏 Size Performance</div>',
                    unsafe_allow_html=True)
        if "Size" in sf.columns and "Units Sold" in sf.columns:
            size_agg = sf.groupby("Size").agg(
                Units_Sold=("Units Sold","sum"),
                In_Stock=("In Stock","sum"),
                Products=("Product Name","nunique")
            ).reset_index()
            size_agg["STR_%"] = size_agg.apply(
                lambda r: calc_str(r["Units_Sold"], r["In_Stock"]), axis=1)
            size_agg["Status"] = size_agg["STR_%"].apply(str_status)
            size_agg = size_agg.sort_values("Units_Sold", ascending=False)

            top_size      = size_agg.iloc[0]["Size"] if len(size_agg) > 0 else "N/A"
            best_str_size = size_agg.loc[size_agg["STR_%"].idxmax(), "Size"] \
                            if len(size_agg) > 0 else "N/A"
            st.markdown(
                f'<div class="insight">💡 Best selling size: <strong>{top_size}</strong> '
                f'by units sold. Best STR: <strong>{best_str_size}</strong>. '
                f'Use this to decide which sizes to reorder most.</div>',
                unsafe_allow_html=True)

            col1, col2 = st.columns([3,2])
            with col1:
                st.markdown("**Units Sold by Size**")
                chart_data = size_agg.set_index("Size")["Units_Sold"]
                ordered = [s for s in SIZE_ORDER if s in chart_data.index]
                others  = [s for s in chart_data.index if s not in SIZE_ORDER]
                st.bar_chart(chart_data.reindex(ordered + others).dropna())
            with col2:
                st.markdown("**STR % by Size**")
                str_data = size_agg.set_index("Size")["STR_%"]
                ordered2 = [s for s in SIZE_ORDER if s in str_data.index]
                others2  = [s for s in str_data.index if s not in SIZE_ORDER]
                st.bar_chart(str_data.reindex(ordered2 + others2).dropna())

            st.markdown("**Size Breakdown Table**")
            display = size_agg[["Size","Units_Sold","In_Stock","STR_%","Status","Products"]].copy()
            display.columns = ["Size","Units Sold","In Stock","STR %","Status","Products"]
            st.dataframe(display, use_container_width=True, hide_index=True)

    elif view == "Color Performance":
        st.markdown('<div class="sec-title">🎨 Color Performance</div>',
                    unsafe_allow_html=True)
        if "Color" in cf.columns and "Units Sold" in cf.columns:
            color_agg = cf.groupby("Color").agg(
                Units_Sold=("Units Sold","sum"),
                In_Stock=("In Stock","sum"),
                Products=("Product Name","nunique")
            ).reset_index()
            color_agg["STR_%"] = color_agg.apply(
                lambda r: calc_str(r["Units_Sold"], r["In_Stock"]), axis=1)
            color_agg["Status"] = color_agg["STR_%"].apply(str_status)
            color_agg = color_agg.sort_values("Units_Sold", ascending=False)

            top_color      = color_agg.iloc[0]["Color"] if len(color_agg) > 0 else "N/A"
            best_str_color = color_agg.loc[color_agg["STR_%"].idxmax(), "Color"] \
                             if len(color_agg) > 0 else "N/A"
            st.markdown(
                f'<div class="insight">💡 Best selling color: <strong>{top_color}</strong> '
                f'by units sold. Best STR: <strong>{best_str_color}</strong>. '
                f'Focus buying on these colors next season.</div>',
                unsafe_allow_html=True)

            col1, col2 = st.columns([3,2])
            with col1:
                st.markdown("**Top 20 Colors by Units Sold**")
                st.bar_chart(color_agg.head(20).set_index("Color")["Units_Sold"])
            with col2:
                st.markdown("**Top 20 Colors by STR %**")
                st.bar_chart(color_agg.nlargest(20,"STR_%").set_index("Color")["STR_%"])

            st.markdown("**Color Breakdown Table**")
            display = color_agg[["Color","Units_Sold","In_Stock","STR_%","Status","Products"]].copy()
            display.columns = ["Color","Units Sold","In Stock","STR %","Status","Products"]
            st.dataframe(display, use_container_width=True, hide_index=True)

    elif view == "Size × Color Matrix":
        st.markdown('<div class="sec-title">📊 Size × Color Performance Matrix</div>',
                    unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Top Sizes by Units Sold**")
            if "Size" in sf.columns:
                sa = sf.groupby("Size")["Units Sold"].sum().sort_values(ascending=False).head(15)
                sa = sa.reindex([s for s in SIZE_ORDER if s in sa.index] +
                                [s for s in sa.index if s not in SIZE_ORDER]).dropna()
                st.dataframe(sa.reset_index(), use_container_width=True, hide_index=True)
        with col2:
            st.markdown("**Top Colors by Units Sold**")
            if "Color" in cf.columns:
                ca = cf.groupby("Color")["Units Sold"].sum().sort_values(ascending=False).head(15)
                st.dataframe(ca.reset_index(), use_container_width=True, hide_index=True)

    elif view == "Top Performers":
        st.markdown('<div class="sec-title">⚡ Top Performing Variants</div>',
                    unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**🏆 Top 20 Sizes — Super Fast STR**")
            if "Status" in size_df.columns:
                top_sf = size_df[size_df["Status"]=="Super Fast"].sort_values(
                    "Units Sold", ascending=False).head(20)
                if len(top_sf) > 0:
                    cols_show = [c for c in ["Product Name","Size","Units Sold","In Stock","STR %","Status"] if c in top_sf.columns]
                    st.dataframe(top_sf[cols_show], use_container_width=True, hide_index=True)
                else:
                    st.info("No Super Fast sizes found")
        with col2:
            st.markdown("**🏆 Top 20 Colors — Super Fast STR**")
            if "Status" in color_df.columns:
                top_cf = color_df[color_df["Status"]=="Super Fast"].sort_values(
                    "Units Sold", ascending=False).head(20)
                if len(top_cf) > 0:
                    cols_show = [c for c in ["Product Name","Color","Units Sold","In Stock","STR %","Status"] if c in top_cf.columns]
                    st.dataframe(top_cf[cols_show], use_container_width=True, hide_index=True)
                else:
                    st.info("No Super Fast colors found")

        st.markdown("---")
        col3, col4 = st.columns(2)
        with col3:
            st.markdown("**🚨 Dead Sizes — Need Clearance**")
            if "Status" in size_df.columns:
                dead_s = size_df[size_df["Status"]=="Dead"].sort_values(
                    "In Stock", ascending=False).head(20)
                if len(dead_s) > 0:
                    cols_show = [c for c in ["Product Name","Size","Units Sold","In Stock","STR %"] if c in dead_s.columns]
                    st.dataframe(dead_s[cols_show], use_container_width=True, hide_index=True)
        with col4:
            st.markdown("**🚨 Dead Colors — Need Clearance**")
            if "Status" in color_df.columns:
                dead_c = color_df[color_df["Status"]=="Dead"].sort_values(
                    "In Stock", ascending=False).head(20)
                if len(dead_c) > 0:
                    cols_show = [c for c in ["Product Name","Color","Units Sold","In Stock","STR %"] if c in dead_c.columns]
                    st.dataframe(dead_c[cols_show], use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()