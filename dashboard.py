import streamlit as st
import pandas as pd
import os
import json
import tempfile
from PIL import Image
import base64
from io import BytesIO
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Salt Fashion — Product Velocity",
    page_icon="👗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fb; }
    .block-container { padding: 1.5rem 2rem; }
    .metric-card {
        background: white; border-radius: 12px;
        padding: 16px 20px; border: 1px solid #e8edf3; text-align: center;
    }
    .metric-val { font-size: 32px; font-weight: 600; margin: 0; }
    .metric-lbl { font-size: 12px; color: #6b7280; margin: 0; margin-top: 4px; }
    .prod-card {
        background: white; border-radius: 12px;
        border: 1px solid #e8edf3; overflow: hidden; margin-bottom: 12px;
    }
    .prod-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.08); }
    .prod-img { width:100%; height:160px; object-fit:cover; display:block; }
    .prod-img-placeholder {
        width:100%; height:160px; background:#f3f4f6;
        display:flex; align-items:center; justify-content:center; font-size:48px;
    }
    .prod-body  { padding: 10px 12px; }
    .prod-name  { font-size:13px; font-weight:600; color:#111827;
                  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .prod-meta  { font-size:11px; color:#6b7280; margin-top:2px; }
    .prod-price { font-size:13px; font-weight:500; color:#374151; margin-top:4px; }
    .badge { display:inline-block; padding:3px 10px; border-radius:20px;
             font-size:11px; font-weight:600; margin-top:6px; }
    .badge-fast   { background:#dcfce7; color:#166534; }
    .badge-medium { background:#fef9c3; color:#854d0e; }
    .badge-slow   { background:#fee2e2; color:#991b1b; }
    .badge-new    { background:#dbeafe; color:#1e40af; }
    .badge-none   { background:#f3f4f6; color:#6b7280; }
    .insight-box {
        background:#eff6ff; border:1px solid #bfdbfe;
        border-radius:10px; padding:12px 16px;
        font-size:13px; color:#1e40af; margin-bottom:16px;
    }
    .divider { border-top:1px solid #e5e7eb; margin:16px 0; }
    .no-data { text-align:center; padding:60px 20px; color:#9ca3af; font-size:15px; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
GDRIVE_FILE_ID  = "1W_Ihk-280YxB_pyj4YkcoR4k84HH086q"
VELOCITY_ORDER  = ["Fast","Medium","Slow","Just Launched","No Sales Data"]
VELOCITY_EMOJI  = {"Fast":"🟢","Medium":"🟡","Slow":"🔴",
                   "Just Launched":"🆕","No Sales Data":"⚪"}
BADGE_CLASS     = {"Fast":"badge-fast","Medium":"badge-medium","Slow":"badge-slow",
                   "Just Launched":"badge-new","No Sales Data":"badge-none"}

# ── Google Drive loader ───────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data_from_gdrive():
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        # Load credentials from Streamlit secrets
        creds_dict = dict(st.secrets["gcp_service_account"])
        scope = [
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets",
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)

        # Download file from Google Drive
        import googleapiclient.discovery
        from googleapiclient.http import MediaIoBaseDownload

        drive_service = googleapiclient.discovery.build("drive", "v3", credentials=creds)
        request = drive_service.files().get_media(fileId=GDRIVE_FILE_ID)
        buf = BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)

        df = pd.read_excel(buf, sheet_name="Products", engine="openpyxl")
        df.columns = [c.strip() for c in df.columns]

        # Clean velocity
        if "Velocity" in df.columns:
            import re
            def clean_vel(x):
                x = re.sub(r"[^a-zA-Z0-9 ]", "", str(x)).strip()
                m = {"fast":"Fast","medium":"Medium","slow":"Slow",
                     "just launched":"Just Launched","justlaunched":"Just Launched",
                     "no sales data":"No Sales Data","nosalesdata":"No Sales Data"}
                return m.get(x.lower().strip(), x.strip() or "No Sales Data")
            df["Velocity"] = df["Velocity"].fillna("No Sales Data").apply(clean_vel)

        return df, None
    except Exception as e:
        return None, str(e)

# ── Local loader (fallback for running locally) ───────────────────────────────
@st.cache_data(ttl=300)
def load_data_local():
    try:
        base = r"C:\Users\Legion\Desktop\odoo_export"
        exports_dir = os.path.join(base, "exports")
        candidates = []
        if os.path.exists(exports_dir):
            for f in Path(exports_dir).glob("odoo_products_*.xlsx"):
                candidates.append(f)
        for name in ["odoo_productsall.xlsx","odoo_products.xlsx","odoo_products_MASTER.xlsx"]:
            p = Path(os.path.join(base, name))
            if p.exists():
                candidates.append(p)
        if not candidates:
            return None, "No Excel file found"
        latest = str(max(candidates, key=lambda f: f.stat().st_mtime))
        df = pd.read_excel(latest, sheet_name="Products", engine="openpyxl")
        df.columns = [c.strip() for c in df.columns]
        if "Velocity" in df.columns:
            import re
            def clean_vel(x):
                x = re.sub(r"[^a-zA-Z0-9 ]", "", str(x)).strip()
                m = {"fast":"Fast","medium":"Medium","slow":"Slow",
                     "just launched":"Just Launched","justlaunched":"Just Launched",
                     "no sales data":"No Sales Data","nosalesdata":"No Sales Data"}
                return m.get(x.lower().strip(), x.strip() or "No Sales Data")
            df["Velocity"] = df["Velocity"].fillna("No Sales Data").apply(clean_vel)
        return df, None
    except Exception as e:
        return None, str(e)

def load_data():
    """Try Google Drive first, fall back to local"""
    try:
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            return load_data_from_gdrive()
    except:
        pass
    return load_data_local()

ODOO_URL      = "https://spos.jeevee.com"
IMAGES_FOLDER = r"C:\Users\Legion\Desktop\odoo_export\product_images"

def get_image_url(product_id):
    """Load image from Odoo server — works everywhere including Streamlit Cloud"""
    try:
        if product_id and str(product_id) not in ("", "nan", "None", "0"):
            pid = int(float(str(product_id)))
            if pid > 0:
                return f"{ODOO_URL}/web/image/product.template/{pid}/image_128"
    except:
        pass
    return None


def get_product_image(sku, name):
    if not os.path.exists(IMAGES_FOLDER):
        return None
    candidates = []
    if sku and str(sku).strip() and str(sku).strip() != "nan":
        candidates.append(str(sku).strip())
    if name and str(name).strip():
        safe = "".join(c for c in str(name) if c.isalnum() or c in "-_")[:60]
        candidates.append(safe)
    for c in candidates:
        path = os.path.join(IMAGES_FOLDER, f"{c}.png")
        if os.path.exists(path):
            return path
    return None

def img_to_base64(path):
    try:
        img = Image.open(path).convert("RGB")
        img.thumbnail((300, 300))
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=85)
        return base64.b64encode(buf.getvalue()).decode()
    except:
        return None

# ── Product card ──────────────────────────────────────────────────────────────
def product_card(row):
    sku      = str(row.get("SKU / Internal Ref","")).strip()
    name     = str(row.get("Product Name","Unknown")).strip()
    brand    = str(row.get("Brand","")).strip()
    cat      = str(row.get("Category","")).strip()
    price    = row.get("Sales Price", 0)
    sold     = row.get("Total Units Sold", 0)
    on_hand  = row.get("On Hand Qty", 0)
    velocity = str(row.get("Velocity","No Sales Data")).strip()
    launch   = str(row.get("Launch Date","")).strip()
    days     = row.get("Days Since Launch","")

    badge_cls = BADGE_CLASS.get(velocity, "badge-none")
    emoji     = VELOCITY_EMOJI.get(velocity, "⚪")

    try:    days_label = f"{int(float(str(days)))} days" if str(days) not in ["","—","nan"] else ""
    except: days_label = ""

    badge_text = f"{emoji} {velocity}" + (f" · {days_label}" if days_label else "")

    # Image: try base64 from Excel first (works everywhere including cloud)
    # Then fall back to local file, then placeholder
    img_b64_raw = row.get("Image_Base64", "")
    img_html = None

    # Option 1: base64 stored in Excel column
    if img_b64_raw and str(img_b64_raw).strip() not in ("", "nan", "None"):
        try:
            raw_b64 = str(img_b64_raw).strip()
            # Decode → resize → re-encode for web
            raw_bytes = base64.b64decode(raw_b64)
            img_obj = Image.open(BytesIO(raw_bytes)).convert("RGB")
            img_obj.thumbnail((300, 300))
            buf = BytesIO()
            img_obj.save(buf, format="JPEG", quality=80)
            web_b64 = base64.b64encode(buf.getvalue()).decode()
            img_html = f'<img class="prod-img" src="data:image/jpeg;base64,{web_b64}" alt="{name}" loading="lazy"/>' 
        except:
            img_html = None

    # Option 2: local image file (when running on your laptop)
    if not img_html and os.path.exists(IMAGES_FOLDER):
        for c in [str(sku).strip(), "".join(ch for ch in str(name) if ch.isalnum() or ch in "-_")[:60]]:
            if c and c != "nan":
                p = os.path.join(IMAGES_FOLDER, f"{c}.png")
                if os.path.exists(p):
                    b64 = img_to_base64(p)
                    if b64:
                        img_html = f'<img class="prod-img" src="data:image/jpeg;base64,{b64}" alt="{name}"/>' 
                    break

    # Option 3: placeholder
    if not img_html:
        img_html = '<div class="prod-img-placeholder">👗</div>'


    price_str   = f"${price:,.2f}"   if isinstance(price,(int,float)) else str(price)
    sold_str    = f"{sold:,.0f} sold"if isinstance(sold,(int,float))  else str(sold)
    onhand_str  = f"{on_hand:,.0f} in stock" if isinstance(on_hand,(int,float)) else ""
    meta_parts  = [x for x in [brand, cat] if x and x != "nan"]
    meta_str    = " · ".join(meta_parts)
    launch_html = f"<div class='prod-meta' style='margin-top:4px'>Launch: {launch}</div>" \
                  if launch and launch not in ("Not sold yet","nan","") else ""

    st.markdown(f"""
    <div class="prod-card">
      {img_html}
      <div class="prod-body">
        <div class="prod-name" title="{name}">{name}</div>
        <div class="prod-meta">{meta_str}</div>
        <div class="prod-price">{price_str} &nbsp;·&nbsp;
          <span style="font-size:11px;color:#6b7280">{sold_str} · {onhand_str}</span>
        </div>
        <span class="badge {badge_cls}">{badge_text}</span>
        {launch_html}
      </div>
    </div>
    """, unsafe_allow_html=True)

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    df, err = load_data()

    if err or df is None:
        st.error(f"Could not load data: {err}")
        st.info("Make sure your Excel file is uploaded to Google Drive or available locally.")
        st.stop()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 👗 Salt Fashion")
        st.markdown("**Product Velocity Dashboard**")
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Brand
        df["Brand"] = df["Brand"].fillna("").astype(str).str.strip()
        brands = sorted([b for b in df["Brand"].unique()
                         if b and b not in ("nan","True","False","None","")])
        selected_brand = st.selectbox("Brand", options=brands, index=0)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Velocity
        st.markdown("**Velocity Filter**")
        df["Velocity"] = df["Velocity"].fillna("No Sales Data").astype(str).str.strip()
        all_vels    = df["Velocity"].unique().tolist()
        vel_options = [v for v in VELOCITY_ORDER if v in all_vels]
        for v in all_vels:
            if v not in vel_options and v not in ("nan","","None"):
                vel_options.append(v)
        selected_vels = []
        for v in vel_options:
            count = len(df[df["Velocity"] == v])
            if st.checkbox(f"{VELOCITY_EMOJI.get(v,'⚪')} {v} ({count:,})", value=True, key=f"v_{v}"):
                selected_vels.append(v)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Category
        brand_df_cats = df[df["Brand"] == selected_brand] if selected_brand else df
        cats = sorted([str(c) for c in brand_df_cats["Category"].dropna().unique()
                       if str(c).strip() not in ("nan","True","False","None","")])
        selected_cats = st.multiselect("Category", options=cats, default=cats)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        search  = st.text_input("Search product", placeholder="e.g. Kurti, Saree...")
        sort_by = st.selectbox("Sort by", [
            "Velocity (Fast first)",
            "Total Units Sold (High)",
            "Sales Price (High)",
            "Sales Price (Low)",
            "Days Since Launch (Recent)",
            "On Hand Qty (High)",
        ])
        per_page = st.selectbox("Per page", [12, 24, 48, 96], index=0)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        if st.button("🔄 Refresh data"):
            st.cache_data.clear()
            st.rerun()

    # ── Filter ────────────────────────────────────────────────────────────────
    filtered = df.copy()
    if selected_brand:
        filtered = filtered[filtered["Brand"] == selected_brand]
    if selected_vels:
        filtered = filtered[filtered["Velocity"].isin(selected_vels)]
    if selected_cats:
        filtered = filtered[filtered["Category"].astype(str).isin(selected_cats)]
    if search.strip():
        filtered = filtered[
            filtered["Product Name"].fillna("").str.contains(search.strip(), case=False)
        ]

    # Sort
    vel_map = {v:i for i,v in enumerate(VELOCITY_ORDER)}
    if sort_by == "Velocity (Fast first)":
        filtered = filtered.copy()
        filtered["_vo"] = filtered["Velocity"].map(vel_map).fillna(99)
        filtered = filtered.sort_values("_vo")
    elif sort_by == "Total Units Sold (High)":
        filtered = filtered.sort_values("Total Units Sold", ascending=False)
    elif sort_by == "Sales Price (High)":
        filtered = filtered.sort_values("Sales Price", ascending=False)
    elif sort_by == "Sales Price (Low)":
        filtered = filtered.sort_values("Sales Price", ascending=True)
    elif sort_by == "Days Since Launch (Recent)":
        filtered = filtered.sort_values("Days Since Launch", ascending=True)
    elif sort_by == "On Hand Qty (High)":
        filtered = filtered.sort_values("On Hand Qty", ascending=False)

    # ── Metrics ───────────────────────────────────────────────────────────────
    bdf = df[df["Brand"] == selected_brand].copy() if selected_brand else df.copy()
    bdf["Velocity"] = bdf["Velocity"].fillna("No Sales Data").astype(str).str.strip()
    fast     = len(bdf[bdf["Velocity"] == "Fast"])
    medium   = len(bdf[bdf["Velocity"] == "Medium"])
    slow     = len(bdf[bdf["Velocity"] == "Slow"])
    new_prod = len(bdf[bdf["Velocity"] == "Just Launched"])
    total    = len(bdf)
    top_sold = pd.to_numeric(bdf.get("Total Units Sold", pd.Series()), errors="coerce").sum()

    st.markdown(f"## {selected_brand or 'All Brands'} — Product Velocity")
    st.markdown(f"Showing **{len(filtered):,}** products · {sort_by}")

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    for col, val, label, color in [
        (c1, total,    "Total Products", "#111827"),
        (c2, fast,     "🟢 Fast",        "#166534"),
        (c3, medium,   "🟡 Medium",      "#854d0e"),
        (c4, slow,     "🔴 Slow",        "#991b1b"),
        (c5, new_prod, "🆕 New",         "#1e40af"),
        (c6, int(top_sold), "Total Sold","#1d4ed8"),
    ]:
        with col:
            st.markdown(
                f'<div class="metric-card">'
                f'<p class="metric-val" style="color:{color}">{val:,}</p>'
                f'<p class="metric-lbl">{label}</p></div>',
                unsafe_allow_html=True
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # Insight
    if slow > 0 and total > 0:
        slow_stock = pd.to_numeric(bdf[bdf["Velocity"]=="Slow"].get("On Hand Qty", pd.Series()), errors="coerce").sum()
        st.markdown(
            f'<div class="insight-box">💡 <strong>Insight:</strong> {slow:,} products '
            f'({round(slow/total*100)}%) are slow movers with '
            f'<strong>{slow_stock:,.0f} units still in stock</strong>. '
            f'Consider markdown pricing or promotions to free up inventory.</div>',
            unsafe_allow_html=True
        )

    # ── Grid ──────────────────────────────────────────────────────────────────
    if len(filtered) == 0:
        st.markdown('<div class="no-data">No products match your filters.</div>', unsafe_allow_html=True)
        st.stop()

    total_pages = max(1, (len(filtered)-1)//per_page+1)
    page = st.number_input(f"Page (1–{total_pages})", min_value=1,
                           max_value=total_pages, value=1) if total_pages > 1 else 1
    page_df = filtered.iloc[(page-1)*per_page : page*per_page]

    COLS = 4
    for r in range((len(page_df)+COLS-1)//COLS):
        cols = st.columns(COLS)
        for c in range(COLS):
            idx = r*COLS+c
            if idx < len(page_df):
                with cols[c]:
                    product_card(page_df.iloc[idx])

    # ── Category breakdown ────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Category Breakdown")
    if "Category" in filtered.columns:
        cat_data = filtered.copy()
        cat_data["Category"] = cat_data["Category"].fillna("Uncategorised").astype(str).str.strip()
        cat_data["Velocity"] = cat_data["Velocity"].fillna("No Sales Data").astype(str).str.strip()
        cat_sum = cat_data.groupby(["Category","Velocity"]).size().unstack(fill_value=0)
        ordered = [v for v in VELOCITY_ORDER if v in cat_sum.columns]
        if ordered:
            cat_sum = cat_sum[ordered]
        cat_sum["Total"] = cat_sum.sum(axis=1)
        st.dataframe(cat_sum.sort_values("Total", ascending=False).head(20),
                     use_container_width=True)

if __name__ == "__main__":
    main()