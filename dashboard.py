import streamlit as st
import pandas as pd
import os
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
        background: white;
        border-radius: 12px;
        padding: 16px 20px;
        border: 1px solid #e8edf3;
        text-align: center;
    }
    .metric-val  { font-size: 32px; font-weight: 600; margin: 0; }
    .metric-lbl  { font-size: 12px; color: #6b7280; margin: 0; margin-top: 4px; }

    .prod-card {
        background: white;
        border-radius: 12px;
        border: 1px solid #e8edf3;
        overflow: hidden;
        margin-bottom: 12px;
        transition: box-shadow 0.2s;
    }
    .prod-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.08); }
    .prod-img {
        width: 100%; height: 160px;
        object-fit: cover; display: block;
    }
    .prod-img-placeholder {
        width: 100%; height: 160px;
        background: #f3f4f6;
        display: flex; align-items: center;
        justify-content: center; font-size: 48px;
    }
    .prod-body  { padding: 10px 12px; }
    .prod-name  { font-size: 13px; font-weight: 600; color: #111827;
                  white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .prod-meta  { font-size: 11px; color: #6b7280; margin-top: 2px; }
    .prod-price { font-size: 13px; font-weight: 500; color: #374151; margin-top: 4px; }
    .prod-sold  { font-size: 11px; color: #6b7280; }

    .badge { display:inline-block; padding:3px 10px; border-radius:20px;
             font-size:11px; font-weight:600; margin-top:6px; }
    .badge-fast   { background:#dcfce7; color:#166534; }
    .badge-medium { background:#fef9c3; color:#854d0e; }
    .badge-slow   { background:#fee2e2; color:#991b1b; }
    .badge-new    { background:#dbeafe; color:#1e40af; }
    .badge-none   { background:#f3f4f6; color:#6b7280; }

    .sidebar-title { font-size: 20px; font-weight: 700; color: #111827; }
    .divider { border-top: 1px solid #e5e7eb; margin: 16px 0; }

    .insight-box {
        background: #eff6ff; border: 1px solid #bfdbfe;
        border-radius: 10px; padding: 12px 16px;
        font-size: 13px; color: #1e40af; margin-bottom: 16px;
    }
    .no-data { text-align:center; padding:60px 20px;
               color:#9ca3af; font-size:15px; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
# ── Auto-find latest Excel file ──────────────────────────────────────────────
def find_latest_excel():
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
        return os.path.join(base, "odoo_productsall.xlsx")
    return str(max(candidates, key=lambda f: f.stat().st_mtime))

# EXCEL_FILE is determined dynamically inside main()
IMAGES_FOLDER = r"C:\Users\Legion\Desktop\odoo_export\product_images"

VELOCITY_ORDER  = ["Fast", "Medium", "Slow", "Just Launched", "No Sales Data"]
VELOCITY_COLORS = {
    "Fast":          "#166534",
    "Medium":        "#854d0e",
    "Slow":          "#991b1b",
    "Just Launched": "#1e40af",
    "No Sales Data": "#6b7280",
}
BADGE_CLASS = {
    "Fast":          "badge-fast",
    "Medium":        "badge-medium",
    "Slow":          "badge-slow",
    "Just Launched": "badge-new",
    "No Sales Data": "badge-none",
}
VELOCITY_EMOJI = {
    "Fast": "🟢", "Medium": "🟡", "Slow": "🔴",
    "Just Launched": "🆕", "No Sales Data": "⚪",
}

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)  # refresh every 5 minutes
def load_data(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name="Products", engine="openpyxl")
        # Normalise column names
        df.columns = [c.strip() for c in df.columns]
        # Clean up velocity column
        if "Velocity" in df.columns:
            df["Velocity"] = df["Velocity"].fillna("No Sales Data").astype(str)
            # Strip emoji prefixes and normalise
            import re as _re
            def clean_vel(x):
                # Remove emoji and extra symbols
                x = _re.sub(r"[^a-zA-Z0-9 ]", "", x).strip()
                mapping = {
                    "fast":          "Fast",
                    "medium":        "Medium",
                    "slow":          "Slow",
                    "just launched": "Just Launched",
                    "no sales data": "No Sales Data",
                    "justlaunched":  "Just Launched",
                    "nosalesdata":   "No Sales Data",
                }
                return mapping.get(x.lower().strip(), x.strip())
            df["Velocity"] = df["Velocity"].apply(clean_vel)
        return df, None
    except Exception as e:
        return None, str(e)

# ── Image loader ──────────────────────────────────────────────────────────────
def get_product_image(sku, name):
    """Try to find product image from the images folder"""
    if not os.path.exists(IMAGES_FOLDER):
        return None
    # Try SKU first, then cleaned name
    candidates = []
    if sku and str(sku).strip():
        candidates.append(str(sku).strip())
    if name and str(name).strip():
        safe = "".join(c for c in str(name) if c.isalnum() or c in "-_")[:60]
        candidates.append(safe)
    for candidate in candidates:
        path = os.path.join(IMAGES_FOLDER, f"{candidate}.png")
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

# ── Render product card ───────────────────────────────────────────────────────
def product_card(row):
    sku      = str(row.get("SKU / Internal Ref", "")).strip()
    name     = str(row.get("Product Name", "Unknown")).strip()
    brand    = str(row.get("Brand", "")).strip()
    cat      = str(row.get("Category", "")).strip()
    price    = row.get("Sales Price", 0)
    sold     = row.get("Total Units Sold", 0)
    on_hand  = row.get("On Hand Qty", 0)
    velocity = str(row.get("Velocity", "No Sales Data")).strip()
    launch   = str(row.get("Launch Date", "")).strip()
    days     = row.get("Days Since Launch", "")
    badge_cls = BADGE_CLASS.get(velocity, "badge-none")
    emoji     = VELOCITY_EMOJI.get(velocity, "⚪")

    # Days label
    if days and str(days).strip() not in ["", "—", "nan"]:
        try:
            days_label = f"{int(float(str(days)))} days"
        except:
            days_label = str(days)
    else:
        days_label = ""

    badge_text = f"{emoji} {velocity}" + (f" · {days_label}" if days_label else "")

    # Image
    img_path = get_product_image(sku, name)
    if img_path:
        b64 = img_to_base64(img_path)
        img_html = f'<img class="prod-img" src="data:image/jpeg;base64,{b64}" alt="{name}"/>' \
                   if b64 else '<div class="prod-img-placeholder">👗</div>'
    else:
        img_html = '<div class="prod-img-placeholder">👗</div>'

    price_str    = f"${price:,.2f}"   if isinstance(price, (int, float)) else str(price)
    sold_str     = f"{sold:,.0f} sold"if isinstance(sold,  (int, float)) else str(sold)
    on_hand_str  = f"{on_hand:,.0f} in stock" if isinstance(on_hand,(int,float)) else ""
    meta_parts   = [x for x in [brand, cat] if x and x != "nan"]
    meta_str     = " · ".join(meta_parts)

    st.markdown(f"""
    <div class="prod-card">
      {img_html}
      <div class="prod-body">
        <div class="prod-name" title="{name}">{name}</div>
        <div class="prod-meta">{meta_str}</div>
        <div class="prod-price">{price_str} &nbsp;·&nbsp; <span class="prod-sold">{sold_str} · {on_hand_str}</span></div>
        <span class="badge {badge_cls}">{badge_text}</span>
        {"<div class='prod-meta' style='margin-top:4px'>Launch: "+launch+"</div>" if launch and launch != "Not sold yet" and launch != "nan" else ""}
      </div>
    </div>
    """, unsafe_allow_html=True)

# ── Main app ──────────────────────────────────────────────────────────────────
def main():
    # Load data
    # Always find the latest file on each page load
    EXCEL_FILE = find_latest_excel()
    df, err = load_data(EXCEL_FILE)

    if err or df is None:
        st.error(f"Could not load Excel file: {err}")
        st.info(f"Expected file at: {find_latest_excel()}")
        st.stop()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown('<div class="sidebar-title">👗 Salt Fashion</div>', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-title" style="font-size:14px;color:#6b7280;font-weight:400">Product Velocity Dashboard</div>', unsafe_allow_html=True)
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Brand filter
        brands = sorted([str(b) for b in df["Brand"].dropna().unique()
                         if str(b).strip() and str(b).strip() not in ("nan","True","False","None")])
        selected_brand = st.selectbox(
            "Brand",
            options=brands,
            index=0,
        )

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Velocity filter — detect all unique values robustly
        st.markdown("**Velocity Filter**")
        # Normalise all velocity values in place
        df["Velocity"] = df["Velocity"].fillna("No Sales Data").astype(str).str.strip()
        all_vels = df["Velocity"].unique().tolist()
        vel_options = [v for v in VELOCITY_ORDER if v in all_vels]
        # Add any extra values not in our list
        for v in all_vels:
            if v not in vel_options and v not in ("nan","","None"):
                vel_options.append(v)
        selected_vels = []
        for v in vel_options:
            emoji = VELOCITY_EMOJI.get(v, "⚪")
            count = len(df[df["Velocity"] == v])
            if st.checkbox(f"{emoji} {v} ({count})", value=True, key=f"vel_{v}"):
                selected_vels.append(v)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Category filter
        filtered_brand = df[df["Brand"] == selected_brand] if selected_brand else df
        cats = sorted([str(c) for c in filtered_brand["Category"].dropna().unique()
                       if str(c).strip() and str(c).strip() not in ("nan","True","False","None")])
        selected_cats = st.multiselect(
            "Category",
            options=cats,
            default=cats,
            placeholder="All categories"
        )

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Search
        search = st.text_input("Search product name", placeholder="e.g. Kurti, Saree...")

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Sort
        sort_by = st.selectbox("Sort by", [
            "Velocity (Fast first)",
            "Total Units Sold (High)",
            "Sales Price (High)",
            "Sales Price (Low)",
            "Days Since Launch (Most recent)",
            "On Hand Qty (High)",
        ])

        # Products per page
        per_page = st.selectbox("Products per page", [12, 24, 48, 96], index=0)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        actual_file = find_latest_excel()
        st.caption(f"Data: {os.path.basename(actual_file)}")
        if st.button("Refresh data"):
            st.cache_data.clear()
            st.rerun()

    # ── Filter data ───────────────────────────────────────────────────────────
    filtered = df.copy()

    if selected_brand:
        filtered = filtered[filtered["Brand"] == selected_brand]

    if selected_vels:
        filtered = filtered[filtered["Velocity"].isin(selected_vels)]

    if selected_cats:
        filtered = filtered[filtered["Category"].isin(selected_cats)]

    if search.strip():
        filtered = filtered[
            filtered["Product Name"].fillna("").str.contains(search.strip(), case=False)
        ]

    # Sort
    if sort_by == "Velocity (Fast first)":
        vel_order_map = {v: i for i, v in enumerate(VELOCITY_ORDER)}
        filtered["_vel_order"] = filtered["Velocity"].map(vel_order_map).fillna(99)
        filtered = filtered.sort_values("_vel_order")
    elif sort_by == "Total Units Sold (High)":
        filtered = filtered.sort_values("Total Units Sold", ascending=False)
    elif sort_by == "Sales Price (High)":
        filtered = filtered.sort_values("Sales Price", ascending=False)
    elif sort_by == "Sales Price (Low)":
        filtered = filtered.sort_values("Sales Price", ascending=True)
    elif sort_by == "Days Since Launch (Most recent)":
        filtered = filtered.sort_values("Days Since Launch", ascending=True)
    elif sort_by == "On Hand Qty (High)":
        filtered = filtered.sort_values("On Hand Qty", ascending=False)

    # ── Header ────────────────────────────────────────────────────────────────
    brand_label = selected_brand or "All Brands"
    st.markdown(f"## {brand_label} — Product Velocity")
    st.markdown(f"Showing **{len(filtered):,}** products · sorted by {sort_by}")

    # ── Metric cards ──────────────────────────────────────────────────────────
    # Normalise Brand column to string for reliable filtering
    df["Brand"] = df["Brand"].fillna("").astype(str).str.strip()
    brand_df = df[df["Brand"] == str(selected_brand).strip()].copy() if selected_brand else df.copy()
    # Normalise velocity for counting
    brand_df["Velocity"] = brand_df["Velocity"].fillna("No Sales Data").astype(str).str.strip()
    fast     = len(brand_df[brand_df["Velocity"] == "Fast"])
    medium   = len(brand_df[brand_df["Velocity"] == "Medium"])
    slow     = len(brand_df[brand_df["Velocity"] == "Slow"])
    no_data  = len(brand_df[brand_df["Velocity"] == "No Sales Data"])
    new_prod = len(brand_df[brand_df["Velocity"] == "Just Launched"])
    total    = len(brand_df)
    top_sold = pd.to_numeric(brand_df["Total Units Sold"], errors="coerce").sum() if "Total Units Sold" in brand_df.columns else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        st.markdown(f'<div class="metric-card"><p class="metric-val" style="color:#111827">{total:,}</p><p class="metric-lbl">Total Products</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><p class="metric-val" style="color:#166534">{fast:,}</p><p class="metric-lbl">🟢 Fast</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><p class="metric-val" style="color:#854d0e">{medium:,}</p><p class="metric-lbl">🟡 Medium</p></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card"><p class="metric-val" style="color:#991b1b">{slow:,}</p><p class="metric-lbl">🔴 Slow</p></div>', unsafe_allow_html=True)
    with c5:
        st.markdown(f'<div class="metric-card"><p class="metric-val" style="color:#1e40af">{new_prod:,}</p><p class="metric-lbl">🆕 New</p></div>', unsafe_allow_html=True)
    with c6:
        st.markdown(f'<div class="metric-card"><p class="metric-val" style="color:#1d4ed8">{top_sold:,.0f}</p><p class="metric-lbl">Total Sold</p></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── AI-style insight ──────────────────────────────────────────────────────
    if slow > 0 and total > 0:
        slow_pct = round(slow / total * 100)
        slow_stock = brand_df[brand_df["Velocity"] == "Slow"]["On Hand Qty"].sum()
        st.markdown(f"""
        <div class="insight-box">
        💡 <strong>Insight:</strong> {slow:,} products ({slow_pct}%) are slow movers
        with <strong>{slow_stock:,.0f} units still in stock</strong>.
        Consider markdown pricing or promotions on these items to free up inventory.
        </div>
        """, unsafe_allow_html=True)

    # ── Products grid ─────────────────────────────────────────────────────────
    if len(filtered) == 0:
        st.markdown('<div class="no-data">No products found matching your filters.</div>',
                    unsafe_allow_html=True)
        st.stop()

    # Pagination
    total_pages = max(1, (len(filtered) - 1) // per_page + 1)
    if total_pages > 1:
        page = st.number_input(f"Page (1 – {total_pages})", min_value=1,
                               max_value=total_pages, value=1, step=1)
    else:
        page = 1

    start_idx = (page - 1) * per_page
    page_df   = filtered.iloc[start_idx: start_idx + per_page]

    # Render grid — 4 columns
    COLS = 4
    rows_needed = (len(page_df) + COLS - 1) // COLS
    for r in range(rows_needed):
        cols = st.columns(COLS)
        for c in range(COLS):
            idx = r * COLS + c
            if idx < len(page_df):
                with cols[c]:
                    product_card(page_df.iloc[idx])

    # ── Bottom stats ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Category Breakdown")
    if "Category" in filtered.columns and "Velocity" in filtered.columns:
        # Normalise both columns before groupby
        cat_data = filtered.copy()
        cat_data["Category"] = cat_data["Category"].fillna("Uncategorised").astype(str).str.strip()
        cat_data["Velocity"] = cat_data["Velocity"].fillna("No Sales Data").astype(str).str.strip()
        cat_summary = cat_data.groupby(["Category","Velocity"]).size().unstack(fill_value=0)
        ordered_cols = [v for v in VELOCITY_ORDER if v in cat_summary.columns]
        if ordered_cols:
            cat_summary = cat_summary[ordered_cols]
        cat_summary["Total"] = cat_summary.sum(axis=1)
        cat_summary = cat_summary.sort_values("Total", ascending=False).head(20)
        st.dataframe(cat_summary, use_container_width=True)

if __name__ == "__main__":
    main()