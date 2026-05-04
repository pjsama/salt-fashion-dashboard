import streamlit as st
import pandas as pd
import os
import base64
from PIL import Image
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Salt Fashion — Intelligence Dashboard",
    page_icon="👗", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container{padding:1.5rem 2rem}
.metric-card{background:white;border-radius:12px;padding:14px 18px;
             border:1px solid #e8edf3;text-align:center;height:90px}
.metric-val{font-size:28px;font-weight:600;margin:0}
.metric-lbl{font-size:11px;color:#6b7280;margin:0;margin-top:3px}
.prod-card{background:white;border-radius:12px;border:1px solid #e8edf3;
           overflow:hidden;margin-bottom:12px}
.prod-card:hover{box-shadow:0 4px 16px rgba(0,0,0,0.08)}
.prod-img{width:100%;height:150px;object-fit:cover;display:block}
.prod-placeholder{width:100%;height:150px;background:#f3f4f6;
                  display:flex;align-items:center;justify-content:center;font-size:40px}
.prod-body{padding:10px 12px}
.prod-name{font-size:12px;font-weight:600;color:#111827;
           white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.prod-meta{font-size:11px;color:#6b7280;margin-top:1px}
.badge{display:inline-block;padding:2px 8px;border-radius:12px;
       font-size:11px;font-weight:600;margin-top:5px;margin-right:3px}
.divider{border-top:1px solid #e5e7eb;margin:12px 0}
.insight{background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;
         padding:10px 14px;font-size:13px;color:#1e40af;margin-bottom:14px}
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
GDRIVE_FILE_ID  = "1kIHUlGCallLjXe9tiBrYDQ16ElQDmLR3"
IMAGES_FOLDER   = r"C:\Users\Legion\Desktop\odoo_export\product_images"

STR_COLORS = {
    "Super Fast": ("#1B5E20","#FFFFFF"),
    "Fast":       ("#43A047","#FFFFFF"),
    "Medium":     ("#F9A825","#000000"),
    "Slow":       ("#E53935","#FFFFFF"),
    "Dead":       ("#424242","#FFFFFF"),
}
ABC_COLORS = {
    "A": ("#1565C0","#FFFFFF"),
    "B": ("#6A1E9A","#FFFFFF"),
    "C": ("#757575","#FFFFFF"),
}
DOC_COLORS = {
    "Reorder Now": ("#B71C1C","#FFFFFF"),
    "Watch":       ("#F57F17","#FFFFFF"),
    "OK":          ("#2E7D32","#FFFFFF"),
    "N/A":         ("#9E9E9E","#FFFFFF"),
}
STR_ORDER = ["Super Fast","Fast","Medium","Slow","Dead"]

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_from_gdrive(file_id):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        import googleapiclient.discovery
        from googleapiclient.http import MediaIoBaseDownload

        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        svc     = googleapiclient.discovery.build("drive","v3",credentials=creds)
        request = svc.files().get_media(fileId=file_id)
        buf     = BytesIO()
        dl      = MediaIoBaseDownload(buf, request)
        done    = False
        while not done: _, done = dl.next_chunk()
        buf.seek(0)
        return pd.read_excel(buf, sheet_name="Products", engine="openpyxl"), None
    except Exception as e:
        return None, str(e)

@st.cache_data(ttl=300)
def load_local():
    base = r"C:\Users\Legion\Desktop\odoo_export"
    dirs = [os.path.join(base,"exports"), base]
    candidates = []
    for d in dirs:
        if os.path.exists(d):
            candidates += list(Path(d).glob("odoo_products*.xlsx"))
    if not candidates:
        return None, "No Excel file found locally"
    latest = str(max(candidates, key=lambda f: f.stat().st_mtime))
    return pd.read_excel(latest, sheet_name="Products", engine="openpyxl"), None

def load_data():
    try:
        if "gcp_service_account" in st.secrets:
            df, err = load_from_gdrive(GDRIVE_FILE_ID)
            if df is not None:
                return df, None
    except: pass
    return load_local()

def clean_df(df):
    df.columns = [c.strip() for c in df.columns]

    # Auto-detect velocity column — handles both old and new export formats
    # New format: "STR Status" column
    # Old format: "Velocity" column
    import re as _re
    def _clean_str_val(x):
        x = _re.sub(r"[^a-zA-Z0-9 ]", "", str(x)).strip()
        m = {
            "super fast":  "Super Fast",
            "fast":        "Fast",
            "medium":      "Medium",
            "slow":        "Slow",
            "dead":        "Dead",
            "just launched":"Just Launched",
            "justlaunched":"Just Launched",
            "no sales data":"Dead",
            "nosalesdata": "Dead",
        }
        return m.get(x.lower().strip(), x.strip() or "Dead")

    if "STR Status" in df.columns:
        df["STR Status"] = df["STR Status"].fillna("Dead").apply(_clean_str_val)
    elif "Velocity" in df.columns:
        # Old format — map velocity to STR status equivalents
        df["STR Status"] = df["Velocity"].fillna("Dead").apply(_clean_str_val)
    else:
        df["STR Status"] = "Dead"
    # Ensure numeric columns
    for col in ["Sales Price","Cost Price","On Hand Qty","Total Units Sold",
                "Revenue","Days of Cover"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    # Sell-Through % — convert from decimal (0.96) to percentage (96.0) if needed
    if "Sell-Through %" in df.columns:
        df["Sell-Through %"] = pd.to_numeric(df["Sell-Through %"], errors="coerce").fillna(0)
        # If max value <= 1.0, it's stored as decimal — multiply by 100
        if df["Sell-Through %"].max() <= 1.0:
            df["Sell-Through %"] = df["Sell-Through %"] * 100
    # Ensure string columns
    for col in ["Brand","Category","ABC Class","DOC Status","STR Status",
                "Product Name","SKU / Internal Ref","Barcode"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df

# ── Image helpers ─────────────────────────────────────────────────────────────
def get_img_html(row, name):
    # Try base64 from Excel first
    b64_raw = row.get("Image_Base64","")
    if b64_raw and str(b64_raw).strip() not in ("","nan","None"):
        try:
            raw   = base64.b64decode(str(b64_raw).strip())
            img   = Image.open(BytesIO(raw)).convert("RGB")
            img.thumbnail((300,300))
            buf   = BytesIO()
            img.save(buf,"JPEG",quality=80)
            web   = base64.b64encode(buf.getvalue()).decode()
            return f'<img class="prod-img" src="data:image/jpeg;base64,{web}" loading="lazy"/>'
        except: pass
    # Local fallback
    if os.path.exists(IMAGES_FOLDER):
        sku = str(row.get("SKU / Internal Ref","")).strip()
        for cand in [sku, "".join(c for c in name if c.isalnum() or c in "-_")[:60]]:
            if cand and cand != "nan":
                p = os.path.join(IMAGES_FOLDER,f"{cand}.png")
                if os.path.exists(p):
                    try:
                        img = Image.open(p).convert("RGB")
                        img.thumbnail((300,300))
                        buf = BytesIO(); img.save(buf,"JPEG",quality=80)
                        web = base64.b64encode(buf.getvalue()).decode()
                        return f'<img class="prod-img" src="data:image/jpeg;base64,{web}" loading="lazy"/>'
                    except: pass
    return '<div class="prod-placeholder">👗</div>'

# ── Product card ──────────────────────────────────────────────────────────────
def product_card(row):
    name    = str(row.get("Product Name","")).strip() or "—"
    brand   = str(row.get("Brand","")).strip()
    cat     = str(row.get("Category","")).strip()
    price   = row.get("Sales Price",0)
    sold    = row.get("Total Units Sold",0)
    onhand  = row.get("On Hand Qty",0)
    str_s   = str(row.get("STR Status","Dead")).strip()
    str_pct = min(float(row.get("Sell-Through %",0) or 0), 100.0)
    abc     = str(row.get("ABC Class","C")).strip()
    doc_s   = str(row.get("DOC Status","N/A")).strip()
    doc_d   = row.get("Days of Cover","")
    revenue = row.get("Revenue",0)
    launch  = str(row.get("Launch Date","")).strip()

    img_html = get_img_html(row, name)

    str_bg, str_fg = STR_COLORS.get(str_s,("#9E9E9E","#FFFFFF"))
    abc_bg, abc_fg = ABC_COLORS.get(abc,  ("#757575","#FFFFFF"))
    doc_bg, doc_fg = DOC_COLORS.get(doc_s,("#9E9E9E","#FFFFFF"))

    price_s   = f"${price:,.0f}"  if isinstance(price,(int,float)) else str(price)
    sold_s    = f"{sold:,.0f}"    if isinstance(sold,(int,float))  else str(sold)
    onhand_s  = f"{onhand:,.0f}"  if isinstance(onhand,(int,float))else str(onhand)
    str_pct_s = f"{min(str_pct,100.0):.1f}%" if isinstance(str_pct,(int,float)) else ""
    rev_s     = f"${revenue:,.0f}"if isinstance(revenue,(int,float))else ""
    doc_s2    = f"{int(doc_d)}d"  if doc_d and str(doc_d) not in ("","nan","0","N/A") else ""
    meta      = " · ".join(x for x in [brand,cat] if x and x!="nan")
    launch_s  = f"<div class='prod-meta'>📅 {launch}</div>" \
                if launch and launch not in ("Not sold yet","nan","") else ""

    st.markdown(f"""
    <div class="prod-card">
      {img_html}
      <div class="prod-body">
        <div class="prod-name" title="{name}">{name}</div>
        <div class="prod-meta">{meta}</div>
        <div class="prod-meta">{price_s} · {sold_s} sold · {onhand_s} stock · {rev_s}</div>
        <span class="badge" style="background:{str_bg};color:{str_fg}">{str_s} {str_pct_s}</span>
        <span class="badge" style="background:{abc_bg};color:{abc_fg}">ABC-{abc}</span>
        <span class="badge" style="background:{doc_bg};color:{doc_fg}">{doc_s} {doc_s2}</span>
        {launch_s}
      </div>
    </div>""", unsafe_allow_html=True)

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    df, err = load_data()
    if df is None:
        st.error(f"Could not load data: {err}")
        st.info("Check that the Google Drive file is shared with the service account and the File ID is correct.")
        st.code(f"File ID being used: {GDRIVE_FILE_ID}")
        st.stop()
    df = clean_df(df)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 👗 Salt Fashion")
        st.markdown("**Intelligence Dashboard**")
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Brand
        df["Brand"] = df["Brand"].fillna("").astype(str).str.strip()
        brands = sorted([b for b in df["Brand"].unique()
                         if b and b not in ("nan","True","False","None","")])
        sel_brand = st.selectbox("Brand", options=brands, index=0)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # STR filter
        st.markdown("**Sell-Through Filter**")
        all_strs = [s for s in STR_ORDER if s in df["STR Status"].unique()]
        for s in df["STR Status"].unique():
            if s not in all_strs and s not in ("nan","","None"):
                all_strs.append(s)
        sel_strs = []
        for s in all_strs:
            cnt = len(df[df["STR Status"]==s])
            bg,fg = STR_COLORS.get(s,("#9E9E9E","#FFFFFF"))
            if st.checkbox(f"{s} ({cnt:,})", value=True, key=f"str_{s}"):
                sel_strs.append(s)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # ABC filter
        st.markdown("**ABC Class**")
        sel_abc = []
        bdf_tmp = df[df["Brand"]==sel_brand] if sel_brand else df
        for abc, label in [("A","A — Top 20% revenue"),
                           ("B","B — Mid 30% revenue"),
                           ("C","C — Low 50% revenue")]:
            cnt = len(bdf_tmp[bdf_tmp["ABC Class"]==abc])
            if st.checkbox(f"{label} ({cnt:,})", value=True, key=f"abc_{abc}"):
                sel_abc.append(abc)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Category
        bdf_cats = df[df["Brand"]==sel_brand] if sel_brand else df
        cats = sorted([str(c) for c in bdf_cats["Category"].unique()
                       if str(c).strip() not in ("nan","True","False","None","")])
        sel_cats = st.multiselect("Category", options=cats, default=cats)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        search  = st.text_input("Search", placeholder="Product name...")
        sort_by = st.selectbox("Sort by", [
            "STR % (High first)",
            "Revenue (High)",
            "Total Units Sold (High)",
            "Days of Cover (Low — urgent first)",
            "Sales Price (High)",
            "Sales Price (Low)",
        ])
        per_page = st.selectbox("Per page", [12,24,48], index=0)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        if st.button("🔄 Refresh"):
            st.cache_data.clear(); st.rerun()

    # ── Filter ────────────────────────────────────────────────────────────────
    f = df.copy()
    if sel_brand:
        f = f[f["Brand"] == sel_brand]
    if sel_strs:
        f = f[f["STR Status"].isin(sel_strs)]
    if sel_abc and "ABC Class" in f.columns:
        f = f[f["ABC Class"].isin(sel_abc)]
    if sel_cats:
        f = f[f["Category"].astype(str).isin(sel_cats)]
    if search.strip():
        f = f[f["Product Name"].str.contains(search.strip(), case=False, na=False)]

    # Sort
    if sort_by == "STR % (High first)":
        f = f.sort_values("Sell-Through %", ascending=False)
    elif sort_by == "Revenue (High)":
        f = f.sort_values("Revenue", ascending=False)
    elif sort_by == "Total Units Sold (High)":
        f = f.sort_values("Total Units Sold", ascending=False)
    elif sort_by == "Days of Cover (Low — urgent first)":
        f["_doc"] = pd.to_numeric(f.get("Days of Cover"), errors="coerce")
        f = f.sort_values("_doc", ascending=True)
    elif sort_by == "Sales Price (High)":
        f = f.sort_values("Sales Price", ascending=False)
    elif sort_by == "Sales Price (Low)":
        f = f.sort_values("Sales Price", ascending=True)

    # ── Metrics ───────────────────────────────────────────────────────────────
    bdf = df[df["Brand"]==sel_brand].copy() if sel_brand else df.copy()
    total    = len(bdf)
    sf       = len(bdf[bdf["STR Status"]=="Super Fast"])
    fast     = len(bdf[bdf["STR Status"]=="Fast"])
    medium   = len(bdf[bdf["STR Status"]=="Medium"])
    slow     = len(bdf[bdf["STR Status"]=="Slow"])
    dead     = len(bdf[bdf["STR Status"]=="Dead"])
    rev_tot  = bdf["Revenue"].sum() if "Revenue" in bdf.columns else 0
    abc_a    = len(bdf[bdf["ABC Class"]=="A"]) if "ABC Class" in bdf.columns else 0
    reorder  = len(bdf[bdf["DOC Status"]=="Reorder Now"]) \
               if "DOC Status" in bdf.columns else 0

    st.markdown(f"## {sel_brand or 'All Brands'} — Product Intelligence")
    st.markdown(f"Showing **{len(f):,}** of {total:,} products · {sort_by}")

    c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
    for col,val,lbl,clr in [
        (c1, total,  "Total",         "#111827"),
        (c2, sf,     "⚡ Super Fast", "#1B5E20"),
        (c3, fast,   "🟢 Fast",       "#43A047"),
        (c4, medium, "🟡 Medium",     "#F9A825"),
        (c5, slow,   "🔴 Slow",       "#E53935"),
        (c6, dead,   "⚫ Dead",       "#424242"),
        (c7, abc_a,  "🔵 ABC-A",      "#1565C0"),
        (c8, reorder,"🚨 Reorder",    "#B71C1C"),
    ]:
        with col:
            st.markdown(
                f'<div class="metric-card">'
                f'<p class="metric-val" style="color:{clr}">{val:,}</p>'
                f'<p class="metric-lbl">{lbl}</p></div>',
                unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Insight
    slow_stock = bdf[bdf["STR Status"].isin(["Slow","Dead"])]["On Hand Qty"].sum() \
                 if "On Hand Qty" in bdf.columns else 0
    reorder_names = bdf[bdf["DOC Status"]=="Reorder Now"]["Product Name"].head(3).tolist() \
                    if "DOC Status" in bdf.columns else []
    insights = []
    if slow + dead > 0:
        insights.append(
            f"<b>{slow+dead:,}</b> slow/dead products with "
            f"<b>{slow_stock:,.0f} units stuck in stock</b> — consider markdown or clearance."
        )
    if reorder > 0:
        names = ", ".join(reorder_names[:3])
        insights.append(
            f"<b>🚨 {reorder} products need reordering now</b> (running out within 14 days): {names}..."
        )
    if rev_tot > 0:
        insights.append(
            f"Total revenue from sold units: <b>${rev_tot:,.0f}</b>"
        )
    if insights:
        st.markdown(
            '<div class="insight">💡 ' + " &nbsp;|&nbsp; ".join(insights) + "</div>",
            unsafe_allow_html=True)

    # ── Products grid ─────────────────────────────────────────────────────────
    if len(f) == 0:
        st.info("No products match your filters."); st.stop()

    pages = max(1,(len(f)-1)//per_page+1)
    page  = st.number_input(f"Page (1–{pages})", min_value=1,
                            max_value=pages, value=1) if pages>1 else 1
    pf    = f.iloc[(page-1)*per_page : page*per_page]

    COLS = 4
    for r in range((len(pf)+COLS-1)//COLS):
        cols = st.columns(COLS)
        for c in range(COLS):
            idx = r*COLS+c
            if idx < len(pf):
                with cols[c]:
                    product_card(pf.iloc[idx])

    # ── Category breakdown ────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Category Breakdown")
    if "STR Status" in f.columns and "Category" in f.columns:
        cd = f.copy()
        cd["Category"] = cd["Category"].fillna("—").astype(str).str.strip()
        pivot = cd.groupby(["Category","STR Status"]).size().unstack(fill_value=0)
        ordered = [s for s in STR_ORDER if s in pivot.columns]
        if ordered: pivot = pivot[ordered]
        pivot["Total"] = pivot.sum(axis=1)
        st.dataframe(pivot.sort_values("Total",ascending=False).head(25),
                     width='stretch')

    # ── ABC Summary ───────────────────────────────────────────────────────────
    if "ABC Class" in f.columns and "Revenue" in f.columns:
        st.markdown("### ABC Revenue Analysis")
        abc_sum = f.groupby("ABC Class").agg(
            Products=("Product Name","count"),
            Revenue=("Revenue","sum"),
            Units_Sold=("Total Units Sold","sum"),
            Avg_STR=("Sell-Through %","mean"),
        ).round(1)
        st.dataframe(abc_sum, width='stretch')

if __name__ == "__main__":
    main()