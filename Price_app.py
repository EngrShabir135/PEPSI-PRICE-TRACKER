"""
PEPSI PRICE TRACKER — Streamlit Dashboard
------------------------------------------
Rebuilds the NTP / CP / Retailer Margin dashboard from a weekly price file.

HOW TO USE
1. pip install -r requirements.txt
2. streamlit run pepsi_price_tracker_app.py
3. Load data two ways (sidebar):
   a) Upload a .xlsx file directly, OR
   b) Paste a Google Drive share link to an Excel file and click "Fetch from Drive"
      (the file must be shared as "Anyone with the link can view").
   Do this every week with the new file / new link -> the app remembers every
   period you've ever loaded (saved to history_store.parquet next to this
   script) and builds the "Last 12 periods" trend charts automatically.

EXPECTED COLUMNS (from your source file):
UNIQUE SKU, REGION, CHANNEL, MCAT, CAT, COMPANY, BRAND, SKUS, PKG,
YEAR, MON, WEEK, PERIOD, PERIOD 2, NTP/Case, NTP/6P, PROMO, REG TP, CP
"""

import os
import re
import io
import numpy as np
import pandas as pd
import requests
import streamlit as st
import plotly.graph_objects as go

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
st.set_page_config(page_title="Pepsi Price Tracker", layout="wide")

HISTORY_FILE = "history_store.parquet"
PKG_ORDER = [                                    # row order, matches screenshot
    "SSRB", "300-350 ML PET", "500ML PET", "1LTR PET",
    "1.5LTR PET", "2LTR PET", "2.25LTR PET", "250ML CAN",
    "200ML TP", "355ML", "600ML PET", "250ML",
]
# Priority order for the main-category filter (CSD/JUICE/WATER/ENERGY/REVIVE
# shown first if present in your data; anything else found gets appended)
MCAT_PRIORITY = ["CSD", "JUICE", "JUICES", "WATER", "ENERGY", "ED", "REVIVE"]

# Retailer Margin formula — EDIT HERE if your real formula differs.
# Current assumption: margin on a per-consumer-unit basis.
def retailer_margin(cp, ntp_per_case, pkg):
    if pd.isna(cp) or pd.isna(ntp_per_case) or pd.isna(pkg) or cp == 0 or pkg == 0:
        return np.nan
    ntp_per_unit = ntp_per_case / pkg
    return (cp - ntp_per_unit) / cp


# --------------------------------------------------------------------------
# GOOGLE DRIVE DOWNLOAD HELPERS
# --------------------------------------------------------------------------
def extract_drive_file_id(url: str):
    """Pull the file ID out of common Google Drive share-link formats."""
    patterns = [
        r"/d/([a-zA-Z0-9_-]+)",       # .../file/d/<ID>/view
        r"id=([a-zA-Z0-9_-]+)",       # ...?id=<ID>
    ]
    for pat in patterns:
        m = re.search(pat, url)
        if m:
            return m.group(1)
    return None


def download_drive_file(url: str) -> bytes:
    """Download a Drive file (handles the 'file too large to scan' confirm token)."""
    file_id = extract_drive_file_id(url)
    if not file_id:
        raise ValueError("Couldn't find a file ID in that link. Make sure it's a "
                          "normal Google Drive share link.")
    session = requests.Session()
    base = "https://drive.google.com/uc?export=download"
    resp = session.get(base, params={"id": file_id}, stream=True)

    # Large files get an interstitial page with a confirm token
    token = None
    for key, value in resp.cookies.items():
        if key.startswith("download_warning"):
            token = value
    if token is None and "text/html" in resp.headers.get("Content-Type", ""):
        m = re.search(r"confirm=([0-9A-Za-z_-]+)", resp.text)
        if m:
            token = m.group(1)

    if token:
        resp = session.get(base, params={"id": file_id, "confirm": token}, stream=True)

    resp.raise_for_status()
    content = resp.content
    if content[:2] not in (b"PK",):  # not a valid .xlsx (zip) file
        raise ValueError("That link didn't return a valid Excel file. Double-check "
                          "sharing is set to 'Anyone with the link' and it points to "
                          "an .xlsx file.")
    return content


# --------------------------------------------------------------------------
# DATA LOADING / HISTORY STORE
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def read_excel_bytes(content: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(content))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_history() -> pd.DataFrame:
    if os.path.exists(HISTORY_FILE):
        return pd.read_parquet(HISTORY_FILE)
    return pd.DataFrame()


def save_history(df: pd.DataFrame):
    df.to_parquet(HISTORY_FILE, index=False)


def merge_into_history(new_df: pd.DataFrame):
    history = load_history()
    if history.empty:
        merged = new_df
    else:
        periods_in_new = new_df["PERIOD"].unique()
        history = history[~history["PERIOD"].isin(periods_in_new)]
        merged = pd.concat([history, new_df], ignore_index=True)
    save_history(merged)
    return merged


# --------------------------------------------------------------------------
# SIDEBAR — data source
# --------------------------------------------------------------------------
st.sidebar.header("📂 Data source")
source_mode = st.sidebar.radio("Load data from", ["Upload file", "Google Drive link"])

if source_mode == "Upload file":
    uploads = st.sidebar.file_uploader(
        "Upload weekly file(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True
    )
    if uploads:
        for f in uploads:
            df_new = read_excel_bytes(f.read())
            merge_into_history(df_new)
        st.sidebar.success(f"Loaded {len(uploads)} file(s) into history.")

else:
    drive_url = st.sidebar.text_input("Paste Google Drive share link to the .xlsx file")
    if st.sidebar.button("Fetch from Drive") and drive_url:
        with st.spinner("Downloading and reading file from Drive..."):
            try:
                content = download_drive_file(drive_url)
                df_new = read_excel_bytes(content)
                merge_into_history(df_new)
                st.sidebar.success("File fetched and added to history.")
            except Exception as e:
                st.sidebar.error(f"Couldn't load that file: {e}")

data = load_history()

if data.empty:
    st.title("PEPSI PRICE TRACKER")
    st.info("Upload a file or paste a Google Drive link in the sidebar to build the dashboard.")
    st.stop()

# Clean up category text case (source data has 'Water' and 'WATER' etc.)
data["MCAT"] = data["MCAT"].astype(str).str.upper().str.strip()
data["CAT"] = data["CAT"].astype(str).str.upper().str.strip()
data["COMPANY"] = data["COMPANY"].astype(str).str.strip()
data["BRAND"] = data["BRAND"].astype(str).str.strip()

all_periods = sorted(data["PERIOD"].unique())

# --------------------------------------------------------------------------
# SIDEBAR — filters
# --------------------------------------------------------------------------
st.sidebar.header("📅 Period")
period = st.sidebar.selectbox("Reporting period", all_periods, index=len(all_periods) - 1)

st.sidebar.header("📡 Channel")
channel = st.sidebar.radio("Channel", sorted(data["CHANNEL"].unique()), horizontal=True)

st.sidebar.header("🗺️ Regions")
all_regions = sorted(data["REGION"].unique())
select_all_regions = st.sidebar.checkbox("Select all regions", value=True)
sel_regions = all_regions if select_all_regions else st.sidebar.multiselect(
    "Choose regions", all_regions, default=all_regions
)

st.sidebar.header("🥤 Main category")
available_mcats = sorted(
    data["MCAT"].unique(),
    key=lambda x: (MCAT_PRIORITY.index(x) if x in MCAT_PRIORITY else 99, x),
)
sel_mcats = st.sidebar.multiselect("Main category", available_mcats, default=available_mcats)

st.sidebar.header("🏷️ Sub-category")
cat_pool = sorted(data.loc[data["MCAT"].isin(sel_mcats), "CAT"].unique())
select_all_cats = st.sidebar.checkbox("Select all sub-categories", value=True)
sel_cats = cat_pool if select_all_cats else st.sidebar.multiselect(
    "Choose sub-categories", cat_pool, default=cat_pool
)

st.sidebar.header("🏢 Companies")
company_pool = sorted(data["COMPANY"].unique())
default_base = "PEP" if "PEP" in company_pool else company_pool[0]
base_company = st.sidebar.selectbox(
    "Base company (used for the index % columns)", company_pool,
    index=company_pool.index(default_base),
)
default_compare = [c for c in ["KO", "MBP", "GF"] if c in company_pool and c != base_company]
comparison_companies = st.sidebar.multiselect(
    "Compare against", [c for c in company_pool if c != base_company],
    default=default_compare,
)
SELECTED_COMPANIES = [base_company] + [c for c in comparison_companies if c != base_company]

st.sidebar.header("🍹 Brands")
brand_pool = sorted(data.loc[data["COMPANY"].isin(SELECTED_COMPANIES), "BRAND"].unique())
select_all_brands = st.sidebar.checkbox("Select all brands", value=True)
sel_brands = brand_pool if select_all_brands else st.sidebar.multiselect(
    "Choose brands", brand_pool, default=brand_pool
)

# --------------------------------------------------------------------------
# FILTER (current period, for the main tables)
# --------------------------------------------------------------------------
mask = (
    (data["PERIOD"] == period)
    & (data["CHANNEL"] == channel)
    & (data["REGION"].isin(sel_regions))
    & (data["MCAT"].isin(sel_mcats))
    & (data["CAT"].isin(sel_cats))
    & (data["COMPANY"].isin(SELECTED_COMPANIES))
    & (data["BRAND"].isin(sel_brands))
)
current = data[mask].copy()

# --------------------------------------------------------------------------
# HEADER
# --------------------------------------------------------------------------
left, right = st.columns([4, 1])
with left:
    st.markdown(
        "<h1 style='background:#0b2545;color:white;padding:14px;"
        "border-radius:6px;text-align:center;'>PEPSI PRICE TRACKER</h1>",
        unsafe_allow_html=True,
    )
with right:
    st.markdown(
        f"<div style='background:#fff8dc;border:2px solid gold;padding:14px;"
        f"text-align:center;border-radius:6px;font-weight:bold;'>{period}</div>",
        unsafe_allow_html=True,
    )

st.caption(
    f"Channel: **{channel}**  |  Regions: **{len(sel_regions)}**  |  "
    f"Main cat: **{', '.join(sel_mcats) if sel_mcats else 'None'}**  |  "
    f"Companies: **{', '.join(SELECTED_COMPANIES)}**  |  "
    f"Brands selected: **{len(sel_brands)}**"
)

if current.empty:
    st.warning("No rows match the current filter combination. Adjust filters in the sidebar.")
    st.stop()


# --------------------------------------------------------------------------
# TABLE BUILDERS
# --------------------------------------------------------------------------
def pivot_metric(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    """SKU rows x Company columns, averaged over selected regions."""
    piv = df.pivot_table(index="SKUS", columns="COMPANY", values=value_col, aggfunc="mean")
    piv = piv.reindex(columns=SELECTED_COMPANIES)
    piv = piv.apply(pd.to_numeric, errors="coerce")
    ordered = [p for p in PKG_ORDER if p in piv.index]
    remaining = [p for p in piv.index if p not in ordered]
    piv = piv.loc[ordered + remaining]
    return piv


def add_index_columns(piv: pd.DataFrame) -> pd.DataFrame:
    out = piv.copy()
    if base_company not in piv.columns:
        return out
    for comp in comparison_companies:
        if comp in piv.columns:
            out[f"{base_company} vs {comp}"] = (piv[base_company] / piv[comp] * 100).round(0)
    return out


def style_index_table(df: pd.DataFrame):
    idx_cols = [c for c in df.columns if " vs " in c]

    def color(v):
        if pd.isna(v):
            return ""
        if v >= 100:
            return "color: #1a7a1a; font-weight:600;"
        return "color: #c0392b; font-weight:600;"

    fmt = {c: "{:.0f}%" for c in idx_cols}
    fmt.update({c: "{:,.0f}" for c in SELECTED_COMPANIES if c in df.columns})
    styler = df.style.format(fmt, na_rep="")
    if hasattr(styler, "map"):
        return styler.map(color, subset=idx_cols)
    return styler.applymap(color, subset=idx_cols)


def trend_chart(hist_df: pd.DataFrame, value_col: str, title=""):
    periods = sorted(hist_df["PERIOD"].unique())[-12:]  # last 12 periods
    sub = hist_df[hist_df["PERIOD"].isin(periods)]
    grp = sub.groupby(["PERIOD", "COMPANY"])[value_col].mean().reset_index()

    fig = go.Figure()
    colors = ["#0b2545", "#c0392b", "#2e8b57", "#e67e22"]
    plot_companies = ([base_company] + comparison_companies[:1])  # base + 1st comparator
    for comp, color in zip(plot_companies, colors):
        vals = [
            grp.loc[(grp["PERIOD"] == p) & (grp["COMPANY"] == comp), value_col].mean()
            for p in periods
        ]
        fig.add_bar(name=comp, x=periods, y=vals, marker_color=color)
    fig.update_layout(
        title=title, barmode="group", height=320,
        margin=dict(l=10, r=10, t=40, b=10), legend=dict(orientation="h"),
    )
    return fig


# --------------------------------------------------------------------------
# SECTION RENDERER
# --------------------------------------------------------------------------
def render_section(title: str, value_col: str, is_margin=False):
    st.subheader(title)
    col_table, col_chart = st.columns([2, 1.3])

    if is_margin:
        base = current.pivot_table(index="SKUS", columns="COMPANY",
                                    values=["CP", "NTP/Case", "PKG"], aggfunc="mean")
        rows = {}
        for sku in base.index:
            row = {}
            for comp in SELECTED_COMPANIES:
                try:
                    cp = base.loc[sku, ("CP", comp)]
                    ntp = base.loc[sku, ("NTP/Case", comp)]
                    pkg = base.loc[sku, ("PKG", comp)]
                    row[comp] = retailer_margin(cp, ntp, pkg)
                except KeyError:
                    row[comp] = np.nan
            rows[sku] = row
        piv = pd.DataFrame.from_dict(rows, orient="index")
        piv = piv.reindex(columns=SELECTED_COMPANIES)
        piv = piv.apply(pd.to_numeric, errors="coerce") * 100
        ordered = [p for p in PKG_ORDER if p in piv.index]
        remaining = [p for p in piv.index if p not in ordered]
        piv = piv.loc[ordered + remaining]
        piv_idx = piv.copy()
        if base_company in piv.columns:
            for comp in comparison_companies:
                if comp in piv.columns:
                    piv_idx[f"{base_company} vs {comp}"] = (piv[base_company] - piv[comp]).round(0)
        with col_table:
            st.dataframe(style_index_table(piv_idx), use_container_width=True)
    else:
        piv = pivot_metric(current, value_col)
        piv_idx = add_index_columns(piv)
        with col_table:
            st.dataframe(style_index_table(piv_idx), use_container_width=True)

    with col_chart:
        hist_slice = data[
            (data["CHANNEL"] == channel)
            & (data["COMPANY"].isin(SELECTED_COMPANIES))
            & (data["BRAND"].isin(sel_brands))
        ]
        fig = trend_chart(
            hist_slice, value_col if not is_margin else "CP",
            title="Last 12 periods trend (avg, filtered SKUs)",
        )
        chart_key = f"trend_chart_{title.replace(' ', '_').lower()}"
        st.plotly_chart(fig, use_container_width=True, key=chart_key)


# --------------------------------------------------------------------------
# RENDER PAGE
# --------------------------------------------------------------------------
render_section("NTP", "NTP/Case")
st.divider()
render_section("CP", "CP")
st.divider()
render_section("RETAILER MARGIN", "CP", is_margin=True)

st.divider()
with st.expander("📄 Filtered raw data (current period)"):
    st.dataframe(current, use_container_width=True)
