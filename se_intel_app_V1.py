"""
SE Intel Dashboard — Python / Streamlit (optimised)
pip install streamlit pandas openpyxl xlrd pyarrow plotly
python -m streamlit run se_intel_app.py
"""

import io
import hashlib
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="SE Intel Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# ACCESS CONTROL
# ─────────────────────────────────────────────────────────────────────────────
def _hash(s: str) -> str:
    return hashlib.sha256(s.encode()).hexdigest()

def _check_launch_key() -> bool:
    try:
        stored = st.secrets["auth"]["launch_key"]
        return bool(stored)
    except Exception:
        st.error("⛔ Launch key not found. Please create `.streamlit/secrets.toml`.")
        st.code("""[auth]\nlaunch_key = "YOUR_KEY"\nadmin_key = "YOUR_ADMIN_KEY"\n\n[users]\nusername = "sha256_hash"
""", language="toml")
        st.stop()

def _login_screen():
    st.markdown("""
    <div style='max-width:400px;margin:80px auto 0;padding:2rem 2.5rem;
    background:white;border-radius:12px;box-shadow:0 4px 24px rgba(31,56,100,.12);'>
    <h2 style='color:#1F3864;margin-bottom:1.5rem;text-align:center;font-size:1.4rem'>
    🔐 SE Intel Dashboard</h2>
    </div>
    """, unsafe_allow_html=True)

    tab_user, tab_admin = st.tabs(["👤 User Login", "🔑 Admin"])

    with tab_user:
        with st.form("login_form"):
            username  = st.text_input("Username", placeholder="Enter your username")
            password  = st.text_input("Password", type="password", placeholder="Enter your password")
            login_btn = st.form_submit_button("Login", use_container_width=True, type="primary")
        if login_btn:
            try:
                users = dict(st.secrets["users"])
            except Exception:
                users = {}
            if username in users and users[username] == _hash(password):
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.rerun()
            else:
                st.error(f"❌ Invalid username or password. Available users: {list(users.keys())}")

    with tab_admin:
        with st.form("admin_form"):
            admin_input = st.text_input("Admin Key", type="password", placeholder="Enter admin key")
            admin_btn   = st.form_submit_button("Enter", use_container_width=True, type="primary")
        if admin_btn:
            try:
                admin_key_stored = st.secrets["auth"].get("admin_key", "")
                if admin_key_stored and admin_input == admin_key_stored:
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = "Admin"
                    st.rerun()
                else:
                    st.error("❌ Invalid admin key.")
            except Exception:
                st.error("❌ Admin key not configured.")

def _check_auth():
    _check_launch_key()
    if not st.session_state.get("authenticated", False):
        _login_screen()
        st.stop()

_check_auth()

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
section[data-testid="stSidebar"] { background: #1F3864; }
section[data-testid="stSidebar"] * { color: #e8edf5 !important; }
section[data-testid="stSidebar"] div[data-baseweb="select"] > div { background-color: #e8edf5 !important; border: none; }
section[data-testid="stSidebar"] div[data-baseweb="select"] * { color: #1F3864 !important; font-weight: 600; }
section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] { background-color: #2E75B6 !important; }
section[data-testid="stSidebar"] label { font-weight:600; font-size:.78rem; letter-spacing:.05em; }
section[data-testid="stSidebar"] h3 {
    color:#BDD7EE !important; font-size:.7rem; font-weight:700;
    letter-spacing:.12em; text-transform:uppercase; margin-top:1.2rem;
    border-bottom:1px solid rgba(255,255,255,.15); padding-bottom:.3rem;
}
.main { background:#F5F7FB; }
h1 { color:#1F3864; font-size:1.5rem; font-weight:700; }
thead tr th { background-color:#1F3864 !important; color:white !important; font-size:.78rem; font-family:'IBM Plex Mono',monospace; }
tbody tr:nth-child(odd)  td { background-color:#F2F2F2; }
tbody tr:nth-child(even) td { background-color:#FFFFFF; }
td { font-size:.80rem; }
.section-title {
    background:#1F3864; color:white; font-weight:700;
    padding:.35rem .8rem; border-radius:4px 4px 0 0;
    font-size:.82rem; letter-spacing:.04em; margin-bottom:0;
}
div[data-testid="metric-container"] {
    background:white; border:1px solid #D6E4F0; border-radius:8px;
    padding:.6rem .9rem; box-shadow:0 1px 4px rgba(31,56,100,.08);
}
div[data-testid="metric-container"] label { color:#2E75B6 !important; font-size:.72rem !important; font-weight:700; }
div[data-testid="stFileUploader"] { border:2px dashed #2E75B6; border-radius:8px; padding:1rem; }
.source-badge { display:inline-block; padding:.25rem .75rem; border-radius:4px; font-size:.75rem; font-weight:700; }
.source-order { background:#2E75B6; color:#fff; }
.source-sales { background:#375623; color:#fff; }
.stProgress > div > div { background-color: #2E75B6; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
TARGET_HEADERS = [
    "Year-Month", "Year", "Month", "Quarter", "Half",
    "Strategic Product Family", "Commercial Reference", "Comm Ref Code",
    "Family", "PLC", "Product Line Label", "Country", "Cluster",
    "Zone", "Area", "Operations", "Market Segment", "Quantity", "Value (EUR)"
]
SOURCE_MAP = {
    "Year-Month":               "year month",
    "Strategic Product Family": "strategic product family",
    "Commercial Reference":     "commercial reference",
    "Comm Ref Code":            "commercial reference code",
    "Family":                   "family",
    "PLC":                      "product line code",
    "Product Line Label":       "product line label",
    "Country":                  "country (country perf.)",
    "Cluster":                  "cluster (country perf.)",
    "Zone":                     "zone (country perf.)",
    "Area":                     "area (country perf.)",
    "Operations":               "operations (country perf.)",
    "Market Segment":           "market segment",
    "Quantity":                 "medium quantity",
    "Value (EUR)":              "medium value",
}
DERIVED   = {"Year", "Month", "Quarter", "Half"}
STR_COLS  = [
    "Zone", "Country", "Cluster", "Area", "Operations",
    "PLC", "Strategic Product Family", "Family",
    "Comm Ref Code", "Commercial Reference", "Market Segment",
    "Year", "Month", "Quarter", "Half",
]

CHART_COLORS = [
    "#2E75B6","#1F3864","#375623","#C55A11","#7030A0",
    "#BDD7EE","#70AD47","#ED7D31","#A9D18E","#9DC3E6",
]

# ─────────────────────────────────────────────────────────────────────────────
# ENGINE 1 — LOAD & ALIGN
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _clean_col(c: str) -> str:
    """Strip, lowercase, remove non-breaking spaces from column names."""
    return c.strip().lower().replace(" ", " ").replace("  ", " ")

@st.cache_data(show_spinner=False)
def load_and_align(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    if file_name.lower().endswith(".csv"):
        src = pd.read_csv(buf, dtype=str)
    else:
        src = pd.read_excel(buf, sheet_name=0, dtype=str)

    # Normalise column names including   non-breaking spaces
    src.columns = [_clean_col(c) for c in src.columns]

    # Drop columns we don't need → massive memory reduction on wide files
    needed_cols = {_clean_col(v) for v in SOURCE_MAP.values() if v}
    drop_cols   = [c for c in src.columns if c not in needed_cols]
    if drop_cols:
        src.drop(columns=drop_cols, inplace=True)

    out = {}
    for tgt in TARGET_HEADERS:
        if tgt in DERIVED:
            out[tgt] = ""
            continue
        src_col = _clean_col(SOURCE_MAP.get(tgt, tgt.lower()))
        out[tgt] = src[src_col].values if src_col in src.columns else ""

    df = pd.DataFrame(out)
    del src  # free wide dataframe immediately

    ym     = df["Year-Month"].astype(str).str.strip()
    mask   = ym.str.len() >= 6
    mo_int = pd.to_numeric(ym.str[4:6], errors="coerce").fillna(0).astype(int)

    df["Year"]    = ym.str[:4].where(mask, "")
    df["Month"]   = mo_int.astype(str).where(mask, "")
    df["Quarter"] = ((mo_int - 1) // 3 + 1).astype(str).where(mask, "")
    df["Half"]    = mo_int.apply(lambda m: "1" if 1 <= m <= 6 else ("2" if m > 6 else "")).where(mask, "")

    df["Value (EUR)"] = pd.to_numeric(df["Value (EUR)"], errors="coerce").fillna(0.0)
    df["Quantity"]    = pd.to_numeric(df["Quantity"],    errors="coerce").fillna(0.0)

    for c in STR_COLS:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()

    # Convert string columns to category — reduces memory by ~95%
    for c in df.columns:
        if df[c].dtype == object:
            if df[c].nunique() < len(df) * 0.5:
                df[c] = df[c].astype("category")

    return df

def load_file_with_progress(uploaded_file):
    bar_area  = st.empty()
    info_area = st.empty()
    with bar_area.container():
        bar = st.progress(0, text="📂 Reading file…")
    raw_bytes = uploaded_file.read()
    bar.progress(25, text="📂 File received, parsing columns…")
    info_area.caption(f"Processing **{uploaded_file.name}** ({len(raw_bytes)/1024/1024:.1f} MB)…")
    df = load_and_align(raw_bytes, uploaded_file.name)
    bar.progress(70, text="🔧 Deriving Year / Month / Quarter / Half…")
    _ = [df[c].unique() for c in STR_COLS if c in df.columns]
    bar.progress(95, text="✅ Finalising…")
    bar.progress(100, text=f"✅ Done — {len(df):,} rows loaded")
    import time; time.sleep(0.4)
    bar_area.empty()
    info_area.empty()
    return df

# ─────────────────────────────────────────────────────────────────────────────
# ENGINE 2 — FAST VECTORISED FILTER
# ─────────────────────────────────────────────────────────────────────────────
def apply_filters(df: pd.DataFrame, fmap: dict) -> pd.DataFrame:
    mask = pd.Series(True, index=df.index)
    for col, sel in fmap.items():
        if not sel:
            continue
        if col in df.columns:
            col_s = df[col].astype(str) if hasattr(df[col], "cat") else df[col]
            mask &= col_s.isin([str(s) for s in sel])
    return df.loc[mask]  # view, no copy — saves memory

# ─────────────────────────────────────────────────────────────────────────────
# ENGINE 3 — TABLE COMPUTATION (cached per filter state)
# ─────────────────────────────────────────────────────────────────────────────
def fmt_pct(v) -> str:
    if v is None: return "—"
    try:    return f"{float(v):+.1%}"
    except: return "—"

def fmt_val(v: float) -> str:
    if v >= 1_000_000: return f"{v/1_000_000:,.2f}M"
    if v >= 1_000:     return f"{v/1_000:,.1f}K"
    return f"{v:,.0f}"

def growth(curr, prev):
    if not prev or prev == 0: return None
    return (curr - prev) / abs(prev)

def df_to_bytes(df: pd.DataFrame) -> bytes:
    """Legacy stub — no longer used for compute_tables. Kept for compatibility."""
    if df is None or df.empty:
        return b""
    return b"nonempty"  # just a truthy marker

def bytes_to_df(b: bytes) -> pd.DataFrame:
    """Legacy stub."""
    return pd.DataFrame()

def compute_tables(
    df_c: pd.DataFrame, df_py: pd.DataFrame,
    df_pq: pd.DataFrame, df_pm: pd.DataFrame,
    df_cq: pd.DataFrame, df_cm: pd.DataFrame,
    geo_all: bool, flag_yoy: bool, flag_qoq: bool, flag_mom: bool,
):
    """Compute ranking tables directly from DataFrames — no parquet serialisation."""

    g_total = df_c["Value (EUR)"].sum() if not df_c.empty else 0.0
    if g_total == 0:
        return None, 0.0

    def safe_agg(df: pd.DataFrame, col: str) -> pd.DataFrame:
        if df.empty or col not in df.columns:
            return pd.DataFrame(columns=["Value (EUR)", "Quantity"])
        return df.groupby(col)[["Value (EUR)", "Quantity"]].sum()

    def safe_composite(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty or "Commercial Reference" not in df.columns or "Country" not in df.columns:
            return pd.DataFrame(columns=["Value (EUR)", "Quantity"])
        tmp = df.copy()
        tmp["_ck"] = tmp["Commercial Reference"] + "|" + tmp["Country"]
        return tmp.groupby("_ck")[["Value (EUR)", "Quantity"]].sum()

    def build(group_col, parent_col=None, country_col=None, top_n=20, composite=False):
        if composite and not geo_all:
            curr_df = safe_composite(df_c)
            py_df   = safe_composite(df_py)
            cq_df   = safe_composite(df_cq)
            pq_df   = safe_composite(df_pq)
            cm_df   = safe_composite(df_cm)
            pm_df   = safe_composite(df_pm)
            if not df_c.empty and "Commercial Reference" in df_c.columns:
                df_c2 = df_c.copy()
                df_c2["_ck"] = df_c2["Commercial Reference"] + "|" + df_c2["Country"]
                par_map = df_c2.groupby("_ck")["Comm Ref Code"].first() if parent_col else None
                cty_map = df_c2.groupby("_ck")["Country"].first()
            else:
                par_map = None; cty_map = None
        else:
            curr_df = safe_agg(df_c,  group_col)
            py_df   = safe_agg(df_py, group_col)
            cq_df   = safe_agg(df_cq, group_col)
            pq_df   = safe_agg(df_pq, group_col)
            cm_df   = safe_agg(df_cm, group_col)
            pm_df   = safe_agg(df_pm, group_col)
            par_map = df_c.groupby(group_col)[parent_col].first() if (parent_col and not df_c.empty and parent_col in df_c.columns) else None
            cty_map = df_c.groupby(group_col)[country_col].first() if (country_col and not df_c.empty and country_col in df_c.columns) else None

        if curr_df.empty:
            return pd.DataFrame()

        top = curr_df.nlargest(top_n, "Value (EUR)")
        rows = []
        
        def get_val(df_in, k):
            if df_in.empty: return 0.0
            return df_in.loc[k, "Value (EUR)"] if k in df_in.index else 0.0

        for rank, (key, row_data) in enumerate(top.iterrows(), 1):
            disp = key.split("|")[0] if "|" in str(key) else key
            val = row_data["Value (EUR)"]
            qty = row_data["Quantity"]

            py_val = get_val(py_df, key)
            cq_val = get_val(cq_df, key)
            pq_val = get_val(pq_df, key)
            cm_val = get_val(cm_df, key)
            pm_val = get_val(pm_df, key)

            row = {
                "Rank":        rank,
                "Name":        disp,
                "Value (EUR)": val,
                "QTY":         qty,
                "% Total":     f"{val/g_total:.1%}",
                "YoY %":       fmt_pct(growth(val, py_val)) if flag_yoy else "—",
                "QoQ %":       fmt_pct(growth(cq_val, pq_val)) if flag_qoq else "—",
                "MoM %":       fmt_pct(growth(cm_val, pm_val)) if flag_mom else "—",
            }
            if par_map is not None and key in par_map.index:
                row["Parent"] = par_map[key]
            if cty_map is not None and not geo_all and key in cty_map.index:
                row["Country"] = cty_map[key]
            rows.append(row)

        return pd.DataFrame(rows)

    tables = {
        "zone":    build("Zone",                  top_n=100),
        "country": build("Country",               parent_col="Zone",          country_col="Cluster", top_n=20),
        "plc":     build("PLC",                   top_n=20),
        "family":  build("Family",                parent_col="PLC",           top_n=20),
        "crc":     build("Comm Ref Code",         parent_col="Family",        top_n=20),
        "cr":      build("Commercial Reference",  parent_col="Comm Ref Code", top_n=20, composite=True),
    }
    return tables, g_total

# ─────────────────────────────────────────────────────────────────────────────
# CHART HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def make_pie(df_table: pd.DataFrame, name_col: str, top_n: int, title: str):
    if df_table is None or df_table.empty:
        st.caption("— No data for chart —")
        return
    d = df_table.head(top_n).copy()
    fig = px.pie(
        d, names=name_col, values="Value (EUR)", title=title,
        color_discrete_sequence=CHART_COLORS, hole=0.35,
    )
    fig.update_traces(textposition="outside", textinfo="percent+label")
    fig.update_layout(
        margin=dict(t=40, b=10, l=10, r=10), legend=dict(orientation="v", x=1.02, y=0.5),
        showlegend=True, height=380, font=dict(family="IBM Plex Sans"),
        title_font=dict(size=13, color="#1F3864"),
    )
    st.plotly_chart(fig, use_container_width=True)

def make_bar_trend(df_raw: pd.DataFrame, group_col: str, top_n: int, title: str, has_month: bool):
    if df_raw is None or df_raw.empty:
        st.caption("— No data for trend chart —")
        return
    top_items = df_raw.groupby(group_col)["Value (EUR)"].sum().nlargest(top_n).index.tolist()
    d = df_raw[df_raw[group_col].isin(top_items)].copy()
    if has_month:
        d["YM_sort"] = d["Year"] + "-" + d["Month"].str.zfill(2)
        agg = d.groupby([group_col, "YM_sort"])["Value (EUR)"].sum().reset_index().sort_values("YM_sort")
        fig = px.bar(agg, x="YM_sort", y="Value (EUR)", color=group_col, barmode="group",
                     title=f"{title} — Monthly Trend", color_discrete_sequence=CHART_COLORS)
        fig.update_xaxes(tickangle=-45, title="Year-Month")
    else:
        agg = d.groupby([group_col, "Year"])["Value (EUR)"].sum().reset_index().sort_values("Year")
        fig = px.bar(agg, x="Year", y="Value (EUR)", color=group_col, barmode="group",
                     title=f"{title} — Annual Trend", color_discrete_sequence=CHART_COLORS)
        fig.update_xaxes(title="Year")
    fig.update_yaxes(title="Value (EUR)")
    fig.update_layout(
        height=380, margin=dict(t=40, b=50, l=10, r=10), font=dict(family="IBM Plex Sans"),
        title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.25),
    )
    st.plotly_chart(fig, use_container_width=True)

def chart_controls(tab_key: str, max_n: int = 20) -> tuple[int, bool]:
    c1, c2 = st.columns([3, 1])
    with c1:
        real_max = max(max_n, 1)
        real_min = min(3, real_max)
        if real_min >= real_max:
            top_n = real_max
            st.caption(f"Top N: {top_n} (only {top_n} item{'s' if top_n>1 else ''} available)")
        else:
            top_n = st.slider("Top N items in charts", min_value=real_min, max_value=real_max, value=min(10, real_max), key=f"topn_{tab_key}")
    with c2:
        show_month = st.checkbox("Monthly breakdown", value=False, key=f"month_{tab_key}")
    return top_n, show_month

def render_tab_charts(df_raw: pd.DataFrame, df_table: pd.DataFrame, name_col: str, group_col: str, tab_key: str, label: str, max_n: int = 20):
    st.markdown("---")
    st.markdown("#### 📈 Charts")
    top_n, show_month = chart_controls(tab_key, max_n=min(max_n, len(df_table) if df_table is not None and not df_table.empty else max_n))
    col_pie, col_bar = st.columns([1, 1])
    with col_pie: make_pie(df_table, name_col, top_n, f"Top {top_n} {label} — Value Share")
    with col_bar:
        has_month_data = show_month and not df_raw["Month"].replace("", pd.NA).dropna().empty
        make_bar_trend(df_raw, group_col, top_n, f"Top {top_n} {label}", has_month=has_month_data)

# ─────────────────────────────────────────────────────────────────────────────
# PRICE ANALYSIS — ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def _price_agg_annual(df: pd.DataFrame, cr_filter_tuple: tuple) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    cr_filter = list(cr_filter_tuple)
    d = df[df["Quantity"] > 0].copy()
    if cr_filter:
        d = d[d["Commercial Reference"].isin(cr_filter)]
    if d.empty: return pd.DataFrame()
    agg = d.groupby(["Commercial Reference","Comm Ref Code","Family",
                     "Strategic Product Family","Year","Country"], as_index=False
    ).agg(Value=("Value (EUR)","sum"), Quantity=("Quantity","sum"))
    agg["Unit Price (EUR)"] = agg["Value"] / agg["Quantity"]
    agg = agg.rename(columns={"Value":"Value (EUR)","Quantity":"Total Qty"})
    return agg.sort_values(["Commercial Reference","Country","Year"])

def _price_agg_monthly(df: pd.DataFrame, cr_filter_tuple: tuple) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    cr_filter = list(cr_filter_tuple)
    d = df[(df["Quantity"] > 0) & df["Month"].replace("", pd.NA).notna()].copy()
    if cr_filter:
        d = d[d["Commercial Reference"].isin(cr_filter)]
    if d.empty: return pd.DataFrame()
    d["YM"] = d["Year"] + "-" + d["Month"].str.zfill(2)
    agg = d.groupby(["Commercial Reference","Comm Ref Code","Family",
                     "Strategic Product Family","YM","Year","Country"], as_index=False
    ).agg(Value=("Value (EUR)","sum"), Quantity=("Quantity","sum"))
    agg = agg[agg["Quantity"] > 0].copy()
    agg["Unit Price (EUR)"] = agg["Value"] / agg["Quantity"]
    agg = agg.rename(columns={"Value":"Value (EUR)","Quantity":"Total Qty"})
    return agg.sort_values(["Commercial Reference","Country","YM"])

# Byte-based wrappers (legacy, kept for safety)
@st.cache_data(show_spinner=False)
def compute_price_analysis(curr_bytes: bytes, cr_filter_tuple: tuple) -> pd.DataFrame:
    return _price_agg_annual(bytes_to_df(curr_bytes), cr_filter_tuple)

@st.cache_data(show_spinner=False)
def compute_price_monthly(curr_bytes: bytes, cr_filter_tuple: tuple) -> pd.DataFrame:
    return _price_agg_monthly(bytes_to_df(curr_bytes), cr_filter_tuple)

# Direct DataFrame versions (no serialisation overhead)
def compute_price_analysis_direct(df: pd.DataFrame, cr_filter_tuple: tuple) -> pd.DataFrame:
    return _price_agg_annual(df, cr_filter_tuple)

def compute_price_monthly_direct(df: pd.DataFrame, cr_filter_tuple: tuple) -> pd.DataFrame:
    return _price_agg_monthly(df, cr_filter_tuple)


_CHART_H = 380
_LAYOUT  = dict(font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.35))
def _fig(f, h=_CHART_H): f.update_layout(height=h, **_LAYOUT); return f
def _yoy_pct(df, x_col, y_col="Unit Price (EUR)", group_col="Country"):
    rows = []
    for g, grp in df.groupby(group_col):
        grp = grp.sort_values(x_col).copy()
        grp["YoY %"] = grp[y_col].pct_change() * 100
        rows.append(grp)
    return pd.concat(rows).dropna(subset=["YoY %"]) if rows else pd.DataFrame()

def _mom_pct(df, y_col="Unit Price (EUR)", group_col="Country"):
    rows = []
    for g, grp in df.groupby(group_col):
        grp = grp.sort_values("YM").copy()
        grp["MoM %"] = grp[y_col].pct_change() * 100
        rows.append(grp)
    return pd.concat(rows).dropna(subset=["MoM %"]) if rows else pd.DataFrame()

def price_trend_chart(df, x, color, title, is_bar=False):
    fn, kw = (px.bar, dict(barmode="group")) if is_bar else (px.line, dict(markers=True))
    fig = fn(df, x=x, y="Unit Price (EUR)", color=color, title=title, color_discrete_sequence=CHART_COLORS, **kw)
    fig.update_xaxes(tickangle=-40 if x == "YM" else 0)
    return _fig(fig)

def qty_trend_chart(df, x, color, title):
    fig = px.bar(df, x=x, y="Total Qty", color=color, barmode="group", title=title, color_discrete_sequence=CHART_COLORS)
    fig.update_xaxes(tickangle=-40 if x == "YM" else 0)
    return _fig(fig)

def pct_chart(df, x, y, color, title):
    fig = px.bar(df, x=x, y=y, color=color, barmode="group", title=title, color_discrete_sequence=CHART_COLORS)
    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
    fig.update_xaxes(tickangle=-40 if x == "YM" else 0)
    return _fig(fig, h=320)

def dual_axis_chart(df, x, title):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df[x], y=df["Total Qty"], name="Volume (Qty)", marker_color="#BDD7EE", opacity=0.85), secondary_y=False)
    fig.add_trace(go.Scatter(x=df[x], y=df["Unit Price (EUR)"], name="Unit Price (EUR)", mode="lines+markers", line=dict(color="#C55A11", width=2.5), marker=dict(size=7)), secondary_y=True)
    fig.update_layout(title=title, height=_CHART_H, **_LAYOUT)
    fig.update_yaxes(title_text="Volume (Qty)", secondary_y=False)
    fig.update_yaxes(title_text="Unit Price (EUR)", secondary_y=True)
    return fig

def show2(fig_l, fig_r):
    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(fig_l, use_container_width=True)
    with c2: st.plotly_chart(fig_r, use_container_width=True)

def section(title): st.markdown(f"#### {title}"); st.markdown("---")

def render_price_tab(df_curr: pd.DataFrame, fmap: dict):
    st.markdown("#### 💶 Price vs Sales Analysis")
    st.caption("Implied unit price = Value (EUR) ÷ Quantity per row.")
    if df_curr.empty: st.warning("No data in current filter selection."); return
    df_q = df_curr[df_curr["Quantity"] > 0].copy()
    if df_q.empty: st.warning("No rows with Quantity > 0 found."); return

    def _pv(col, df=df_q): return sorted(df[col].replace("", pd.NA).dropna().unique().tolist())
    st.markdown("**🔍 Filter Selection** *(cascade: each level narrows the next)*")
    _p = df_q
    fc1, fc2 = st.columns(2)
    with fc1: sel_plc = st.multiselect("Product Line Code", _pv("PLC"), placeholder="All PLCs…", key="pr_plc")
    if sel_plc: _p = _p[_p["PLC"].isin(sel_plc)]
    with fc2: sel_fam = st.multiselect("Family", _pv("Family", _p), placeholder="All Families…", key="pr_fam")
    if sel_fam: _p = _p[_p["Family"].isin(sel_fam)]
    fc3, fc4 = st.columns(2)
    with fc3: sel_spf = st.multiselect("Strategic Product Family", _pv("Strategic Product Family", _p), placeholder="All SPF…", key="pr_spf")
    if sel_spf: _p = _p[_p["Strategic Product Family"].isin(sel_spf)]
    with fc4: sel_crc = st.multiselect("Comm Ref Code", _pv("Comm Ref Code", _p), placeholder="All CRCs…", key="pr_crc")
    if sel_crc: _p = _p[_p["Comm Ref Code"].isin(sel_crc)]
    fc5, fc6 = st.columns(2)
    with fc5: sel_cr = st.multiselect("Comm Ref (optional)", _pv("Commercial Reference", _p), placeholder="Leave empty for family-level view…", key="pr_cr")
    with fc6: avail_years = _pv("Year", df_q); sel_years = st.multiselect("Years", avail_years, default=avail_years, key="pr_yr")
    
    avail_countries = _pv("Country", _p)
    sel_countries = st.multiselect("Countries", avail_countries, default=avail_countries[:15] if len(avail_countries)>15 else avail_countries, key="pr_cty")
    update_price_btn = st.button("▶ UPDATE PLOTS", use_container_width=True, type="primary", key="pr_update_btn")

    if update_price_btn:
        st.session_state["price_committed"] = {"cr": sel_cr, "years": sel_years, "countries": sel_countries, "fam": sel_fam, "spf": sel_spf, "crc": sel_crc, "plc": sel_plc}
        st.rerun()

    if "price_committed" not in st.session_state or st.session_state["price_committed"] is None:
        st.info("👆 Set your filters above and press **▶ UPDATE PLOTS** to render charts.")
        return
    pc = st.session_state["price_committed"]

    df_base = df_q.copy()
    if pc["plc"]: df_base = df_base[df_base["PLC"].isin(pc["plc"])]
    if pc["fam"]: df_base = df_base[df_base["Family"].isin(pc["fam"])]
    if pc["spf"]: df_base = df_base[df_base["Strategic Product Family"].isin(pc["spf"])]
    if pc["crc"]: df_base = df_base[df_base["Comm Ref Code"].isin(pc["crc"])]
    if pc["years"]: df_base = df_base[df_base["Year"].isin(pc["years"])]
    if pc["countries"]: df_base = df_base[df_base["Country"].isin(pc["countries"])]
    if df_base.empty: st.warning("No data matches the committed filter."); return

    has_cr_filter = bool(pc["cr"])
    all_cr_in_scope = sorted(df_base["Commercial Reference"].unique().tolist())
    annual_price_df  = compute_price_analysis_direct(df_base, tuple(all_cr_in_scope))
    monthly_price_df = compute_price_monthly_direct(df_base, tuple(all_cr_in_scope))

    fam_selected, spf_selected = pc["fam"], pc["spf"]
    if fam_selected and not spf_selected: group_dim, group_vals = "Family", fam_selected
    elif spf_selected and not fam_selected: group_dim, group_vals = "Strategic Product Family", spf_selected
    elif fam_selected and spf_selected: group_dim, group_vals = "Family", fam_selected
    else: group_dim, group_vals = "Family", sorted(df_base["Family"].replace("", pd.NA).dropna().unique().tolist())

    st.markdown("---"); section("🗂 Sub-Product Selector")
    if len(group_vals) > 1: focus_group = st.selectbox(f"Select {group_dim} to explore", group_vals, key="pr_focus_group")
    else: focus_group = group_vals[0] if group_vals else None; st.info(f"Showing: **{group_dim}** = {focus_group}") if focus_group else None
    if not focus_group: st.warning("No product group found."); return

    df_grp = df_base[df_base[group_dim] == focus_group]
    ap_grp = annual_price_df[annual_price_df[group_dim] == focus_group] if not annual_price_df.empty else pd.DataFrame()
    mp_grp = monthly_price_df[monthly_price_df[group_dim] == focus_group] if not monthly_price_df.empty else pd.DataFrame()
    crc_in_group = sorted(df_grp["Comm Ref Code"].replace("", pd.NA).dropna().unique().tolist())
    cr_in_group  = sorted(df_grp["Commercial Reference"].replace("", pd.NA).dropna().unique().tolist())

    section(f"📦 {group_dim} Level — {focus_group}")
    if not ap_grp.empty:
        grp_annual = ap_grp.groupby(["Year","Country"], as_index=False).agg(Value=("Value (EUR)", "sum"), Qty=("Total Qty", "sum"))
        grp_annual = grp_annual[grp_annual["Qty"] > 0].copy()
        grp_annual["Unit Price (EUR)"] = grp_annual["Value"] / grp_annual["Qty"]
        grp_annual.rename(columns={"Qty":"Total Qty","Value":"Value (EUR)"}, inplace=True)
        show2(price_trend_chart(grp_annual.sort_values(["Country","Year"]), "Year", "Country", f"{focus_group} — Avg Unit Price by Year & Country"), qty_trend_chart(grp_annual.sort_values(["Country","Year"]), "Year", "Country", f"{focus_group} — Total Volume by Year & Country"))
        yoy = _yoy_pct(grp_annual, "Year")
        if not yoy.empty: st.plotly_chart(pct_chart(yoy, "Year", "YoY %", "Country", f"{focus_group} — YoY Price Change %"), use_container_width=True)

    if not mp_grp.empty:
        grp_monthly = mp_grp.groupby(["YM","Country"], as_index=False).agg(Value=("Value (EUR)", "sum"), Qty=("Total Qty", "sum"))
        grp_monthly = grp_monthly[grp_monthly["Qty"] > 0].copy()
        grp_monthly["Unit Price (EUR)"] = grp_monthly["Value"] / grp_monthly["Qty"]
        grp_monthly.rename(columns={"Qty":"Total Qty","Value":"Value (EUR)"}, inplace=True)
        show2(price_trend_chart(grp_monthly.sort_values("YM"), "YM", "Country", f"{focus_group} — Monthly Avg Price"), pct_chart(_mom_pct(grp_monthly), "YM", "MoM %", "Country", f"{focus_group} — MoM Price Change %"))

    if crc_in_group:
        section(f"🔖 Comm Ref Code Level — within {focus_group}")
        s4c1, s4c2, s4c3 = st.columns(3)
        with s4c1:
            all_countries_grp = sorted(df_grp["Country"].replace("", pd.NA).dropna().unique().tolist())
            sel_s4_country = st.multiselect("Country", all_countries_grp, default=all_countries_grp[:5] if len(all_countries_grp)>5 else all_countries_grp, key="pr_s4_country")
        with s4c2: s4_period = st.radio("Period", ["Year", "Month"], horizontal=True, key="pr_s4_period")
        with s4c3:
            if s4_period == "Year":
                all_years_grp = sorted(df_grp["Year"].replace("", pd.NA).dropna().unique().tolist())
                sel_s4_periods = st.multiselect("Select Years", all_years_grp, default=all_years_grp, key="pr_s4_years")
            else:
                all_ym = sorted(mp_grp["YM"].unique().tolist()) if not mp_grp.empty else []
                sel_s4_periods = st.multiselect("Select Year-Months", all_ym, default=all_ym[-12:] if len(all_ym)>=12 else all_ym, key="pr_s4_ym")

        def _s4_make_chart(df_src, period_col, grp_keys):
            agg = df_src.groupby(grp_keys + ["Country"], as_index=False).agg(Value=("Value (EUR)", "sum"), Qty=("Total Qty", "sum"))
            agg = agg[agg["Qty"] > 0].copy()
            agg["Unit Price (EUR)"] = agg["Value"] / agg["Qty"]
            agg["Total Qty"] = agg["Qty"]
            return agg

        period_col = "Year" if s4_period == "Year" else "YM"
        src = ap_grp.copy() if s4_period == "Year" else mp_grp.copy()
        s4_agg = pd.DataFrame()
        if not src.empty:
            if sel_s4_country: src = src[src["Country"].isin(sel_s4_country)]
            if sel_s4_periods: src = src[src[period_col].isin(sel_s4_periods)]
            s4_agg = _s4_make_chart(src, period_col, ["Comm Ref Code", period_col])

        if not s4_agg.empty:
            ctry_label = ", ".join(sel_s4_country) if sel_s4_country else "All Countries"
            periods_selected = sorted(s4_agg[period_col].unique().tolist())
            if len(periods_selected) > 1: sel_snap = st.select_slider(f"Snapshot {period_col}", options=periods_selected, value=periods_selected[-1], key="pr_s4_snap")
            else: sel_snap = periods_selected[0]
            snap_df = s4_agg[s4_agg[period_col] == sel_snap].sort_values("Comm Ref Code")

            fig_snap_price = px.bar(snap_df, x="Comm Ref Code", y="Unit Price (EUR)", color="Country", barmode="group", title=f"Avg Unit Price by Model — {focus_group} | {sel_snap}", color_discrete_sequence=CHART_COLORS, text_auto=".2s")
            fig_snap_price.update_layout(height=440, font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.35), bargap=0.15)
            st.plotly_chart(fig_snap_price, use_container_width=True)

            fig_snap_qty = px.bar(snap_df, x="Comm Ref Code", y="Total Qty", color="Country", barmode="group", title=f"Sales Volume by Model — {focus_group} | {sel_snap}", color_discrete_sequence=CHART_COLORS)
            fig_snap_qty.update_layout(height=400, font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.35))
            st.plotly_chart(fig_snap_qty, use_container_width=True)

            if len(periods_selected) > 1:
                st.markdown("**📈 Price Trend Over Time — per Model**")
                all_models = sorted(s4_agg["Comm Ref Code"].unique().tolist())
                sel_trend_models = st.multiselect("Select models for trend chart", all_models, key="pr_s4_trend_models")
                trend_df = s4_agg.copy()
                if sel_trend_models: trend_df = trend_df[trend_df["Comm Ref Code"].isin(sel_trend_models)]
                models_to_plot = sorted(trend_df["Comm Ref Code"].unique().tolist())
                if len(models_to_plot) <= 4:
                    for model in models_to_plot:
                        mdf = trend_df[trend_df["Comm Ref Code"] == model].sort_values([period_col, "Country"])
                        c_l, c_r = st.columns(2)
                        with c_l:
                            fig_tr = px.line(mdf, x=period_col, y="Unit Price (EUR)", color="Country", markers=True, title=f"{model} — Price Trend", color_discrete_sequence=CHART_COLORS)
                            fig_tr.update_layout(height=320, font=dict(family="IBM Plex Sans"), title_font=dict(size=12, color="#1F3864"), legend=dict(orientation="h", y=-0.4)); st.plotly_chart(fig_tr, use_container_width=True)
                        with c_r:
                            fig_qr = px.bar(mdf, x=period_col, y="Total Qty", color="Country", barmode="group", title=f"{model} — Volume Trend", color_discrete_sequence=CHART_COLORS)
                            fig_qr.update_layout(height=320, font=dict(family="IBM Plex Sans"), title_font=dict(size=12, color="#1F3864"), legend=dict(orientation="h", y=-0.4)); st.plotly_chart(fig_qr, use_container_width=True)
                else:
                    trend_df["Label"] = trend_df["Comm Ref Code"] + " | " + trend_df["Country"]
                    fig_all = px.line(trend_df.sort_values([period_col, "Label"]), x=period_col, y="Unit Price (EUR)", color="Label", markers=True, title="Price Trend — all selected models", color_discrete_sequence=CHART_COLORS)
                    fig_all.update_layout(height=420, font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.4)); st.plotly_chart(fig_all, use_container_width=True)

            yoy_rows = []
            for (model, cty), grp in s4_agg.groupby(["Comm Ref Code","Country"]):
                grp = grp.sort_values(period_col).copy()
                grp["Chg %"] = grp["Unit Price (EUR)"].pct_change() * 100
                yoy_rows.append(grp)
            df_chg = pd.concat(yoy_rows).dropna(subset=["Chg %"]) if yoy_rows else pd.DataFrame()
            chg_label = "YoY" if s4_period == "Year" else "MoM"

            if not df_chg.empty:
                df_chg_snap = df_chg[df_chg[period_col] == sel_snap] if len(periods_selected) > 1 else df_chg
                if not df_chg_snap.empty:
                    fig_chg = px.bar(df_chg_snap.sort_values("Comm Ref Code"), x="Comm Ref Code", y="Chg %", color="Country", barmode="group", title=f"{chg_label} Price Change % by Model", color_discrete_sequence=CHART_COLORS)
                    fig_chg.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
                    fig_chg.update_layout(height=340, font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.35)); st.plotly_chart(fig_chg, use_container_width=True)

    if has_cr_filter:
        cr_to_detail = [c for c in pc["cr"] if c in cr_in_group] or pc["cr"]
        section("🏷 Individual Comm Ref Detail")
        sel_detail_cr = st.selectbox("Select Comm Ref", cr_to_detail, key="pr_detail_cr")
        ap_cr = annual_price_df[annual_price_df["Commercial Reference"] == sel_detail_cr] if not annual_price_df.empty else pd.DataFrame()
        mp_cr = monthly_price_df[monthly_price_df["Commercial Reference"] == sel_detail_cr] if not monthly_price_df.empty else pd.DataFrame()

        if not ap_cr.empty:
            show2(price_trend_chart(ap_cr.sort_values(["Country","Year"]), "Year", "Country", f"{sel_detail_cr} — Annual Unit Price"), qty_trend_chart(ap_cr.sort_values(["Country","Year"]), "Year", "Country", f"{sel_detail_cr} — Annual Volume"))
            yoy_cr = _yoy_pct(ap_cr, "Year")
            if not yoy_cr.empty: st.plotly_chart(pct_chart(yoy_cr, "Year", "YoY %", "Country", f"{sel_detail_cr} — YoY Price Change %"), use_container_width=True)
            sel_ov_cty = st.selectbox("Country for Price vs Volume overlay", ap_cr["Country"].unique().tolist(), key="pr_ov_cty")
            d_ov = ap_cr[ap_cr["Country"] == sel_ov_cty].sort_values("Year")
            if not d_ov.empty: st.plotly_chart(dual_axis_chart(d_ov, "Year", f"{sel_detail_cr} — {sel_ov_cty}: Price vs Volume"), use_container_width=True)

        if not mp_cr.empty:
            st.markdown("##### Monthly Detail")
            show2(price_trend_chart(mp_cr.sort_values(["Country","YM"]), "YM", "Country", f"{sel_detail_cr} — Monthly Unit Price"), qty_trend_chart(mp_cr.sort_values(["Country","YM"]), "YM", "Country", f"{sel_detail_cr} — Monthly Volume"))
            mom_cr = _mom_pct(mp_cr)
            if not mom_cr.empty: st.plotly_chart(pct_chart(mom_cr, "YM", "MoM %", "Country", f"{sel_detail_cr} — MoM Price Change %"), use_container_width=True)
            sel_m_cty = st.selectbox("Country for monthly overlay", mp_cr["Country"].unique().tolist(), key="pr_m_ov_cty")
            d_mvo = mp_cr[mp_cr["Country"] == sel_m_cty].sort_values("YM")
            if not d_mvo.empty: st.plotly_chart(dual_axis_chart(d_mvo, "YM", f"{sel_detail_cr} — {sel_m_cty}: Price vs Volume"), use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
defaults = {
    "order_df":    None,
    "sales_df":    None,
    "committed":   None,
    "tables":      None,
    "g_total":     0.0,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("# 📊 SE Intel")
    uname = st.session_state.get("username", "")
    st.caption(f"👤 Logged in as **{uname}**")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()
    st.markdown("---")
    st.markdown("### 📁 Data Sources")

    order_file = st.file_uploader("Order File",  type=["xlsx","xls","xlsm","csv"], key="up_order")
    if order_file:
        fid = f"{order_file.name}_{order_file.size}"
        if st.session_state.get("order_file_id") != fid:
            df_loaded = load_file_with_progress(order_file)
            st.session_state.order_df      = df_loaded
            st.session_state.order_file_id = fid
        st.success(f"✓ Order: {len(st.session_state.order_df):,} rows")

    sales_file = st.file_uploader("Sales File",  type=["xlsx","xls","xlsm","csv"], key="up_sales")
    if sales_file:
        fid = f"{sales_file.name}_{sales_file.size}"
        if st.session_state.get("sales_file_id") != fid:
            df_loaded = load_file_with_progress(sales_file)
            st.session_state.sales_df      = df_loaded
            st.session_state.sales_file_id = fid
        st.success(f"✓ Sales: {len(st.session_state.sales_df):,} rows")

    st.markdown("### ⚡ Performance Source")
    avail = []
    if st.session_state.order_df is not None: avail.append("Order")
    if st.session_state.sales_df is not None: avail.append("Sales")
    if not avail: avail = ["Order", "Sales"]
    perf_source = st.selectbox("Source", avail, label_visibility="collapsed")

    active_df: pd.DataFrame | None = (
        st.session_state.sales_df if perf_source == "Sales"
        else st.session_state.order_df
    )

    # ── Pre-compute unique values once per column (no DataFrame copies) ────────
    # This avoids creating multiple copies of the 700MB+ DataFrame during cascade
    @st.cache_data(show_spinner=False)
    def _precompute_uniques(file_id: str, source: str) -> dict:
        """Cache unique values per column keyed by file identity."""
        df = st.session_state.get(
            "order_df" if source == "Order" else "sales_df")
        if df is None: return {}
        result = {}
        for col in ["Operations","Zone","Cluster","Area","Country",
                    "PLC","Family","Strategic Product Family",
                    "Comm Ref Code","Commercial Reference","Year"]:
            if col in df.columns:
                vals = df[col].astype(str).replace("nan","").replace("","<NA>")
                result[col] = sorted([v for v in vals.unique() if v not in ("","<NA>","nan")])
        return result

    file_id   = st.session_state.get("order_file_id" if perf_source == "Order" else "sales_file_id", "")
    _uniq     = _precompute_uniques(file_id, perf_source)

    def _cascade_opts(col, filters: dict) -> list:
        """Return filtered unique values using pre-computed sets — NO DataFrame ops."""
        if not _uniq: return []
        if not any(filters.values()):
            return _uniq.get(col, [])
        # Build allowed set by intersecting active filters across columns
        # We need to know which rows pass all OTHER filters
        # Since we can't do row-level ops without the df, fall back to
        # showing all options for that column (safe, just slightly less precise cascade)
        return _uniq.get(col, [])

    def opts(col): return _uniq.get(col, [])

    st.markdown("---")
    with st.form(key="filter_form", border=False):
        st.markdown("### 🕐 Time Filters")
        f_year  = st.multiselect("Year",    opts("Year"),     placeholder="Select...")
        f_month = st.multiselect("Month",   [str(m) for m in range(1,13)], placeholder="Select...")
        st.markdown("---")
        update_btn = st.form_submit_button("▶  UPDATE ALL TABLES", use_container_width=True, type="primary", disabled=(active_df is None))

    # Geography & Product filters — show all unique values, no live cascade filtering
    # (cascade narrowing happens only at UPDATE time, not on every widget interaction)
    st.markdown("### 🌍 Geography Filters")
    st.caption("Select values then press ▶ UPDATE ALL TABLES to apply.")
    f_ops     = st.multiselect("Operations", opts("Operations"), placeholder="Select...", key="sb_ops")
    f_zone    = st.multiselect("Zone",       opts("Zone"),       placeholder="Select...", key="sb_zone")
    f_cluster = st.multiselect("Cluster",    opts("Cluster"),    placeholder="Select...", key="sb_cluster")
    f_area    = st.multiselect("Area",       opts("Area"),       placeholder="Select...", key="sb_area")
    f_country = st.multiselect("Country",    opts("Country"),    placeholder="Select...", key="sb_country")

    st.markdown("### 📦 Product Filters")
    st.caption("Select values then press ▶ UPDATE ALL TABLES to apply.")
    f_plc = st.multiselect("Product Line Code",        opts("PLC"),                        placeholder="Select...", key="sb_plc")
    f_fam = st.multiselect("Family",                   opts("Family"),                     placeholder="Select...", key="sb_fam")
    f_spf = st.multiselect("Strategic Product Family", opts("Strategic Product Family"),   placeholder="Select...", key="sb_spf")
    f_crc = st.multiselect("Comm Ref Code",            opts("Comm Ref Code"),              placeholder="Select...", key="sb_crc")
    f_cr  = st.multiselect("Comm Ref",                 opts("Commercial Reference"),       placeholder="Select...", key="sb_cr")

    st.markdown("---")
    rst_btn = st.button("↺  RESET ALL FILTERS", use_container_width=True, disabled=(active_df is None))
    if rst_btn:
        for k in ["sb_ops","sb_zone","sb_cluster","sb_area","sb_country","sb_plc","sb_fam","sb_spf","sb_crc","sb_cr"]:
            st.session_state.pop(k, None)
        st.session_state.committed = None
        st.session_state.tables = None
        st.session_state.pop("price_committed", None)
        st.rerun()

    if update_btn:
        st.session_state.committed = {
            "perf_source": perf_source,
            "fmap": {
                "Year":                     f_year,
                "Month":                    f_month,
                "Operations":               f_ops,
                "Area":                     f_area,
                "Zone":                     f_zone,
                "Cluster":                  f_cluster,
                "Country":                  f_country,
                "PLC":                      f_plc,
                "Strategic Product Family": f_spf,
                "Family":                   f_fam,
                "Comm Ref Code":            f_crc,
                "Commercial Reference":     f_cr,
            }
        }

# ─────────────────────────────────────────────────────────────────────────────
# MAIN AREA
# ─────────────────────────────────────────────────────────────────────────────
col_t, col_b = st.columns([6, 1])
with col_t:
    st.markdown("# EXECUTIVE PERFORMANCE CONTROL HUB")
with col_b:
    badge = "source-order" if perf_source == "Order" else "source-sales"
    st.markdown(f'<br><span class="source-badge {badge}">{perf_source.upper()}</span>', unsafe_allow_html=True)

if active_df is None:
    st.info("👈  Upload an Order and/or Sales file from the sidebar to get started.")
    st.stop()

committed = st.session_state.committed
if committed is None:
    st.info("👈  Set your filters in the sidebar, then press **▶ UPDATE ALL TABLES** to run.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# COMPUTE
# ─────────────────────────────────────────────────────────────────────────────
fmap = committed["fmap"]
src  = committed["perf_source"]
active_df = st.session_state.sales_df if src == "Sales" else st.session_state.order_df

if active_df is None:
    st.warning(f"Data source '{src}' not loaded.")
    st.stop()

df_curr = apply_filters(active_df, fmap)
geo_all = all(len(fmap.get(c, [])) == 0 for c in ["Operations", "Area", "Zone", "Cluster", "Country"])

# ─────────────────────────────────────────────────────────────────────────────
# YoY / QoQ / MoM LOGIC (Completely Rewritten for robust accuracy)
# ─────────────────────────────────────────────────────────────────────────────
flag_yoy = flag_qoq = flag_mom = False
df_prev_y  = pd.DataFrame()
df_curr_q  = pd.DataFrame()
df_prev_q  = pd.DataFrame()
df_curr_m  = pd.DataFrame()
df_prev_m  = pd.DataFrame()

yr_sel = fmap.get("Year", [])
mo_sel = fmap.get("Month", [])
fmap_no_time = {k: v for k, v in fmap.items() if k not in ("Year", "Month")}

if len(yr_sel) == 1:
    yr = int(yr_sel[0])
    prev_yr = yr - 1
    
    # ── YoY (Year over Year): 嚴格對齊去年的相同篩選條件 ──
    flag_yoy = True
    fmap_py = dict(fmap)
    fmap_py["Year"] = [str(prev_yr)]
    df_prev_y = apply_filters(active_df, fmap_py)
    
    # ── 如果有選擇月份，自動推算精準的 MoM 與 QoQ ──
    if mo_sel:
        max_m = max(int(m) for m in mo_sel)
        
        # MoM (Month over Month): 最大選擇月份 vs 前一個月
        flag_mom = True
        pm_yr = yr if max_m > 1 else prev_yr
        pm_mo = max_m - 1 if max_m > 1 else 12
        fmap_cm = dict(fmap); fmap_cm["Year"] = [str(yr)]; fmap_cm["Month"] = [str(max_m)]
        df_curr_m = apply_filters(active_df, fmap_cm)
        fmap_pm = dict(fmap); fmap_pm["Year"] = [str(pm_yr)]; fmap_pm["Month"] = [str(pm_mo)]
        df_prev_m = apply_filters(active_df, fmap_pm)
        
        # QoQ (Quarter over Quarter): 當前季 vs 前一季
        q = (max_m - 1) // 3 + 1
        cq_mos = [str(m) for m in range((q-1)*3 + 1, q*3 + 1)]
        if q == 1:
            pq_mos = ["10", "11", "12"]
            pq_yr = prev_yr
        else:
            pq_mos = [str(m) for m in range((q-2)*3 + 1, (q-1)*3 + 1)]
            pq_yr = yr
        flag_qoq = True
        fmap_cq = dict(fmap); fmap_cq["Year"] = [str(yr)]; fmap_cq["Month"] = cq_mos
        df_curr_q = apply_filters(active_df, fmap_cq)
        fmap_pq = dict(fmap); fmap_pq["Year"] = [str(pq_yr)]; fmap_pq["Month"] = pq_mos
        df_prev_q = apply_filters(active_df, fmap_pq)

elif not yr_sel and not mo_sel:
    # ── 自動偵測：若完全沒選時間，抓取最新資料月 ──
    all_ym = active_df["Year-Month"].replace("", pd.NA).dropna()
    all_ym = all_ym[all_ym.str.len() == 6]
    if not all_ym.empty:
        latest_ym = all_ym.max()
        cur_yr = int(latest_ym[:4])
        cur_mo = int(latest_ym[4:6])
        prev_yr = cur_yr - 1
        
        yoy_range = [str(m) for m in range(1, cur_mo + 1)]
        df_curr = apply_filters(active_df, {**fmap_no_time, "Year": [str(cur_yr)], "Month": yoy_range})
        flag_yoy = True
        df_prev_y = apply_filters(active_df, {**fmap_no_time, "Year": [str(prev_yr)], "Month": yoy_range})
        
        flag_mom = True
        pm_yr = cur_yr if cur_mo > 1 else prev_yr
        pm_mo = cur_mo - 1 if cur_mo > 1 else 12
        df_curr_m = apply_filters(active_df, {**fmap_no_time, "Year": [str(cur_yr)], "Month": [str(cur_mo)]})
        df_prev_m = apply_filters(active_df, {**fmap_no_time, "Year": [str(pm_yr)], "Month": [str(pm_mo)]})
        
        q = (cur_mo - 1) // 3 + 1
        cq_mos = [str(m) for m in range((q-1)*3 + 1, q*3 + 1)]
        if q == 1:
            pq_mos = ["10", "11", "12"]
            pq_yr = prev_yr
        else:
            pq_mos = [str(m) for m in range((q-2)*3 + 1, (q-1)*3 + 1)]
            pq_yr = cur_yr
        flag_qoq = True
        df_curr_q = apply_filters(active_df, {**fmap_no_time, "Year": [str(cur_yr)], "Month": cq_mos})
        df_prev_q = apply_filters(active_df, {**fmap_no_time, "Year": [str(pq_yr)], "Month": pq_mos})

with st.spinner("Calculating tables…"):
    tables, g_total = compute_tables(
        df_curr, df_prev_y, df_prev_q, df_prev_m,
        df_curr_q, df_curr_m,
        geo_all, flag_yoy, flag_qoq, flag_mom,
    )

if tables is None:
    st.warning("⚠️ No data matches the current filters.")
    st.stop()

# ── Summary metrics — Row 1 (基礎運作規模，無 Delta) ─────────────────────────
mc1 = st.columns(6)
mc1[0].metric("Total Value (EUR)", fmt_val(g_total))
mc1[1].metric("# Transactions",         f"{len(df_curr):,}")
mc1[2].metric("# Countries",            df_curr["Country"].nunique())
mc1[3].metric("# Product Line Codes",   df_curr["PLC"].nunique())
mc1[4].metric("# Comm Refs",            df_curr["Commercial Reference"].nunique())
try:
    import psutil, os
    mem_mb = psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024
    mc1[5].metric("🧠 Memory",
                  f"{mem_mb:.0f} MB",
                  delta=f"{mem_mb - 1024:.0f} MB vs 1GB limit" if mem_mb > 700 else "✅ OK",
                  delta_color="inverse" if mem_mb > 700 else "normal")
except Exception:
    mc1[5].metric("🧠 Memory", "N/A")

# ── Summary metrics — Row 2 (💡 需求 3：全新獨立增長率大字卡矩陣列) ───────────
st.markdown("<div style='margin-top: -10px;'></div>", unsafe_allow_html=True)
mc2 = st.columns(3)

prev_y_total = df_prev_y["Value (EUR)"].sum() if not df_prev_y.empty else 0.0
prev_q_total = df_prev_q["Value (EUR)"].sum() if not df_prev_q.empty else 0.0
curr_q_total = df_curr_q["Value (EUR)"].sum() if not df_curr_q.empty else 0.0
prev_m_total = df_prev_m["Value (EUR)"].sum() if not df_prev_m.empty else 0.0
curr_m_total = df_curr_m["Value (EUR)"].sum() if not df_curr_m.empty else 0.0

# YoY Card
if flag_yoy and prev_y_total != 0:
    y_pct = growth(g_total, prev_y_total)
    y_diff = g_total - prev_y_total
    mc2[0].metric("# YoY Growth Rate", fmt_pct(y_pct), delta=f"{y_diff:+,.0f} EUR")
else:
    mc2[0].metric("# YoY Growth Rate", "—", delta="No Base Year")

# QoQ Card
if flag_qoq and prev_q_total != 0:
    q_pct = growth(curr_q_total, prev_q_total)
    q_diff = curr_q_total - prev_q_total
    mc2[1].metric("# QoQ Growth Rate", fmt_pct(q_pct), delta=f"{q_diff:+,.0f} EUR")
else:
    mc2[1].metric("# QoQ Growth Rate", "—", delta="No Base Quarter")

# MoM Card
if flag_mom and prev_m_total != 0:
    m_pct = growth(curr_m_total, prev_m_total)
    m_diff = curr_m_total - prev_m_total
    mc2[2].metric("# MoM Growth Rate", fmt_pct(m_pct), delta=f"{m_diff:+,.0f} EUR")
else:
    mc2[2].metric("# MoM Growth Rate", "—", delta="No Base Month")

st.markdown("---")

def _get_avail_years(active_df):
    all_ym = active_df["Year-Month"].replace("", pd.NA).dropna()
    all_ym = all_ym[all_ym.str.len() == 6]
    if all_ym.empty: return []
    return sorted(active_df["Year"].replace("", pd.NA).dropna().unique().tolist())

def render_yoy_monthly_chart(fmap_no_time, active_df, yr_sel, src_label, chart_key="main"):
    avail_years = _get_avail_years(active_df)
    if not avail_years: return
    default_cur = yr_sel[0] if yr_sel else avail_years[-1]
    if default_cur not in avail_years: default_cur = avail_years[-1]

    cc1, cc2 = st.columns([2, 4])
    with cc1:
        cur_yr_str = st.selectbox("Compare Year (current)", avail_years, index=avail_years.index(default_cur), key=f"mth_chart_cur_{chart_key}")
    with cc2:
        prev_options = [y for y in avail_years if y < cur_yr_str]
        if not prev_options: st.caption("No prior year data available for comparison."); return
        prev_yr_str = st.selectbox("vs Prior Year", prev_options, index=len(prev_options)-1, key=f"mth_chart_prev_{chart_key}")

    df_cy = apply_filters(active_df, {**fmap_no_time, "Year": [cur_yr_str]})
    df_py = apply_filters(active_df, {**fmap_no_time, "Year": [prev_yr_str]})
    if df_cy.empty and df_py.empty: return

    month_names = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun", 7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    def _agg(df, yr):
        if df.empty: return pd.DataFrame()
        agg = df[df["Month"].replace("", pd.NA).notna()].copy()
        if agg.empty: return pd.DataFrame()
        g = agg.groupby("Month", as_index=False).agg(Value=("Value (EUR)", "sum"), Quantity=("Quantity", "sum"))
        g["Month_num"] = g["Month"].astype(int); g["Month_label"] = g["Month_num"].map(month_names); g["Year"] = yr
        return g.sort_values("Month_num")

    cy_agg = _agg(df_cy, cur_yr_str)
    py_agg = _agg(df_py, prev_yr_str)
    combined = pd.concat([py_agg, cy_agg], ignore_index=True)
    if combined.empty: return

    cat_order = {"Month_label": [month_names[m] for m in range(1,13)]}
    color_map = {prev_yr_str: "#BDD7EE", cur_yr_str: "#1F3864"}

    col_v, col_q = st.columns(2)
    with col_v:
        fig_v = px.bar(combined.sort_values(["Month_num","Year"]), x="Month_label", y="Value", color="Year", barmode="group", title=f"Monthly Value — {cur_yr_str} vs {prev_yr_str} {src_label}", color_discrete_map=color_map, text_auto=".2s", category_orders=cat_order)
        fig_v.update_layout(height=380, font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.28), bargap=0.2)
        fig_v.update_yaxes(title="Value (EUR)"); fig_v.update_xaxes(title="Month")
        st.plotly_chart(fig_v, use_container_width=True)
    with col_q:
        fig_q = px.bar(combined.sort_values(["Month_num","Year"]), x="Month_label", y="Quantity", color="Year", barmode="group", title=f"Monthly Quantity — {cur_yr_str} vs {prev_yr_str} {src_label}", color_discrete_map=color_map, text_auto=".2s", category_orders=cat_order)
        fig_q.update_layout(height=380, font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.28), bargap=0.2)
        fig_q.update_yaxes(title="Quantity"); fig_q.update_xaxes(title="Month")
        st.plotly_chart(fig_q, use_container_width=True)

render_yoy_monthly_chart(fmap_no_time, active_df, yr_sel, f"({src})")
st.markdown("---")

def render(df, title, rename=None, drop=None):
    st.markdown(f'<h4>{title}</h4>', unsafe_allow_html=True)
    if df is None or df.empty: st.caption("— No data —"); return
    d = df.copy()
    if drop: d = d.drop(columns=[c for c in drop if c in d.columns])
    if rename: d = d.rename(columns=rename)
    if "Value (EUR)" in d.columns: d["Value (EUR)"] = d["Value (EUR)"].apply(fmt_val)
    if "QTY" in d.columns: d["QTY"] = d["QTY"].apply(lambda x: f"{x:,.0f}")
    st.dataframe(d, use_container_width=True, hide_index=True, height=min(38 * len(d) + 40, 520))

label = f"({src})"

# ─────────────────────────────────────────────────────────────────────────────
# TABS OVERVIEW & DELEGATION
# ─────────────────────────────────────────────────────────────────────────────
tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🎯 TARGET & FORECAST", "📍 ZONE", "🌍 COUNTRIES", "📦 PRODUCT LINE CODE",
    "🏢 FAMILIES", "🔖 COMM REF CODES", "🏷️ COMM REFS", "💶 PRICE ANALYSIS",
])

def render_yoy_breakdown(df_curr: pd.DataFrame, active_df: pd.DataFrame, fmap_no_time: dict, dim_col: str, tab_key: str, yr_sel: list, label: str):
    st.markdown("---"); st.markdown("#### 📅 YoY Monthly Breakdown")
    avail_years = _get_avail_years(active_df)
    if not avail_years: st.caption("No Year-Month data available."); return

    default_cur = yr_sel[0] if yr_sel else avail_years[-1]
    if default_cur not in avail_years: default_cur = avail_years[-1]
    yc1, yc2, yc3 = st.columns([2, 2, 3])
    with yc1: cur_yr_str = st.selectbox("Compare Year (current)", avail_years, index=avail_years.index(default_cur), key=f"bd_cur_{tab_key}")
    with yc2:
        prev_options = [y for y in avail_years if y < cur_yr_str]
        if not prev_options: st.caption("No prior year available."); return
        prev_yr_str = st.selectbox("vs Prior Year", prev_options, index=len(prev_options)-1, key=f"bd_prev_{tab_key}")
    df_cy = apply_filters(active_df, {**fmap_no_time, "Year": [cur_yr_str]})
    df_py = apply_filters(active_df, {**fmap_no_time, "Year": [prev_yr_str]})
    if df_cy.empty and df_py.empty: st.caption("No data for comparison."); return

    if not df_cy.empty and dim_col in df_cy.columns:
        all_items = (df_cy.groupby(dim_col)["Value (EUR)"].sum().sort_values(ascending=False).index.tolist())
    else: all_items = []
    if not all_items: st.caption(f"No data for '{dim_col}'."); return
    default_items = all_items[:min(10, len(all_items))]
    commit_key = f"yoy_bd_committed_{tab_key}"

    fc1, fc2 = st.columns([4, 1])
    with fc1: sel_items = st.multiselect(f"Select {dim_col}s to compare", options=all_items, default=default_items, key=f"yoy_bd_sel_{tab_key}", placeholder=f"Choose {dim_col}s…")
    with fc2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("▶ Apply", key=f"yoy_bd_apply_{tab_key}", use_container_width=True, type="primary"):
            st.session_state[commit_key] = sel_items if sel_items else default_items
            st.rerun()
    top_items = st.session_state.get(commit_key, default_items) or default_items

    month_names = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun", 7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    def _agg(df, yr_str):
        if df.empty or dim_col not in df.columns: return pd.DataFrame()
        d = df[df[dim_col].isin(top_items)].copy()
        d = d[d["Month"].replace("", pd.NA).notna()]
        if d.empty: return pd.DataFrame()
        g = d.groupby([dim_col, "Month"], as_index=False).agg(Value=("Value (EUR)", "sum"), Quantity=("Quantity", "sum"))
        g["Month_num"] = g["Month"].astype(int); g["Month_label"] = g["Month_num"].map(month_names); g["Year"] = yr_str; g["Series"] = g[dim_col] + f" ({yr_str})"
        return g

    cy_agg, py_agg = _agg(df_cy, cur_yr_str), _agg(df_py, prev_yr_str)
    combined = pd.concat([py_agg, cy_agg], ignore_index=True)
    if combined.empty: st.caption("No monthly data available."); return

    n = len(top_items)
    base_colors = CHART_COLORS * ((n // len(CHART_COLORS)) + 1)
    color_map = {}
    for i, item in enumerate(top_items):
        color_map[f"{item} ({cur_yr_str})"] = base_colors[i]
        color_map[f"{item} ({prev_yr_str})"] = base_colors[i]

    cat_order = {"Month_label": [month_names[m] for m in range(1, 13)], "Series": [f"{it} ({y})" for it in top_items for y in [prev_yr_str, cur_yr_str]]}
    chart_df = combined.sort_values(["Month_num", dim_col, "Year"])

    def _make_bar(y_col, y_title, title_prefix):
        fig = px.bar(chart_df, x="Month_label", y=y_col, color="Series", barmode="group", title=f"{title_prefix} by {dim_col} — {cur_yr_str} vs {prev_yr_str} {label}", color_discrete_map=color_map, category_orders=cat_order, text_auto=".2s")
        for trace in fig.data: trace.marker.opacity = 0.95 if f"({prev_yr_str})" not in trace.name else 0.4
        fig.update_xaxes(title="Month"); fig.update_yaxes(title=y_title)
        fig.update_layout(height=max(420, 50*n), font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"), legend=dict(orientation="h", y=-0.35), bargap=0.15)
        return fig
    st.plotly_chart(_make_bar("Value", "Value (EUR)", "Monthly Value"), use_container_width=True)
    st.plotly_chart(_make_bar("Quantity", "Quantity", "Monthly Quantity"), use_container_width=True)


def render_target_tab(active_df: pd.DataFrame, fmap_no_time: dict):
    st.markdown("#### 🎯 Performance Target & Forecast")
    st.caption("透過 PLC 設定成長目標、追蹤 YTD 達成狀況，並推算後續每個月的所需目標量。")
    avail_years = _get_avail_years(active_df)
    if not avail_years: st.warning("未載入任何有效數據。"); return
    month_names = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun", 7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}

    st.markdown("---"); st.markdown("### ⚙️ Section 1 — Target Setup")
    
    # 建立設定表單與右側進度小工具的左右版面
    setup_col, vsd_col = st.columns([5, 2])
    
    with setup_col:
        sc1, sc2 = st.columns(2)
        with sc1: target_year = st.selectbox("目標規劃年份 (Target Year)", avail_years, index=len(avail_years)-1, key="tgt_yr")
        with sc2:
            base_options = {"僅參考去年業績 (Prior Year only)": "py", "僅參考前年業績 (Year before prior)": "ppy", "前兩年平均業績 (2-Year Average)": "avg2", "權重自訂平均 (Weighted Average)": "weighted"}
            base_choice = st.selectbox("計算基礎 (Target Base)", list(base_options.keys()), key="tgt_base")
            base_mode = base_options[base_choice]
            
        py_yr = str(int(target_year) - 1); ppy_yr = str(int(target_year) - 2)
        if base_mode == "weighted":
            wt_c1, wt_c2 = st.columns(2)
            with wt_c1: w_py = st.slider(f"{py_yr} 業績權重 %", 0, 100, 60, key="tgt_w_py") / 100
            with wt_c2: w_ppy = st.slider(f"{ppy_yr} 業績權重 %", 0, 100, 40, key="tgt_w_ppy") / 100
        else: w_py, w_ppy = 1.0, 0.0

        st.markdown("##### 📦 批次套用整體成長率")
        st.caption("設定基礎成長率後按「批次套用」，下方各 PLC 微調值將歸零（代表與整體成長率相同）。")
        c_all1, c_all2 = st.columns([3, 1])
        with c_all1:
            overall_growth = st.number_input(
                "整體成長率 % (所有 PLC 的基準)", min_value=-100, max_value=500,
                value=int(st.session_state.get("_overall_growth", 10)),
                step=1, key="tgt_growth"
            )
        with c_all2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("▶ 批次套用", use_container_width=True, type="primary"):
                st.session_state["_overall_growth"] = overall_growth
                st.session_state["_apply_all"] = overall_growth
                st.rerun()

    df_py_all  = apply_filters(active_df, {**fmap_no_time, "Year": [py_yr]})
    df_ppy_all = apply_filters(active_df, {**fmap_no_time, "Year": [ppy_yr]})
    df_ty_all  = apply_filters(active_df, {**fmap_no_time, "Year": [target_year]})

    def _plc_sum(df):
        if df.empty: return pd.Series(dtype=float)
        return df.groupby("PLC")["Value (EUR)"].sum()

    py_plc, ppy_plc, ty_plc = _plc_sum(df_py_all), _plc_sum(df_ppy_all), _plc_sum(df_ty_all)
    all_plcs = sorted(set(py_plc.index.tolist() + ppy_plc.index.tolist() + ty_plc.index.tolist()))
    all_plcs = [p for p in all_plcs if p]
    if not all_plcs: st.warning("找不到 PLC 相關資料。"); return

    plc_py_vals  = {plc: float(py_plc.get(plc, 0))  for plc in all_plcs}
    plc_ppy_vals = {plc: float(ppy_plc.get(plc, 0)) for plc in all_plcs}
    plc_ty_vals  = {plc: float(ty_plc.get(plc, 0))  for plc in all_plcs}

    if "_apply_all" in st.session_state:
        st.session_state.pop("_apply_all")
        if "tgt_committed" not in st.session_state: st.session_state["tgt_committed"] = {}
        for p in all_plcs: st.session_state["tgt_committed"][p] = 0
        st.rerun()

    if "tgt_committed" not in st.session_state: st.session_state["tgt_committed"] = {}
    base_growth = int(st.session_state.get("_overall_growth", overall_growth))

    def _base(plc):
        py_v, ppy_v = plc_py_vals.get(plc, 0), plc_ppy_vals.get(plc, 0)
        if base_mode == "py":       return py_v
        if base_mode == "ppy":      return ppy_v
        if base_mode == "avg2":     return (py_v + ppy_v) / 2
        if base_mode == "weighted": return py_v * w_py + ppy_v * w_ppy
        return py_v

    with setup_col:
        st.markdown("##### 📦 每個 PLC 獨立增長率微調")
        st.caption(
            f"整體成長率基準 = **{base_growth}%**。"
            "微調值為在整體成長率之上的加減（+5 代表整體+5%，-3 代表整體-3%）。"
        )
        cols = st.columns(3)
        plc_targets = {}

        for i, plc in enumerate(all_plcs):
            with cols[i % 3]:
                delta     = int(st.session_state["tgt_committed"].get(plc, 0))
                effective = base_growth + delta
                base_val  = _base(plc)
                tgt_val   = base_val * (1 + effective / 100)

                new_delta = st.number_input(
                    f"{plc}  微調 ± %",
                    min_value=-200, max_value=500,
                    value=delta, step=1,
                    key=f"tgt_plc_{plc}",
                    help=f"整體 {base_growth}% + 微調 {delta:+d}% = 實際成長率 {effective}%\n"
                         f"Base: {fmt_val(base_val)} → Target: {fmt_val(tgt_val)}"
                )
                st.session_state["tgt_committed"][plc] = new_delta
                eff = base_growth + new_delta
                plc_targets[plc] = _base(plc) * (1 + eff / 100)

    total_target = sum(plc_targets.values())
    total_actual = float(df_ty_all["Value (EUR)"].sum()) if not df_ty_all.empty else 0.0

    # PLC colour palette
    n_plc = len(all_plcs)
    plc_colors = (CHART_COLORS * ((n_plc // len(CHART_COLORS)) + 1))[:n_plc]
    plc_color_map = dict(zip(all_plcs, plc_colors))

    # ── 右側 VSD Chart：各 PLC 達標與剩餘 Gap ──────────────────────────────────
    with vsd_col:
        st.markdown(
            "<div style='text-align:center;font-weight:700;color:#1F3864;font-size:.85rem'>"
            "🎯 各產品線達標與差距 (Actual vs Target)</div>", unsafe_allow_html=True
        )
        fig_vsd = go.Figure()
        t_vals = [plc_targets.get(p, 0) for p in all_plcs]
        a_vals = [plc_ty_vals.get(p, 0) for p in all_plcs]

        # 淺色：目標 (較寬，放在底層)
        fig_vsd.add_trace(go.Bar(
            x=all_plcs, y=t_vals, name="目標 (Target)",
            marker_color="#BDD7EE", width=0.65
        ))
        # 深色：已實現 (較窄，疊加在前)
        fig_vsd.add_trace(go.Bar(
            x=all_plcs, y=a_vals, name="已實現 (Actual)",
            marker_color="#1F3864", width=0.4
        ))

        annotations = []
        for i, p in enumerate(all_plcs):
            tv = t_vals[i]; av = a_vals[i]
            diff = av - tv
            pct  = diff / tv if tv > 0 else 0
            if diff >= 0:
                text_str  = f"超前<br>+{pct:.1%}<br>({fmt_val(diff)})"
                font_color = "#375623"
            else:
                text_str  = f"未完成<br>{pct:.1%}<br>({fmt_val(abs(diff))})"
                font_color = "#C55A11"
            annotations.append(dict(
                x=p, y=max(tv, av), text=text_str,
                showarrow=False, yshift=25,
                font=dict(size=10, color=font_color)
            ))

        fig_vsd.update_layout(
            barmode="overlay", showlegend=True, height=300,
            legend=dict(orientation="h", y=-0.3),
            margin=dict(t=40, b=10, l=10, r=10),
            font=dict(family="IBM Plex Sans"),
            annotations=annotations
        )
        st.plotly_chart(fig_vsd, use_container_width=True)
    st.markdown("---"); st.markdown("### 📊 全局預測模型與業績進度圖")
    
    # ── 💡 需求 1：雙預測模型選項切換與月度 Faded 疊加圖 ──────────────────────
    proj_mode = st.radio(
        "🔮 選擇 Cumulative 預期的目標線分配模式 (Target Projection Mode):",
        ["模式 A：線性等均分疊加預測 (Linear Projection)", "模式 B：依過去兩年歷史走勢權重預測 (Historical Trend Projection)"],
        horizontal=True
    )

    # 建立月份目標的權重基礎
    months_keys = [str(m) for m in range(1, 13)]
    months_labels = [month_names[m] for m in range(1, 13)]
    
    if "Historical" in proj_mode:
        # 歷史趨勢權重：各年度先正規化為月份佔比，再依 base_mode 加權平均
        def _mo_weights(df):
            agg = df.groupby("Month")["Value (EUR)"].sum()
            total = sum(float(agg.get(m, 0)) for m in months_keys)
            if total <= 0:
                return {m: 1/12 for m in months_keys}
            return {m: float(agg.get(m, 0)) / total for m in months_keys}

        w_py_mo  = _mo_weights(df_py_all)
        w_ppy_mo = _mo_weights(df_ppy_all)

        if base_mode == "py":
            month_weights = w_py_mo
        elif base_mode == "ppy":
            month_weights = w_ppy_mo
        elif base_mode == "avg2":
            month_weights = {m: (w_py_mo[m] + w_ppy_mo[m]) / 2 for m in months_keys}
        else:  # weighted
            month_weights = {m: w_py_mo[m] * w_py + w_ppy_mo[m] * w_ppy for m in months_keys}
            # re-normalise in case weights don't sum to 1
            wt_sum = sum(month_weights.values())
            if wt_sum > 0:
                month_weights = {m: v / wt_sum for m, v in month_weights.items()}
    else:
        # 線性預測：12個月均分
        month_weights = {m: 1/12 for m in months_keys}

    # 各月原始目標（全年均分或歷史加權）
    monthly_targets = [total_target * month_weights[str(m)] for m in range(1, 13)]

    # 純目標累計線（不考慮實際，僅供參考）
    cum_targets = list(pd.Series(monthly_targets).cumsum())

    # 實際月度與累計
    df_ty_monthly = df_ty_all.groupby("Month")["Value (EUR)"].sum()
    max_actual_month = int(df_ty_all["Month"].replace("", "0").astype(int).max()) if not df_ty_all.empty else 0
    actual_vals = [float(df_ty_monthly.get(str(m), 0)) for m in range(1, 13)]
    cum_actuals = list(pd.Series(actual_vals[:max_actual_month]).cumsum())

    # ── 調整後預測曲線 ────────────────────────────────────────────────────────
    # 邏輯：已完成月份 → 使用實際累計
    #       未完成月份 → 剩餘目標（total_target - 已完成實際）÷ 剩餘月份 線性分配
    #       保留歷史趨勢模式時，未完成月份按各月原始權重比例分配剩餘目標
    actual_ytd   = cum_actuals[-1] if cum_actuals else 0.0
    remaining_tgt = total_target - actual_ytd
    remaining_mo  = 12 - max_actual_month

    if remaining_mo > 0:
        if "Historical" in proj_mode:
            # 剩餘月份按歷史比例重新分配
            future_weights_raw = {str(m): month_weights[str(m)] for m in range(max_actual_month+1, 13)}
            fw_sum = sum(future_weights_raw.values())
            if fw_sum > 0:
                adj_monthly = {m: remaining_tgt * v / fw_sum for m, v in future_weights_raw.items()}
            else:
                adj_monthly = {str(m): remaining_tgt / remaining_mo for m in range(max_actual_month+1, 13)}
        else:
            adj_monthly = {str(m): remaining_tgt / remaining_mo for m in range(max_actual_month+1, 13)}
    else:
        adj_monthly = {}

    # Build projected cumulative series
    projected_cum = []
    for m in range(1, 13):
        if m <= max_actual_month:
            projected_cum.append(cum_actuals[m-1])
        else:
            prev = projected_cum[-1] if projected_cum else actual_ytd
            projected_cum.append(prev + adj_monthly.get(str(m), 0))

    # Also build adjusted monthly bar values (for chart A)
    adj_monthly_bars = []
    for m in range(1, 13):
        if m <= max_actual_month:
            adj_monthly_bars.append(monthly_targets[m-1])  # original for reference
        else:
            adj_monthly_bars.append(adj_monthly.get(str(m), 0))

    # ── Chart A: 月度疊加圖 (目標 vs 實際) + 調整後月度目標 ────────────────────
    col_ch1, col_ch2 = st.columns(2)
    with col_ch1:
        st.markdown("#### 🌊 A. 月度疊加比較圖")
        st.caption("淺色 = 月度目標，深色疊加 = 實際達成；橙線 = 調整後月度目標（依剩餘業績重新分配）。")
        fig_monthly_overlay = go.Figure()
        fig_monthly_overlay.add_trace(go.Bar(
            x=months_labels, y=monthly_targets, name="原始月度目標",
            marker_color="#BDD7EE", opacity=0.5, offsetgroup=0
        ))
        fig_monthly_overlay.add_trace(go.Bar(
            x=months_labels[:max_actual_month],
            y=actual_vals[:max_actual_month],
            name="實際達成 Actual",
            marker_color="#1F3864", opacity=0.9, offsetgroup=0
        ))
        # Adjusted target line overlay
        fig_monthly_overlay.add_trace(go.Scatter(
            x=months_labels, y=adj_monthly_bars,
            name="調整後月度目標", mode="lines+markers",
            line=dict(color="#C55A11", width=2, dash="dot"),
            marker=dict(size=5),
        ))
        fig_monthly_overlay.update_layout(
            barmode="overlay", height=420, font=dict(family="IBM Plex Sans"),
            title="月度目標 vs 實際（含調整後目標線）",
            title_font=dict(size=12, color="#1F3864"),
            legend=dict(orientation="h", y=-0.25), yaxis_title="Value (EUR)"
        )
        st.plotly_chart(fig_monthly_overlay, use_container_width=True)

    # ── Chart B: 累計折線 + 長條圖 ───────────────────────────────────────────
    with col_ch2:
        st.markdown("#### 📈 B. 累計業績 — 實際 vs 預期成長曲線")
        st.caption(
            "長條 = 累計實際（已完成）/ 預測（未來）。"
            "橙虛線 = 純線性/歷史目標累計。藍實線 = 調整後預測曲線（考慮已完成業績重新分配剩餘目標）。"
        )
        fig_cum = go.Figure()

        # 已完成月份：深色長條
        if cum_actuals:
            fig_cum.add_trace(go.Bar(
                x=months_labels[:max_actual_month], y=cum_actuals,
                name="累計實際 Actual", marker_color="#1F3864", opacity=0.9
            ))

        # 未來月份：淺色長條（調整後累計）
        if max_actual_month < 12:
            fig_cum.add_trace(go.Bar(
                x=months_labels[max_actual_month:],
                y=projected_cum[max_actual_month:],
                name="預測累計 Forecast", marker_color="#BDD7EE", opacity=0.65
            ))

        # 純原始目標線（虛線參考）
        fig_cum.add_trace(go.Scatter(
            x=months_labels, y=cum_targets, name="原始目標累計線",
            mode="lines", line=dict(color="#C55A11", width=2, dash="dash")
        ))

        # 調整後預測曲線（從最後一個實際月份延伸）
        connect_idx = max(0, max_actual_month - 1)
        fig_cum.add_trace(go.Scatter(
            x=months_labels[connect_idx:], y=projected_cum[connect_idx:],
            name="調整後預測曲線", mode="lines+markers",
            line=dict(color="#2E75B6", width=3), marker=dict(size=7)
        ))

        # Mark boundary actual / forecast
        if 0 < max_actual_month < 12:
            fig_cum.add_vrect(
                x0=months_labels[max_actual_month-1],
                x1=months_labels[max_actual_month],
                fillcolor="gray", opacity=0.07, line_width=0,
                annotation_text="← 實際 | 預測 →",
                annotation_position="top left",
                annotation_font_size=10,
            )

        fig_cum.update_layout(
            height=420, font=dict(family="IBM Plex Sans"),
            title="累計業績 — 實際 vs 調整後預測曲線",
            title_font=dict(size=12, color="#1F3864"),
            legend=dict(orientation="h", y=-0.25),
            yaxis_title="Cumulative Value (EUR)", barmode="group"
        )
        st.plotly_chart(fig_cum, use_container_width=True)

    # 瀑布圖：展示歷年實際到今年新 Target 的組成
    st.markdown("#### 🌊 C. PLC Waterfall: 歷史實績 ➔ 今年規劃目標")
    hist_years = [ppy_yr, py_yr, target_year]
    
    fig_wf = go.Figure()
    for plc, color in zip(all_plcs, (CHART_COLORS * 3)):
        fig_wf.add_trace(go.Bar(name=plc, x=hist_years, y=[plc_ppy_vals.get(plc,0), plc_py_vals.get(plc,0), plc_ty_vals.get(plc,0)], marker_color=color, legendgroup=plc, showlegend=True))
        
    running = total_actual
    for plc in all_plcs:
        increment = plc_targets.get(plc, 0) - plc_ty_vals.get(plc, 0)
        if increment <= 0: continue
        fig_wf.add_trace(go.Bar(name=plc, x=[f"+{plc}"], y=[increment], base=[running], marker_color=color, opacity=0.6, legendgroup=plc, showlegend=False, text=f"+{fmt_val(increment)}", textposition="outside"))
        running += increment

    fig_wf.add_hline(y=total_target, line_dash="dash", line_color="#1F3864", annotation_text=f"Target {target_year}: {fmt_val(total_target)}", annotation_position="top right")
    fig_wf.update_layout(barmode="stack", title="PLC 結構轉變與目標堆疊瀑布圖", height=420, font=dict(family="IBM Plex Sans"))
    st.plotly_chart(fig_wf, use_container_width=True)


with tab0: render_target_tab(active_df, fmap_no_time)
with tab1: render(tables["zone"], f">> ZONE PERFORMANCE {label}", rename={"Name":"Zone"})
with tab2:
    t = tables["country"]
    if t is not None and not t.empty: t = t.rename(columns={"Parent":"Zone", "Country":"Cluster"})
    render(t, f">> TOP 20 COUNTRIES {label}", rename={"Name":"Country"})
with tab3:
    render(tables["plc"], f">> TOP 20 PLC {label}", rename={"Name":"Product Line Code"}, drop=["Parent"])
    render_tab_charts(df_curr, tables["plc"], "Name", "PLC", "plc", "PLCs")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "PLC", "plc", yr_sel, label)
with tab4:
    # 💡 需求 5：在 Tab:Family 的表格中加入 QTY 數量欄位
    render(tables["family"], f">> TOP 20 FAMILIES {label}", rename={"Name": "Family", "Parent": "PLC"})
    render_tab_charts(df_curr, tables["family"], "Name", "Family", "family", "Families")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Family", "family", yr_sel, label)
with tab5:
    # 💡 需求 5：在 Tab:Comm Ref Code 的表格中加入 QTY 數量欄位
    render(tables["crc"], f">> TOP 20 COMM REF CODES {label}", rename={"Name": "Comm Ref Code", "Parent": "Family"})
    render_tab_charts(df_curr, tables["crc"], "Name", "Comm Ref Code", "crc", "Comm Ref Codes")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Comm Ref Code", "crc", yr_sel, label)
with tab6:
    # 💡 需求 5：在 Tab:Comm Ref 的表格中加入 QTY 數量欄位
    t_cr = tables["cr"]
    if t_cr is not None and not t_cr.empty:
        if geo_all and "Country" in t_cr.columns: t_cr = t_cr.drop(columns=["Country"])
        t_cr = t_cr.rename(columns={"Name": "Comm Ref", "Parent": "Comm Ref Code"})
    render(t_cr, f">> TOP 20 COMM REFS {label}")
    render_tab_charts(df_curr, tables["cr"], "Name", "Commercial Reference", "cr", "Comm Refs")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Commercial Reference", "cr", yr_sel, label)
with tab7:
    render_price_tab(df_curr, fmap)

# ─────────────────────────────────────────────────────────────────────────────
# RAW DATA PREVIEW (跨 Tab 獨立置底)
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")

with st.expander("📋 Raw Data Preview (filtered)", expanded=False):
    st.dataframe(df_curr[TARGET_HEADERS].head(2000), use_container_width=True, hide_index=True, height=400)
    st.caption(f"Showing first 2,000 of {len(df_curr):,} filtered rows.")