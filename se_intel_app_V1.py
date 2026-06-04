"""
SE Intel Dashboard — Python / Streamlit (optimised)  v1.0
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
    """Verify the launch key stored in .streamlit/secrets.toml."""
    try:
        stored = st.secrets["auth"]["launch_key"]
        return bool(stored)
    except Exception:
        st.error("⛔ Launch key not found. Please create `.streamlit/secrets.toml` with the required credentials.")
        st.code("""# .streamlit/secrets.toml
[auth]
launch_key = "YOUR_LAUNCH_KEY_HERE"

[users]
alice = "hashed_password_here"
bob   = "hashed_password_here"
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
            users: dict = dict(st.secrets.get("users", {}))
            pw_hash = _hash(password)
            if username in users and users[username] == pw_hash:
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.rerun()
            else:
                st.error("❌ Invalid username or password.")

    with tab_admin:
        with st.form("admin_form"):
            admin_input = st.text_input("Admin Key", type="password",
                                        placeholder="Enter admin key")
            admin_btn = st.form_submit_button("Enter", use_container_width=True,
                                              type="primary")

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
                st.error("❌ Admin key not configured in secrets.toml.")

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

# Plotly color palette (IBM-inspired)
CHART_COLORS = [
    "#2E75B6","#1F3864","#375623","#C55A11","#7030A0",
    "#BDD7EE","#70AD47","#ED7D31","#A9D18E","#9DC3E6",
]

# ─────────────────────────────────────────────────────────────────────────────
# ENGINE 1 — LOAD & ALIGN
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_and_align(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    if file_name.lower().endswith(".csv"):
        src = pd.read_csv(buf, dtype=str)
    else:
        src = pd.read_excel(buf, sheet_name=0, dtype=str)

    src.columns = [c.strip().lower() for c in src.columns]

    out = {}
    for tgt in TARGET_HEADERS:
        if tgt in DERIVED:
            out[tgt] = ""
            continue
        src_col = SOURCE_MAP.get(tgt, tgt.lower())
        out[tgt] = src[src_col].values if src_col in src.columns else ""

    df = pd.DataFrame(out)

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
            mask &= df[col].isin(sel)
    return df.loc[mask]


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
    if df is None or df.empty:
        return b""
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    return buf.getvalue()

def bytes_to_df(b: bytes) -> pd.DataFrame:
    if not b:
        return pd.DataFrame()
    return pd.read_parquet(io.BytesIO(b))


@st.cache_data(show_spinner=False)
def compute_tables(
    curr_bytes: bytes, pvy_bytes: bytes,
    pvq_bytes: bytes,  pvm_bytes: bytes,
    cq_bytes: bytes,   cm_bytes: bytes,
    geo_all: bool, flag_yoy: bool, flag_qoq: bool, flag_mom: bool,
):
    """
    curr_bytes : current period (YoY window, 01..max_m)
    pvy_bytes  : YoY previous  (same window, prior year)
    pvq_bytes  : QoQ previous  (last completed quarter, prior year)
    pvm_bytes  : MoM previous  (same single month, prior year)
    cq_bytes   : QoQ current   (last completed quarter, current year)
    cm_bytes   : MoM current   (same single month, current year)
    """
    df_c   = bytes_to_df(curr_bytes)
    df_py  = bytes_to_df(pvy_bytes)
    df_pq  = bytes_to_df(pvq_bytes)
    df_pm  = bytes_to_df(pvm_bytes)
    df_cq  = bytes_to_df(cq_bytes)   # QoQ current window
    df_cm  = bytes_to_df(cm_bytes)   # MoM current window

    g_total = df_c["Value (EUR)"].sum() if not df_c.empty else 0.0
    if g_total == 0:
        return None, 0.0

    def safe_agg(df: pd.DataFrame, col: str) -> pd.Series:
        if df.empty or col not in df.columns:
            return pd.Series(dtype=float)
        return df.groupby(col)["Value (EUR)"].sum()

    def safe_composite(df: pd.DataFrame) -> pd.Series:
        if df.empty or "Commercial Reference" not in df.columns or "Country" not in df.columns:
            return pd.Series(dtype=float)
        tmp = df.copy()
        tmp["_ck"] = tmp["Commercial Reference"] + "|" + tmp["Country"]
        return tmp.groupby("_ck")["Value (EUR)"].sum()

    def build(group_col, parent_col=None, country_col=None, top_n=20, composite=False):
        if composite and not geo_all:
            curr_s  = safe_composite(df_c)
            py_s    = safe_composite(df_py)
            cq_s    = safe_composite(df_cq)
            pq_s    = safe_composite(df_pq)
            cm_s    = safe_composite(df_cm)
            pm_s    = safe_composite(df_pm)
            if not df_c.empty and "Commercial Reference" in df_c.columns:
                df_c2 = df_c.copy()
                df_c2["_ck"] = df_c2["Commercial Reference"] + "|" + df_c2["Country"]
                par_map = df_c2.groupby("_ck")["Comm Ref Code"].first() if parent_col else None
                cty_map = df_c2.groupby("_ck")["Country"].first()
            else:
                par_map = None; cty_map = None
        else:
            curr_s  = safe_agg(df_c,  group_col)
            py_s    = safe_agg(df_py, group_col)
            cq_s    = safe_agg(df_cq, group_col)
            pq_s    = safe_agg(df_pq, group_col)
            cm_s    = safe_agg(df_cm, group_col)
            pm_s    = safe_agg(df_pm, group_col)
            par_map = df_c.groupby(group_col)[parent_col].first()  if (parent_col  and not df_c.empty and parent_col  in df_c.columns) else None
            cty_map = df_c.groupby(group_col)[country_col].first() if (country_col and not df_c.empty and country_col in df_c.columns) else None

        if curr_s.empty:
            return pd.DataFrame()

        top = curr_s.nlargest(top_n)
        rows = []
        for rank, (key, val) in enumerate(top.items(), 1):
            disp = key.split("|")[0] if "|" in str(key) else key

            # QoQ: growth(cq_curr, pq_prev)
            cq_val = cq_s.get(key, 0)
            pq_val = pq_s.get(key, 0)
            # MoM: growth(cm_curr, pm_prev)
            cm_val = cm_s.get(key, 0)
            pm_val = pm_s.get(key, 0)

            row = {
                "Rank":        rank,
                "Name":        disp,
                "Value (EUR)": val,
                "% Total":     f"{val/g_total:.1%}",
                "YoY %":       fmt_pct(growth(val, py_s.get(key))) if flag_yoy else "—",
                "QoQ %":       fmt_pct(growth(cq_val, pq_val))     if flag_qoq else "—",
                "MoM %":       fmt_pct(growth(cm_val, pm_val))     if flag_mom else "—",
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
    """Render a Plotly pie chart from a ranked table."""
    if df_table is None or df_table.empty:
        st.caption("— No data for chart —")
        return
    d = df_table.head(top_n).copy()
    # Strip formatted strings if needed; Value (EUR) is still numeric here
    fig = px.pie(
        d,
        names=name_col,
        values="Value (EUR)",
        title=title,
        color_discrete_sequence=CHART_COLORS,
        hole=0.35,
    )
    fig.update_traces(textposition="outside", textinfo="percent+label")
    fig.update_layout(
        margin=dict(t=40, b=10, l=10, r=10),
        legend=dict(orientation="v", x=1.02, y=0.5),
        showlegend=True,
        height=380,
        font=dict(family="IBM Plex Sans"),
        title_font=dict(size=13, color="#1F3864"),
    )
    st.plotly_chart(fig, use_container_width=True)


def make_bar_trend(df_raw: pd.DataFrame, group_col: str, top_n: int, title: str, has_month: bool):
    """Bar chart — Value by Year (and optionally Year-Month) for top N items."""
    if df_raw is None or df_raw.empty:
        st.caption("— No data for trend chart —")
        return

    top_items = (
        df_raw.groupby(group_col)["Value (EUR)"]
        .sum()
        .nlargest(top_n)
        .index.tolist()
    )
    d = df_raw[df_raw[group_col].isin(top_items)].copy()

    if has_month:
        # Year-Month trend
        d["YM_sort"] = d["Year"] + "-" + d["Month"].str.zfill(2)
        agg = d.groupby([group_col, "YM_sort"])["Value (EUR)"].sum().reset_index()
        agg = agg.sort_values("YM_sort")
        fig = px.bar(
            agg, x="YM_sort", y="Value (EUR)", color=group_col,
            barmode="group",
            title=f"{title} — Monthly Trend",
            color_discrete_sequence=CHART_COLORS,
        )
        fig.update_xaxes(tickangle=-45, title="Year-Month")
    else:
        agg = d.groupby([group_col, "Year"])["Value (EUR)"].sum().reset_index()
        agg = agg.sort_values("Year")
        fig = px.bar(
            agg, x="Year", y="Value (EUR)", color=group_col,
            barmode="group",
            title=f"{title} — Annual Trend",
            color_discrete_sequence=CHART_COLORS,
        )
        fig.update_xaxes(title="Year")

    fig.update_yaxes(title="Value (EUR)")
    fig.update_layout(
        height=380,
        margin=dict(t=40, b=50, l=10, r=10),
        font=dict(family="IBM Plex Sans"),
        title_font=dict(size=13, color="#1F3864"),
        legend=dict(orientation="h", y=-0.25),
    )
    st.plotly_chart(fig, use_container_width=True)


def chart_controls(tab_key: str, max_n: int = 20) -> tuple[int, bool]:
    """Sidebar-style inline controls: Top N slider + show-month toggle."""
    c1, c2 = st.columns([3, 1])
    with c1:
        real_max = max(max_n, 1)
        real_min = min(3, real_max)
        if real_min >= real_max:
            top_n = real_max
            st.caption(f"Top N: {top_n} (only {top_n} item{'s' if top_n>1 else ''} available)")
        else:
            top_n = st.slider(
                "Top N items in charts", min_value=real_min, max_value=real_max,
                value=min(10, real_max),
                key=f"topn_{tab_key}"
            )
    with c2:
        show_month = st.checkbox("Monthly breakdown", value=False, key=f"month_{tab_key}")
    return top_n, show_month


def render_tab_charts(df_raw: pd.DataFrame, df_table: pd.DataFrame,
                      name_col: str, group_col: str,
                      tab_key: str, label: str, max_n: int = 20):
    """Full chart block: controls + pie + bar."""
    st.markdown("---")
    st.markdown("#### 📈 Charts")

    top_n, show_month = chart_controls(tab_key, max_n=min(max_n, len(df_table) if df_table is not None and not df_table.empty else max_n))

    col_pie, col_bar = st.columns([1, 1])
    with col_pie:
        make_pie(df_table, name_col, top_n, f"Top {top_n} {label} — Value Share")
    with col_bar:
        # check if multi-year
        years_in_data = df_raw["Year"].replace("", pd.NA).dropna().unique().tolist() if not df_raw.empty else []
        multi_year = len(years_in_data) > 1
        has_month_data = show_month and not df_raw["Month"].replace("", pd.NA).dropna().empty
        make_bar_trend(df_raw, group_col, top_n, f"Top {top_n} {label}", has_month=has_month_data)


# ─────────────────────────────────────────────────────────────────────────────
# PRICE ANALYSIS ENGINE
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)

# ─────────────────────────────────────────────────────────────────────────────
# PRICE ANALYSIS — ENGINE
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def compute_price_analysis(curr_bytes: bytes, cr_filter_tuple: tuple) -> pd.DataFrame:
    """Implied unit price per Comm Ref × Year × Country."""
    df = bytes_to_df(curr_bytes)
    if df.empty:
        return pd.DataFrame()
    cr_filter = list(cr_filter_tuple)
    if cr_filter:
        df = df[df["Commercial Reference"].isin(cr_filter)]
    df = df[df["Quantity"] > 0].copy()
    if df.empty:
        return pd.DataFrame()
    agg = df.groupby(
        ["Commercial Reference", "Comm Ref Code", "Family",
         "Strategic Product Family", "Year", "Country"], as_index=False
    ).agg(Value=("Value (EUR)", "sum"), Quantity=("Quantity", "sum"))
    agg["Unit Price (EUR)"] = agg["Value"] / agg["Quantity"]
    agg = agg.rename(columns={"Value": "Value (EUR)", "Quantity": "Total Qty"})
    return agg.sort_values(["Commercial Reference", "Country", "Year"])


@st.cache_data(show_spinner=False)
def compute_price_monthly(curr_bytes: bytes, cr_filter_tuple: tuple) -> pd.DataFrame:
    """Monthly implied unit price per Comm Ref × Year-Month × Country."""
    df = bytes_to_df(curr_bytes)
    if df.empty:
        return pd.DataFrame()
    cr_filter = list(cr_filter_tuple)
    if cr_filter:
        df = df[df["Commercial Reference"].isin(cr_filter)]
    df = df[df["Quantity"] > 0].copy()
    df = df[df["Month"].replace("", pd.NA).notna()].copy()
    if df.empty:
        return pd.DataFrame()
    df["YM"] = df["Year"] + "-" + df["Month"].str.zfill(2)
    agg = df.groupby(
        ["Commercial Reference", "Comm Ref Code", "Family",
         "Strategic Product Family", "YM", "Year", "Country"], as_index=False
    ).agg(Value=("Value (EUR)", "sum"), Quantity=("Quantity", "sum"))
    agg = agg[agg["Quantity"] > 0].copy()
    agg["Unit Price (EUR)"] = agg["Value"] / agg["Quantity"]
    agg = agg.rename(columns={"Value": "Value (EUR)", "Quantity": "Total Qty"})
    return agg.sort_values(["Commercial Reference", "Country", "YM"])


# ─────────────────────────────────────────────────────────────────────────────
# PRICE ANALYSIS — CHART HELPERS
# ─────────────────────────────────────────────────────────────────────────────
_CHART_H = 380
_LAYOUT  = dict(font=dict(family="IBM Plex Sans"), title_font=dict(size=13, color="#1F3864"),
                legend=dict(orientation="h", y=-0.35))

def _fig(f, h=_CHART_H):
    f.update_layout(height=h, **_LAYOUT)
    return f

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
    fn = px.bar if is_bar else px.line
    kw = dict(barmode="group") if is_bar else dict(markers=True)
    fig = fn(df, x=x, y="Unit Price (EUR)", color=color,
             title=title, color_discrete_sequence=CHART_COLORS, **kw)
    if not is_bar:
        pass
    fig.update_xaxes(tickangle=-40 if x == "YM" else 0)
    return _fig(fig)

def qty_trend_chart(df, x, color, title):
    fig = px.bar(df, x=x, y="Total Qty", color=color, barmode="group",
                 title=title, color_discrete_sequence=CHART_COLORS)
    fig.update_xaxes(tickangle=-40 if x == "YM" else 0)
    return _fig(fig)

def pct_chart(df, x, y, color, title):
    fig = px.bar(df, x=x, y=y, color=color, barmode="group",
                 title=title, color_discrete_sequence=CHART_COLORS)
    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
    fig.update_xaxes(tickangle=-40 if x == "YM" else 0)
    return _fig(fig, h=320)

def dual_axis_chart(df, x, title):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df[x], y=df["Total Qty"], name="Volume (Qty)",
        marker_color="#BDD7EE", opacity=0.85), secondary_y=False)
    fig.add_trace(go.Scatter(x=df[x], y=df["Unit Price (EUR)"], name="Unit Price (EUR)",
        mode="lines+markers", line=dict(color="#C55A11", width=2.5),
        marker=dict(size=7)), secondary_y=True)
    fig.update_layout(title=title, height=_CHART_H, **_LAYOUT)
    fig.update_yaxes(title_text="Volume (Qty)", secondary_y=False)
    fig.update_yaxes(title_text="Unit Price (EUR)", secondary_y=True)
    return fig

def show2(fig_l, fig_r):
    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(fig_l, use_container_width=True)
    with c2: st.plotly_chart(fig_r, use_container_width=True)

def section(title):
    st.markdown(f"#### {title}")
    st.markdown("---")


# ─────────────────────────────────────────────────────────────────────────────
# PRICE ANALYSIS — RENDER
# ─────────────────────────────────────────────────────────────────────────────
def render_price_tab(df_curr: pd.DataFrame, fmap: dict):
    st.markdown("#### 💶 Price vs Sales Analysis")
    st.caption("Implied unit price = Value (EUR) ÷ Quantity per row.")

    if df_curr.empty:
        st.warning("No data in current filter selection.")
        return

    df_q = df_curr[df_curr["Quantity"] > 0].copy()
    if df_q.empty:
        st.warning("No rows with Quantity > 0 found.")
        return

    # ── Helper ───────────────────────────────────────────────────────────────
    def _pv(col, df=df_q):
        return sorted(df[col].replace("", pd.NA).dropna().unique().tolist())

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 1 — FILTER PANEL (cascade, commit on button)
    # ─────────────────────────────────────────────────────────────────────────
    st.markdown("**🔍 Filter Selection** *(cascade: each level narrows the next)*")

    _p = df_q  # narrowing cursor

    fc1, fc2 = st.columns(2)
    with fc1:
        sel_plc = st.multiselect("Product Line Code", _pv("PLC"), placeholder="All PLCs…", key="pr_plc")
    if sel_plc: _p = _p[_p["PLC"].isin(sel_plc)]

    with fc2:
        sel_fam = st.multiselect("Family", _pv("Family", _p), placeholder="All Families…", key="pr_fam")
    if sel_fam: _p = _p[_p["Family"].isin(sel_fam)]

    fc3, fc4 = st.columns(2)
    with fc3:
        sel_spf = st.multiselect("Strategic Product Family", _pv("Strategic Product Family", _p),
                                 placeholder="All SPF…", key="pr_spf")
    if sel_spf: _p = _p[_p["Strategic Product Family"].isin(sel_spf)]

    with fc4:
        sel_crc = st.multiselect("Comm Ref Code", _pv("Comm Ref Code", _p),
                                 placeholder="All CRCs…", key="pr_crc")
    if sel_crc: _p = _p[_p["Comm Ref Code"].isin(sel_crc)]

    fc5, fc6 = st.columns(2)
    with fc5:
        avail_cr = _pv("Commercial Reference", _p)
        sel_cr   = st.multiselect("Comm Ref (optional — for detailed view)",
                                  avail_cr, placeholder="Leave empty for family-level view…", key="pr_cr")
    with fc6:
        avail_years = _pv("Year", df_q)
        sel_years   = st.multiselect("Years", avail_years, default=avail_years, key="pr_yr")

    avail_countries = _pv("Country", _p)
    sel_countries   = st.multiselect(
        "Countries",
        avail_countries,
        default=avail_countries[:15] if len(avail_countries) > 15 else avail_countries,
        key="pr_cty",
    )

    update_price_btn = st.button("▶  UPDATE PLOTS", use_container_width=True,
                                 type="primary", key="pr_update_btn")

    if update_price_btn:
        st.session_state["price_committed"] = {
            "cr":        sel_cr,
            "years":     sel_years,
            "countries": sel_countries,
            "fam":       sel_fam,
            "spf":       sel_spf,
            "crc":       sel_crc,
            "plc":       sel_plc,
        }
        st.rerun()

    if "price_committed" not in st.session_state or st.session_state["price_committed"] is None:
        st.info("👆 Set your filters above and press **▶ UPDATE PLOTS** to render charts.")
        return

    pc = st.session_state["price_committed"]

    # ── Build working datasets ────────────────────────────────────────────────
    df_base = df_q.copy()
    if pc["plc"]:     df_base = df_base[df_base["PLC"].isin(pc["plc"])]
    if pc["fam"]:     df_base = df_base[df_base["Family"].isin(pc["fam"])]
    if pc["spf"]:     df_base = df_base[df_base["Strategic Product Family"].isin(pc["spf"])]
    if pc["crc"]:     df_base = df_base[df_base["Comm Ref Code"].isin(pc["crc"])]
    if pc["years"]:   df_base = df_base[df_base["Year"].isin(pc["years"])]
    if pc["countries"]: df_base = df_base[df_base["Country"].isin(pc["countries"])]

    if df_base.empty:
        st.warning("No data matches the committed filter.")
        return

    has_cr_filter = bool(pc["cr"])

    # all CRs in scope
    all_cr_in_scope = sorted(df_base["Commercial Reference"].unique().tolist())

    # precompute annual + monthly price dfs for all CRs in scope
    annual_price_df  = compute_price_analysis(
        df_to_bytes(df_base), tuple(all_cr_in_scope))
    monthly_price_df = compute_price_monthly(
        df_to_bytes(df_base), tuple(all_cr_in_scope))

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 2 — SUB-PRODUCT SELECTOR
    # Determine the grouping dimension: Family or SPF
    # ─────────────────────────────────────────────────────────────────────────
    fam_selected = pc["fam"]
    spf_selected = pc["spf"]

    # decide primary grouping dim
    if fam_selected and not spf_selected:
        group_dim = "Family"
        group_vals = fam_selected
    elif spf_selected and not fam_selected:
        group_dim = "Strategic Product Family"
        group_vals = spf_selected
    elif fam_selected and spf_selected:
        group_dim = "Family"
        group_vals = fam_selected
    else:
        group_dim = "Family"
        group_vals = sorted(df_base["Family"].replace("", pd.NA).dropna().unique().tolist())

    st.markdown("---")
    section("🗂 Sub-Product Selector")

    if len(group_vals) > 1:
        focus_group = st.selectbox(
            f"Select {group_dim} to explore",
            group_vals, key="pr_focus_group"
        )
    else:
        focus_group = group_vals[0] if group_vals else None
        if focus_group:
            st.info(f"Showing: **{group_dim}** = {focus_group}")

    if not focus_group:
        st.warning("No product group found.")
        return

    # filter datasets to focus group
    df_grp  = df_base[df_base[group_dim] == focus_group]
    ap_grp  = annual_price_df[annual_price_df[group_dim] == focus_group]  if not annual_price_df.empty else pd.DataFrame()
    mp_grp  = monthly_price_df[monthly_price_df[group_dim] == focus_group] if not monthly_price_df.empty else pd.DataFrame()

    # CRCs within focus group
    crc_in_group = sorted(df_grp["Comm Ref Code"].replace("", pd.NA).dropna().unique().tolist())
    cr_in_group  = sorted(df_grp["Commercial Reference"].replace("", pd.NA).dropna().unique().tolist())

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 3 — FAMILY / SPF LEVEL: price trend by Year & Country
    # ─────────────────────────────────────────────────────────────────────────
    section(f"📦 {group_dim} Level — {focus_group}")

    # aggregate to group×year×country
    if not ap_grp.empty:
        grp_annual = ap_grp.groupby(["Year","Country"], as_index=False).agg(
            Value=("Value (EUR)", "sum"), Qty=("Total Qty", "sum"))
        grp_annual = grp_annual[grp_annual["Qty"] > 0].copy()
        grp_annual["Unit Price (EUR)"] = grp_annual["Value"] / grp_annual["Qty"]
        grp_annual.rename(columns={"Qty":"Total Qty","Value":"Value (EUR)"}, inplace=True)

        show2(
            price_trend_chart(grp_annual.sort_values(["Country","Year"]),
                              "Year", "Country", f"{focus_group} — Avg Unit Price by Year & Country"),
            qty_trend_chart(grp_annual.sort_values(["Country","Year"]),
                            "Year", "Country", f"{focus_group} — Total Volume by Year & Country"),
        )

        # YoY for group
        yoy = _yoy_pct(grp_annual, "Year")
        if not yoy.empty:
            st.plotly_chart(
                pct_chart(yoy, "Year", "YoY %", "Country",
                          f"{focus_group} — YoY Price Change %"),
                use_container_width=True)

    if not mp_grp.empty:
        grp_monthly = mp_grp.groupby(["YM","Country"], as_index=False).agg(
            Value=("Value (EUR)", "sum"), Qty=("Total Qty", "sum"))
        grp_monthly = grp_monthly[grp_monthly["Qty"] > 0].copy()
        grp_monthly["Unit Price (EUR)"] = grp_monthly["Value"] / grp_monthly["Qty"]
        grp_monthly.rename(columns={"Qty":"Total Qty","Value":"Value (EUR)"}, inplace=True)
        grp_monthly = grp_monthly.sort_values("YM")

        show2(
            price_trend_chart(grp_monthly, "YM", "Country",
                              f"{focus_group} — Monthly Avg Price"),
            pct_chart(_mom_pct(grp_monthly), "YM", "MoM %", "Country",
                      f"{focus_group} — MoM Price Change %"),
        )

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 4 — COMM REF CODE LEVEL inside focus group
    # X-axis = Comm Ref Code (model), Y-axis = Avg Unit Price
    # Filters: Country, then Year or Month
    # ─────────────────────────────────────────────────────────────────────────
    if crc_in_group:
        section(f"🔖 Comm Ref Code Level — within {focus_group}")

        # ── Filters ──────────────────────────────────────────────────────────
        s4c1, s4c2, s4c3 = st.columns(3)

        with s4c1:
            all_countries_grp = sorted(
                df_grp["Country"].replace("", pd.NA).dropna().unique().tolist())
            sel_s4_country = st.multiselect(
                "Country", all_countries_grp,
                default=all_countries_grp[:5] if len(all_countries_grp) > 5 else all_countries_grp,
                key="pr_s4_country",
            )
        with s4c2:
            s4_period = st.radio("Period", ["Year", "Month"], horizontal=True, key="pr_s4_period")
        with s4c3:
            all_years_grp = sorted(df_grp["Year"].replace("", pd.NA).dropna().unique().tolist())
            if s4_period == "Year":
                sel_s4_periods = st.multiselect(
                    "Select Years", all_years_grp, default=all_years_grp, key="pr_s4_years")
            else:
                # build YM options from monthly df
                all_ym = sorted(mp_grp["YM"].unique().tolist()) if not mp_grp.empty else []
                sel_s4_periods = st.multiselect(
                    "Select Year-Months", all_ym,
                    default=all_ym[-12:] if len(all_ym) >= 12 else all_ym,
                    key="pr_s4_ym")

        # ── Build aggregated data: X=CRC, colour=Country, facet=period ─────
        multi_country = len(sel_s4_country) > 1

        def _s4_make_chart(df_src, period_col, grp_keys, ctry_label):
            """Aggregate src → CRC × period × Country, return chart-ready df."""
            agg = df_src.groupby(grp_keys + ["Country"], as_index=False).agg(
                Value=("Value (EUR)", "sum"), Qty=("Total Qty", "sum"))
            agg = agg[agg["Qty"] > 0].copy()
            agg["Unit Price (EUR)"] = agg["Value"] / agg["Qty"]
            agg["Total Qty"] = agg["Qty"]
            return agg

        if s4_period == "Year":
            period_col = "Year"
            src = ap_grp.copy() if not ap_grp.empty else pd.DataFrame()
            if not src.empty:
                if sel_s4_country: src = src[src["Country"].isin(sel_s4_country)]
                if sel_s4_periods: src = src[src["Year"].isin(sel_s4_periods)]
                s4_agg = _s4_make_chart(src, period_col, ["Comm Ref Code", "Year"], "")
            else:
                s4_agg = pd.DataFrame()
        else:
            period_col = "YM"
            src = mp_grp.copy() if not mp_grp.empty else pd.DataFrame()
            if not src.empty:
                if sel_s4_country: src = src[src["Country"].isin(sel_s4_country)]
                if sel_s4_periods: src = src[src["YM"].isin(sel_s4_periods)]
                s4_agg = _s4_make_chart(src, period_col, ["Comm Ref Code", "YM"], "")
            else:
                s4_agg = pd.DataFrame()

        # ── Charts ────────────────────────────────────────────────────────────
        if not s4_agg.empty:
            ctry_label = ", ".join(sel_s4_country) if sel_s4_country else "All Countries"
            periods_selected = sorted(s4_agg[period_col].unique().tolist())

            # If multiple periods, let user pick one period for the X-axis snapshot
            # AND show trend charts below
            if len(periods_selected) > 1:
                sel_snap = st.select_slider(
                    f"Snapshot {period_col} (X-axis bar chart)",
                    options=periods_selected,
                    value=periods_selected[-1],
                    key="pr_s4_snap",
                )
            else:
                sel_snap = periods_selected[0]

            snap_df = s4_agg[s4_agg[period_col] == sel_snap].sort_values("Comm Ref Code")

            # ── Snapshot bar: X=CRC, Y=price, colour=Country ─────────────────
            fig_snap_price = px.bar(
                snap_df,
                x="Comm Ref Code",
                y="Unit Price (EUR)",
                color="Country",
                barmode="group",
                title=f"Avg Unit Price by Model — {focus_group} | {sel_snap} | {ctry_label}",
                color_discrete_sequence=CHART_COLORS,
                text_auto=".2s",
            )
            fig_snap_price.update_xaxes(tickangle=-35, title="Comm Ref Code (Model)")
            fig_snap_price.update_yaxes(title="Avg Unit Price (EUR)")
            fig_snap_price.update_layout(
                height=440, font=dict(family="IBM Plex Sans"),
                title_font=dict(size=13, color="#1F3864"),
                legend=dict(orientation="h", y=-0.35), bargap=0.15,
            )
            st.plotly_chart(fig_snap_price, use_container_width=True)

            # ── Snapshot bar: X=CRC, Y=volume, colour=Country ────────────────
            fig_snap_qty = px.bar(
                snap_df,
                x="Comm Ref Code",
                y="Total Qty",
                color="Country",
                barmode="group",
                title=f"Sales Volume by Model — {focus_group} | {sel_snap} | {ctry_label}",
                color_discrete_sequence=CHART_COLORS,
            )
            fig_snap_qty.update_xaxes(tickangle=-35, title="Comm Ref Code (Model)")
            fig_snap_qty.update_yaxes(title="Volume (Qty)")
            fig_snap_qty.update_layout(
                height=400, font=dict(family="IBM Plex Sans"),
                title_font=dict(size=13, color="#1F3864"),
                legend=dict(orientation="h", y=-0.35),
            )
            st.plotly_chart(fig_snap_qty, use_container_width=True)

            # ── Trend over time per model (line, colour=Country) ─────────────
            if len(periods_selected) > 1:
                st.markdown("**📈 Price Trend Over Time — per Model**")
                # Let user pick which models to trend
                all_models = sorted(s4_agg["Comm Ref Code"].unique().tolist())
                sel_trend_models = st.multiselect(
                    "Select models for trend chart (leave empty = all)",
                    all_models,
                    key="pr_s4_trend_models",
                )
                trend_df = s4_agg.copy()
                if sel_trend_models:
                    trend_df = trend_df[trend_df["Comm Ref Code"].isin(sel_trend_models)]

                # one tab per model, or facet if ≤4 models
                models_to_plot = sorted(trend_df["Comm Ref Code"].unique().tolist())
                if len(models_to_plot) <= 4:
                    for model in models_to_plot:
                        mdf = trend_df[trend_df["Comm Ref Code"] == model].sort_values([period_col, "Country"])
                        c_l, c_r = st.columns(2)
                        with c_l:
                            fig_tr = px.line(mdf, x=period_col, y="Unit Price (EUR)",
                                color="Country", markers=True,
                                title=f"{model} — Price Trend",
                                color_discrete_sequence=CHART_COLORS)
                            fig_tr.update_xaxes(tickangle=-40)
                            fig_tr.update_layout(height=320, font=dict(family="IBM Plex Sans"),
                                title_font=dict(size=12, color="#1F3864"),
                                legend=dict(orientation="h", y=-0.4))
                            st.plotly_chart(fig_tr, use_container_width=True)
                        with c_r:
                            fig_qr = px.bar(mdf, x=period_col, y="Total Qty",
                                color="Country", barmode="group",
                                title=f"{model} — Volume Trend",
                                color_discrete_sequence=CHART_COLORS)
                            fig_qr.update_xaxes(tickangle=-40)
                            fig_qr.update_layout(height=320, font=dict(family="IBM Plex Sans"),
                                title_font=dict(size=12, color="#1F3864"),
                                legend=dict(orientation="h", y=-0.4))
                            st.plotly_chart(fig_qr, use_container_width=True)
                else:
                    # too many models: combined line, colour=Country, one line per CRC+Country combo
                    trend_df["Label"] = trend_df["Comm Ref Code"] + " | " + trend_df["Country"]
                    fig_all = px.line(trend_df.sort_values([period_col, "Label"]),
                        x=period_col, y="Unit Price (EUR)", color="Label",
                        markers=True,
                        title=f"Price Trend — all selected models | {ctry_label}",
                        color_discrete_sequence=CHART_COLORS)
                    fig_all.update_xaxes(tickangle=-40)
                    fig_all.update_layout(height=420, font=dict(family="IBM Plex Sans"),
                        title_font=dict(size=13, color="#1F3864"),
                        legend=dict(orientation="h", y=-0.4))
                    st.plotly_chart(fig_all, use_container_width=True)

            # ── YoY / MoM per model × country ────────────────────────────────
            if s4_period == "Year":
                yoy_rows = []
                for (model, cty), grp in s4_agg.groupby(["Comm Ref Code","Country"]):
                    grp = grp.sort_values("Year").copy()
                    grp["YoY %"] = grp["Unit Price (EUR)"].pct_change() * 100
                    yoy_rows.append(grp)
                df_chg = pd.concat(yoy_rows).dropna(subset=["YoY %"]) if yoy_rows else pd.DataFrame()
                chg_col, chg_label = "YoY %", "YoY"
            else:
                mom_rows = []
                for (model, cty), grp in s4_agg.groupby(["Comm Ref Code","Country"]):
                    grp = grp.sort_values("YM").copy()
                    grp["MoM %"] = grp["Unit Price (EUR)"].pct_change() * 100
                    mom_rows.append(grp)
                df_chg = pd.concat(mom_rows).dropna(subset=["MoM %"]) if mom_rows else pd.DataFrame()
                chg_col, chg_label = "MoM %", "MoM"

            if not df_chg.empty:
                # snapshot period for % chart
                df_chg_snap = df_chg[df_chg[period_col] == sel_snap] if len(periods_selected) > 1 else df_chg
                if not df_chg_snap.empty:
                    fig_chg = px.bar(
                        df_chg_snap.sort_values("Comm Ref Code"),
                        x="Comm Ref Code", y=chg_col, color="Country",
                        barmode="group",
                        title=f"{chg_label} Price Change % by Model — {focus_group} | {sel_snap} | {ctry_label}",
                        color_discrete_sequence=CHART_COLORS,
                    )
                    fig_chg.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
                    fig_chg.update_xaxes(tickangle=-35)
                    fig_chg.update_layout(
                        height=340, font=dict(family="IBM Plex Sans"),
                        title_font=dict(size=13, color="#1F3864"),
                        legend=dict(orientation="h", y=-0.35),
                    )
                    st.plotly_chart(fig_chg, use_container_width=True)
        else:
            st.caption("— No data for the selected Country / Period combination —")

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 5 — SINGLE COMM REF DETAIL (only if user specified CRs)
    # ─────────────────────────────────────────────────────────────────────────
    if has_cr_filter:
        cr_to_detail = [c for c in pc["cr"] if c in cr_in_group]
        if not cr_to_detail:
            cr_to_detail = pc["cr"]

        section("🏷 Individual Comm Ref Detail")
        sel_detail_cr = st.selectbox("Select Comm Ref", cr_to_detail, key="pr_detail_cr")

        ap_cr  = annual_price_df[annual_price_df["Commercial Reference"] == sel_detail_cr] \
            if not annual_price_df.empty else pd.DataFrame()
        mp_cr  = monthly_price_df[monthly_price_df["Commercial Reference"] == sel_detail_cr] \
            if not monthly_price_df.empty else pd.DataFrame()

        if not ap_cr.empty:
            # Annual price + volume
            show2(
                price_trend_chart(ap_cr.sort_values(["Country","Year"]),
                                  "Year", "Country",
                                  f"{sel_detail_cr} — Annual Unit Price"),
                qty_trend_chart(ap_cr.sort_values(["Country","Year"]),
                                "Year", "Country",
                                f"{sel_detail_cr} — Annual Volume"),
            )
            # YoY
            yoy_cr = _yoy_pct(ap_cr, "Year")
            if not yoy_cr.empty:
                st.plotly_chart(
                    pct_chart(yoy_cr, "Year", "YoY %", "Country",
                              f"{sel_detail_cr} — YoY Price Change %"),
                    use_container_width=True)

            # Dual-axis per country
            countries_cr = ap_cr["Country"].unique().tolist()
            sel_ov_cty = st.selectbox("Country for Price vs Volume overlay",
                                      countries_cr, key="pr_ov_cty")
            d_ov = ap_cr[ap_cr["Country"] == sel_ov_cty].sort_values("Year")
            if not d_ov.empty:
                st.plotly_chart(
                    dual_axis_chart(d_ov, "Year",
                                    f"{sel_detail_cr} — {sel_ov_cty}: Price vs Volume"),
                    use_container_width=True)

        if not mp_cr.empty:
            st.markdown("##### Monthly Detail")
            show2(
                price_trend_chart(mp_cr.sort_values(["Country","YM"]),
                                  "YM", "Country",
                                  f"{sel_detail_cr} — Monthly Unit Price"),
                qty_trend_chart(mp_cr.sort_values(["Country","YM"]),
                                "YM", "Country",
                                f"{sel_detail_cr} — Monthly Volume"),
            )
            mom_cr = _mom_pct(mp_cr)
            if not mom_cr.empty:
                st.plotly_chart(
                    pct_chart(mom_cr, "YM", "MoM %", "Country",
                              f"{sel_detail_cr} — MoM Price Change %"),
                    use_container_width=True)

            # monthly dual-axis per country
            countries_m = mp_cr["Country"].unique().tolist()
            sel_m_cty = st.selectbox("Country for monthly Price vs Volume overlay",
                                     countries_m, key="pr_m_ov_cty")
            d_mvo = mp_cr[mp_cr["Country"] == sel_m_cty].sort_values("YM")
            if not d_mvo.empty:
                st.plotly_chart(
                    dual_axis_chart(d_mvo, "YM",
                                    f"{sel_detail_cr} — {sel_m_cty}: Monthly Price vs Volume"),
                    use_container_width=True)

        # Raw detail table
        with st.expander("📋 Raw Annual Price Table", expanded=False):
            if not ap_cr.empty:
                disp = ap_cr.copy()
                disp["Unit Price (EUR)"] = disp["Unit Price (EUR)"].map(lambda x: f"{x:,.2f}")
                disp["Total Qty"]        = disp["Total Qty"].map(lambda x: f"{x:,.0f}")
                disp["Value (EUR)"]      = disp["Value (EUR)"].map(fmt_val)
                st.dataframe(disp, use_container_width=True, hide_index=True)

    # ── Pivot overview ────────────────────────────────────────────────────────
    with st.expander("📊 All Comm Refs in Scope — Annual Price Pivot", expanded=False):
        if not annual_price_df.empty:
            pivot = annual_price_df.pivot_table(
                index=["Commercial Reference","Country"],
                columns="Year",
                values="Unit Price (EUR)",
                aggfunc="mean",
            ).reset_index()
            pivot.columns = [str(c) for c in pivot.columns]
            st.dataframe(pivot, use_container_width=True, hide_index=True)

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
    uname = st.session_state.get("username","")
    st.caption(f"👤 Logged in as **{uname}**")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()
    st.markdown("---")

    st.markdown("### 📁 Data Sources")

    order_file = st.file_uploader("Order File",  type=["xlsx","xls","xlsm","csv"], key="up_order")
    if order_file:
        df_loaded = load_file_with_progress(order_file)
        st.session_state.order_df = df_loaded
        st.success(f"✓ Order: {len(df_loaded):,} rows")

    sales_file = st.file_uploader("Sales File",  type=["xlsx","xls","xlsm","csv"], key="up_sales")
    if sales_file:
        df_loaded = load_file_with_progress(sales_file)
        st.session_state.sales_df = df_loaded
        st.success(f"✓ Sales: {len(df_loaded):,} rows")

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

    # ── helper: get unique sorted values from a df column ───────────────────
    def col_vals(df, col):
        if df is None or df.empty: return []
        return sorted(df[col].replace("", pd.NA).dropna().unique().tolist())

    def opts(col):
        return col_vals(active_df, col)

    st.markdown("---")

    with st.form(key="filter_form", border=False):
        # ── Time ─────────────────────────────────────────────────────────────
        st.markdown("### 🕐 Time Filters")
        f_year  = st.multiselect("Year",    opts("Year"),     placeholder="Select...")
        f_month = st.multiselect("Month",   [str(m) for m in range(1,13)], placeholder="Select...")

        st.markdown("---")
        update_btn = st.form_submit_button("▶  UPDATE ALL TABLES", use_container_width=True, type="primary", disabled=(active_df is None))

    # ── Geography Filters OUTSIDE form → cascade Operations→Zone→Cluster→Area→Country
    st.markdown("### 🌍 Geography Filters")
    st.caption("⬇ Cascade: each selection narrows the next.")

    _df_geo = active_df if active_df is not None else pd.DataFrame(columns=STR_COLS)

    f_ops = st.multiselect("Operations", col_vals(_df_geo, "Operations"), placeholder="Select...", key="sb_ops")
    if f_ops and not _df_geo.empty:
        _df_geo = _df_geo[_df_geo["Operations"].isin(f_ops)]

    f_zone = st.multiselect("Zone", col_vals(_df_geo, "Zone"), placeholder="Select...", key="sb_zone")
    if f_zone and not _df_geo.empty:
        _df_geo = _df_geo[_df_geo["Zone"].isin(f_zone)]

    f_cluster = st.multiselect("Cluster", col_vals(_df_geo, "Cluster"), placeholder="Select...", key="sb_cluster")
    if f_cluster and not _df_geo.empty:
        _df_geo = _df_geo[_df_geo["Cluster"].isin(f_cluster)]

    f_area = st.multiselect("Area", col_vals(_df_geo, "Area"), placeholder="Select...", key="sb_area")
    if f_area and not _df_geo.empty:
        _df_geo = _df_geo[_df_geo["Area"].isin(f_area)]

    f_country = st.multiselect("Country", col_vals(_df_geo, "Country"), placeholder="Select...", key="sb_country")

    # ── Product Filters OUTSIDE form → cascade PLC→Family→SPF→CRC→CR ─────────
    st.markdown("### 📦 Product Filters")
    st.caption("⬇ Cascade: each selection narrows the next.")

    _df_sb = active_df if active_df is not None else pd.DataFrame(columns=STR_COLS)

    f_plc = st.multiselect("Product Line Code", col_vals(_df_sb, "PLC"), placeholder="Select...", key="sb_plc")
    if f_plc and not _df_sb.empty:
        _df_sb = _df_sb[_df_sb["PLC"].isin(f_plc)]

    f_fam = st.multiselect("Family", col_vals(_df_sb, "Family"), placeholder="Select...", key="sb_fam")
    if f_fam and not _df_sb.empty:
        _df_sb = _df_sb[_df_sb["Family"].isin(f_fam)]

    f_spf = st.multiselect("Strategic Product Family", col_vals(_df_sb, "Strategic Product Family"), placeholder="Select...", key="sb_spf")
    if f_spf and not _df_sb.empty:
        _df_sb = _df_sb[_df_sb["Strategic Product Family"].isin(f_spf)]

    f_crc = st.multiselect("Comm Ref Code", col_vals(_df_sb, "Comm Ref Code"), placeholder="Select...", key="sb_crc")
    if f_crc and not _df_sb.empty:
        _df_sb = _df_sb[_df_sb["Comm Ref Code"].isin(f_crc)]

    f_cr = st.multiselect("Comm Ref", col_vals(_df_sb, "Commercial Reference"), placeholder="Select...", key="sb_cr")

    st.markdown("---")
    rst_btn = st.button("↺  RESET ALL FILTERS", use_container_width=True, disabled=(active_df is None))

    if rst_btn:
        for k in ["sb_ops","sb_zone","sb_cluster","sb_area","sb_country",
                  "sb_plc","sb_fam","sb_spf","sb_crc","sb_cr"]:
            st.session_state.pop(k, None)
        st.session_state.committed     = None
        st.session_state.tables        = None
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
    st.markdown(f'<br><span class="source-badge {badge}">{perf_source.upper()}</span>',
                unsafe_allow_html=True)

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
# YoY / QoQ / MoM LOGIC
# Rules:
#  A) No year+month filters → use latest YM in data, auto-derive intervals
#  B) Year + months selected, months are consecutive from 01 → use them
#  C) Any other combination → suppress all three metrics
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

def _consecutive_from_01(months):
    if not months: return False
    nums = sorted(int(m) for m in months)
    return nums[0] == 1 and nums == list(range(1, nums[-1] + 1))

def _last_complete_quarter_months(month_num):
    """Full months of the last completed quarter, capped at month_num.
    e.g. month=5 → Q1=[1,2,3]; month=8 → Q2=[4,5,6]; month=1,2,3 → Q4=[10,11,12] prev yr
    Returns (months_list, year_offset)
    """
    completed_q = (month_num - 1) // 3
    if completed_q == 0:
        return [str(m) for m in range(10, 13)], -1
    return [str(m) for m in range((completed_q-1)*3+1, completed_q*3+1)], 0

def _af(year, months):
    return apply_filters(active_df, {**fmap_no_time, "Year": [str(year)], "Month": months})

# ── Case A: no year/month filter → anchor on latest YM in full dataset ───
if not yr_sel and not mo_sel:
    all_ym = active_df["Year-Month"].replace("", pd.NA).dropna()
    all_ym = all_ym[all_ym.str.len() == 6]
    if not all_ym.empty:
        latest_ym = all_ym.max()
        cur_yr  = int(latest_ym[:4])
        cur_mo  = int(latest_ym[4:6])
        prev_yr = cur_yr - 1
        yoy_range = [str(m) for m in range(1, cur_mo + 1)]

        df_curr = _af(cur_yr, yoy_range)

        # YoY: cur_yr 01..cur_mo  vs  prev_yr 01..cur_mo
        flag_yoy  = True
        df_prev_y = _af(prev_yr, yoy_range)

        # QoQ: last completed quarter cur_yr  vs  same quarter prev_yr
        q_mos, yr_off = _last_complete_quarter_months(cur_mo)
        flag_qoq  = True
        df_curr_q = _af(cur_yr  + yr_off, q_mos)
        df_prev_q = _af(prev_yr + yr_off, q_mos)

        # MoM: cur_yr/cur_mo  vs  prev_yr/cur_mo (same month, prior year)
        flag_mom  = True
        df_curr_m = _af(cur_yr,  [str(cur_mo)])
        df_prev_m = _af(prev_yr, [str(cur_mo)])

# ── Case B: single year + (single month OR consecutive months from 01) ───
elif len(yr_sel) == 1 and mo_sel and (
        len(mo_sel) == 1 or _consecutive_from_01(mo_sel)):
    yr      = int(yr_sel[0])
    prev_yr = yr - 1
    max_m   = max(int(m) for m in mo_sel)
    yoy_range = [str(m) for m in range(1, max_m + 1)]

    # Realign df_curr to YoY window 01..max_m
    df_curr = _af(yr, yoy_range)

    # YoY: yr 01..max_m  vs  prev_yr 01..max_m
    flag_yoy  = True
    df_prev_y = _af(prev_yr, yoy_range)

    # QoQ: last completed quarter yr  vs  same quarter prev_yr
    q_mos, yr_off = _last_complete_quarter_months(max_m)
    flag_qoq  = True
    df_curr_q = _af(yr      + yr_off, q_mos)
    df_prev_q = _af(prev_yr + yr_off, q_mos)

    # MoM: yr/max_m  vs  prev_yr/max_m (same month, prior year)
    flag_mom  = True
    df_curr_m = _af(yr,      [str(max_m)])
    df_prev_m = _af(prev_yr, [str(max_m)])

# ── Case C: anything else → suppress all metrics ─────────────────────────

with st.spinner("Calculating tables…"):
    tables, g_total = compute_tables(
        df_to_bytes(df_curr),
        df_to_bytes(df_prev_y),
        df_to_bytes(df_prev_q),
        df_to_bytes(df_prev_m),
        df_to_bytes(df_curr_q),
        df_to_bytes(df_curr_m),
        geo_all, flag_yoy, flag_qoq, flag_mom,
    )

if tables is None:
    st.warning("⚠️ No data matches the current filters.")
    st.stop()

# ── Summary metrics
mc = st.columns(5)
mc[0].metric("Total Value (EUR)",      fmt_val(g_total))
mc[1].metric("# Transactions",         f"{len(df_curr):,}")
mc[2].metric("# Countries",            df_curr["Country"].nunique())
mc[3].metric("# Product Line Codes",   df_curr["PLC"].nunique())
mc[4].metric("# Comm Refs",            df_curr["Commercial Reference"].nunique())

st.markdown("---")

# ── Year-over-Year Monthly Performance Chart ──────────────────────────────────
def render_yoy_monthly_chart(df_base, fmap_no_time, active_df, yr_sel, mo_sel, src_label):
    """Bar chart comparing current year vs prior year month by month."""

    # Determine which years to compare
    all_ym = active_df["Year-Month"].replace("", pd.NA).dropna()
    all_ym = all_ym[all_ym.str.len() == 6]
    if all_ym.empty:
        return

    if yr_sel:
        cur_yr  = int(yr_sel[0])
    else:
        cur_yr  = int(all_ym.max()[:4])
    prev_yr = cur_yr - 1

    # Pull monthly data for both years, respecting non-time filters
    df_cy = apply_filters(active_df, {**fmap_no_time, "Year": [str(cur_yr)]})
    df_py = apply_filters(active_df, {**fmap_no_time, "Year": [str(prev_yr)]})

    if df_cy.empty and df_py.empty:
        return

    def _monthly_agg(df, year):
        if df.empty:
            return pd.DataFrame(columns=["Month_num","Month","Value (EUR)","Year"])
        agg = df.groupby("Month", as_index=False)["Value (EUR)"].sum()
        agg = agg[agg["Month"].replace("", pd.NA).notna()]
        agg["Month_num"] = agg["Month"].astype(int)
        agg["Year"] = str(year)
        return agg.sort_values("Month_num")

    cy_agg = _monthly_agg(df_cy, cur_yr)
    py_agg = _monthly_agg(df_py, prev_yr)
    combined = pd.concat([py_agg, cy_agg], ignore_index=True)

    if combined.empty:
        return

    # Month labels
    month_names = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                   7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    combined["Month_label"] = combined["Month_num"].map(month_names)

    # Sort by month number so x-axis is in order
    combined = combined.sort_values(["Month_num","Year"])

    fig = px.bar(
        combined,
        x="Month_label",
        y="Value (EUR)",
        color="Year",
        barmode="group",
        title=f"Monthly Performance: {cur_yr} vs {prev_yr}  {src_label}",
        color_discrete_map={
            str(prev_yr): "#BDD7EE",
            str(cur_yr):  "#1F3864",
        },
        text_auto=".2s",
        category_orders={"Month_label": [month_names[m] for m in range(1,13)]},
    )
    fig.update_yaxes(title="Value (EUR)")
    fig.update_xaxes(title="Month")
    fig.update_layout(
        height=400,
        font=dict(family="IBM Plex Sans"),
        title_font=dict(size=14, color="#1F3864"),
        legend=dict(orientation="h", y=-0.25),
        bargap=0.2,
    )
    st.plotly_chart(fig, use_container_width=True)

render_yoy_monthly_chart(df_curr, fmap_no_time, active_df, yr_sel, mo_sel, f"({src})")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# RENDER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def render(df, title, rename=None, drop=None):
    st.markdown(f'<h4>{title}</h4>', unsafe_allow_html=True)
    if df is None or df.empty:
        st.caption("— No data —")
        return
    d = df.copy()
    if drop:
        d = d.drop(columns=[c for c in drop if c in d.columns])
    if rename:
        d = d.rename(columns=rename)
    if "Value (EUR)" in d.columns:
        d["Value (EUR)"] = d["Value (EUR)"].apply(fmt_val)
    st.dataframe(d, use_container_width=True, hide_index=True, height=min(38 * len(d) + 40, 520))

label = f"({src})"

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📍 ZONE",
    "🌍 COUNTRIES",
    "📦 PRODUCT LINE CODE",
    "🏢 FAMILIES",
    "🔖 COMM REF CODES",
    "🏷️ COMM REFS",
    "💶 PRICE ANALYSIS",
])

# ─────────────────────────────────────────────────────────────────────────────
# YoY MONTHLY BREAKDOWN HELPER
# ─────────────────────────────────────────────────────────────────────────────
def render_yoy_breakdown(df_curr: pd.DataFrame, active_df: pd.DataFrame,
                         fmap_no_time: dict, dim_col: str,
                         tab_key: str, yr_sel: list, label: str):
    """
    YoY monthly bar chart broken down by dim_col.
    Multiselect + Apply button controls which items to show.
    Current year = solid; prior year = same colour, dimmed.
    """
    st.markdown("---")
    st.markdown("#### 📅 YoY Monthly Breakdown")

    # Determine years
    all_ym = active_df["Year-Month"].replace("", pd.NA).dropna()
    all_ym = all_ym[all_ym.str.len() == 6]
    if all_ym.empty:
        st.caption("No Year-Month data available.")
        return

    cur_yr  = int(yr_sel[0]) if yr_sel else int(all_ym.max()[:4])
    prev_yr = cur_yr - 1

    # Pull both years respecting non-time filters
    df_cy = apply_filters(active_df, {**fmap_no_time, "Year": [str(cur_yr)]})
    df_py = apply_filters(active_df, {**fmap_no_time, "Year": [str(prev_yr)]})

    if df_cy.empty and df_py.empty:
        st.caption("No data for comparison.")
        return

    # All available items ranked by current year value
    if not df_cy.empty and dim_col in df_cy.columns:
        all_items = (df_cy.groupby(dim_col)["Value (EUR)"].sum()
                     .sort_values(ascending=False).index.tolist())
    else:
        all_items = []

    if not all_items:
        st.caption(f"No data for dimension '{dim_col}'.")
        return

    default_items = all_items[:min(10, len(all_items))]
    commit_key    = f"yoy_bd_committed_{tab_key}"

    # ── Filter controls + Apply button ───────────────────────────────────────
    fc1, fc2 = st.columns([4, 1])
    with fc1:
        sel_items = st.multiselect(
            f"Select {dim_col}s to compare",
            options=all_items,
            default=default_items,
            key=f"yoy_bd_sel_{tab_key}",
            placeholder=f"Choose {dim_col}s…",
        )
    with fc2:
        st.markdown("<br>", unsafe_allow_html=True)
        apply_btn = st.button("▶ Apply", key=f"yoy_bd_apply_{tab_key}",
                              use_container_width=True, type="primary")

    if apply_btn:
        st.session_state[commit_key] = sel_items if sel_items else default_items
        st.rerun()

    # Use committed selection, fallback to default on first load
    top_items = st.session_state.get(commit_key, default_items)
    if not top_items:
        top_items = default_items

    # ── Monthly aggregation ───────────────────────────────────────────────────
    month_names = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                   7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}

    def _agg(df, year):
        if df.empty or dim_col not in df.columns:
            return pd.DataFrame()
        d = df[df[dim_col].isin(top_items)].copy()
        d = d[d["Month"].replace("", pd.NA).notna()]
        if d.empty: return pd.DataFrame()
        agg = d.groupby([dim_col, "Month"], as_index=False)["Value (EUR)"].sum()
        agg["Month_num"]   = agg["Month"].astype(int)
        agg["Month_label"] = agg["Month_num"].map(month_names)
        agg["Year"]        = str(year)
        agg["Series"]      = agg[dim_col] + f" ({year})"
        return agg

    cy_agg = _agg(df_cy, cur_yr)
    py_agg = _agg(df_py, prev_yr)
    combined = pd.concat([py_agg, cy_agg], ignore_index=True)

    if combined.empty:
        st.caption("No monthly data available.")
        return

    # ── Colour map: same colour per item, dimmed for prior year ──────────────
    n = len(top_items)
    base_colors = CHART_COLORS * ((n // len(CHART_COLORS)) + 1)
    color_map = {}
    for i, item in enumerate(top_items):
        color_map[f"{item} ({cur_yr})"]  = base_colors[i]
        color_map[f"{item} ({prev_yr})"] = base_colors[i]

    fig = px.bar(
        combined.sort_values(["Month_num", dim_col, "Year"]),
        x="Month_label",
        y="Value (EUR)",
        color="Series",
        barmode="group",
        title=f"YoY Monthly Breakdown by {dim_col} — {cur_yr} vs {prev_yr}  {label}",
        color_discrete_map=color_map,
        category_orders={
            "Month_label": [month_names[m] for m in range(1, 13)],
            "Series": [f"{it} ({y})"
                       for it in top_items
                       for y in [prev_yr, cur_yr]],
        },
        text_auto=".2s",
    )

    for trace in fig.data:
        trace.marker.opacity = 0.95 if f"({prev_yr})" not in trace.name else 0.4

    fig.update_xaxes(title="Month")
    fig.update_yaxes(title="Value (EUR)")
    fig.update_layout(
        height=max(440, 55 * n),
        font=dict(family="IBM Plex Sans"),
        title_font=dict(size=13, color="#1F3864"),
        legend=dict(orientation="h", y=-0.35),
        bargap=0.15,
    )
    st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    render(tables["zone"], f">> ZONE PERFORMANCE {label}", rename={"Name": "Zone"})
    render_tab_charts(df_curr, tables["zone"], "Name", "Zone", "zone", "Zones")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Zone", "zone", yr_sel, label)

with tab2:
    t = tables["country"]
    if t is not None and not t.empty:
        t = t.rename(columns={"Parent": "Zone", "Country": "Cluster"})
    render(t, f">> TOP 20 COUNTRIES {label}", rename={"Name": "Country"})
    render_tab_charts(df_curr, tables["country"], "Name", "Country", "country", "Countries")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Country", "country", yr_sel, label)

with tab3:
    render(tables["plc"], f">> TOP 20 PLC {label}", rename={"Name": "Product Line Code"}, drop=["Parent"])
    render_tab_charts(df_curr, tables["plc"], "Name", "PLC", "plc", "PLCs")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "PLC", "plc", yr_sel, label)

with tab4:
    render(tables["family"], f">> TOP 20 FAMILIES {label}", rename={"Name": "Family", "Parent": "PLC"})
    render_tab_charts(df_curr, tables["family"], "Name", "Family", "family", "Families")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Family", "family", yr_sel, label)

with tab5:
    render(tables["crc"], f">> TOP 20 COMM REF CODES {label}", rename={"Name": "Comm Ref Code", "Parent": "Family"})
    render_tab_charts(df_curr, tables["crc"], "Name", "Comm Ref Code", "crc", "Comm Ref Codes")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Comm Ref Code", "crc", yr_sel, label)

with tab6:
    t_cr = tables["cr"]
    if t_cr is not None and not t_cr.empty:
        if geo_all and "Country" in t_cr.columns:
            t_cr = t_cr.drop(columns=["Country"])
        t_cr = t_cr.rename(columns={"Name": "Comm Ref", "Parent": "Comm Ref Code"})
    render(t_cr, f">> TOP 20 COMM REFS {label}")
    render_tab_charts(df_curr, tables["cr"], "Name", "Commercial Reference", "cr", "Comm Refs")
    render_yoy_breakdown(df_curr, active_df, fmap_no_time, "Commercial Reference", "cr", yr_sel, label)

with tab7:
    render_price_tab(df_curr, fmap)

st.markdown("---")

with st.expander("📋 Raw Data Preview (filtered)", expanded=False):
    st.dataframe(df_curr[TARGET_HEADERS].head(2000), use_container_width=True, hide_index=True, height=400)
    st.caption(f"Showing first 2,000 of {len(df_curr):,} filtered rows.")