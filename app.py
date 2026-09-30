"""
FlashBI - a fast, self-service analytics + ETL workspace.

Install:  pip install streamlit pandas plotly openpyxl xlrd numpy
Run:      streamlit run app.py
(xlrd is only needed for legacy .xls files.)
"""
import hmac
import html
import io
import json
import operator
import os
import re
import time
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from pandas.api.types import (is_bool_dtype, is_datetime64_any_dtype,
                              is_numeric_dtype)

st.set_page_config(page_title="FlashBI", page_icon="⚡", layout="wide")

# Streamlit >= 1.50 uses width="stretch"; older versions use use_container_width.
_v = tuple(int(p) for p in re.findall(r"\d+", st.__version__)[:2])
STRETCH = {"width": "stretch"} if _v >= (1, 50) else {"use_container_width": True}
# Fragments rerun only their own block when a widget inside them changes (much faster).
fragment = getattr(st, "fragment", None) or getattr(st, "experimental_fragment", None) or (lambda f: f)

# ───────────────────────── Theme & CSS ─────────────────────────
with st.sidebar:
    st.markdown("## ⚡ FlashBI")
    dark = st.toggle("Dark mode", value=True)

if dark:
    T = dict(bg="#0a0f1f", card="#131a30", text="#e9edfa", muted="#8d99c2",
             border="rgba(255,255,255,.08)", shadow="0 10px 30px rgba(0,0,0,.4)")
else:
    T = dict(bg="#f3f5fc", card="#ffffff", text="#1a2040", muted="#68729a",
             border="rgba(25,35,90,.09)", shadow="0 10px 30px rgba(50,60,120,.13)")
TEMPLATE = "plotly_dark" if dark else "plotly_white"
PALETTE = ["#7c5cff", "#22d3ee", "#f472b6", "#facc15", "#34d399", "#fb923c", "#60a5fa", "#a78bfa"]
CFG = {"displaylogo": False}

ROOT = (f":root{{--bg:{T['bg']};--card:{T['card']};--text:{T['text']};--muted:{T['muted']};"
        f"--border:{T['border']};--shadow:{T['shadow']};}}")
CSS = """
.stApp{background:var(--bg);transition:background .4s ease}
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp label,
[data-testid="stMarkdownContainer"] p,[data-testid="stCaptionContainer"]{color:var(--text)!important}
[data-testid="stSidebar"]{background:var(--card);border-right:1px solid var(--border)}
[data-baseweb="select"]>div,[data-baseweb="input"],input,textarea{background:var(--card)!important;color:var(--text)!important;border-radius:10px!important}
@keyframes fadeUp{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:none}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
@keyframes glow{0%,100%{box-shadow:0 0 0 0 rgba(124,92,255,.35)}50%{box-shadow:0 0 0 14px rgba(124,92,255,0)}}
.hero{padding:1.6rem 0 .6rem;animation:fadeUp .7s both}
.hero h1{font-size:3rem;font-weight:800;margin:0;letter-spacing:-.03em;
 background:linear-gradient(90deg,#7c5cff,#22d3ee,#f472b6,#7c5cff);background-size:200% auto;
 -webkit-background-clip:text;-webkit-text-fill-color:transparent;animation:shimmer 6s linear infinite}
.hero div{color:var(--muted);font-size:1.1rem;margin-top:.3rem}
[data-testid="stFileUploaderDropzone"]{border:2px dashed #7c5cff!important;border-radius:22px!important;
 background:var(--card)!important;padding:2.6rem!important;transition:all .3s ease;animation:glow 3s infinite}
[data-testid="stFileUploaderDropzone"]:hover{transform:scale(1.01);border-color:#22d3ee!important}
.kpi,.insight,.feature{background:var(--card);border:1px solid var(--border);border-radius:20px;
 padding:1.1rem 1.3rem;box-shadow:var(--shadow);animation:fadeUp .6s both;transition:transform .25s,box-shadow .25s}
.kpi:hover,.insight:hover,.feature:hover{transform:translateY(-4px)}
.kpi .ic{font-size:1.5rem}
.kpi .lb,.insight .tt{color:var(--muted);font-size:.85rem;font-weight:600}
.kpi .vl{font-size:1.7rem;overflow-wrap:anywhere;line-height:1.25;font-weight:800;color:var(--text);
 background:linear-gradient(90deg,#7c5cff,#22d3ee);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.insight{margin-bottom:1rem;border-left:5px solid #7c5cff}
.insight .bd{color:var(--text);font-size:1.05rem;line-height:1.55;margin-top:.3rem}
.feature{text-align:center;color:var(--text)}
.story{background:var(--card);border:1px solid var(--border);border-left:5px solid #22d3ee;border-radius:20px;
 padding:1.2rem 1.5rem;box-shadow:var(--shadow);margin:1rem 0;animation:fadeUp .6s both}
.story .stt{font-weight:800;font-size:1.1rem;color:var(--text);margin-bottom:.4rem}
.story .sl{color:var(--text);line-height:1.65;margin:.5rem 0}
[data-testid="stPlotlyChart"]{background:var(--card);border:1px solid var(--border);border-radius:20px;
 padding:8px;box-shadow:var(--shadow);animation:fadeUp .7s both;transition:transform .25s}
[data-testid="stPlotlyChart"]:hover{transform:translateY(-3px)}
div[data-baseweb="tab-list"]{gap:6px;background:var(--card);padding:6px;border-radius:18px;border:1px solid var(--border);
 box-shadow:var(--shadow);overflow-x:auto}
div[data-baseweb="tab-highlight"],div[data-baseweb="tab-border"]{display:none}
button[data-baseweb="tab"]{border-radius:13px!important;padding:9px 20px;font-weight:600;transition:all .25s;white-space:nowrap}
button[data-baseweb="tab"]:hover{background:rgba(124,92,255,.14)}
button[data-baseweb="tab"][aria-selected="true"]{background:linear-gradient(135deg,#7c5cff,#22d3ee)!important;box-shadow:0 6px 18px rgba(124,92,255,.4)}
.stApp button[aria-selected="true"] [data-testid="stMarkdownContainer"] p{color:#fff!important}
[data-testid="stTabs"] [role="tabpanel"]{padding-top:1.2rem;animation:fadeUp .5s both}
.lp{position:relative;overflow:hidden;border-radius:28px;padding:3rem 2.4rem;min-height:540px;color:#fff;
 background:linear-gradient(145deg,#2b1a7a,#5b3df0 55%,#1aa6c4);box-shadow:0 24px 60px rgba(91,61,240,.4);animation:fadeUp .7s both}
.lp .blob{position:absolute;border-radius:50%;filter:blur(40px);opacity:.55;animation:float 9s ease-in-out infinite}
.lp .b1{width:220px;height:220px;background:#f472b6;top:-60px;right:-40px}
.lp .b2{width:260px;height:260px;background:#22d3ee;bottom:-90px;left:-70px;animation-delay:-4s}
@keyframes float{50%{transform:translate(20px,-24px) scale(1.1)}}
.lp-in{position:relative}
.lp .login-logo{background:rgba(255,255,255,.2);box-shadow:none;animation:none;margin:0 0 18px}
.lp-t{font-size:2.6rem;font-weight:800;letter-spacing:-.03em;color:#fff}
.lp-sub{color:rgba(255,255,255,.85);font-size:1.05rem;margin:.4rem 0 1.6rem}
.lp-li{color:#fff;padding:.55rem 0;border-top:1px solid rgba(255,255,255,.18);font-size:.98rem}
.login-h{font-size:2rem;font-weight:800;color:var(--text);margin-top:7vh;letter-spacing:-.02em;animation:fadeUp .6s both}
.login-s{color:var(--muted);margin:.2rem 0 1.2rem}
input:focus{box-shadow:0 0 0 3px rgba(124,92,255,.35)!important}
.splash{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:70vh;gap:1.4rem;animation:fadeUp .4s both}
.ring.big{width:84px;height:84px;border-width:7px}
.sp-t{color:var(--text);font-size:1.2rem;font-weight:700}
.sp-bar{width:240px;height:6px;border-radius:6px;background:linear-gradient(90deg,transparent,#7c5cff,#22d3ee,transparent);
 background-size:200% 100%;animation:shimmer 1.2s linear infinite}
.user-chip{display:inline-block;margin-top:1.5rem;padding:9px 16px;border-radius:999px;background:var(--card);
 border:1px solid var(--border);color:var(--text);font-weight:600;box-shadow:var(--shadow)}
.stApp .stButton button,.stApp .stDownloadButton button,.stApp [data-testid="stFormSubmitButton"] button,
.stApp [data-testid="stFileUploaderDropzone"] button,.stApp [data-testid^="stBaseButton-secondary"],
.stApp [data-testid^="stBaseButton-primary"]{border:0!important;border-radius:12px!important;font-weight:600;
 background:linear-gradient(135deg,#7c5cff,#22d3ee)!important;color:#fff!important;transition:transform .2s,box-shadow .2s}
.stApp .stButton button *,.stApp .stDownloadButton button *,.stApp [data-testid="stFormSubmitButton"] button *,
.stApp [data-testid="stFileUploaderDropzone"] button *,.stApp [data-testid^="stBaseButton-secondary"] *,
.stApp [data-testid^="stBaseButton-primary"] *{color:#fff!important}
.stApp .stButton button:hover,.stApp .stDownloadButton button:hover,.stApp [data-testid="stFormSubmitButton"] button:hover{
 transform:translateY(-2px);box-shadow:0 8px 22px rgba(124,92,255,.45)}
.stApp button:disabled{opacity:.45;transform:none;box-shadow:none}
.st-key-nav [role="radiogroup"]{gap:6px;background:var(--card);padding:6px;border-radius:18px;border:1px solid var(--border);
 box-shadow:var(--shadow);flex-wrap:wrap}
.st-key-nav label{padding:9px 20px;border-radius:13px;cursor:pointer;transition:all .25s;margin:0}
.st-key-nav label>div:first-child{display:none}
.st-key-nav label:hover{background:rgba(124,92,255,.14)}
.st-key-nav label:has(input:checked){background:linear-gradient(135deg,#7c5cff,#22d3ee);box-shadow:0 6px 18px rgba(124,92,255,.4)}
.st-key-nav label:has(input:checked) *{color:#fff!important}
.sug{background:var(--bg);border:1px solid var(--border);border-left:4px solid #facc15;border-radius:12px;
 padding:.6rem .8rem;margin:.5rem 0;font-size:.86rem;color:var(--text);line-height:1.45}
.sug b{display:block;margin-bottom:.15rem}
.loader{display:flex;align-items:center;gap:1rem;padding:1.2rem 1.5rem;border-radius:18px;
 background:var(--card);border:1px solid var(--border);box-shadow:var(--shadow);margin:1rem 0}
.ring{width:38px;height:38px;border-radius:50%;border:4px solid rgba(124,92,255,.2);
 border-top-color:#7c5cff;border-right-color:#22d3ee;animation:spin .8s linear infinite}
.loader .lt{color:var(--text);font-weight:600}
[data-testid="stForm"]{background:var(--card);border:1px solid var(--border)!important;border-radius:22px;
 padding:1.6rem;box-shadow:var(--shadow);animation:fadeUp .7s .1s both}
.login-top{text-align:center;padding:4vh 0 1.2rem;animation:fadeUp .6s both}
.login-logo{width:72px;height:72px;margin:0 auto 14px;border-radius:22px;display:flex;align-items:center;
 justify-content:center;font-size:2rem;background:linear-gradient(135deg,#7c5cff,#22d3ee);
 box-shadow:0 12px 30px rgba(124,92,255,.45);animation:glow 3s infinite}
.login-top h1{font-size:2.4rem;font-weight:800;margin:0;letter-spacing:-.03em;
 background:linear-gradient(90deg,#7c5cff,#22d3ee,#f472b6);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.login-top div.sub{color:var(--muted);margin-top:.3rem}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
"""
st.markdown(f"<style>{ROOT}{CSS}</style>", unsafe_allow_html=True)


# ───────────────────────── Login ─────────────────────────
# Default login is admin / admin. Change it here, or set the FLASHBI_USER and
# FLASHBI_PASS environment variables before starting the app.
APP_USER = os.environ.get("FLASHBI_USER", "admin")
APP_PASS = os.environ.get("FLASHBI_PASS", "admin")


def check_login(user, pwd):
    """Constant-time comparison of the submitted credentials."""
    ok_u = hmac.compare_digest(user.strip().encode(), APP_USER.encode())
    ok_p = hmac.compare_digest(pwd.encode(), APP_PASS.encode())
    return ok_u and ok_p


def login_screen():
    # Hide the sidebar while logged out
    st.markdown("<style>[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],"
                "[data-testid='collapsedControl']{display:none}</style>", unsafe_allow_html=True)
    left, right = st.columns([1.15, 1], gap="large")
    with left:
        st.markdown('<div class="lp"><div class="blob b1"></div><div class="blob b2"></div><div class="lp-in">'
                    '<div class="login-logo">⚡</div><div class="lp-t">FlashBI</div>'
                    '<div class="lp-sub">Analytics and ETL without the enterprise weight.</div>'
                    '<div class="lp-li">📊 Instant charts from any Excel or CSV file</div>'
                    '<div class="lp-li">🧩 Self-service cleaning and transformation</div>'
                    '<div class="lp-li">🤖 Plain-language insights on your data</div></div></div>',
                    unsafe_allow_html=True)
    with right:
        st.markdown('<div class="login-h">Welcome back 👋</div><div class="login-s">Sign in to open your workspace</div>',
                    unsafe_allow_html=True)
        with st.form("login"):
            user = st.text_input("Username", placeholder="Enter username")
            pwd = st.text_input("Password", type="password", placeholder="Enter password")
            go = st.form_submit_button("Sign in →", **STRETCH)
        if go:
            if check_login(user, pwd):
                st.session_state["auth"] = True
                st.session_state["user"] = user.strip()
                st.session_state["splash"] = True
                st.rerun()
            time.sleep(1)  # slows down password guessing
            st.error("Incorrect username or password.")


if not st.session_state.get("auth"):
    login_screen()
    st.stop()

# Splash spinner shown once, right after signing in
if st.session_state.pop("splash", False):
    _sp = st.empty()
    _sp.markdown('<div class="splash"><div class="ring big"></div><div class="sp-t">Preparing your workspace…</div>'
                 '<div class="sp-bar"></div></div>', unsafe_allow_html=True)
    time.sleep(1.0)
    _sp.empty()


# ───────────────────────── Helpers ─────────────────────────
def loader(ph, msg):
    """Render the animated spinner into a placeholder."""
    ph.markdown(f'<div class="loader"><div class="ring"></div><div class="lt">{html.escape(msg)}</div></div>',
                unsafe_allow_html=True)


def is_text(s):
    return not (is_numeric_dtype(s) or is_datetime64_any_dtype(s) or is_bool_dtype(s))


def col_types(df):
    """Return (numeric, datetime, categorical) column lists (cached on the DataFrame)."""
    key = tuple(map(str, df.columns))
    hit = df.attrs.get("_ct")
    if hit and hit[0] == key:
        return hit[1]
    sample = df if len(df) <= 100_000 else df.iloc[:100_000]
    num = [c for c in df.columns if is_numeric_dtype(df[c]) and not is_bool_dtype(df[c])]
    dt = [c for c in df.columns if is_datetime64_any_dtype(df[c])]
    cat = [c for c in df.columns if c not in num + dt and 1 < sample[c].nunique() <= 50]
    df.attrs["_ct"] = (key, (num, dt, cat))
    return num, dt, cat


def main_num(df, num):
    """First numeric column that does not look like a unique ID."""
    for c in num:
        if len(df) < 20 or df[c].nunique() < len(df) * 0.98:
            return c
    return num[0]


def arrow_safe(df):
    """Make mixed-type columns displayable in st.dataframe."""
    d = df.copy()
    for c in d.columns:
        if is_text(d[c]):
            d[c] = d[c].astype(str)
    return d


def style(fig, h=380):
    fig.update_layout(template=TEMPLATE, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      colorway=PALETTE, height=h, margin=dict(l=20, r=20, t=55, b=20),
                      font=dict(color=T["text"]), title_font_size=16)
    return fig


def show(fig):
    st.plotly_chart(fig, config=CFG, **STRETCH)


def kpis(items):
    for i, (col, (icon, label, val)) in enumerate(zip(st.columns(len(items)), items)):
        col.markdown(f'<div class="kpi" style="animation-delay:{i * .08}s"><div class="ic">{icon}</div>'
                     f'<div class="lb">{html.escape(label)}</div><div class="vl">{html.escape(str(val))}</div></div>',
                     unsafe_allow_html=True)


# ───────────────────────── Loading & auto-cleaning ─────────────────────────
@st.cache_data(show_spinner=False)
def sheet_names(name, data):
    if name.lower().endswith(".csv"):
        return []
    if name.lower().endswith(".xlsx"):
        try:  # read-only mode lists sheets without loading the whole workbook
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception:
            pass
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


def read_csv_fast(data):
    """Detect the delimiter, then use the fastest parser available."""
    first = data[:20000].decode("utf-8", errors="ignore").split("\n", 1)[0]
    sep = max(",;\t|", key=first.count)
    for enc in ("utf-8-sig", "latin-1"):
        for eng in ("pyarrow", "c"):
            try:
                return pd.read_csv(io.BytesIO(data), sep=sep, encoding=enc, engine=eng)
            except UnicodeDecodeError:
                break
            except Exception:
                continue
    raise ValueError("Could not parse this CSV file.")


def read_file(name, data, sheet):
    if name.lower().endswith(".csv"):
        return read_csv_fast(data)
    try:  # python-calamine (optional) reads Excel several times faster
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet or 0, engine="calamine")
    except Exception:
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet or 0)


BLANKS = {"", "nan", "NaN", "None", "NaT", "<NA>", "null", "NULL", "N/A", "n/a"}


def auto_clean(raw):
    """Normalise headers, infer numbers/dates (on unique values only, for speed), fill gaps, drop duplicates."""
    df = raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()].dropna(how="all").dropna(axis=1, how="all")

    for c in df.columns:  # pass 1: type inference
        if not is_text(df[c]):
            continue
        u = pd.Series(df[c].dropna().unique())
        if u.empty:
            continue
        uc = u.astype(str).str.strip()
        uc = uc.where(~uc.isin(BLANKS), np.nan)
        valid = uc.dropna()
        if valid.empty:
            df[c] = np.nan
            continue
        num_u = pd.to_numeric(valid.str.replace(",", "", regex=False), errors="coerce")
        if num_u.notna().mean() >= 0.9:
            df[c] = df[c].map(dict(zip(u[valid.index], num_u))).astype("float64")
            continue
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # try month-first and day-first, keep whichever parses more values
            samp = valid.head(200)
            ratios = [pd.to_datetime(samp, errors="coerce", dayfirst=d).notna().mean() for d in (False, True)]
            if max(ratios) >= 0.8:
                parsed = pd.to_datetime(valid, errors="coerce", dayfirst=bool(ratios[1] > ratios[0]))
                df[c] = pd.to_datetime(df[c].map(dict(zip(u[valid.index], parsed))))
                continue
        df[c] = df[c].map(dict(zip(u, uc)))
    df = df.dropna(axis=1, how="all")

    before_dup = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    dups = before_dup - len(df)
    nulls_before = int(df.isna().sum().sum())

    for c in df.columns:  # pass 2: fill gaps
        if is_numeric_dtype(df[c]) and not is_bool_dtype(df[c]):
            df[c] = df[c].fillna(df[c].median())
        elif is_text(df[c]) and not is_datetime64_any_dtype(df[c]):
            df[c] = df[c].fillna("Unknown").astype(str)
    fixed = nulls_before - int(df.isna().sum().sum())
    return df, {"fixed": fixed, "dups": dups}


@st.cache_resource(show_spinner=False, max_entries=2)
def load_and_clean(name, data, sheet):
    """Parse + clean once per file (cached, no copying of big frames)."""
    raw = read_file(name, data, sheet)
    clean, stats = auto_clean(raw)
    return raw, clean, stats


# ───────────────────────── Slicers ─────────────────────────
def sidebar_filters(df, key):
    num, dt, cat = col_types(df)
    mask = pd.Series(True, index=df.index)
    active = False
    with st.sidebar:
        st.markdown("### 🎛️ Slicers")
        if dt:
            c = dt[0]
            lo, hi = df[c].min(), df[c].max()
            if pd.notna(lo) and lo != hi:
                rng = st.date_input(c, (lo.date(), hi.date()), min_value=lo.date(),
                                    max_value=hi.date(), key=f"d{key}_{c}")
                if isinstance(rng, (tuple, list)) and len(rng) == 2:
                    if rng[0] > lo.date() or rng[1] < hi.date():
                        active = True
                        mask &= df[c].between(pd.Timestamp(rng[0]), pd.Timestamp(rng[1]) + pd.Timedelta(days=1))
        for c in cat:
            opts = sorted(df[c].dropna().unique().tolist(), key=str)
            sel = st.multiselect(c, opts, key=f"f{key}_{c}")
            if sel:
                active = True
                mask &= df[c].isin(sel)
        if not (dt or cat):
            st.caption("No categorical or date columns to slice.")
    return df[mask] if active else df


# ───────────────────────── Auto visuals ─────────────────────────
def trend_series(df, dcol, ncol):
    s = df[[dcol, ncol]].dropna().set_index(dcol)[ncol].sort_index()
    if s.empty:
        return None
    span = (s.index.max() - s.index.min()).days
    return s.resample("D" if span <= 90 else "W" if span <= 730 else "MS").sum()


def auto_visuals(df):
    num, dt, cat = col_types(df)
    n = main_num(df, num) if num else None
    if dt and n:
        s = trend_series(df, dt[0], n)
        if s is not None and len(s) > 1:
            fig = px.line(s.reset_index(), x=dt[0], y=n, title=f"📈 {n} over time")
            fig.update_traces(fill="tozeroy", line=dict(width=3), fillcolor="rgba(124,92,255,.15)")
            fig.update_layout(hovermode="x unified")
            show(style(fig))
    left, right = st.columns(2)
    if cat:
        c = next((x for x in cat if df[x].nunique() <= 20), cat[0])
        g = df.groupby(c)[n].sum() if n else df[c].value_counts()
        g = g.nlargest(10).sort_values().reset_index()
        g.columns = [c, n or "count"]
        fig = px.bar(g, x=g.columns[1], y=c, orientation="h", color=g.columns[1],
                     color_continuous_scale=["#22d3ee", "#7c5cff"],
                     title=f"📊 Top {c}" + (f" by {n}" if n else ""))
        fig.update_layout(coloraxis_showscale=False)
        with left:
            show(style(fig))
    if len(num) >= 2:
        corr = df[num[:12]].corr()
        fig = px.imshow(corr, text_auto=".2f", zmin=-1, zmax=1, aspect="auto",
                        color_continuous_scale="RdBu_r", title="🔗 Correlation matrix")
        with right:
            show(style(fig))
    if not (dt and n) and not cat and len(num) < 2:
        st.info("Not enough numeric, date or categorical columns for automatic charts.")


# ───────────────────────── Business KPIs ─────────────────────────
METRIC_HINTS = ("revenue", "sales", "amount", "total", "profit", "income", "price", "cost", "value", "spend")


def fmt_num(x):
    """Compact number: 1,250,000 -> 1.25M."""
    for div, suf in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if abs(x) >= div:
            return f"{x / div:,.2f}{suf}"
    return f"{x:,.0f}" if abs(x) >= 100 else f"{x:,.2f}"


def guess_metric(df, num):
    """Pick the column most likely to be money/sales, else the first non-ID numeric column."""
    for c in num:
        if any(h in str(c).lower() for h in METRIC_HINTS):
            return c
    return main_num(df, num)


def business_kpis(df, metric):
    """Return [(icon, label, value)] business KPIs for the chosen metric."""
    num, dt, cat = col_types(df)
    out = [("💰", f"Total {metric}", fmt_num(df[metric].sum())),
           ("🧮", f"Average {metric} per row", fmt_num(df[metric].mean()))]
    if cat:
        g = df.groupby(cat[0])[metric].sum().drop("Unknown", errors="ignore").sort_values(ascending=False)
        if len(g) and g.sum() > 0:
            out.append(("🏆", f"Top {cat[0]}", f"{str(g.index[0])[:18]} · {g.iloc[0] / g.sum() * 100:.0f}%"))
    if dt:
        s = trend_series(df, dt[0], metric)
        if s is not None and len(s) >= 2:
            full = s.iloc[:-1] if len(s) >= 3 else s  # the newest period is often incomplete
            if full.iloc[-2] != 0:
                ch = (full.iloc[-1] - full.iloc[-2]) / abs(full.iloc[-2]) * 100
                out.append(("📈" if ch >= 0 else "📉", "Latest full period vs previous", f"{'▲' if ch >= 0 else '▼'} {abs(ch):.1f}%"))
            out.append(("⭐", "Best period", f"{s.idxmax():%d %b %Y}"))
    return out[:5]


# ───────────────────────── Storytelling ─────────────────────────
def pct(a, b):
    """Percent change from b to a (None when undefined)."""
    if b in (0, None) or pd.isna(a) or pd.isna(b):
        return None
    return (a - b) / abs(b) * 100


def story_box(title, lines, accent="#22d3ee"):
    if not lines:
        return
    body = "".join(f'<div class="sl">{html.escape(str(l))}</div>' for l in lines)
    st.markdown(f'<div class="story" style="border-left-color:{accent}"><div class="stt">{title}</div>{body}</div>',
                unsafe_allow_html=True)


def current_metric(df):
    """The metric chosen on the Overview tab (or a sensible guess)."""
    num, _, _ = col_types(df)
    if not num:
        return None
    v = st.session_state.get("metric_sel")
    return v if v in num else guess_metric(df, num)


def story_overview(df, metric, stats):
    """Narrative + recommended actions for the selected data."""
    num, dt, cat = col_types(df)
    story, acts = [], []
    span = f" from {df[dt[0]].min():%d %b %Y} to {df[dt[0]].max():%d %b %Y}" if dt and df[dt[0]].notna().any() else ""
    story.append(f"Your data holds {len(df):,} records across {df.shape[1]} columns{span}.")
    total, mean, med = df[metric].sum(), df[metric].mean(), df[metric].median()
    line = f"{metric} totals {fmt_num(total)}, averaging {fmt_num(mean)} per record"
    if med > 0 and mean > 1.25 * med:
        line += f", yet the typical (median) record is only {fmt_num(med)}, so a few very large records pull the average up."
        acts.append(f"A handful of large records drive {metric}. Review the biggest ones and decide whether they repeat or are one-offs.")
    else:
        line += f", with a median of {fmt_num(med)}, so values are fairly evenly spread."
    story.append(line)
    if dt:
        s = trend_series(df, dt[0], metric)
        if s is not None and len(s) >= 4:
            h = len(s) // 2
            ch = pct(s.iloc[h:].sum(), s.iloc[:h].sum())
            if ch is not None:
                story.append(f"Over time, the second half of the period {'grew' if ch >= 0 else 'fell'} {abs(ch):.0f}% compared with the first half. "
                             f"The strongest period was {s.idxmax():%d %b %Y} ({fmt_num(s.max())}) and the weakest {s.idxmin():%d %b %Y} ({fmt_num(s.min())}).")
                if ch <= -10:
                    acts.append(f"{metric} is declining. Compare the recent period with the earlier one to find what changed.")
                elif ch >= 10:
                    acts.append(f"{metric} is growing. Identify what drove the strongest periods and repeat it.")
    if cat:
        c = cat[0]
        g = df.groupby(c)[metric].sum().drop("Unknown", errors="ignore").sort_values(ascending=False)
        if len(g) >= 2 and g.sum() > 0:
            top = g.iloc[0] / g.sum() * 100
            story.append(f"By {c}, “{g.index[0]}” contributes {top:.0f}% of {metric}, while “{g.index[-1]}” adds only "
                         f"{g.iloc[-1] / g.sum() * 100:.0f}%. " + (f"The top 3 together make up {g.iloc[:3].sum() / g.sum() * 100:.0f}%." if len(g) > 4 else ""))
            if top > 50:
                acts.append(f"Over half of {metric} depends on one {c}. Protect that segment and grow the others to reduce risk.")
            elif len(g) >= 4:
                acts.append(f"“{g.index[-1]}” is your weakest {c}. Check whether it needs more attention or should be deprioritised.")
    q1, q3 = df[metric].quantile([.25, .75])
    if q3 > q1:
        out = int(((df[metric] < q1 - 1.5 * (q3 - q1)) | (df[metric] > q3 + 1.5 * (q3 - q1))).sum())
        if out and out / len(df) > 0.01:
            acts.append(f"{out:,} records have unusually high or low {metric}. Check them for entry errors or special events.")
    if stats["fixed"] or stats["dups"]:
        story.append(f"Auto-clean filled {stats['fixed']:,} empty cells and removed {stats['dups']:,} duplicate row(s) before this analysis.")
        if stats["fixed"]:
            acts.append("Filled values are estimates. Fix the empty cells at the source for fully reliable numbers.")
    if not acts:
        acts.append("No warning signs found. Use the Data Analytics tab to dig into specific segments.")
    return story, acts


def selection_story(work, sel, metric):
    """Compare the slicer-selected rows with the unselected rest."""
    if len(sel) == len(work):
        return [f"No slicer is active, so all {len(work):,} records are selected. Pick a segment in the sidebar "
                "(a region, product or date range) and I will compare it with the unselected rest."]
    if sel.empty:
        return ["Your slicers select 0 records. Loosen a filter to see the story."]
    rest = work.loc[~work.index.isin(sel.index)]
    out = [f"You selected {len(sel):,} of {len(work):,} records ({len(sel) / len(work) * 100:.0f}%); the other {len(rest):,} are unselected."]
    num, dt, cat = col_types(work)
    if metric and not rest.empty:
        a, b = sel[metric].sum(), rest[metric].sum()
        if a + b:
            out.append(f"The selection holds {fmt_num(a)} of {metric}, which is {a / (a + b) * 100:.0f}% of the total, against {fmt_num(b)} in the unselected data.")
        d = pct(sel[metric].mean(), rest[metric].mean())
        if d is not None:
            out.append(f"Per record, the selection averages {fmt_num(sel[metric].mean())} versus {fmt_num(rest[metric].mean())} for the rest, "
                       f"{abs(d):.0f}% {'higher' if d >= 0 else 'lower'}. " +
                       ("This segment performs about the same as the rest." if abs(d) < 5 else
                        f"This segment is {'stronger' if d >= 0 else 'weaker'} than the rest."))
    for c in cat[:2]:
        if rest.empty:
            break
        f = (lambda d_: d_.groupby(c)[metric].sum() if metric else d_[c].value_counts())
        a, b = f(sel).drop("Unknown", errors="ignore"), f(rest).drop("Unknown", errors="ignore")
        if len(a) and len(b):
            out.append(f"Within the selection “{a.idxmax()}” leads {c}; " +
                       ("the unselected data has the same leader." if a.idxmax() == b.idxmax() else f"in the unselected data “{b.idxmax()}” leads."))
    if dt and sel[dt[0]].notna().any():
        out.append(f"The selected records span {sel[dt[0]].min():%d %b %Y} to {sel[dt[0]].max():%d %b %Y}.")
    return out


def profile_story(df):
    """Plain-language description of the dataset view being inspected."""
    num, dt, cat = col_types(df)
    other = df.shape[1] - len(num) - len(dt) - len(cat)
    out = [f"This view has {len(df):,} rows and {df.shape[1]} columns: {len(num)} numeric, {len(dt)} date, "
           f"{len(cat)} category-style and {other} free-text or ID columns."]
    nulls = df.isna().sum()
    if nulls.sum() == 0:
        out.append("Every cell is filled, so this data is 100% complete.")
    else:
        w = nulls.idxmax()
        out.append(f"{nulls.sum() / df.size * 100:.1f}% of cells are empty. The gappiest column is “{w}” ({nulls[w] / len(df) * 100:.0f}% empty).")
    const = [c for c in df.columns if df[c].nunique(dropna=True) <= 1]
    if const:
        out.append(f"{', '.join(map(str, const[:3]))} hold a single value, so they add no information. Consider dropping them in the ETL tab.")
    ids = [c for c in df.columns if len(df) > 20 and df[c].nunique() == len(df) and not is_datetime64_any_dtype(df[c])]
    if ids:
        out.append(f"{', '.join(map(str, ids[:3]))} look like unique identifiers, which are not useful for grouping.")
    sk = {c: df[c].skew() for c in num if df[c].nunique() > 5}
    sk = {c: v for c, v in sk.items() if pd.notna(v)}
    if sk:
        c = max(sk, key=lambda k: abs(sk[k]))
        if abs(sk[c]) > 1:
            out.append(f"{c} is strongly skewed towards {'high' if sk[c] > 0 else 'low'} values, so its average can mislead. Compare it with the median.")
    return out


def chart_story(data, kind, x, y):
    """Describe what the chart built in the Data Analytics tab shows."""
    out = []
    lab = lambda v: f"{v:%d %b %Y}" if isinstance(v, pd.Timestamp) else str(v)
    try:
        if kind == "Scatter":
            if is_numeric_dtype(data[x]) and is_numeric_dtype(data[y]):
                r = data[x].corr(data[y])
                if abs(r) < 0.2:
                    out.append(f"{x} and {y} show almost no linear relationship (r = {r:+.2f}).")
                else:
                    w = "strong" if abs(r) >= .7 else "moderate" if abs(r) >= .4 else "weak"
                    out.append(f"{x} and {y} show a {w} {'positive' if r > 0 else 'negative'} relationship (r = {r:+.2f}): "
                               f"as {x} rises, {y} tends to {'rise' if r > 0 else 'fall'}.")
            return out
        if not is_numeric_dtype(data[y]):
            return out
        if kind in ("Line", "Area"):
            g = data.groupby(x)[y].sum().sort_index()
            if len(g) >= 2:
                ch = pct(g.iloc[-1], g.iloc[0])
                out.append(f"From {lab(g.index[0])} to {lab(g.index[-1])}, {y} goes from {fmt_num(g.iloc[0])} to {fmt_num(g.iloc[-1])}"
                           + (f" ({'up' if ch >= 0 else 'down'} {abs(ch):.0f}%)." if ch is not None else "."))
                out.append(f"The peak is at {lab(g.idxmax())} ({fmt_num(g.max())}) and the lowest point at {lab(g.idxmin())} ({fmt_num(g.min())}).")
        else:
            g = data.groupby(x)[y].sum().sort_values(ascending=False)
            if len(g) >= 2 and g.sum() > 0:
                out.append(f"“{g.index[0]}” leads {x} with {fmt_num(g.iloc[0])} ({g.iloc[0] / g.sum() * 100:.0f}% of the total); "
                           f"“{g.index[-1]}” is last with {fmt_num(g.iloc[-1])}.")
                if g.iloc[-1] > 0:
                    out.append(f"The top item is {g.iloc[0] / g.iloc[-1]:.1f}× the smallest, and the top 3 together account for {g.iloc[:3].sum() / g.sum() * 100:.0f}%.")
    except Exception:
        pass
    return out


# ───────────────────────── Tab 1: inspector ─────────────────────────
@fragment
def tab_inspect(work, sel=None):
    ss = st.session_state
    view = st.radio("Dataset view", ["Current (after ETL)", "Auto-cleaned", "Original upload"], horizontal=True)
    df = {"Current (after ETL)": work, "Auto-cleaned": ss["clean"], "Original upload": ss["raw"]}[view]
    num, dt, cat = col_types(df)
    kpis([("🔢", "Numeric columns", len(num)), ("🏷️", "Categorical columns", len(cat)),
          ("📅", "Date columns", len(dt)), ("💾", "Memory", f"{df.memory_usage(deep=len(df) <= 200_000).sum() / 1e6:.2f} MB")])
    story_box("📖 Story of this dataset view", profile_story(df))
    if sel is not None:
        story_box("🧭 Selected vs unselected data", selection_story(work, sel, current_metric(work)), "#f472b6")
    st.markdown("#### Data")
    st.dataframe(arrow_safe(df.head(1000)), height=380, **STRETCH)
    if len(df) > 1000:
        st.caption(f"Showing the first 1,000 of {len(df):,} rows to keep the app fast. Use the ETL tab to download everything.")
    st.markdown("#### Profile")
    prof = pd.DataFrame({"Column": df.columns, "Type": [str(t) for t in df.dtypes],
                         "Non-null": df.notna().sum().values, "Nulls": df.isna().sum().values,
                         "Null %": (df.isna().mean() * 100).round(1).values, "Unique": (df if len(df) <= 100_000 else df.iloc[:100_000]).nunique().values})
    st.dataframe(prof, hide_index=True, **STRETCH)
    if num:
        with st.expander("Numeric summary"):
            st.dataframe(df[num].describe().T, **STRETCH)


# ───────────────────────── Tab 2: chart builder ─────────────────────────
@fragment
def tab_builder(df, work=None):
    if df.empty:
        st.warning("No rows match the current slicers.")
        return
    num, dt, cat = col_types(df)
    cols = df.columns.tolist()
    a, b, c, d = st.columns(4)
    kind = a.selectbox("Chart type", ["Bar", "Line", "Scatter", "Pie", "Area"])
    x = b.selectbox("X-axis / category", cols)
    y = c.selectbox("Y-axis / value", num or cols)
    agg = d.selectbox("Aggregation", ["sum", "mean", "median", "count", "max", "min", "none"],
                      disabled=(kind == "Scatter"))
    e, f = st.columns(2)
    color = e.selectbox("Color / group by", ["None"] + [k for k in cat if k != x])
    top_n = f.slider("Max categories (Bar/Pie)", 3, 50, 15)
    color = None if color == "None" else color

    if kind == "Scatter":
        agg = "none"
    if agg not in ("count", "none") and not is_numeric_dtype(df[y]):
        st.warning("Pick a numeric Y-axis for this aggregation (or use count).")
        return
    if x == y and agg != "none":
        st.info("Choose different X and Y columns.")
        return

    data = df
    if agg != "none":
        keys = [x] + ([color] if color and kind != "Pie" else [])
        data = df.groupby(keys)[y].agg(agg).reset_index()
        if kind in ("Line", "Area"):
            data = data.sort_values(x)
        else:
            data = data.sort_values(y, ascending=False)
            if kind in ("Bar", "Pie") and not color:
                data = data.head(top_n)
    elif kind in ("Line", "Area"):
        data = df.sort_values(x)
    if len(data) > 20000 and kind == "Scatter":
        data = data.sample(20000, random_state=1)

    title = f"{agg + ' of ' if agg != 'none' else ''}{y} by {x}"
    if kind == "Bar":
        fig = px.bar(data, x=x, y=y, color=color, barmode="group", title=title)
    elif kind == "Line":
        fig = px.line(data, x=x, y=y, color=color, markers=True, title=title)
    elif kind == "Area":
        fig = px.area(data, x=x, y=y, color=color, title=title)
    elif kind == "Scatter":
        fig = px.scatter(data, x=x, y=y, color=color, opacity=0.75, title=title)
    else:
        fig = px.pie(data, names=x, values=y, hole=0.45, title=title)
    show(style(fig, 460))
    story_box("📖 What this chart says", chart_story(data, kind, x, y))
    if work is not None:
        story_box("🧭 Selected vs unselected data", selection_story(work, df, current_metric(work)), "#f472b6")
    with st.expander("Chart data"):
        st.dataframe(arrow_safe(data), **STRETCH)


# ───────────────────────── Tab 3: ETL ─────────────────────────
OPS = {"==": operator.eq, "!=": operator.ne, ">": operator.gt, ">=": operator.ge,
       "<": operator.lt, "<=": operator.le}


def apply_filter(df, col, op, val):
    s = df[col]
    if op == "contains":
        return df[s.astype(str).str.contains(val, case=False, na=False, regex=False)]
    try:
        v = float(val) if is_numeric_dtype(s) else pd.to_datetime(val) if is_datetime64_any_dtype(s) else val
        return df[OPS[op](s, v)]
    except Exception:
        raise ValueError(f"Cannot compare '{col}' {op} '{val}'.")


def run_etl(base, cfg, miss, dedupe, conds):
    df = base.copy()
    for col, op, val in conds:
        if val.strip():
            df = apply_filter(df, col, op, val.strip())
    df = df.copy()
    if miss == "Drop rows with nulls":
        df = df.dropna()
    elif miss != "Leave as is":
        nums = [c for c in df.columns if is_numeric_dtype(df[c]) and not is_bool_dtype(df[c])]
        if "0" in miss:
            df[nums] = df[nums].fillna(0)
        elif "mean" in miss:
            df[nums] = df[nums].fillna(df[nums].mean())
        else:
            df[nums] = df[nums].fillna(df[nums].median())
        for c in df.columns:
            if is_text(df[c]) and not is_datetime64_any_dtype(df[c]):
                df[c] = df[c].fillna("Unknown")
    if dedupe:
        df = df.drop_duplicates()
    keep = cfg[cfg["Keep"]]
    if keep.empty:
        raise ValueError("Keep at least one column.")
    names = keep["Rename to"].fillna("").astype(str).str.strip().tolist()
    if any(not n for n in names) or len(set(names)) != len(names):
        raise ValueError("New column names must be non-empty and unique.")
    df = df[keep["Column"].tolist()]
    df.columns = names
    if df.empty:
        raise ValueError("Your filters removed every row.")
    return df.reset_index(drop=True)


def to_xlsx(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


@fragment
def tab_etl(base, work):
    ss = st.session_state
    fid = ss["fid"]
    st.markdown("#### 1 · Columns")
    st.caption("Untick Keep to drop a column; edit “Rename to” to rename it.")
    cfg = pd.DataFrame({"Keep": True, "Column": base.columns.tolist(), "Rename to": base.columns.tolist()})
    cfg = st.data_editor(cfg, key=f"cols{fid}", hide_index=True, disabled=["Column"], **STRETCH)

    st.markdown("#### 2 · Missing values")
    a, b = st.columns(2)
    miss = a.selectbox("Strategy", ["Leave as is", "Drop rows with nulls", "Fill with 0",
                                    "Fill with mean", "Fill with median"], key=f"miss{fid}")
    dedupe = b.checkbox("Remove duplicate rows", key=f"dd{fid}")

    st.markdown("#### 3 · Row filters")
    n = st.number_input("Number of conditions", 0, 5, 0, key=f"nc{fid}")
    conds = []
    for i in range(int(n)):
        c1, c2, c3 = st.columns([2, 1, 2])
        col = c1.selectbox("Column", base.columns.tolist(), key=f"fc{fid}_{i}")
        op = c2.selectbox("Operator", [*OPS, "contains"], key=f"fo{fid}_{i}")
        val = c3.text_input("Value", key=f"fv{fid}_{i}")
        conds.append((col, op, val))

    b1, b2 = st.columns(2)
    if b1.button("⚡ Apply pipeline", **STRETCH):
        ph = st.empty()
        loader(ph, "Running ETL pipeline…")
        try:
            out = run_etl(base, cfg, miss, dedupe, conds)
        except Exception as e:
            ph.empty()
            st.error(f"Pipeline failed: {e}")
            return
        ss["etl"] = out
        ss["ev"] += 1
        st.rerun()
    if b2.button("↩️ Reset to auto-cleaned", **STRETCH):
        ss["etl"], ss["fid"], ss["ev"] = None, fid + 1, ss["ev"] + 1
        st.rerun()

    st.markdown("#### Result")
    if ss["etl"] is not None:
        st.success(f"Pipeline applied: {len(base):,} → {len(work):,} rows, {work.shape[1]} columns. "
                   "Every chart and insight now uses this data.")
    else:
        st.info("Showing auto-cleaned data. Apply the pipeline to reshape it.")
    st.dataframe(arrow_safe(work.head(200)), height=300, **STRETCH)
    st.markdown("#### Download")
    stamp = (len(work), tuple(map(str, work.columns)), ss["ev"])
    if st.button("📦 Prepare download files", **STRETCH):
        ph = st.empty()
        loader(ph, "Preparing your files…")
        ss["dl"] = dict(stamp=stamp, csv=work.to_csv(index=False).encode("utf-8"),
                        xlsx=to_xlsx(work) if len(work) < 1_000_000 else None)
        ph.empty()
    dl = ss.get("dl")
    if dl and dl["stamp"] == stamp:
        d1, d2 = st.columns(2)
        if dl["xlsx"]:
            d1.download_button("⬇️ Download cleaned .xlsx", dl["xlsx"], "cleaned_data.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **STRETCH)
        else:
            d1.caption("Too many rows for one Excel sheet. Use the CSV.")
        d2.download_button("⬇️ Download cleaned .csv", dl["csv"], "cleaned_data.csv", "text/csv", **STRETCH)


# ───────────────────────── Tab 4: AI insights ─────────────────────────
def build_insights(df, stats):
    num, dt, cat = col_types(df)
    out = []
    n = main_num(df, num) if num else None
    if dt and n:
        s = trend_series(df, dt[0], n)
        if s is not None and len(s) >= 3:
            chg = (s.iloc[-1] - s.iloc[0]) / abs(s.iloc[0]) * 100 if s.iloc[0] else None
            txt = f"{n} peaks on {s.idxmax():%d %b %Y} at {s.max():,.0f}, with the lowest point on {s.idxmin():%d %b %Y}."
            if chg is not None:
                txt += f" The latest period is {chg:+.1f}% versus the first."
            out.append(("📈", "Trend", txt))
    if cat:
        c = cat[0]
        g = (df.groupby(c)[n].sum() if n else df[c].value_counts()).drop("Unknown", errors="ignore")
        g = g.sort_values(ascending=False)
        if len(g) >= 2 and g.sum() > 0:
            out.append(("🏆", "Top performer",
                        f"“{g.index[0]}” leads {c} with {g.iloc[0] / g.sum() * 100:.0f}% of total "
                        f"{n or 'rows'}, ahead of “{g.index[1]}” at {g.iloc[1] / g.sum() * 100:.0f}%."))
    best = None
    for c in num:
        q1, q3 = df[c].quantile([.25, .75])
        iqr = q3 - q1
        if iqr > 0:
            cnt = int(((df[c] < q1 - 1.5 * iqr) | (df[c] > q3 + 1.5 * iqr)).sum())
            if cnt and (best is None or cnt > best[1]):
                best = (c, cnt)
    if best:
        out.append(("🚨", "Anomalies", f"{best[0]} has {best[1]:,} outliers ({best[1] / len(df) * 100:.1f}% of rows) "
                                       "outside the normal IQR range. Worth checking for errors or special events."))
    if len(num) >= 2:
        corr = df[num[:12]].corr().abs()
        pair = corr.where(np.triu(np.ones(corr.shape, dtype=bool), 1)).stack()
        if not pair.empty and pair.max() >= 0.5:
            a, b = pair.idxmax()
            r = df[a].corr(df[b])
            out.append(("🔗", "Relationship", f"{a} and {b} move {'together' if r > 0 else 'in opposite directions'} "
                                              f"(correlation {r:+.2f})."))
    out.append(("🧹", "Data quality", f"{len(df):,} rows × {df.shape[1]} columns. Auto-clean filled "
                                      f"{stats['fixed']:,} missing values and removed {stats['dups']:,} duplicates."))
    return out[:4]


@fragment
def tab_insights(df, stats):
    if df.empty:
        st.warning("No rows match the current slicers.")
        return
    items = build_insights(df, stats)
    for i, (icon, title, text) in enumerate(items):
        st.markdown(f'<div class="insight" style="animation-delay:{i * .12}s"><div class="tt">{icon} {title}</div>'
                    f'<div class="bd">{html.escape(text)}</div></div>', unsafe_allow_html=True)
    st.caption("Insights are generated from the currently filtered data using statistical rules, not an external model.")


# ───────────────────────── Smart suggestions (sidebar) ─────────────────────────
def suggestions(df, metric):
    """Data-driven ideas to grow the metric, built from the segments found in THIS dataset."""
    num, dt, cat = col_types(df)
    if not metric or not cat or len(df) < 5:
        return []
    out, n, avg_all = [], len(df), df[metric].mean()
    first = True
    for c in cat[:2]:
        g = df.groupby(c)[metric].agg(["sum", "count", "mean"]).drop("Unknown", errors="ignore")
        if len(g) < 2 or g["sum"].sum() <= 0:
            continue
        g = g.sort_values("sum", ascending=False)
        share = g["sum"] / g["sum"].sum() * 100
        best, worst = g.index[0], g.index[-1]
        out.append((f"🏆 Scale “{best}”", f"Best {c}, with {share.iloc[0]:.0f}% of {metric}. " +
                    ("That is a lot in one place, so grow the others to reduce dependence." if share.iloc[0] > 50 else
                     "Study what works here (offers, pricing, channels) and copy it to the other segments.")))
        if first:
            lag = g[g["mean"] < avg_all]
            up = ((avg_all - lag["mean"]) * lag["count"]).sum()
            if up > 0:
                out.append(("🎯 Biggest opportunity", f"If every below-average {c} reached the average {metric} per record "
                            f"({fmt_num(avg_all)}), the total could rise by about {fmt_num(up)} ({up / g['sum'].sum() * 100:.0f}%)."))
            first = False
        vol, val = g["count"].median(), g["mean"].median()
        hv = g[(g["count"] >= vol) & (g["mean"] < val) & (g["mean"] < avg_all)]
        if len(hv):
            gap = (avg_all - hv["mean"]) * hv["count"]
            sg = gap.idxmax()
            out.append((f"💵 Raise value in “{sg}”", f"{c} “{sg}” has many records but earns only {fmt_num(g.loc[sg, 'mean'])} each "
                        f"versus {fmt_num(avg_all)} on average. Lifting it could add about {fmt_num(gap[sg])}. Try pricing, bundles or upselling."))
        lv = g[(g["count"] < vol) & (g["mean"] > val) & (g["mean"] > avg_all)]
        if len(lv):
            sg = lv["mean"].idxmax()
            out.append((f"📣 Grow volume in “{sg}”", f"{c} “{sg}” earns {fmt_num(g.loc[sg, 'mean'])} per record, above average, "
                        "but has few records. Increase reach there: marketing, stock, outlets or sales staff."))
        out.append((f"🔧 Fix “{worst}”", f"Lowest {c}, with {share.iloc[-1]:.0f}% of {metric}. Compare it with “{best}”: check pricing, "
                    "availability, staffing, campaigns and customer mix before investing more or scaling down."))
        if dt:
            span = (df[dt[0]].max() - df[dt[0]].min()).days
            p = df.groupby([pd.Grouper(key=dt[0], freq="D" if span <= 90 else "W" if span <= 730 else "MS"), c])[metric].sum().unstack(fill_value=0)
            if len(p) >= 4:
                p = p.iloc[:-1]  # newest period is often incomplete
                ch = ((p.iloc[-1] - p.iloc[-2]) / p.iloc[-2].replace(0, np.nan) * 100).dropna().drop("Unknown", errors="ignore")
                if len(ch) and ch.min() < -10:
                    out.append((f"⚠️ “{ch.idxmin()}” is slipping", f"{c} “{ch.idxmin()}” fell {abs(ch.min()):.0f}% versus the previous period. "
                                "Find the cause (lost customers, stock-outs, pricing, competitors) and act early."))
                if len(ch) and ch.max() > 10:
                    out.append((f"🚀 “{ch.idxmax()}” is rising", f"{c} “{ch.idxmax()}” grew {ch.max():.0f}% versus the previous period. "
                                "Back it with more budget and stock while the momentum lasts."))
    if len(cat) >= 2:
        c0, c1 = cat[:2]
        gg = df.groupby([c0, c1])[metric].agg(["mean", "count"])
        gg = gg[gg["count"] >= max(5, n * 0.01)].drop("Unknown", level=0, errors="ignore")
        if len(gg) and avg_all:
            a, b = gg["mean"].idxmax()
            out.append(("🧩 Best combination", f"{c0} “{a}” with {c1} “{b}” earns {fmt_num(gg['mean'].max())} per record "
                        f"({gg['mean'].max() / avg_all:.1f}× the average). Replicate this mix elsewhere."))
    return out[:9]


def sidebar_suggestions(df, metric):
    ss = st.session_state
    st.markdown("### 💡 Smart suggestions")
    ck = (metric, ss.get("fid"), ss.get("ev"))
    if ss.get("sug_key") != ck:
        try:
            ss["sug_items"] = suggestions(df, metric)
        except Exception:
            ss["sug_items"] = []
        ss["sug_key"] = ck
    if not ss["sug_items"]:
        st.caption("Suggestions appear when the data has a numeric metric and a category column.")
        return
    st.caption(f"Ideas to improve “{metric}”, generated from your data.")
    for t, x in ss["sug_items"]:
        st.markdown(f'<div class="sug"><b>{html.escape(t)}</b>{html.escape(x)}</div>', unsafe_allow_html=True)


# ───────────────────────── Ask your data (search bar) ─────────────────────────
KIND_WORDS = [("Pie", ("pie", "donut", "doughnut", "share", "proportion")), ("Scatter", ("scatter", "correlat", "relationship")),
              ("Area", ("area",)), ("Line", ("line", "trend", "over time", "growth")),
              ("Bar", ("bar", "column chart", "compare", "comparison", "rank", "histogram")), ("Table", ("table", "list", "rows"))]
AGG_WORDS = [("mean", ("average", "avg", "mean")), ("median", ("median",)), ("max", ("maximum", "max")),
             ("min", ("minimum", "min")), ("count", ("count", "how many", "number of"))]
AGG_NAME = {"sum": "total", "mean": "average", "median": "median", "max": "maximum", "min": "minimum", "count": "number of records"}
SYN = [{"sales", "revenue", "income", "turnover", "amount", "earning"}, {"profit", "margin"},
       {"units", "quantity", "qty", "volume", "sold"}, {"cost", "expense", "spend"}]
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
KINDS = {"Bar", "Line", "Area", "Pie", "Scatter", "Table", "Number"}


def norm_words(t):
    return {w[:-1] if w.endswith("s") and len(w) > 3 else w for w in re.findall(r"[a-z0-9]+", str(t).lower())}


def match_cols(q, cols):
    """Columns mentioned in the question (by name, plural form or synonym), best match first."""
    qw, hits = norm_words(q), []
    for c in cols:
        cw = norm_words(c)
        score = (2 if str(c).lower() in q else 0) + len(cw & qw) + sum(1 for g in SYN if (qw & g) and (cw & g))
        if score:
            hits.append((score, c))
    return [c for _, c in sorted(hits, key=lambda t: -t[0])]


def parse_query(q, df, metric):
    """Rule-based understanding of a plain-English question -> chart spec."""
    ql = q.lower().strip()
    num, dt, cat = col_types(df)
    sp = dict(kind=None, y=metric, x=None, agg="sum", filters={}, top_n=None, ascending=False, grain=None,
              year=None, month=None, understood=False)
    for a, ws in AGG_WORDS:
        if any(re.search(r"\b" + w, ql) for w in ws):
            sp["agg"], sp["understood"] = a, True
            break
    m = match_cols(ql, num)
    if m:
        sp["y"], sp["understood"] = m[0], True
    for c in cat:  # values named in the question become filters
        vals = [v for v in df[c].dropna().unique()
                if len(str(v)) > 1 and re.search(r"\b" + re.escape(str(v).lower()) + r"\b", ql)]
        if vals:
            sp["filters"][c], sp["understood"] = vals, True
    for k, ws in KIND_WORDS:
        if any(re.search(r"\b" + w, ql) for w in ws):
            sp["kind"], sp["understood"] = k, True
            break
    yr = re.search(r"\b((?:19|20)\d{2})\b", ql)
    if yr:
        sp["year"], sp["understood"] = int(yr.group(1)), True
    mo = re.search(r"\b(?:in|of|during|for)\s+(" + "|".join(MONTHS) + r")\b", ql)
    if mo:
        sp["month"], sp["understood"] = MONTHS.index(mo.group(1)) + 1, True
    g = re.search(r"\b(daily|weekly|monthly|quarterly|yearly|annual)\b|\b(?:by|per|each)\s+(day|week|month|quarter|year)\b", ql)
    if g:
        w = (g.group(1) or g.group(2))[0].upper()
        sp["grain"] = {"A": "Y"}.get(w, w)
        sp["understood"] = True
    # ranking words
    n = re.search(r"\b(?:top|best|first|highest|bottom|worst|lowest)\s+(\d+)", ql)
    asc = bool(re.search(r"\b(bottom|worst|lowest|least|minimum|smallest)\b", ql))
    top = bool(re.search(r"\b(top|best|highest|most|maximum|largest|biggest)\b", ql))
    question = bool(re.search(r"\b(which|what|who|where)\b", ql))
    if n:
        sp["top_n"], sp["ascending"], sp["understood"] = int(n.group(1)), asc, True
    elif question and (asc or top):
        sp["top_n"], sp["ascending"], sp["understood"] = 1, asc, True
    # what to group by
    cols = match_cols(ql, cat + dt)
    if cols:
        sp["x"] = cols[0]
    elif sp["filters"]:
        sp["x"] = next(iter(sp["filters"]))
    elif sp["kind"] in ("Line", "Area") or sp["grain"]:
        sp["x"] = dt[0] if dt else None
    elif (sp["kind"] in ("Pie", "Bar") or sp["top_n"]) and cat:
        sp["x"] = cat[0]
    if sp["kind"] == "Scatter" and len(m) >= 2:
        sp["x"], sp["y"] = m[0], m[1]
    if sp["kind"] is None:
        sp["kind"] = "Number" if sp["x"] is None else ("Line" if sp["x"] in dt else "Bar")
    if re.search(r"\b(how many|number of|count)\b", ql) and not m:
        sp["agg"] = "count"
    return sp


def sanitize(sp, df):
    """Make sure a spec (from rules or from an LLM) only refers to real columns and values."""
    num, dt, cat = col_types(df)
    sp = dict(sp)
    sp["kind"] = sp.get("kind") if sp.get("kind") in KINDS else "Bar"
    sp["agg"] = sp.get("agg") if sp.get("agg") in AGG_NAME else "sum"
    if sp.get("y") not in num:
        sp["y"] = num[0] if num else None
        if not num:
            sp["agg"] = "count"
    if sp.get("x") not in df.columns:
        sp["x"] = None
    fl = {}
    for c, vals in (sp.get("filters") or {}).items():
        if c in df.columns:
            lookup = {str(v).lower(): v for v in df[c].dropna().unique()}
            ok = [lookup[str(v).lower()] for v in (vals if isinstance(vals, list) else [vals]) if str(v).lower() in lookup]
            if ok:
                fl[c] = ok
    sp["filters"] = fl
    sp["grain"] = sp.get("grain") if sp.get("grain") in ("D", "W", "M", "Q", "Y") else None
    for k in ("top_n", "year", "month"):
        try:
            sp[k] = int(sp[k]) if sp.get(k) else None
        except (TypeError, ValueError):
            sp[k] = None
    sp["ascending"] = bool(sp.get("ascending"))
    return sp


def llm_spec(q, df):
    """Optional: let Claude interpret the question (only column names and category values are sent, never rows)."""
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return None
    try:
        import anthropic
        num, dt, cat = col_types(df)
        schema = {"numeric": num, "date": dt, "categories": {c: [str(v) for v in df[c].dropna().unique()[:20]] for c in cat}}
        system = ("You turn a business question into a chart spec for a dataset. Reply with ONE JSON object only, keys: "
                  "kind (Bar|Line|Area|Pie|Scatter|Table|Number), y (numeric column), x (column to group by or null), "
                  "agg (sum|mean|median|max|min|count), filters (object: column -> list of values), top_n (int or null), "
                  "ascending (bool), grain (D|W|M|Q|Y or null), year (int or null), month (1-12 or null). Use only the columns given.")
        msg = anthropic.Anthropic(api_key=key).messages.create(
            model="claude-haiku-4-5-20251001", max_tokens=400, system=system,
            messages=[{"role": "user", "content": f"Columns: {json.dumps(schema)}\nQuestion: {q}"}])
        spec = json.loads(re.search(r"\{.*\}", msg.content[0].text, re.S).group(0))
        spec = sanitize(spec, df)
        spec["understood"] = True
        return spec
    except Exception:
        return None


def describe_spec(sp):
    what = "records" if sp["agg"] == "count" else f"{AGG_NAME[sp['agg']]} {sp['y']}"
    txt = f"{sp['kind']} of {what}" + (f" by {sp['x']}" if sp["x"] else "")
    if sp["kind"] == "Scatter" and sp["x"]:
        txt = f"Scatter of {sp['y']} vs {sp['x']}"
    cond = [f"{c} = {', '.join(map(str, v))}" for c, v in sp["filters"].items()]
    cond += [f"year {sp['year']}"] if sp["year"] else []
    cond += [f"month {sp['month']}"] if sp["month"] else []
    return txt + (" where " + "; ".join(cond) if cond else "") + (f", top {sp['top_n']}" if sp["top_n"] else "")


def run_query(df, sp):
    """Execute a spec and return {text, fig, table, big, label}."""
    num, dt, cat = col_types(df)
    d = df
    for c, vals in sp["filters"].items():
        d = d[d[c].isin(vals)]
    if dt and sp["year"]:
        d = d[d[dt[0]].dt.year == sp["year"]]
    if dt and sp["month"]:
        d = d[d[dt[0]].dt.month == sp["month"]]
    if d.empty:
        return {"text": ["No records match those conditions."]}
    y, agg, x, kind = sp["y"], sp["agg"], sp["x"], sp["kind"]
    yl = "Records" if agg == "count" else y
    what = "records" if agg == "count" else f"{AGG_NAME[agg]} {y}"
    if kind == "Table":
        return {"text": [f"Showing {min(len(d), 500):,} of {len(d):,} matching records."], "table": d.head(500)}
    if kind == "Scatter" and x in num and y in num:
        d = d.sample(20000, random_state=1) if len(d) > 20000 else d
        fig = px.scatter(d, x=x, y=y, opacity=.75, title=f"{y} vs {x}")
        return {"text": chart_story(d, "Scatter", x, y), "fig": style(fig, 440), "table": d[[x, y]].head(500)}
    if x is None or kind == "Number":
        val = len(d) if agg == "count" else getattr(d[y], agg)()
        text = (f"There are {len(d):,} matching records." if agg == "count" else
                f"The {what} is {fmt_num(val)}" + (f" across {len(d):,} matching records." if d is not df else "."))
        return {"text": [text],
                "big": fmt_num(val), "label": what.capitalize()}
    xc = x
    if x in dt:
        span = (d[x].max() - d[x].min()).days
        d = d.assign(_x=d[x].dt.to_period(sp["grain"] or ("M" if span > 90 else "D")).dt.to_timestamp())
        xc = "_x"
    grp = d.groupby(xc)
    g = (grp.size() if agg == "count" else grp[y].agg(agg)).rename(yl)
    g = g.sort_index() if (x in dt or kind in ("Line", "Area")) and not sp["top_n"] else g.sort_values(ascending=sp["ascending"])
    if sp["top_n"]:
        g = g.head(sp["top_n"])
    t = g.reset_index()
    t.columns = [x, yl]
    title = f"{what.capitalize()} by {x}"
    if kind == "Pie":
        fig = px.pie(t, names=x, values=yl, hole=.45, title=title)
    elif kind == "Line":
        fig = px.line(t, x=x, y=yl, markers=True, title=title)
    elif kind == "Area":
        fig = px.area(t, x=x, y=yl, title=title)
    else:
        fig = px.bar(t, x=x, y=yl, color=yl, color_continuous_scale=["#22d3ee", "#7c5cff"], title=title)
        fig.update_layout(coloraxis_showscale=False)
    text = chart_story(t, kind, x, yl)
    if agg not in ("sum", "count") and kind in ("Bar", "Pie") and len(t) >= 2:  # shares of averages are meaningless
        text = [f"“{t[x].iloc[0]}” has the {'lowest' if sp['ascending'] else 'highest'} {what} ({fmt_num(t[yl].iloc[0])}), "
                f"while “{t[x].iloc[-1]}” has the {'highest' if sp['ascending'] else 'lowest'} ({fmt_num(t[yl].iloc[-1])})."]
    if sp["top_n"] == 1 and len(t):
        text.insert(0, f"{t[x].iloc[0]} has the {'lowest' if sp['ascending'] else 'highest'} {what}: {fmt_num(t[yl].iloc[0])}.")
    return {"text": text or [f"Here is the {what} by {x}."], "fig": style(fig, 440), "table": t}


@fragment
def ask_bar(work):
    """Search bar: ask a question in plain English and get a chart plus a written answer."""
    ss = st.session_state
    num, dt, cat = col_types(work)
    metric = current_metric(work)
    st.markdown("### 🔎 Ask your data")
    with st.form("ask"):
        c1, c2 = st.columns([6, 1])
        q = c1.text_input("Ask", key="ask_input", label_visibility="collapsed",
                          placeholder='e.g. "show lahore and karachi sales in pie chart"')
        go = c2.form_submit_button("Ask ✨", **STRETCH)
    ex = []
    if cat and metric:
        vals = [str(v) for v in work[cat[0]].dropna().unique()[:2]]
        ex.append(f"show {' and '.join(vals).lower()} {metric} in pie chart")
        ex.append(f"top 5 {cat[0]} by {metric}")
    if dt and metric:
        ex.append(f"{metric} trend monthly")
    ex.append(f"total {metric}" if metric else "how many records")
    mode = "Claude AI mode" if os.environ.get("ANTHROPIC_API_KEY") else "smart rules mode"
    st.caption("Try: " + "  ·  ".join(f"“{e}”" for e in ex[:4]) + f"   |   Running in {mode}. Answers use all rows, ignoring sidebar slicers.")
    if go:
        ss["ask_q"] = q.strip()
    qq = ss.get("ask_q")
    if not qq:
        return
    ck = (qq, metric, ss.get("fid"), ss.get("ev"))
    if ss.get("ask_key") != ck:  # compute once per question
        with st.spinner("Thinking…"):
            spec = llm_spec(qq, work) or sanitize(parse_query(qq, work, metric), work)
            res = run_query(work, spec) if spec.get("understood") else None
        ss["ask_key"], ss["ask_out"] = ck, (spec, res)
    spec, res = ss["ask_out"]
    if res is None:
        story_box("🤖 I did not understand that", ["Mention a column, a value or a chart type. For example: " +
                  " or ".join(f"“{e}”" for e in ex[:3]) + "."], "#f472b6")
        return
    if res.get("big"):
        kpis([("💡", res["label"], res["big"])])
    story_box("🤖 Answer", res["text"], "#7c5cff")
    if res.get("fig") is not None:
        show(res["fig"])
    if res.get("table") is not None:
        with st.expander("Data behind this answer"):
            st.dataframe(arrow_safe(res["table"].head(1000)), **STRETCH)
    st.caption("Interpreted as: " + describe_spec(spec))


# ───────────────────────── Main ─────────────────────────
_c1, _c2, _c3, _c4 = st.columns([4.2, 1.5, 1.5, 1.3])
_c1.markdown('<div class="hero"><h1>FlashBI</h1><div>Drop a spreadsheet. Get answers in seconds.</div></div>',
             unsafe_allow_html=True)
_c2.markdown(f'<div class="user-chip">👤 {html.escape(st.session_state.get("user", "admin"))}</div>',
             unsafe_allow_html=True)
_c3.markdown('<div style="height:1.5rem"></div>', unsafe_allow_html=True)
if _c3.button("🗑️ Clear data", disabled="fkey" not in st.session_state, **STRETCH):
    for _k in ("fkey", "skey", "sheets", "raw", "clean", "stats", "etl", "dl", "ask_q", "ask_key", "metric_sel", "metric_w"):
        st.session_state.pop(_k, None)
    st.session_state["ev"] = 0
    st.session_state["fid"] = st.session_state.get("fid", 0) + 1
    st.session_state["up_key"] = st.session_state.get("up_key", 0) + 1  # new key = empty uploader
    st.session_state["msg"] = "Data cleared. Upload a new file to start again."
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()
_c4.markdown('<div style="height:1.5rem"></div>', unsafe_allow_html=True)
if _c4.button("🔒 Log out", **STRETCH):
    st.session_state.clear()
    st.rerun()
up = st.file_uploader("Upload data", type=["xlsx", "xls", "csv"], label_visibility="collapsed",
                      key=f"up{st.session_state.get('up_key', 0)}")
_msg = st.session_state.pop("msg", None)
if _msg:
    st.success(_msg)

if not up:
    for col, (ic, t, d) in zip(st.columns(3), [
            ("🧹", "Auto-clean", "Types, dates, gaps and duplicates fixed on upload"),
            ("📊", "Instant charts", "Trends, comparisons and correlations with zero clicks"),
            ("🧩", "Self-service ETL", "Reshape data, then export it as Excel or CSV")]):
        col.markdown(f'<div class="feature"><div style="font-size:1.8rem">{ic}</div><b>{t}</b><br>{d}</div>',
                     unsafe_allow_html=True)
    st.stop()

ss = st.session_state
skey = (up.name, up.size)
if ss.get("skey") != skey:  # list sheets once per file
    try:
        ss["sheets"] = sheet_names(up.name, up.getvalue())
    except Exception as e:
        st.error(f"Could not open this file: {e}")
        st.stop()
    ss["skey"] = skey
sheets = ss["sheets"]
sheet = st.sidebar.selectbox("Sheet", sheets) if len(sheets) > 1 else None
fkey = (up.name, up.size, sheet)

if ss.get("fkey") != fkey:  # new file → parse and clean once
    ph = st.empty()
    loader(ph, "Parsing and auto-cleaning your data…")
    try:
        raw, clean, stats = load_and_clean(up.name, up.getvalue(), sheet)
    except Exception as e:
        ph.empty()
        st.error(f"Could not read this file: {e}")
        st.stop()
    ph.empty()
    if clean.empty:
        st.error("This file has no data rows.")
        st.stop()
    for _k in ("metric_sel", "metric_w", "dl", "ask_key", "sug_key"):
        ss.pop(_k, None)
    ss.update(fkey=fkey, raw=raw, clean=clean, stats=stats, etl=None, fid=ss.get("fid", 0) + 1, ev=0)

work = ss["etl"] if ss["etl"] is not None else ss["clean"]
fdf = sidebar_filters(work, f"{ss['fid']}_{ss['ev']}")
st.sidebar.caption(f"📄 {up.name}  ·  {len(fdf):,} of {len(work):,} rows shown")
with st.sidebar:
    sidebar_suggestions(work, current_metric(work))


def overview(fdf, work):
    """Overview section: business KPIs, story, data health, auto visuals."""
    ss = st.session_state
    _num, _, _ = col_types(fdf)
    if _num and not fdf.empty:
        st.markdown("### 💼 Business snapshot")
        if ss.get("metric_w") not in _num:  # restore the choice even after visiting other sections
            ss["metric_w"] = ss.get("metric_sel") if ss.get("metric_sel") in _num else guess_metric(fdf, _num)
        metric = st.selectbox("Business metric (sales, revenue, profit...)", _num, key="metric_w")
        ss["metric_sel"] = metric
        kpis(business_kpis(fdf, metric))
        _story, _acts = story_overview(fdf, metric, ss["stats"])
        story_box("📖 The story of your data", _story)
        story_box("💡 Useful insights: what to do next", _acts, "#facc15")
        if len(fdf) != len(work):
            story_box("🧭 Selected vs unselected data", selection_story(work, fdf, metric), "#f472b6")
    else:
        st.info("Business KPIs need at least one numeric column, such as sales or revenue.")
    st.markdown("### 🩺 Data health")
    kpis([("📦", "Total rows", f"{len(fdf):,}"), ("🧱", "Columns", fdf.shape[1]),
          ("🩹", "Missing values fixed", f"{ss['stats']['fixed']:,}"),
          ("♻️", "Duplicates removed", f"{ss['stats']['dups']:,}")])
    st.markdown("### ✨ Auto visuals")
    if fdf.empty:
        st.warning("No rows match the current slicers.")
    else:
        auto_visuals(fdf)


ask_bar(work)
# Only the selected section is computed, so switching sections stays fast even on big files.
NAV = ["🏠 Overview", "🔍 Data Analysis", "📊 Data Analytics", "🧩 ETL & Transformation", "🤖 AI Insights"]
nav = st.radio("Section", NAV, horizontal=True, key="nav", label_visibility="collapsed")
if nav == NAV[0]:
    overview(fdf, work)
elif nav == NAV[1]:
    tab_inspect(work, fdf)
elif nav == NAV[2]:
    tab_builder(fdf, work)
elif nav == NAV[3]:
    tab_etl(ss["clean"], work)
else:
    tab_insights(fdf, ss["stats"])
