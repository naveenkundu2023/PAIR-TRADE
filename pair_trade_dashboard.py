"""
Pair Trading Dashboard - Streamlit App (Enhanced v3.3 - Correlation + Speed) - FIXED
=======================================================================
Automated pair-trading signal dashboard using independent analytics only.
No Excel dependency.

How to run:
    1. pip install streamlit pandas numpy yfinance plotly openpyxl scipy
    2. streamlit run pair_trade_dashboard.py

Features:
    - Dark theme, high-contrast UI
    - ~40 curated pairs with latest NSE F&O lot sizes (auto month-aware Oct 2026+)
    - Live prices via yfinance (NSE) — SINGLE BATCHED DOWNLOAD for speed
    - Correlation (30D / 60D / 90D) with hard filter (≥0.80 on both 30D & 60D)
    - Z-score, Ratio (current / 30D / 60D min-max / avg)
    - Cash-neutrality lot recommendation
    - Profit potential at different ratio targets
    - Last-5-day divergence dots (green/red/gray) + values + day-over-day trend
    - Z-score interpretation tooltip
    - Historical win-rate simulation (12 months) sorted by win rate
    - Row coloring: light green (ENTER), light yellow (WATCH), light pink (EXIT/HOLD)
    - Signal timestamp (approximate entry time of current signal)
    - Signal board: 1-based index, sorted ENTER (freshest first) → EXIT/HOLD → WATCH
    - Charts: 4 panels (Normalized Prices, Ratio, Z-Score, Rolling Correlation)
    - Independent analytics ranked by confidence (gradient green)
    - Speed: one yf.download for all tickers + 5-min cache
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, timezone
import time
import os
import warnings
warnings.filterwarnings("ignore")

# Force IST (UTC+5:30) for all displayed timestamps
IST = timezone(timedelta(hours=5, minutes=30))

def now_ist():
    return datetime.now(IST)

try:
    import yfinance as yf
    HAS_YF = True
except ImportError:
    HAS_YF = False

try:
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

try:
    from scipy import stats
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG + DARK THEME
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Dark theme CSS – light text, colourful accents
st.markdown("""
<style>
    /* Global dark */
    .stApp {
        background-color: #0E1117;
        color: #E0E0E0;
    }
    h1, h2, h3, h4, h5, h6, p, label, span, div {
        color: #E8EAF0 !important;
    }
    .main-header {
        font-size: 2.3rem;
        font-weight: 800;
        background: linear-gradient(90deg, #00E5FF, #76FF03, #FFEA00);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.15rem;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #90A4AE !important;
        margin-bottom: 1.2rem;
    }
    /* Metric cards */
    [data-testid="stMetric"] {
        background: #1A1F2E;
        border-radius: 10px;
        padding: 12px 16px;
        border-left: 4px solid #00E5FF;
        box-shadow: 0 2px 8px rgba(0,0,0,0.4);
    }
    [data-testid="stMetric"] label {
        color: #90CAF9 !important;
        font-size: 0.8rem !important;
    }
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #FFFFFF !important;
        font-weight: 700;
    }
    /* Signal badges */
    .badge-enter {
        background: linear-gradient(135deg, #00C853, #1B5E20);
        color: #fff; padding: 6px 14px; border-radius: 8px;
        font-weight: 700; font-size: 0.95rem; display: inline-block;
    }
    .badge-watch {
        background: linear-gradient(135deg, #FFD600, #F57F17);
        color: #000; padding: 6px 14px; border-radius: 8px;
        font-weight: 700; font-size: 0.95rem; display: inline-block;
    }
    .badge-exit {
        background: linear-gradient(135deg, #546E7A, #263238);
        color: #fff; padding: 6px 14px; border-radius: 8px;
        font-weight: 700; font-size: 0.95rem; display: inline-block;
    }
    /* Ratio reference table */
    .ratio-box {
        background: #1A1F2E;
        border: 1px solid #37474F;
        border-radius: 10px;
        padding: 14px 18px;
        margin: 8px 0 16px 0;
    }
    .ratio-box table {
        width: 100%;
        border-collapse: collapse;
        color: #E0E0E0;
    }
    .ratio-box th {
        background: #263238;
        color: #80D8FF;
        padding: 8px 12px;
        text-align: center;
        font-size: 0.85rem;
    }
    .ratio-box td {
        padding: 7px 12px;
        text-align: center;
        border-bottom: 1px solid #37474F;
        font-size: 0.9rem;
    }
    /* Divergence dots – use !important so global span rule does not override */
    .dot-green { color: #00E676 !important; font-size: 1.4rem !important; }
    .dot-red   { color: #FF5252 !important; font-size: 1.4rem !important; }
    .dot-gray  { color: #78909C !important; font-size: 1.4rem !important; }
    .dot-white { color: #ECEFF1 !important; font-size: 1.4rem !important; }
    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #12161F;
    }
    [data-testid="stSidebar"] * {
        color: #E0E0E0 !important;
    }
    /* Dataframe */
    .stDataFrame {
        background: #1A1F2E;
    }
    /* Info / warning boxes */
    .stAlert {
        background-color: #1A2332 !important;
        color: #E0E0E0 !important;
    }
    /* Selectbox / inputs */
    .stSelectbox label, .stSlider label, .stMultiSelect label {
        color: #90CAF9 !important;
    }
    /* Make selectbox dropdown options clearly legible */
    div[data-baseweb="select"] > div {
        background-color: #1A1F2E !important;
        color: #FFFFFF !important;
        border: 1px solid #37474F !important;
    }
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div {
        color: #FFFFFF !important;
    }
    ul[data-testid="stSelectboxVirtualDropdown"] li,
    div[role="listbox"] li,
    div[role="option"] {
        background-color: #1A1F2E !important;
        color: #FFFFFF !important;
    }
    div[role="option"]:hover,
    li[aria-selected="true"] {
        background-color: #263238 !important;
        color: #00E5FF !important;
    }
    /* Horizontal rules */
    hr {
        border-color: #37474F !important;
    }
    /* Caption */
    .stCaption, small {
        color: #78909C !important;
    }
    /* Independent analytics confidence gradient cards */
    .conf-card {
        border-radius: 10px;
        padding: 12px 16px;
        margin: 8px 0;
        border-left: 5px solid;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# DEFAULT PAIRS (~40) with latest NSE F&O lot sizes
# Base reference: Oct 2026 (equityx / stockmojo / NSE masters ~15 Oct 2026)
# Monthly revision logic applied via get_active_lot_sizes()
# ─────────────────────────────────────────────────────────────
DEFAULT_PAIRS = [
    # Banking
    {"name": "ICICIBANK-HDFCBANK", "s1": "ICICIBANK.NS", "s2": "HDFCBANK.NS",
     "lot1": 700, "lot2": 650, "sector": "Banking"},
    {"name": "AXISBANK-ICICIBANK", "s1": "AXISBANK.NS", "s2": "ICICIBANK.NS",
     "lot1": 625, "lot2": 700, "sector": "Banking"},
    {"name": "HDFCBANK-SBIN", "s1": "HDFCBANK.NS", "s2": "SBIN.NS",
     "lot1": 650, "lot2": 750, "sector": "Banking"},
    {"name": "BANKBARODA-CANBK", "s1": "BANKBARODA.NS", "s2": "CANBK.NS",
     "lot1": 2925, "lot2": 6750, "sector": "Banking"},
    {"name": "SBIN-BANKBARODA", "s1": "SBIN.NS", "s2": "BANKBARODA.NS",
     "lot1": 750, "lot2": 2925, "sector": "Banking"},
    {"name": "AXISBANK-HDFCBANK", "s1": "AXISBANK.NS", "s2": "HDFCBANK.NS",
     "lot1": 625, "lot2": 650, "sector": "Banking"},
    {"name": "INDUSINDBK-AXISBANK", "s1": "INDUSINDBK.NS", "s2": "AXISBANK.NS",
     "lot1": 700, "lot2": 625, "sector": "Banking"},
    {"name": "KOTAKBANK-HDFCBANK", "s1": "KOTAKBANK.NS", "s2": "HDFCBANK.NS",
     "lot1": 2000, "lot2": 650, "sector": "Banking"},

    # Pharma
    {"name": "CIPLA-SUNPHARMA", "s1": "CIPLA.NS", "s2": "SUNPHARMA.NS",
     "lot1": 425, "lot2": 350, "sector": "Pharma"},
    {"name": "ALKEM-AUROPHARMA", "s1": "ALKEM.NS", "s2": "AUROPHARMA.NS",
     "lot1": 125, "lot2": 550, "sector": "Pharma"},
    {"name": "DRREDDY-SUNPHARMA", "s1": "DRREDDY.NS", "s2": "SUNPHARMA.NS",
     "lot1": 625, "lot2": 350, "sector": "Pharma"},
    {"name": "CIPLA-DRREDDY", "s1": "CIPLA.NS", "s2": "DRREDDY.NS",
     "lot1": 425, "lot2": 625, "sector": "Pharma"},
    {"name": "LUPIN-CIPLA", "s1": "LUPIN.NS", "s2": "CIPLA.NS",
     "lot1": 425, "lot2": 425, "sector": "Pharma"},

    # Oil & Gas / Energy
    {"name": "BPCL-HINDPETRO", "s1": "BPCL.NS", "s2": "HINDPETRO.NS",
     "lot1": 1975, "lot2": 2025, "sector": "Oil & Gas"},
    {"name": "IGL-GAIL", "s1": "IGL.NS", "s2": "GAIL.NS",
     "lot1": 1375, "lot2": 3550, "sector": "Oil & Gas"},
    {"name": "COALINDIA-GAIL", "s1": "COALINDIA.NS", "s2": "GAIL.NS",
     "lot1": 1350, "lot2": 3550, "sector": "Energy"},
    {"name": "ONGC-BPCL", "s1": "ONGC.NS", "s2": "BPCL.NS",
     "lot1": 2250, "lot2": 1975, "sector": "Oil & Gas"},
    {"name": "IOC-BPCL", "s1": "IOC.NS", "s2": "BPCL.NS",
     "lot1": 4875, "lot2": 1975, "sector": "Oil & Gas"},

    # Cement
    {"name": "ACC-AMBUJACEM", "s1": "ACC.NS", "s2": "AMBUJACEM.NS",
     "lot1": 300, "lot2": 1200, "sector": "Cement"},
    {"name": "ULTRACEMCO-AMBUJACEM", "s1": "ULTRACEMCO.NS", "s2": "AMBUJACEM.NS",
     "lot1": 50, "lot2": 1200, "sector": "Cement"},
    {"name": "SHREECEM-ULTRACEMCO", "s1": "SHREECEM.NS", "s2": "ULTRACEMCO.NS",
     "lot1": 25, "lot2": 50, "sector": "Cement"},

    # IT
    {"name": "TCS-HCLTECH", "s1": "TCS.NS", "s2": "HCLTECH.NS",
     "lot1": 225, "lot2": 400, "sector": "IT"},
    {"name": "INFY-TCS", "s1": "INFY.NS", "s2": "TCS.NS",
     "lot1": 400, "lot2": 225, "sector": "IT"},
    {"name": "WIPRO-INFY", "s1": "WIPRO.NS", "s2": "INFY.NS",
     "lot1": 3000, "lot2": 400, "sector": "IT"},
    {"name": "TECHM-HCLTECH", "s1": "TECHM.NS", "s2": "HCLTECH.NS",
     "lot1": 600, "lot2": 400, "sector": "IT"},
    {"name": "INFY-HCLTECH", "s1": "INFY.NS", "s2": "HCLTECH.NS",
     "lot1": 400, "lot2": 400, "sector": "IT"},

    # Metals
    {"name": "TATASTEEL-JSWSTEEL", "s1": "TATASTEEL.NS", "s2": "JSWSTEEL.NS",
     "lot1": 2750, "lot2": 675, "sector": "Metals"},
    {"name": "HINDALCO-JSWSTEEL", "s1": "HINDALCO.NS", "s2": "JSWSTEEL.NS",
     "lot1": 700, "lot2": 675, "sector": "Metals"},
    {"name": "SAIL-TATASTEEL", "s1": "SAIL.NS", "s2": "TATASTEEL.NS",
     "lot1": 4700, "lot2": 2750, "sector": "Metals"},
    {"name": "JINDALSTEL-JSWSTEEL", "s1": "JINDALSTEL.NS", "s2": "JSWSTEEL.NS",
     "lot1": 625, "lot2": 675, "sector": "Metals"},

    # FMCG
    {"name": "NESTLEIND-HINDUNILVR", "s1": "NESTLEIND.NS", "s2": "HINDUNILVR.NS",
     "lot1": 500, "lot2": 300, "sector": "FMCG"},
    {"name": "BRITANNIA-HINDUNILVR", "s1": "BRITANNIA.NS", "s2": "HINDUNILVR.NS",
     "lot1": 125, "lot2": 300, "sector": "FMCG"},
    {"name": "ITC-HINDUNILVR", "s1": "ITC.NS", "s2": "HINDUNILVR.NS",
     "lot1": 1725, "lot2": 300, "sector": "FMCG"},
    {"name": "DABUR-MARICO", "s1": "DABUR.NS", "s2": "MARICO.NS",
     "lot1": 1250, "lot2": 1200, "sector": "FMCG"},
    {"name": "GODREJCP-HINDUNILVR", "s1": "GODREJCP.NS", "s2": "HINDUNILVR.NS",
     "lot1": 500, "lot2": 300, "sector": "FMCG"},

    # Finance
    {"name": "BAJFINANCE-BAJAJFINSV", "s1": "BAJFINANCE.NS", "s2": "BAJAJFINSV.NS",
     "lot1": 750, "lot2": 300, "sector": "Finance"},
    {"name": "RECLTD-PFC", "s1": "RECLTD.NS", "s2": "PFC.NS",
     "lot1": 1575, "lot2": 1300, "sector": "Finance"},
    {"name": "CHOLAFIN-BAJFINANCE", "s1": "CHOLAFIN.NS", "s2": "BAJFINANCE.NS",
     "lot1": 625, "lot2": 750, "sector": "Finance"},
    {"name": "SBILIFE-HDFCLIFE", "s1": "SBILIFE.NS", "s2": "HDFCLIFE.NS",
     "lot1": 375, "lot2": 1100, "sector": "Finance"},

    # Paints / Others
    {"name": "BERGEPAINT-ASIANPAINT", "s1": "BERGEPAINT.NS", "s2": "ASIANPAINT.NS",
     "lot1": 1320, "lot2": 250, "sector": "Paints"},
    {"name": "ASIANPAINT-PIDILITIND", "s1": "ASIANPAINT.NS", "s2": "PIDILITIND.NS",
     "lot1": 250, "lot2": 250, "sector": "Paints"},

    # Index
    {"name": "NIFTY-BANKNIFTY", "s1": "^NSEI", "s2": "^NSEBANK",
     "lot1": 65, "lot2": 30, "sector": "Index"},
    {"name": "HDFCBANK-BANKNIFTY", "s1": "HDFCBANK.NS", "s2": "^NSEBANK",
     "lot1": 650, "lot2": 30, "sector": "Banking"},
]

# ─────────────────────────────────────────────────────────────
# MONTH-AWARE LOT SIZE HELPER
# ─────────────────────────────────────────────────────────────
LOT_REVISIONS = {
    # Oct 2026 base is already in DEFAULT_PAIRS.
    # Add future revisions here as NSE publishes circulars.
}

def _current_fo_month():
    now = now_ist()
    y, m, d = now.year, now.month, now.day
    if d >= 28:
        if m == 12:
            return (y + 1, 1)
        return (y, m + 1)
    return (y, m)


def get_active_lot_sizes(pair: dict) -> tuple:
    fo_y, fo_m = _current_fo_month()
    t1 = pair["s1"].replace(".NS", "").replace("^", "")
    t2 = pair["s2"].replace(".NS", "").replace("^", "")
    lot1, lot2 = pair["lot1"], pair["lot2"]

    applicable = [
        (k, v) for k, v in LOT_REVISIONS.items()
        if k <= (fo_y, fo_m)
    ]
    if applicable:
        applicable.sort(key=lambda x: x[0], reverse=True)
        rev = applicable[0][1]
        if t1 in rev:
            lot1 = rev[t1]
        if t2 in rev:
            lot2 = rev[t2]
    return lot1, lot2


def fo_month_label() -> str:
    y, m = _current_fo_month()
    return datetime(y, m, 1).strftime("%b %Y")


# ─────────────────────────────────────────────────────────────
# DATA HELPERS — SPEED OPTIMISED (single batched download)
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)   # 5-minute cache — most refreshes are instant
def fetch_all_closes(tickers: tuple, period: str = "1y") -> dict:
    """
    Download ALL unique tickers in ONE call (threads=True).
    Returns {ticker: Series of Close prices} or empty dict on failure.
    """
    if not HAS_YF or not tickers:
        return {}
    try:
        data = yf.download(
            list(tickers),
            period=period,
            group_by="ticker",
            auto_adjust=True,
            threads=True,
            progress=False,
        )
        result = {}
        if len(tickers) == 1:
            # single ticker → columns are simple
            t = tickers[0]
            if "Close" in data.columns:
                s = data["Close"].dropna()
                s.index = pd.to_datetime(s.index).tz_localize(None)
                result[t] = s
        else:
            # multi-ticker → multi-index columns
            for t in tickers:
                try:
                    if t in data.columns.get_level_values(0):
                        s = data[t]["Close"].dropna()
                        s.index = pd.to_datetime(s.index).tz_localize(None)
                        result[t] = s
                except Exception:
                    continue
        return result
    except Exception:
        return {}


def get_pair_df(closes: dict, t1: str, t2: str) -> pd.DataFrame:
    """Build aligned DataFrame from pre-downloaded closes."""
    if t1 not in closes or t2 not in closes:
        return pd.DataFrame()
    s1 = closes[t1]
    s2 = closes[t2]
    df = pd.DataFrame({"price1": s1, "price2": s2}).dropna()
    return df


def compute_pair_metrics(df: pd.DataFrame, lookback: int = 60) -> dict:
    if df.empty or len(df) < 30:
        return {}

    ratio = df["price1"] / df["price2"]
    window = min(lookback, len(ratio) - 1)

    ratio_mean = ratio.rolling(window).mean()
    ratio_std = ratio.rolling(window).std()
    zscore = (ratio - ratio_mean) / ratio_std

    corr_30 = df["price1"].tail(30).corr(df["price2"].tail(30)) if len(df) >= 30 else np.nan
    corr_60 = df["price1"].tail(60).corr(df["price2"].tail(60)) if len(df) >= 60 else np.nan
    corr_90 = df["price1"].tail(90).corr(df["price2"].tail(90)) if len(df) >= 90 else np.nan

    # Rolling correlation series for the 4th chart
    corr_roll_30 = df["price1"].rolling(30).corr(df["price2"])
    corr_roll_60 = df["price1"].rolling(60).corr(df["price2"])

    cur_ratio = float(ratio.iloc[-1])
    cur_z = float(zscore.iloc[-1]) if not np.isnan(zscore.iloc[-1]) else 0.0
    cur_p1 = float(df["price1"].iloc[-1])
    cur_p2 = float(df["price2"].iloc[-1])

    sma20_1 = float(df["price1"].tail(20).mean())
    sma20_2 = float(df["price2"].tail(20).mean())
    std20_1 = float(df["price1"].tail(20).std())
    std20_2 = float(df["price2"].tail(20).std())

    recent60 = df.tail(60)
    support1 = float(recent60["price1"].min())
    resist1 = float(recent60["price1"].max())
    support2 = float(recent60["price2"].min())
    resist2 = float(recent60["price2"].max())

    r30 = ratio.tail(30)
    r60 = ratio.tail(60)
    ratio_max_30 = float(r30.max())
    ratio_min_30 = float(r30.min())
    ratio_max_60 = float(r60.max())
    ratio_min_60 = float(r60.min())
    ratio_avg_30 = float(r30.mean())
    ratio_avg_60 = float(r60.mean())
    ratio_mean_60 = ratio_avg_60
    ratio_avg_90 = float(ratio.tail(90).mean()) if len(ratio) >= 90 else ratio_avg_60

    half_life = np.nan
    if HAS_SCIPY and len(ratio) > 30:
        try:
            lag = ratio.shift(1).dropna()
            delta = ratio.diff().dropna()
            lag = lag.loc[delta.index]
            slope, _, _, _, _ = stats.linregress(lag, delta)
            if slope < 0:
                half_life = -np.log(2) / slope
        except Exception:
            pass

    ratio_sma20 = ratio.rolling(20).mean()
    divergence = (ratio - ratio_sma20) / ratio_sma20
    last5_div = divergence.tail(5).tolist() if len(divergence) >= 5 else []
    last5_dates = [d.strftime("%d-%b") for d in divergence.tail(5).index] if len(divergence) >= 5 else []

    last5_abs_div = [abs(x) if not np.isnan(x) else np.nan for x in last5_div]
    last5_div_delta = []
    for i in range(len(last5_abs_div)):
        if i == 0 or np.isnan(last5_abs_div[i]) or np.isnan(last5_abs_div[i - 1]):
            last5_div_delta.append(np.nan)
        else:
            last5_div_delta.append(last5_abs_div[i] - last5_abs_div[i - 1])

    last5_z = zscore.tail(5).tolist() if len(zscore) >= 5 else []

    signal_since = None
    if len(zscore) > 5:
        for i in range(len(zscore) - 2, max(0, len(zscore) - 60), -1):
            if abs(zscore.iloc[i]) < 1.5:
                signal_since = zscore.index[i + 1]
                break
        if signal_since is None and abs(cur_z) >= 1.5:
            signal_since = zscore.index[max(0, len(zscore) - 30)]

    return {
        "ratio": cur_ratio,
        "zscore": cur_z,
        "corr_30": corr_30,
        "corr_60": corr_60,
        "corr_90": corr_90,
        "corr_roll_30": corr_roll_30,
        "corr_roll_60": corr_roll_60,
        "price1": cur_p1,
        "price2": cur_p2,
        "sma20_1": sma20_1,
        "sma20_2": sma20_2,
        "upper1": sma20_1 + 2 * std20_1,
        "lower1": sma20_1 - 2 * std20_1,
        "upper2": sma20_2 + 2 * std20_2,
        "lower2": sma20_2 - 2 * std20_2,
        "support1": support1,
        "resist1": resist1,
        "support2": support2,
        "resist2": resist2,
        "ratio_max_30": ratio_max_30,
        "ratio_min_30": ratio_min_30,
        "ratio_max_60": ratio_max_60,
        "ratio_min_60": ratio_min_60,
        "ratio_avg_30": ratio_avg_30,
        "ratio_avg_60": ratio_avg_60,
        "ratio_avg_90": ratio_avg_90,
        "ratio_mean": ratio_mean_60,
        "half_life": half_life,
        "ratio_series": ratio,
        "zscore_series": zscore,
        "last5_div": last5_div,
        "last5_dates": last5_dates,
        "last5_div_delta": last5_div_delta,
        "last5_z": last5_z,
        "signal_since": signal_since,
        "df": df,
    }


def generate_signal(metrics: dict, z_entry: float = 1.5, z_exit: float = 0.5) -> dict:
    if not metrics:
        return {"action": "NO DATA", "leg1": "—", "leg2": "—", "confidence": 0,
                "reason": "Insufficient data", "deep_reason": ""}

    z = metrics["zscore"]
    corr_30 = metrics.get("corr_30")
    corr_60 = metrics.get("corr_60")
    abs_z = abs(z)

    # Strong correlation filter: both 30D and 60D must be ≥ 0.80
    corr_30_ok = (not np.isnan(corr_30)) and abs(corr_30) >= 0.80
    corr_60_ok = (not np.isnan(corr_60)) and abs(corr_60) >= 0.80
    corr_ok = corr_30_ok and corr_60_ok

    deep = ""
    if abs_z >= z_entry and corr_ok:
        if z > 0:
            action, leg1, leg2 = "ENTER PAIR", "SELL", "BUY"
            reason = (f"Z-score +{z:.2f} (ratio stretched high). "
                      f"Expect mean-reversion → Sell leg1 / Buy leg2")
            deep = (
                f"The price ratio is {abs_z:.1f} standard deviations above its "
                f"{60}-day mean. Historically, such stretches reverse toward the mean. "
                f"30-day correlation is {corr_30:.2f} and 60-day is {corr_60:.2f} "
                f"(both ≥ 0.80), confirming the pair still moves tightly together. "
                f"Selling the richer leg and buying the cheaper leg captures the reversion."
            )
        else:
            action, leg1, leg2 = "ENTER PAIR", "BUY", "SELL"
            reason = (f"Z-score {z:.2f} (ratio stretched low). "
                      f"Expect mean-reversion → Buy leg1 / Sell leg2")
            deep = (
                f"The price ratio is {abs_z:.1f} standard deviations below its "
                f"rolling mean. Mean-reversion probability is elevated. "
                f"Correlation (30D={corr_30:.2f}, 60D={corr_60:.2f}) supports pair quality. "
                f"Buying the cheaper leg and shorting the richer leg is the classic setup."
            )
        confidence = min(95, int(40 + abs_z * 15 + (min(abs(corr_30), abs(corr_60)) * 25)))
    elif abs_z <= z_exit:
        action, leg1, leg2 = "EXIT / HOLD", "CLOSE", "CLOSE"
        reason = f"Z-score near mean ({z:.2f}). No edge / take-profit zone"
        deep = "Ratio is close to its average; edge has largely been realized or never appeared."
        confidence = 30
    else:
        action, leg1, leg2 = "WATCH", "—", "—"
        reason = f"Z-score {z:.2f} inside neutral band. Waiting for stretch"
        deep = "Wait for |Z| to exceed the entry threshold before committing capital."
        confidence = 20

    if not corr_ok and action == "ENTER PAIR":
        action = "WATCH (Low Corr)"
        c30_str = f"{corr_30:.2f}" if not np.isnan(corr_30) else "n/a"
        c60_str = f"{corr_60:.2f}" if not np.isnan(corr_60) else "n/a"
        reason += f" | Corr30={c30_str} / Corr60={c60_str} (need both ≥ 0.80)"
        confidence = max(10, confidence - 35)
        deep += " Low correlation on 30D or 60D reduces confidence in mean-reversion of the ratio."

    return {
        "action": action, "leg1": leg1, "leg2": leg2,
        "confidence": confidence, "reason": reason, "deep_reason": deep,
        "z": z, "corr_30": corr_30, "corr_60": corr_60,
    }


def recommend_cash_neutral_lots(metrics: dict, lot1: int, lot2: int) -> dict:
    if not metrics:
        return {}
    p1, p2 = metrics["price1"], metrics["price2"]
    v1 = p1 * lot1
    v2 = p2 * lot2
    best = None
    for n1 in range(1, 6):
        for n2 in range(1, 6):
            diff = abs(n1 * v1 - n2 * v2)
            pct = diff / max(n1 * v1, n2 * v2) * 100
            if best is None or pct < best["imbalance_pct"]:
                best = {
                    "lots1": n1, "lots2": n2,
                    "shares1": n1 * lot1, "shares2": n2 * lot2,
                    "value1": round(n1 * v1, 0),
                    "value2": round(n2 * v2, 0),
                    "imbalance": round(n1 * v1 - n2 * v2, 0),
                    "imbalance_pct": round(pct, 2),
                }
    best["single_value1"] = round(v1, 0)
    best["single_value2"] = round(v2, 0)
    best["single_imbalance"] = round(v1 - v2, 0)
    best["single_imbalance_pct"] = round(abs(v1 - v2) / max(v1, v2) * 100, 2)
    return best


def profit_potential(metrics: dict, signal: dict, lots_info: dict) -> list:
    if not metrics or not lots_info:
        return []
    p1, p2 = metrics["price1"], metrics["price2"]
    cur_r = metrics["ratio"]
    targets = [
        ("Mean (60D avg)", metrics["ratio_mean"]),
        ("30D Avg", metrics["ratio_avg_30"]),
        ("Halfway to mean", (cur_r + metrics["ratio_mean"]) / 2),
        ("30D Min", metrics["ratio_min_30"]),
        ("30D Max", metrics["ratio_max_30"]),
    ]
    sh1 = lots_info.get("shares1", 0)
    sh2 = lots_info.get("shares2", 0)
    if sh1 == 0:
        return []

    rows = []
    direction = signal.get("leg1", "")
    for label, tgt_r in targets:
        if tgt_r <= 0 or cur_r <= 0:
            continue
        p1_new_a = tgt_r * p2
        p2_new_b = p1 / tgt_r

        if direction == "SELL":
            pnl_a = sh1 * (p1 - p1_new_a) + sh2 * 0
            pnl_b = sh1 * 0 + sh2 * (p2_new_b - p2)
        elif direction == "BUY":
            pnl_a = sh1 * (p1_new_a - p1) + sh2 * 0
            pnl_b = sh1 * 0 + sh2 * (p2 - p2_new_b)
        else:
            pnl_a = pnl_b = 0

        pnl = (pnl_a + pnl_b) / 2
        move_pct = abs(tgt_r - cur_r) / cur_r * 100
        rows.append({
            "Target": label,
            "Target Ratio": round(tgt_r, 4),
            "Ratio Move %": round(move_pct, 2),
            "Est. P&L (₹)": round(pnl, 0),
        })
    return rows


def interpret_z_trend(last5_z: list) -> str:
    if len(last5_z) < 3:
        return "Insufficient Z-score history."
    recent = [z for z in last5_z if not np.isnan(z)]
    if len(recent) < 3:
        return "Insufficient Z-score history."
    delta = recent[-1] - recent[0]
    abs_now = abs(recent[-1])
    if delta > 0.3:
        return (f"Z-score is **rising** (Δ ≈ {delta:+.2f} over 5 days). "
                f"The ratio is stretching further from its mean. "
                f"If already extended (|Z|={abs_now:.2f}), a reversion entry may be approaching; "
                f"if still near zero, the pair is starting to diverge – watch for entry threshold.")
    elif delta < -0.3:
        return (f"Z-score is **falling** (Δ ≈ {delta:+.2f} over 5 days). "
                f"The ratio is moving back toward (or past) its mean. "
                f"If you are in a mean-reversion trade, this is favourable; "
                f"if |Z| is collapsing toward 0, consider taking profits.")
    else:
        return (f"Z-score is **stable / oscillating** (Δ ≈ {delta:+.2f}). "
                f"No strong directional stretch. Wait for a clearer move beyond the entry band.")


def historical_pair_stats(df: pd.DataFrame, z_entry: float = 1.5, lookback: int = 60) -> dict:
    if df.empty or len(df) < lookback + 20:
        return {"trades": 0, "wins": 0, "win_rate": None, "avg_pnl_pct": None, "details": []}

    ratio = df["price1"] / df["price2"]
    rmean = ratio.rolling(lookback).mean()
    rstd = ratio.rolling(lookback).std()
    z = (ratio - rmean) / rstd

    trades = []
    i = lookback
    while i < len(z) - 5:
        zi = z.iloc[i]
        if np.isnan(zi):
            i += 1
            continue
        if abs(zi) >= z_entry:
            entry_ratio = ratio.iloc[i]
            entry_z = zi
            exit_i = None
            for j in range(i + 1, min(i + 31, len(z))):
                if abs(z.iloc[j]) <= 0.5:
                    exit_i = j
                    break
            if exit_i is None:
                exit_i = min(i + 30, len(z) - 1)
            exit_ratio = ratio.iloc[exit_i]
            if entry_z > 0:
                pnl_pct = (entry_ratio - exit_ratio) / entry_ratio * 100
            else:
                pnl_pct = (exit_ratio - entry_ratio) / entry_ratio * 100
            trades.append({
                "entry_date": z.index[i].strftime("%Y-%m-%d"),
                "exit_date": z.index[exit_i].strftime("%Y-%m-%d"),
                "entry_z": round(entry_z, 2),
                "pnl_pct": round(pnl_pct, 2),
                "win": pnl_pct > 0,
            })
            i = exit_i + 1
        else:
            i += 1

    if not trades:
        return {"trades": 0, "wins": 0, "win_rate": None, "avg_pnl_pct": None, "details": []}
    wins = sum(1 for t in trades if t["win"])
    return {
        "trades": len(trades),
        "wins": wins,
        "win_rate": round(wins / len(trades) * 100, 1),
        "avg_pnl_pct": round(np.mean([t["pnl_pct"] for t in trades]), 2),
        "details": trades[-12:],
    }


def recommend_prices(metrics: dict, signal: dict) -> dict:
    if not metrics or signal["action"] not in ("ENTER PAIR",):
        return {}
    z = metrics["zscore"]
    p1, p2 = metrics["price1"], metrics["price2"]
    ratio = metrics["ratio"]
    mean_r = metrics["ratio_mean"]
    target_ratio = mean_r
    if z > 0:
        stop_ratio = metrics["ratio_max_60"] * 1.02
    else:
        stop_ratio = metrics["ratio_min_60"] * 0.98

    if signal["leg1"] == "SELL":
        entry1, target1 = p1, p1 * (target_ratio / ratio) * 0.98
        stop1 = p1 * (stop_ratio / ratio) * 1.01
        entry2, target2 = p2, p2 * (ratio / target_ratio) * 1.02
        stop2 = p2 * (ratio / stop_ratio) * 0.99
    else:
        entry1, target1 = p1, p1 * (target_ratio / ratio) * 1.02
        stop1 = p1 * (stop_ratio / ratio) * 0.99
        entry2, target2 = p2, p2 * (ratio / target_ratio) * 0.98
        stop2 = p2 * (ratio / stop_ratio) * 1.01

    return {
        "entry1": round(entry1, 2), "target1": round(target1, 2), "stop1": round(stop1, 2),
        "entry2": round(entry2, 2), "target2": round(target2, 2), "stop2": round(stop2, 2),
        "support1": round(metrics["support1"], 2), "resist1": round(metrics["resist1"], 2),
        "support2": round(metrics["support2"], 2), "resist2": round(metrics["resist2"], 2),
    }


# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Controls")
    st.markdown("---")

    z_entry = st.slider("Z-score Entry Threshold", 1.0, 3.0, 1.5, 0.1,
                        help="Enter when |Z| exceeds this")
    z_exit = st.slider("Z-score Exit Threshold", 0.1, 1.5, 0.5, 0.1)
    lookback = st.selectbox("Z-score Lookback (days)", [30, 60, 90], index=1)

    sector_filter = st.multiselect(
        "Filter by Sector",
        options=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
        default=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
    )

    auto_refresh = st.checkbox("Auto-refresh every 60s", value=False)
    show_charts = st.checkbox("Show detailed charts", value=True)

    st.markdown("---")
    st.caption("Data: Yahoo Finance (batched + 5-min cache)")
    st.caption(f"Lot sizes: NSE F&O — active series {fo_month_label()}")
    st.caption(f"Clock: {now_ist().strftime('%Y-%m-%d %H:%M:%S')} IST")
    st.caption("Corr filter: |Corr30| ≥ 0.80 AND |Corr60| ≥ 0.80 for ENTER")

    if st.button("🔄 Refresh Now (clear cache)"):
        st.cache_data.clear()
        st.rerun()


# ─────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────
_now = now_ist()
st.markdown('<p class="main-header">📈 Automated Pair Trading Dashboard</p>', unsafe_allow_html=True)
st.markdown(
    f'<p class="sub-header">Fully Independent Analytics • Dark theme • '
    f'Updated {_now.strftime("%Y-%m-%d %H:%M:%S")} IST</p>',
    unsafe_allow_html=True,
)
st.caption(
    f"📂 Independent engine only — no Excel dependency. "
    f"Lot sizes from NSE F&O references (active series: {fo_month_label()}). "
    f"After each monthly expiry the engine rolls to the next month; "
    f"update LOT_REVISIONS when NSE publishes new circulars. "
    f"Correlation filter: both 30-day and 60-day |corr| ≥ 0.80 required for ENTER. "
    f"Speed: single batched yfinance download + 5-minute cache."
)

if not HAS_YF:
    st.error("yfinance is not installed. Run: `pip install yfinance`")
    st.stop()

# ─────────────────────────────────────────────────────────────
# SCAN ALL PAIRS — SPEED PATH
# ─────────────────────────────────────────────────────────────
pairs_to_scan = [p for p in DEFAULT_PAIRS if p["sector"] in sector_filter]

# Collect unique tickers once
all_tickers = sorted(set(
    [p["s1"] for p in pairs_to_scan] + [p["s2"] for p in pairs_to_scan]
))
tickers_tuple = tuple(all_tickers)   # hashable for cache

progress = st.progress(0, text="Fetching live data (batched)…")
closes = fetch_all_closes(tickers_tuple, period="1y")
progress.progress(0.4, text="Computing metrics…")

results = []
n = len(pairs_to_scan)
for i, pair in enumerate(pairs_to_scan):
    progress.progress(0.4 + 0.6 * (i + 1) / n, text=f"Scanning {pair['name']}…")
    hist = get_pair_df(closes, pair["s1"], pair["s2"])
    metrics = compute_pair_metrics(hist, lookback=lookback)
    signal = generate_signal(metrics, z_entry=z_entry, z_exit=z_exit)
    prices = recommend_prices(metrics, signal)
    active_lot1, active_lot2 = get_active_lot_sizes(pair)
    lots_info = recommend_cash_neutral_lots(metrics, active_lot1, active_lot2) if metrics else {}
    hist_stats = historical_pair_stats(hist, z_entry=z_entry, lookback=lookback) if not hist.empty else {}

    results.append({
        "pair": pair,
        "active_lot1": active_lot1,
        "active_lot2": active_lot2,
        "metrics": metrics,
        "signal": signal,
        "prices": prices,
        "lots_info": lots_info,
        "hist_stats": hist_stats,
        "hist": hist,
    })

progress.empty()

# ─────────────────────────────────────────────────────────────
# SUMMARY TABLE
# ─────────────────────────────────────────────────────────────
st.subheader("🎯 Live Signal Board — All Pairs")

rows = []
for r in results:
    m = r["metrics"]
    s = r["signal"]
    p = r["pair"]
    if not m:
        continue

    sig_since = ""
    sig_since_ts = None
    if m.get("signal_since") is not None and "ENTER" in s["action"]:
        try:
            sig_since_ts = m["signal_since"]
            sig_since = sig_since_ts.strftime("%d-%b %H:%M")
        except Exception:
            sig_since = str(m["signal_since"])[:16]
    elif "ENTER" in s["action"]:
        sig_since_ts = now_ist().replace(tzinfo=None)
        sig_since = now_ist().strftime("%d-%b %H:%M")

    rows.append({
        "Pair": p["name"],
        "Sector": p["sector"],
        "Lot1": r["active_lot1"],
        "Lot2": r["active_lot2"],
        "Price1": round(m["price1"], 2),
        "Price2": round(m["price2"], 2),
        "Ratio": round(m["ratio"], 4),
        "Z-Score": round(m["zscore"], 2),
        "Corr30": round(m["corr_30"], 2) if not np.isnan(m["corr_30"]) else "—",
        "Corr60": round(m["corr_60"], 2) if not np.isnan(m["corr_60"]) else "—",
        "Corr90": round(m["corr_90"], 2) if not np.isnan(m["corr_90"]) else "—",
        "Action": s["action"],
        "Leg1": s["leg1"],
        "Leg2": s["leg2"],
        "Conf%": s["confidence"],
        "Signal Since": sig_since if sig_since else "—",
        "HalfLife": round(m["half_life"], 1) if not np.isnan(m.get("half_life", np.nan)) else "—",
        "_sig_since_ts": sig_since_ts,
        "_action_rank": (
            0 if "ENTER" in s["action"] else
            1 if ("EXIT" in s["action"] or "HOLD" in s["action"]) else
            2
        ),
    })

if not rows:
    st.warning("No data returned. Check internet / ticker symbols.")
    st.stop()

df_signals = pd.DataFrame(rows)

df_signals = df_signals.sort_values(
    by=["_action_rank", "_sig_since_ts"],
    ascending=[True, False],
    na_position="last",
).reset_index(drop=True)

df_signals.insert(0, "#", range(1, len(df_signals) + 1))

display_cols = [c for c in df_signals.columns if not c.startswith("_")]
df_display = df_signals[display_cols].copy()

def color_action_row(row):
    action = str(row.get("Action", ""))
    styles = [""] * len(row)
    if "ENTER" in action:
        styles = ["background-color: rgba(0, 200, 83, 0.12)"] * len(row)
    elif "WATCH" in action:
        styles = ["background-color: rgba(255, 214, 0, 0.12)"] * len(row)
    elif "EXIT" in action or "HOLD" in action:
        styles = ["background-color: rgba(255, 82, 82, 0.10)"] * len(row)
    return styles

def color_action_cell(val):
    v = str(val)
    if "ENTER" in v:
        return "background-color: #1B5E20; color: #FFFFFF; font-weight: bold"
    if "EXIT" in v or "HOLD" in v:
        return "background-color: #37474F; color: #FFFFFF"
    if "WATCH" in v:
        return "background-color: #F9A825; color: #000000; font-weight: bold"
    return ""

styled = (df_display.style
          .apply(color_action_row, axis=1)
          .map(color_action_cell, subset=["Action"])
          .format({
              "Price1": "{:.2f}", "Price2": "{:.2f}",
              "Ratio": "{:.4f}", "Z-Score": "{:.2f}",
          }))

st.dataframe(styled, use_container_width=True, height=920)

enter_count = sum(1 for r in rows if "ENTER" in r["Action"])
watch_count = sum(1 for r in rows if "WATCH" in r["Action"])
exit_count = sum(1 for r in rows if "EXIT" in r["Action"] or "HOLD" in r["Action"])
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Pairs Scanned", len(rows))
col2.metric("🟢 Enter Signals", enter_count)
col3.metric("🟡 Watchlist", watch_count)
col4.metric("⚪ Exit / Hold", exit_count)
col5.metric("Avg |Z|", round(np.mean([abs(r["Z-Score"]) for r in rows]), 2))

# ─────────────────────────────────────────────────────────────
# DETAILED VIEW
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("🔍 Pair Deep-Dive & Trade Recommendations")

pair_names = [r["pair"]["name"] for r in results if r["metrics"]]
selected = st.selectbox("Select pair for detailed analysis", pair_names)

selected_result = next(r for r in results if r["pair"]["name"] == selected)
m = selected_result["metrics"]
s = selected_result["signal"]
pr = selected_result["prices"]
pair_info = selected_result["pair"]
lots_info = selected_result["lots_info"]
hist_stats = selected_result["hist_stats"]
active_lot1 = selected_result["active_lot1"]
active_lot2 = selected_result["active_lot2"]

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Current Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}",
          delta="Stretched" if abs(m["zscore"]) > z_entry else "Normal")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if not np.isnan(m["corr_60"]) else "—")
k5.metric("90D Corr", f"{m['corr_90']:.2f}" if not np.isnan(m["corr_90"]) else "—")
k6.metric("Half-Life (d)", f"{m['half_life']:.1f}" if not np.isnan(m.get("half_life", np.nan)) else "—")

if "ENTER" in s["action"]:
    action_color = "#1B5E20"
elif "WATCH" in s["action"]:
    action_color = "#F9A825"
else:
    action_color = "#37474F"

sig_time_note = ""
if m.get("signal_since") is not None and "ENTER" in s["action"]:
    try:
        sig_time_note = f'<p style="margin:6px 0 0 0;opacity:0.9;">Signal since ≈ {m["signal_since"].strftime("%d-%b %Y %H:%M")}</p>'
    except Exception:
        pass

st.markdown(
    f"""
    <div style="background:{action_color};color:white;padding:18px;border-radius:12px;margin:12px 0;">
        <h3 style="margin:0;color:white;">{s['action']} — Confidence {s['confidence']}%</h3>
        <p style="margin:8px 0 0 0;font-size:1.1rem;">
            <b>{pair_info['s1']}</b> → <b>{s['leg1']}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>{pair_info['s2']}</b> → <b>{s['leg2']}</b>
        </p>
        <p style="margin:8px 0 0 0;opacity:0.95;">{s['reason']}</p>
        {sig_time_note}
    </div>
    """,
    unsafe_allow_html=True,
)

if s.get("deep_reason"):
    st.markdown("#### 🧠 Independent Deep Rational")
    st.info(s["deep_reason"])

st.markdown("#### 📐 Ratio Reference Levels")
st.markdown(
    f"""
    <div class="ratio-box">
    <table>
        <tr>
            <th>Metric</th>
            <th>Current</th>
            <th>30-Day</th>
            <th>60-Day</th>
            <th>90-Day</th>
        </tr>
        <tr>
            <td><b>Ratio</b></td>
            <td style="color:#00E5FF;font-weight:700;">{m['ratio']:.4f}</td>
            <td>—</td>
            <td>—</td>
            <td>—</td>
        </tr>
        <tr>
            <td>Average</td>
            <td>—</td>
            <td>{m['ratio_avg_30']:.4f}</td>
            <td>{m['ratio_avg_60']:.4f}</td>
            <td>{m.get('ratio_avg_90', m['ratio_avg_60']):.4f}</td>
        </tr>
        <tr>
            <td>Maximum</td>
            <td>—</td>
            <td style="color:#FF8A80;">{m['ratio_max_30']:.4f}</td>
            <td style="color:#FF8A80;">{m['ratio_max_60']:.4f}</td>
            <td>—</td>
        </tr>
        <tr>
            <td>Minimum</td>
            <td>—</td>
            <td style="color:#69F0AE;">{m['ratio_min_30']:.4f}</td>
            <td style="color:#69F0AE;">{m['ratio_min_60']:.4f}</td>
            <td>—</td>
        </tr>
    </table>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("#### 📦 Lot Sizes & Cash-Neutrality Recommendation")
c1, c2 = st.columns(2)
with c1:
    st.markdown(f"**Exchange lot sizes (NSE F&O — active series {fo_month_label()})**")
    st.write(f"| Leg | Ticker | Lot Size |")
    st.write(f"|-----|--------|----------|")
    st.write(f"| 1 | {pair_info['s1']} | **{active_lot1}** |")
    st.write(f"| 2 | {pair_info['s2']} | **{active_lot2}** |")
    st.caption(
        "💡 Tip: NSE revises lot sizes periodically. "
        "When NSE publishes a new circular, add the new sizes to LOT_REVISIONS."
    )

with c2:
    if lots_info:
        st.markdown("**Best cash-neutral lot combination**")
        st.write(
            f"Use **{lots_info['lots1']} × lot1** ({lots_info['shares1']} shares) of {pair_info['s1']}  \n"
            f"and **{lots_info['lots2']} × lot2** ({lots_info['shares2']} shares) of {pair_info['s2']}"
        )
        st.write(
            f"Value1 ≈ ₹{lots_info['value1']:,.0f} &nbsp;|&nbsp; "
            f"Value2 ≈ ₹{lots_info['value2']:,.0f}  \n"
            f"Imbalance: ₹{lots_info['imbalance']:,.0f} "
            f"({lots_info['imbalance_pct']}%)"
        )
        st.caption(
            f"Single-lot imbalance would be {lots_info['single_imbalance_pct']}% "
            f"(₹{lots_info['single_imbalance']:,.0f}). Multi-lot reduces residual risk."
        )

st.markdown("#### 💰 Profit Potential at Different Ratio Targets")
pnl_rows = profit_potential(m, s, lots_info)
if pnl_rows:
    st.dataframe(pd.DataFrame(pnl_rows), use_container_width=True, hide_index=True)
    st.caption("Estimates assume the recommended cash-neutral position. Not a guarantee.")
else:
    st.caption("Profit grid available when an ENTER signal and lot data are present.")

if pr:
    st.markdown("#### 🎯 Suggested Entry / Target / Stop")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{pair_info['s1']}** (Lot {active_lot1})")
        st.write(f"| Entry | Target | Stop | Support | Resistance |")
        st.write(f"|-------|--------|------|---------|------------|")
        st.write(f"| {pr['entry1']} | {pr['target1']} | {pr['stop1']} | {pr['support1']} | {pr['resist1']} |")
    with c2:
        st.markdown(f"**{pair_info['s2']}** (Lot {active_lot2})")
        st.write(f"| Entry | Target | Stop | Support | Resistance |")
        st.write(f"|-------|--------|------|---------|------------|")
        st.write(f"| {pr['entry2']} | {pr['target2']} | {pr['stop2']} | {pr['support2']} | {pr['resist2']} |")
    st.info(
        "Targets assume ratio mean-reversion. Stops sit beyond recent ratio extremes. "
        "Educational only — not financial advice."
    )

st.markdown("#### 📅 Last 5 Trading Days — Divergence Trend")
if m.get("last5_div") and m.get("last5_dates"):
    dots_html = []
    deltas = m.get("last5_div_delta") or [np.nan] * len(m["last5_div"])
    for d, div, dlt in zip(m["last5_dates"], m["last5_div"], deltas):
        if np.isnan(div):
            color, label = "#78909C", "●"
            val_str = "—"
        else:
            if not np.isnan(dlt) and dlt > 0.0005:
                color, label = "#00E676", "●"
            elif not np.isnan(dlt) and dlt < -0.0005:
                color, label = "#FF5252", "●"
            elif abs(div) > 0.001:
                color, label = ("#00E676", "●") if div > 0 else ("#FF5252", "●")
            else:
                color, label = "#ECEFF1", "●"
            val_str = f"{div:+.4f}"
        dots_html.append(
            f'<div style="display:inline-block;text-align:center;margin:0 12px;">'
            f'<span style="color:{color} !important;font-size:1.5rem;">{label}</span><br>'
            f'<small style="color:#90A4AE;">{d}</small><br>'
            f'<small style="color:#B0BEC5;">{val_str}</small></div>'
        )
    st.markdown("".join(dots_html), unsafe_allow_html=True)

    divs = [x for x in m["last5_div"] if not np.isnan(x)]
    if len(divs) >= 2:
        abs_start, abs_end = abs(divs[0]), abs(divs[-1])
        if abs_end > abs_start + 0.001:
            st.caption("Divergence is **growing** → stretch building.")
        elif abs_end < abs_start - 0.001:
            st.caption("Divergence is **shrinking** → mean-reversion in progress.")
        else:
            st.caption("Divergence is roughly **stable**.")
    st.caption(
        "🟢 Green = |divergence| growing · 🔴 Red = |divergence| shrinking · ⚪ White/Gray = flat"
    )
else:
    st.caption("Not enough history for 5-day divergence dots.")

# ── Charts (4 panels) ──
if show_charts and HAS_PLOTLY and not selected_result["hist"].empty:
    hist = selected_result["hist"]
    ratio_s = m["ratio_series"]
    z_s = m["zscore_series"]
    corr30_s = m.get("corr_roll_30")
    corr60_s = m.get("corr_roll_60")

    x_start = hist.index[0]
    x_end = hist.index[-1]
    n_days = (x_end - x_start).days

    fig = make_subplots(
        rows=4, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.06,
        subplot_titles=(
            f"Normalized Prices — {pair_info['s1']} vs {pair_info['s2']}",
            "Ratio (Price1 / Price2) + Reference Levels",
            "Z-Score of Ratio",
            "Rolling Correlation (30D & 60D)",
        ),
        row_heights=[0.28, 0.28, 0.22, 0.22],
    )

    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100
    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=pair_info["s1"],
                             line=dict(color="#00E5FF", width=1.5)), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=pair_info["s2"],
                             line=dict(color="#FF5252", width=1.5)), row=1, col=1)

    fig.add_trace(go.Scatter(x=ratio_s.index, y=ratio_s, name="Ratio",
                             line=dict(color="#69F0AE", width=1.8)), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00", row=2, col=1)
    fig.add_hline(y=m["ratio_max_30"], line_dash="dot", line_color="#FF8A80", row=2, col=1)
    fig.add_hline(y=m["ratio_min_30"], line_dash="dot", line_color="#80D8FF", row=2, col=1)
    fig.add_hline(y=m["ratio_max_60"], line_dash="dot", line_color="#FF1744", row=2, col=1)
    fig.add_hline(y=m["ratio_min_60"], line_dash="dot", line_color="#00B0FF", row=2, col=1)

    fig.add_trace(go.Scatter(x=z_s.index, y=z_s, name="Z-Score",
                             line=dict(color="#E040FB", width=1.8)), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676", row=3, col=1)
    fig.add_hline(y=0, line_color="#78909C", row=3, col=1)

    if corr30_s is not None and not corr30_s.empty:
        fig.add_trace(go.Scatter(x=corr30_s.index, y=corr30_s, name="Corr 30D",
                                 line=dict(color="#00E676", width=1.6)), row=4, col=1)
    if corr60_s is not None and not corr60_s.empty:
        fig.add_trace(go.Scatter(x=corr60_s.index, y=corr60_s, name="Corr 60D",
                                 line=dict(color="#FFD600", width=1.6)), row=4, col=1)
    fig.add_hline(y=0.80, line_dash="dash", line_color="#00C853", row=4, col=1)
    fig.add_hline(y=-0.80, line_dash="dash", line_color="#00C853", row=4, col=1)
    fig.add_hline(y=0, line_color="#78909C", row=4, col=1)

    fig.update_layout(
        height=980,
        template="plotly_dark",
        paper_bgcolor="#0E1117",
        plot_bgcolor="#1A1F2E",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#E0E0E0")),
        font=dict(color="#E0E0E0"),
        margin=dict(l=50, r=20, t=70, b=50),
        dragmode=False,
        hovermode="x unified",
    )
    fig.update_xaxes(
        title_text=f"Date  ({x_start.strftime('%d-%b-%Y')} → {x_end.strftime('%d-%b-%Y')}  ·  ~{n_days} calendar days / ~{len(hist)} trading days)",
        tickformat="%d-%b",
        nticks=8,
        showgrid=True,
        gridcolor="#37474F",
        row=4, col=1,
    )
    for r in [1, 2, 3]:
        fig.update_xaxes(showgrid=True, gridcolor="#37474F", row=r, col=1)

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={
            "displayModeBar": False,
            "scrollZoom": False,
            "doubleClick": False,
            "responsive": True,
            "staticPlot": False,
        },
    )

    # ---- FIXED f-string section (the previous crash) ----
    c30 = m["corr_30"]
    c60 = m["corr_60"]
    c90 = m["corr_90"]
    corr_ok = (not np.isnan(c30) and abs(c30) >= 0.80) and (not np.isnan(c60) and abs(c60) >= 0.80)
    corr_color = "#00E676" if corr_ok else "#FF8A80"
    corr_label = "PASS (≥0.80 both)" if corr_ok else "FAIL (need both ≥0.80)"

    c30_str = f"{c30:.2f}" if not np.isnan(c30) else "—"
    c60_str = f"{c60:.2f}" if not np.isnan(c60) else "—"
    c90_str = f"{c90:.2f}" if not np.isnan(c90) else "—"

    st.markdown(
        f"""
        <div style="background:#1A1F2E;border-radius:8px;padding:8px 12px;margin-top:-6px;font-size:0.78rem;line-height:1.55;color:#B0BEC5;">
        <div>
          <span style="color:#90CAF9;font-weight:600;">Period</span> &nbsp;
          {x_start.strftime('%d-%b-%Y')} → {x_end.strftime('%d-%b-%Y')}
          &nbsp;·&nbsp; ~{n_days} calendar days / {len(hist)} trading days
        </div>
        <div>
          <span style="color:#90CAF9;font-weight:600;">Normalized</span> &nbsp;
          <span style="color:#00E5FF;">{pair_info['s1'].replace('.NS','')}: {norm1.iloc[-1]:.1f}</span>
          &nbsp;|&nbsp;
          <span style="color:#FF5252;">{pair_info['s2'].replace('.NS','').replace('^','')}: {norm2.iloc[-1]:.1f}</span>
        </div>
        <div>
          <span style="color:#90CAF9;font-weight:600;">Ratio</span> &nbsp;
          <span style="color:#69F0AE;font-weight:600;">{m['ratio']:.4f}</span>
          &nbsp;·&nbsp; 60D Avg {m['ratio_mean']:.4f}
          &nbsp;·&nbsp; 30D Avg {m['ratio_avg_30']:.4f}
          &nbsp;·&nbsp; 30D Max <span style="color:#FF8A80;">{m['ratio_max_30']:.4f}</span>
          &nbsp;·&nbsp; 30D Min <span style="color:#80D8FF;">{m['ratio_min_30']:.4f}</span>
          &nbsp;·&nbsp; 60D Max <span style="color:#FF1744;">{m['ratio_max_60']:.4f}</span>
          &nbsp;·&nbsp; 60D Min <span style="color:#00B0FF;">{m['ratio_min_60']:.4f}</span>
        </div>
        <div>
          <span style="color:#90CAF9;font-weight:600;">Z-Score</span> &nbsp;
          <span style="color:#E040FB;font-weight:600;">{m['zscore']:.2f}</span>
          &nbsp;·&nbsp; Entry ±{z_entry}
          &nbsp;·&nbsp; Exit ±{z_exit}
        </div>
        <div>
          <span style="color:#90CAF9;font-weight:600;">Correlation</span> &nbsp;
          <span style="color:{corr_color};font-weight:600;">
            30D: {c30_str} &nbsp;·&nbsp;
            60D: {c60_str} &nbsp;·&nbsp;
            90D: {c90_str}
          </span>
          &nbsp;→&nbsp; <span style="color:{corr_color};font-weight:700;">{corr_label}</span>
        </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("##### 📖 Z-Score Trend Interpretation")
    st.markdown(interpret_z_trend(m.get("last5_z", [])))

# ─────────────────────────────────────────────────────────────
# INDEPENDENT ANALYTICS
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("🧠 Independent Analytics — High-Conviction Entries")

enter_notes = []
for r in results:
    if not r["metrics"]:
        continue
    s = r["signal"]
    m = r["metrics"]
    p = r["pair"]
    if "ENTER" in s["action"]:
        enter_notes.append({
            "name": p["name"],
            "action": s["action"],
            "conf": s["confidence"],
            "z": m["zscore"],
            "corr_30": m["corr_30"],
            "corr_60": m["corr_60"],
            "reason": s["reason"],
            "deep": s.get("deep_reason", ""),
        })

enter_notes.sort(key=lambda x: x["conf"], reverse=True)

if enter_notes:
    n = len(enter_notes)
    for idx, note in enumerate(enter_notes):
        frac = idx / max(n - 1, 1)
        bg_alpha = 0.28 - 0.20 * frac
        border_g = int(200 - 80 * frac)
        border_color = f"rgb(0,{border_g},83)"
        rank_label = f"#{idx + 1}"
        st.markdown(
            f"""
            <div class="conf-card" style="
                background: rgba(0, 200, 83, {bg_alpha:.2f});
                border-left-color: {border_color};
            ">
                <div style="font-weight:700;color:#E8F5E9 !important;margin-bottom:4px;">
                    {rank_label} &nbsp; {note['name']} &nbsp;—&nbsp; {note['action']}
                    &nbsp;(conf {note['conf']}%)
                </div>
                <div style="font-size:0.92rem;color:#C8E6C9 !important;">
                    Z={note['z']:.2f}, Corr30={note['corr_30']:.2f}, Corr60={note['corr_60']:.2f}. {note['reason']}
                </div>
                <div style="font-size:0.85rem;color:#A5D6A7 !important;margin-top:6px;font-style:italic;">
                    Deep rational: {note['deep']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
else:
    st.write("No high-conviction entry signals at the moment. Pairs are inside normal Z-score bands or fail the correlation filter (|Corr30| ≥ 0.80 AND |Corr60| ≥ 0.80).")

# ─────────────────────────────────────────────────────────────
# HISTORICAL WIN / LOSS
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📊 Historical Pair Performance (Simulated — Last ~12 Months)")

st.caption(
    "Simple back-test: enter when |Z| ≥ entry threshold, exit when |Z| ≤ 0.5 or after 30 trading days. "
    "P&L measured on the ratio. Past results do not guarantee future performance. "
    "Sorted by Win Rate % (highest first)."
)

hist_rows = []
for r in results:
    hs = r.get("hist_stats") or {}
    if not hs or hs.get("trades", 0) == 0:
        continue
    hist_rows.append({
        "Pair": r["pair"]["name"],
        "Opportunities": hs["trades"],
        "Wins": hs["wins"],
        "Win Rate %": hs["win_rate"],
        "Avg P&L % (ratio)": hs["avg_pnl_pct"],
    })

if hist_rows:
    df_hist = pd.DataFrame(hist_rows)
    df_hist = df_hist.sort_values("Win Rate %", ascending=False).reset_index(drop=True)
    st.dataframe(df_hist, use_container_width=True, hide_index=True)

    sel_hs = selected_result.get("hist_stats") or {}
    if sel_hs.get("details"):
        st.markdown(f"**Recent simulated trades for {selected}**")
        st.dataframe(pd.DataFrame(sel_hs["details"]), use_container_width=True, hide_index=True)
else:
    st.write("Not enough history to compute win-rate statistics.")

st.markdown("""
**Methodology (Independent engine):**
1. Fetch 1-year daily adjusted closes via Yahoo Finance (**single batched download**).
2. Ratio = Stock1 / Stock2.
3. Z-score = (Ratio − RollingMean) / RollingStd (default 60-day).
4. Enter when |Z| > threshold **and** |Corr30| ≥ 0.80 **and** |Corr60| ≥ 0.80.
5. Direction: positive Z → sell expensive leg / buy cheap leg.
6. Support / Resistance from 60-day high/low of each leg and of the ratio.
7. Half-life via AR(1) / Ornstein-Uhlenbeck regression.
8. Cash-neutrality: search integer lot multiples that minimise |Value1 − Value2|.
9. Historical stats: count mean-reversion opportunities and win rate over available history.
10. Lot sizes are month-aware (base Oct 2026 + LOT_REVISIONS map).
11. Correlation filter is applied both to live signals and to the high-conviction board.
12. Speed: one `yf.download` for all tickers + 5-minute cache → most refreshes are near-instant.

> ⚠️ Educational & research use only. Past mean-reversion does not guarantee future results. Always manage risk and position size carefully.
""")

# ─────────────────────────────────────────────────────────────
# AUTO REFRESH
# ─────────────────────────────────────────────────────────────
if auto_refresh:
    time.sleep(60)
    st.rerun()
