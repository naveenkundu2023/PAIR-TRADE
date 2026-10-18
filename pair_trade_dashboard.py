"""
Pair Trading Dashboard - Streamlit App (Enhanced v2)
====================================================
Automated pair-trading signal dashboard with dual confirmation
(Excel SIGNAL sheet + Independent analytics).

How to run:
    1. pip install streamlit pandas numpy yfinance plotly openpyxl scipy
    2. Place "PAIR TRADE -NEW (1).xlsx" in the same folder (or update EXCEL_PATH)
    3. streamlit run pair_trade_dashboard.py

Features:
    - Dark theme, high-contrast UI
    - Loads latest lot sizes + Excel signals from your workbook
    - Live prices via yfinance (NSE)
    - Correlation, Z-score, Ratio (current / 30D min-max / 60D min-max / avg)
    - Cash-neutrality lot recommendation
    - Profit potential at different ratio targets
    - Last-5-day divergence dots (green/red)
    - Z-score interpretation tooltip
    - Historical win-rate simulation (12 months)
    - Dual confirmation: Excel signal + Independent engine
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time
import os
import warnings
warnings.filterwarnings("ignore")

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

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

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
    .badge-confirm {
        background: linear-gradient(135deg, #00B0FF, #01579B);
        color: #fff; padding: 4px 10px; border-radius: 6px;
        font-weight: 600; font-size: 0.85rem;
    }
    .badge-excel {
        background: linear-gradient(135deg, #AB47BC, #4A148C);
        color: #fff; padding: 4px 10px; border-radius: 6px;
        font-weight: 600; font-size: 0.85rem;
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
    /* Divergence dots */
    .dot-green { color: #00E676; font-size: 1.3rem; }
    .dot-red   { color: #FF5252; font-size: 1.3rem; }
    .dot-gray  { color: #78909C; font-size: 1.3rem; }
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
    /* Horizontal rules */
    hr {
        border-color: #37474F !important;
    }
    /* Caption */
    .stCaption, small {
        color: #78909C !important;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# PATH TO EXCEL (user can change)
# ─────────────────────────────────────────────────────────────
EXCEL_PATH = os.environ.get(
    "PAIR_EXCEL_PATH",
    os.path.join(os.path.dirname(__file__) if "__file__" in dir() else ".", "PAIR TRADE -NEW (1).xlsx")
)
# Fallback paths
_CANDIDATES = [
    EXCEL_PATH,
    "/home/workdir/attachments/PAIR TRADE -NEW (1).xlsx",
    "PAIR TRADE -NEW (1).xlsx",
    os.path.join(os.getcwd(), "PAIR TRADE -NEW (1).xlsx"),
]

# ─────────────────────────────────────────────────────────────
# DEFAULT PAIRS (latest lot sizes from Excel Oct-2026)
# ─────────────────────────────────────────────────────────────
DEFAULT_PAIRS = [
    {"name": "ICICIBANK-HDFCBANK", "s1": "ICICIBANK.NS", "s2": "HDFCBANK.NS",
     "lot1": 700, "lot2": 550, "sector": "Banking", "excel_key": "ICICIBANK"},
    {"name": "AXISBANK-ICICIBANK", "s1": "AXISBANK.NS", "s2": "ICICIBANK.NS",
     "lot1": 625, "lot2": 700, "sector": "Banking", "excel_key": "AXISBANK"},
    {"name": "CIPLA-SUNPHARMA", "s1": "CIPLA.NS", "s2": "SUNPHARMA.NS",
     "lot1": 650, "lot2": 350, "sector": "Pharma", "excel_key": "CIPLA"},
    {"name": "BPCL-HINDPETRO", "s1": "BPCL.NS", "s2": "HINDPETRO.NS",
     "lot1": 1975, "lot2": 2025, "sector": "Oil & Gas", "excel_key": "BPCL"},
    {"name": "ACC-AMBUJACEM", "s1": "ACC.NS", "s2": "AMBUJACEM.NS",
     "lot1": 300, "lot2": 900, "sector": "Cement", "excel_key": "ACC"},
    {"name": "TCS-HCLTECH", "s1": "TCS.NS", "s2": "HCLTECH.NS",
     "lot1": 175, "lot2": 350, "sector": "IT", "excel_key": "TCS"},
    {"name": "TATASTEEL-JSWSTEEL", "s1": "TATASTEEL.NS", "s2": "JSWSTEEL.NS",
     "lot1": 5500, "lot2": 675, "sector": "Metals", "excel_key": "TATASTEEL"},
    {"name": "NESTLE-HINDUNILVR", "s1": "NESTLEIND.NS", "s2": "HINDUNILVR.NS",
     "lot1": 200, "lot2": 300, "sector": "FMCG", "excel_key": "NESTLEIND"},
    {"name": "NIFTY-BANKNIFTY", "s1": "^NSEI", "s2": "^NSEBANK",
     "lot1": 25, "lot2": 15, "sector": "Index", "excel_key": "INDEXNIFTY_50"},
    {"name": "BRITANNIA-HINDUNILVR", "s1": "BRITANNIA.NS", "s2": "HINDUNILVR.NS",
     "lot1": 200, "lot2": 300, "sector": "FMCG", "excel_key": "BRITANNIA"},
    {"name": "BAJFINANCE-BAJAJFINSV", "s1": "BAJFINANCE.NS", "s2": "BAJAJFINSV.NS",
     "lot1": 100, "lot2": 500, "sector": "Finance", "excel_key": "BAJFINANCE"},
    {"name": "HDFCBANK-BANKNIFTY", "s1": "HDFCBANK.NS", "s2": "^NSEBANK",
     "lot1": 550, "lot2": 15, "sector": "Banking", "excel_key": "HDFCBANK"},
    {"name": "RECLTD-PFC", "s1": "RECLTD.NS", "s2": "PFC.NS",
     "lot1": 2000, "lot2": 1300, "sector": "Finance", "excel_key": "RECLTD"},
    {"name": "IGL-GAIL", "s1": "IGL.NS", "s2": "GAIL.NS",
     "lot1": 1375, "lot2": 4575, "sector": "Oil & Gas", "excel_key": "IGL"},
    {"name": "ALKEM-AUROPHARMA", "s1": "ALKEM.NS", "s2": "AUROPHARMA.NS",
     "lot1": 100, "lot2": 550, "sector": "Pharma", "excel_key": "ALKEM"},
    {"name": "BERGEPAINT-ASIANPAINT", "s1": "BERGEPAINT.NS", "s2": "ASIANPAINT.NS",
     "lot1": 1320, "lot2": 200, "sector": "Paints", "excel_key": "BERGEPAINT"},
    {"name": "BANKBARODA-CANBK", "s1": "BANKBARODA.NS", "s2": "CANBK.NS",
     "lot1": 2925, "lot2": 6750, "sector": "Banking", "excel_key": "BANKBARODA"},
    {"name": "COALINDIA-GAIL", "s1": "COALINDIA.NS", "s2": "GAIL.NS",
     "lot1": 2925, "lot2": 6750, "sector": "Energy", "excel_key": "COALINDIA"},
]

# ─────────────────────────────────────────────────────────────
# LOAD EXCEL SIGNAL DATA
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_excel_signals():
    """Parse SIGNAL sheet for latest lot sizes, Excel z-score, ratio, Buy/Sell flags."""
    path = None
    for c in _CANDIDATES:
        if os.path.isfile(c):
            path = c
            break
    if not path or not HAS_OPENPYXL:
        return {}, "Excel not found – using built-in defaults"

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        # sheet name may have trailing space
        sheet_name = None
        for s in wb.sheetnames:
            if s.strip().upper() == "SIGNAL":
                sheet_name = s
                break
        if not sheet_name:
            return {}, "SIGNAL sheet not found"

        ws = wb[sheet_name]
        excel_map = {}
        row = 5
        while row < 80:
            a = ws.cell(row=row, column=1).value
            b = ws.cell(row=row, column=2).value
            if a is not None and b is not None and isinstance(a, (int, float)):
                lot1 = int(a)
                raw1 = str(b)
                s1 = (raw1.replace("NSE:", "")
                          .replace("INDEXNSE:", "")
                          .replace("MCX:", "")
                          .strip())
                a2 = ws.cell(row=row + 1, column=1).value
                b2 = ws.cell(row=row + 1, column=2).value
                if a2 is not None and b2 is not None:
                    lot2 = int(a2)
                    raw2 = str(b2)
                    s2 = (raw2.replace("NSE:", "")
                              .replace("INDEXNSE:", "")
                              .replace("MCX:", "")
                              .strip())
                    z_val = ws.cell(row=row, column=12).value   # L CURRENT z
                    ratio_val = ws.cell(row=row, column=15).value  # O
                    corr30 = ws.cell(row=row, column=5).value
                    avg_r = ws.cell(row=row, column=9).value
                    max60 = ws.cell(row=row, column=10).value
                    min30 = ws.cell(row=row, column=11).value
                    excel_buy = ws.cell(row=row, column=30).value  # AD
                    # also check individual leg signals
                    sell1 = str(ws.cell(row=row, column=3).value or "")
                    buy1 = str(ws.cell(row=row, column=4).value or "")
                    sell2 = str(ws.cell(row=row + 1, column=3).value or "")
                    buy2 = str(ws.cell(row=row + 1, column=4).value or "")

                    # Derive Excel pair action
                    excel_action = "NEUTRAL"
                    if excel_buy and "Buy" in str(excel_buy) and "Do Not" not in str(excel_buy):
                        excel_action = "BUY SIGNAL"
                    # Classic: one Buy + one Sell
                    leg1_buy = "BUY" in buy1.upper() and "DO NOT" not in buy1.upper()
                    leg1_sell = "SELL" in sell1.upper() and "DO NOT" not in sell1.upper()
                    leg2_buy = "BUY" in buy2.upper() and "DO NOT" not in buy2.upper()
                    leg2_sell = "SELL" in sell2.upper() and "DO NOT" not in sell2.upper()
                    if (leg1_buy and leg2_sell) or (leg1_sell and leg2_buy):
                        excel_action = "PAIR ENTRY"

                    key = s1.upper()
                    excel_map[key] = {
                        "lot1": lot1, "lot2": lot2,
                        "s1": s1, "s2": s2,
                        "excel_z": float(z_val) if z_val is not None else None,
                        "excel_ratio": float(ratio_val) if ratio_val is not None else None,
                        "excel_corr30": float(corr30) if corr30 is not None else None,
                        "excel_avg_ratio": float(avg_r) if avg_r is not None else None,
                        "excel_max60": float(max60) if max60 is not None else None,
                        "excel_min30": float(min30) if min30 is not None else None,
                        "excel_action": excel_action,
                        "excel_buy_flag": str(excel_buy or ""),
                        "leg1_signal": f"{sell1} / {buy1}",
                        "leg2_signal": f"{sell2} / {buy2}",
                    }
            row += 3
        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        return excel_map, f"Loaded from Excel ({path}) @ {ts}"
    except Exception as e:
        return {}, f"Excel load error: {e}"


def merge_excel_lots(pairs, excel_map):
    """Update DEFAULT_PAIRS lot sizes from Excel if available."""
    updated = []
    for p in pairs:
        key = p.get("excel_key", p["name"].split("-")[0]).upper()
        # also try bare ticker
        key2 = p["s1"].replace(".NS", "").replace("^", "").upper()
        info = excel_map.get(key) or excel_map.get(key2)
        new_p = dict(p)
        if info:
            new_p["lot1"] = info["lot1"]
            new_p["lot2"] = info["lot2"]
            new_p["excel"] = info
        else:
            new_p["excel"] = {}
        updated.append(new_p)
    return updated


# ─────────────────────────────────────────────────────────────
# DATA HELPERS
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=55, show_spinner=False)
def fetch_pair_history(ticker1: str, ticker2: str, period: str = "1y") -> pd.DataFrame:
    if not HAS_YF:
        return pd.DataFrame()
    try:
        h1 = yf.Ticker(ticker1).history(period=period, auto_adjust=True)
        h2 = yf.Ticker(ticker2).history(period=period, auto_adjust=True)
        if h1.empty or h2.empty:
            return pd.DataFrame()
        df = pd.DataFrame({"price1": h1["Close"], "price2": h2["Close"]}).dropna()
        df.index = pd.to_datetime(df.index).tz_localize(None)
        return df
    except Exception:
        return pd.DataFrame()


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

    # Ratio stats – 30D / 60D / average
    r30 = ratio.tail(30)
    r60 = ratio.tail(60)
    ratio_max_30 = float(r30.max())
    ratio_min_30 = float(r30.min())
    ratio_max_60 = float(r60.max())
    ratio_min_60 = float(r60.min())
    ratio_avg_30 = float(r30.mean())
    ratio_avg_60 = float(r60.mean())
    ratio_mean_60 = ratio_avg_60

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

    # Last 5 days divergence (ratio vs its 20-SMA)
    ratio_sma20 = ratio.rolling(20).mean()
    divergence = (ratio - ratio_sma20) / ratio_sma20
    last5_div = divergence.tail(5).tolist() if len(divergence) >= 5 else []
    last5_dates = [d.strftime("%d-%b") for d in divergence.tail(5).index] if len(divergence) >= 5 else []

    # Z-score trend (last 5 days)
    last5_z = zscore.tail(5).tolist() if len(zscore) >= 5 else []

    return {
        "ratio": cur_ratio,
        "zscore": cur_z,
        "corr_30": corr_30,
        "corr_60": corr_60,
        "corr_90": corr_90,
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
        "ratio_mean": ratio_mean_60,
        "half_life": half_life,
        "ratio_series": ratio,
        "zscore_series": zscore,
        "last5_div": last5_div,
        "last5_dates": last5_dates,
        "last5_z": last5_z,
        "df": df,
    }


def generate_signal(metrics: dict, z_entry: float = 1.5, z_exit: float = 0.5) -> dict:
    if not metrics:
        return {"action": "NO DATA", "leg1": "—", "leg2": "—", "confidence": 0,
                "reason": "Insufficient data", "deep_reason": ""}

    z = metrics["zscore"]
    corr = metrics.get("corr_30") or metrics.get("corr_60") or 0
    abs_z = abs(z)
    corr_ok = abs(corr) >= 0.5 if not np.isnan(corr) else False

    deep = ""
    if abs_z >= z_entry and corr_ok:
        if z > 0:
            action, leg1, leg2 = "ENTER PAIR", "SELL", "BUY"
            reason = (f"Z-score +{z:.2f} (ratio stretched high). "
                      f"Expect mean-reversion → Sell leg1 / Buy leg2")
            deep = (
                f"The price ratio is {abs_z:.1f} standard deviations above its "
                f"{60}-day mean. Historically, such stretches reverse toward the mean. "
                f"30-day correlation is {corr:.2f}, indicating the pair still moves together. "
                f"Selling the richer leg and buying the cheaper leg captures the reversion."
            )
        else:
            action, leg1, leg2 = "ENTER PAIR", "BUY", "SELL"
            reason = (f"Z-score {z:.2f} (ratio stretched low). "
                      f"Expect mean-reversion → Buy leg1 / Sell leg2")
            deep = (
                f"The price ratio is {abs_z:.1f} standard deviations below its "
                f"rolling mean. Mean-reversion probability is elevated. "
                f"Correlation ({corr:.2f}) supports pair quality. "
                f"Buying the cheaper leg and shorting the richer leg is the classic setup."
            )
        confidence = min(95, int(40 + abs_z * 15 + (abs(corr) * 20)))
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
        reason += f" | Correlation only {corr:.2f} – pair quality weak"
        confidence = max(10, confidence - 30)
        deep += " Low correlation reduces confidence in mean-reversion of the ratio."

    return {
        "action": action, "leg1": leg1, "leg2": leg2,
        "confidence": confidence, "reason": reason, "deep_reason": deep,
        "z": z, "corr": corr,
    }


def recommend_cash_neutral_lots(metrics: dict, lot1: int, lot2: int) -> dict:
    """Suggest lot multiples that minimise residual cash exposure."""
    if not metrics:
        return {}
    p1, p2 = metrics["price1"], metrics["price2"]
    # Value of 1 lot each
    v1 = p1 * lot1
    v2 = p2 * lot2
    # Ideal hedge ratio (shares of stock2 per share of stock1 for cash neutrality)
    # Using lot sizes: find integer n1, n2 minimising |n1*v1 - n2*v2|
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
    # Also single-lot reference
    best["single_value1"] = round(v1, 0)
    best["single_value2"] = round(v2, 0)
    best["single_imbalance"] = round(v1 - v2, 0)
    best["single_imbalance_pct"] = round(abs(v1 - v2) / max(v1, v2) * 100, 2)
    return best


def profit_potential(metrics: dict, signal: dict, lots_info: dict) -> list:
    """Estimate P&L if ratio reverts to various targets."""
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
    n1 = lots_info.get("lots1", 1)
    n2 = lots_info.get("lots2", 1)
    lot1 = lots_info.get("shares1", 0) // max(n1, 1) if n1 else 0
    # Use the recommended multi-lot position
    sh1 = lots_info.get("shares1", 0)
    sh2 = lots_info.get("shares2", 0)
    if sh1 == 0:
        return []

    rows = []
    direction = signal.get("leg1", "")
    for label, tgt_r in targets:
        if tgt_r <= 0 or cur_r <= 0:
            continue
        # Approximate: assume leg2 price fixed, leg1 moves with ratio
        # (or vice-versa). Use average of both assumptions for robustness.
        # Method A: p2 fixed → p1_new = tgt_r * p2
        p1_new_a = tgt_r * p2
        # Method B: p1 fixed → p2_new = p1 / tgt_r
        p2_new_b = p1 / tgt_r

        if direction == "SELL":
            # Short leg1, Long leg2
            pnl_a = sh1 * (p1 - p1_new_a) + sh2 * 0   # only leg1 moves
            pnl_b = sh1 * 0 + sh2 * (p2_new_b - p2)
        elif direction == "BUY":
            pnl_a = sh1 * (p1_new_a - p1) + sh2 * 0
            pnl_b = sh1 * 0 + sh2 * (p2 - p2_new_b)
        else:
            pnl_a = pnl_b = 0

        # Average of two assumptions
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
    """Simulate mean-reversion trades over the available history (~12 months)."""
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
            # Look forward up to 30 days for reversion to |Z|<0.5 or 30d timeout
            exit_i = None
            for j in range(i + 1, min(i + 31, len(z))):
                if abs(z.iloc[j]) <= 0.5:
                    exit_i = j
                    break
            if exit_i is None:
                exit_i = min(i + 30, len(z) - 1)
            exit_ratio = ratio.iloc[exit_i]
            # P&L direction: if entry Z>0 we short ratio (profit when ratio falls)
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
        "details": trades[-12:],  # last 12 trades
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
# SIDEBAR (horizontal-friendly controls, no vertical-only clutter)
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
    st.caption("Data: Yahoo Finance + your Excel SIGNAL sheet")
    st.caption("Independent engine confirms / challenges Excel signals")
    st.caption(f"Clock: {datetime.now().strftime('%H:%M:%S')} IST")

    if st.button("🔄 Refresh Now"):
        st.cache_data.clear()
        st.rerun()


# ─────────────────────────────────────────────────────────────
# LOAD EXCEL + MERGE
# ─────────────────────────────────────────────────────────────
excel_map, excel_status = load_excel_signals()
ALL_PAIRS = merge_excel_lots(DEFAULT_PAIRS, excel_map)

# ─────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────
st.markdown('<p class="main-header">📈 Automated Pair Trading Dashboard</p>', unsafe_allow_html=True)
st.markdown(
    f'<p class="sub-header">Dual confirmation (Excel + Independent) • Dark theme • '
    f'Updated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} IST</p>',
    unsafe_allow_html=True,
)
st.caption(f"📂 {excel_status}")

if not HAS_YF:
    st.error("yfinance is not installed. Run: `pip install yfinance`")
    st.stop()

# ─────────────────────────────────────────────────────────────
# SCAN ALL PAIRS
# ─────────────────────────────────────────────────────────────
pairs_to_scan = [p for p in ALL_PAIRS if p["sector"] in sector_filter]

progress = st.progress(0, text="Fetching live data…")
results = []

for i, pair in enumerate(pairs_to_scan):
    progress.progress((i + 1) / len(pairs_to_scan), text=f"Scanning {pair['name']}…")
    hist = fetch_pair_history(pair["s1"], pair["s2"], period="1y")
    metrics = compute_pair_metrics(hist, lookback=lookback)
    signal = generate_signal(metrics, z_entry=z_entry, z_exit=z_exit)
    prices = recommend_prices(metrics, signal)
    lots_info = recommend_cash_neutral_lots(metrics, pair["lot1"], pair["lot2"]) if metrics else {}
    hist_stats = historical_pair_stats(hist, z_entry=z_entry, lookback=lookback) if not hist.empty else {}

    results.append({
        "pair": pair,
        "metrics": metrics,
        "signal": signal,
        "prices": prices,
        "lots_info": lots_info,
        "hist_stats": hist_stats,
        "hist": hist,
    })

progress.empty()

# ─────────────────────────────────────────────────────────────
# SUMMARY TABLE – all pairs visible (no vertical slider issue)
# ─────────────────────────────────────────────────────────────
st.subheader("🎯 Live Signal Board — All Pairs")

rows = []
for r in results:
    m = r["metrics"]
    s = r["signal"]
    p = r["pair"]
    if not m:
        continue
    excel_info = p.get("excel", {})
    excel_act = excel_info.get("excel_action", "—")
    # Dual confirmation
    dual = ""
    if "ENTER" in s["action"] and excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
        dual = "✅ DUAL CONFIRM"
    elif "ENTER" in s["action"]:
        dual = "⚡ Independent only"
    elif excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
        dual = "📋 Excel only"
    else:
        dual = "—"

    rows.append({
        "Pair": p["name"],
        "Sector": p["sector"],
        "Lot1": p["lot1"],
        "Lot2": p["lot2"],
        "Price1": round(m["price1"], 2),
        "Price2": round(m["price2"], 2),
        "Ratio": round(m["ratio"], 4),
        "Z-Score": round(m["zscore"], 2),
        "Corr30": round(m["corr_30"], 2) if not np.isnan(m["corr_30"]) else "—",
        "Action": s["action"],
        "Leg1": s["leg1"],
        "Leg2": s["leg2"],
        "Conf%": s["confidence"],
        "Excel": excel_act,
        "Dual": dual,
        "HalfLife": round(m["half_life"], 1) if not np.isnan(m.get("half_life", np.nan)) else "—",
    })

if not rows:
    st.warning("No data returned. Check internet / ticker symbols.")
    st.stop()

df_signals = pd.DataFrame(rows)

def color_action(val):
    v = str(val)
    if "ENTER" in v:
        return "background-color: #1B5E20; color: #FFFFFF; font-weight: bold"
    if "EXIT" in v or "HOLD" in v:
        return "background-color: #37474F; color: #FFFFFF"
    if "WATCH" in v:
        return "background-color: #F9A825; color: #000000; font-weight: bold"
    return ""

def color_dual(val):
    if "DUAL" in str(val):
        return "background-color: #01579B; color: #FFFFFF; font-weight: bold"
    if "Independent" in str(val):
        return "background-color: #00695C; color: #FFFFFF"
    if "Excel" in str(val):
        return "background-color: #6A1B9A; color: #FFFFFF"
    return ""

styled = (df_signals.style
          .map(color_action, subset=["Action"])
          .map(color_dual, subset=["Dual"])
          .format({
              "Price1": "{:.2f}", "Price2": "{:.2f}",
              "Ratio": "{:.4f}", "Z-Score": "{:.2f}",
          }))

# Height large enough to show all ~18 pairs without vertical scrollbar frustration
st.dataframe(styled, use_container_width=True, height=680)

enter_count = sum(1 for r in rows if "ENTER" in r["Action"])
dual_count = sum(1 for r in rows if "DUAL" in r["Dual"])
watch_count = sum(1 for r in rows if "WATCH" in r["Action"])
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Pairs Scanned", len(rows))
col2.metric("🟢 Enter Signals", enter_count)
col3.metric("✅ Dual Confirm", dual_count)
col4.metric("🟡 Watchlist", watch_count)
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
excel_info = pair_info.get("excel", {})

# KPI row
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Current Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}",
          delta="Stretched" if abs(m["zscore"]) > z_entry else "Normal")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if not np.isnan(m["corr_60"]) else "—")
k5.metric("Half-Life (d)", f"{m['half_life']:.1f}" if not np.isnan(m.get("half_life", np.nan)) else "—")

# Signal box
if "ENTER" in s["action"]:
    action_color = "#1B5E20"
elif "WATCH" in s["action"]:
    action_color = "#F9A825"
else:
    action_color = "#37474F"

excel_line = ""
if excel_info:
    excel_line = (f'<p style="margin:6px 0 0 0;opacity:0.95;">'
                  f'<span class="badge-excel">Excel: {excel_info.get("excel_action", "—")}</span> '
                  f'&nbsp; Z(excel)={excel_info.get("excel_z")} &nbsp; '
                  f'Flag: {excel_info.get("excel_buy_flag", "")}</p>')

dual_note = ""
if "ENTER" in s["action"] and excel_info.get("excel_action") in ("BUY SIGNAL", "PAIR ENTRY"):
    dual_note = '<p style="margin:8px 0 0 0;"><span class="badge-confirm">✅ DUAL CONFIRMATION — Excel + Independent agree</span></p>'

st.markdown(
    f"""
    <div style="background:{action_color};color:white;padding:18px;border-radius:12px;margin:12px 0;">
        <h3 style="margin:0;color:white;">{s['action']} — Confidence {s['confidence']}%</h3>
        <p style="margin:8px 0 0 0;font-size:1.1rem;">
            <b>{pair_info['s1']}</b> → <b>{s['leg1']}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>{pair_info['s2']}</b> → <b>{s['leg2']}</b>
        </p>
        <p style="margin:8px 0 0 0;opacity:0.95;">{s['reason']}</p>
        {excel_line}
        {dual_note}
    </div>
    """,
    unsafe_allow_html=True,
)

# Deep rational
if s.get("deep_reason"):
    st.markdown("#### 🧠 Independent Deep Rational")
    st.info(s["deep_reason"])

# ── Ratio reference box (current / 30D / 60D / avg) ──
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
        </tr>
        <tr>
            <td><b>Ratio</b></td>
            <td style="color:#00E5FF;font-weight:700;">{m['ratio']:.4f}</td>
            <td>—</td>
            <td>—</td>
        </tr>
        <tr>
            <td>Average</td>
            <td>—</td>
            <td>{m['ratio_avg_30']:.4f}</td>
            <td>{m['ratio_avg_60']:.4f}</td>
        </tr>
        <tr>
            <td>Maximum</td>
            <td>—</td>
            <td style="color:#FF8A80;">{m['ratio_max_30']:.4f}</td>
            <td style="color:#FF8A80;">{m['ratio_max_60']:.4f}</td>
        </tr>
        <tr>
            <td>Minimum</td>
            <td>—</td>
            <td style="color:#69F0AE;">{m['ratio_min_30']:.4f}</td>
            <td style="color:#69F0AE;">{m['ratio_min_60']:.4f}</td>
        </tr>
    </table>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Lot sizes + Cash Neutrality ──
st.markdown("#### 📦 Lot Sizes & Cash-Neutrality Recommendation")
c1, c2 = st.columns(2)
with c1:
    st.markdown(f"**Exchange lot sizes (from Excel / latest)**")
    st.write(f"| Leg | Ticker | Lot Size |")
    st.write(f"|-----|--------|----------|")
    st.write(f"| 1 | {pair_info['s1']} | **{pair_info['lot1']}** |")
    st.write(f"| 2 | {pair_info['s2']} | **{pair_info['lot2']}** |")
    st.caption("💡 Tip: On month-end, re-check NSE lot-size circulars and update the Excel SIGNAL sheet. "
               "This dashboard re-reads the sheet on every refresh.")

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

# ── Profit potential ──
st.markdown("#### 💰 Profit Potential at Different Ratio Targets")
pnl_rows = profit_potential(m, s, lots_info)
if pnl_rows:
    st.dataframe(pd.DataFrame(pnl_rows), use_container_width=True, hide_index=True)
    st.caption("Estimates assume the recommended cash-neutral position. "
               "Actual P&L depends on which leg moves. Not a guarantee.")
else:
    st.caption("Profit grid available when an ENTER signal and lot data are present.")

# ── Entry / Target / Stop ──
if pr:
    st.markdown("#### 🎯 Suggested Entry / Target / Stop")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{pair_info['s1']}** (Lot {pair_info['lot1']})")
        st.write(f"| Entry | Target | Stop | Support | Resistance |")
        st.write(f"|-------|--------|------|---------|------------|")
        st.write(f"| {pr['entry1']} | {pr['target1']} | {pr['stop1']} | {pr['support1']} | {pr['resist1']} |")
    with c2:
        st.markdown(f"**{pair_info['s2']}** (Lot {pair_info['lot2']})")
        st.write(f"| Entry | Target | Stop | Support | Resistance |")
        st.write(f"|-------|--------|------|---------|------------|")
        st.write(f"| {pr['entry2']} | {pr['target2']} | {pr['stop2']} | {pr['support2']} | {pr['resist2']} |")
    st.info(
        "Targets assume ratio mean-reversion. Stops sit beyond recent ratio extremes. "
        "Size for approximate cash neutrality using the lot recommendation above. "
        "Educational only — not financial advice."
    )

# ── Last 5 days divergence dots ──
st.markdown("#### 📅 Last 5 Trading Days — Divergence Trend")
if m.get("last5_div") and m.get("last5_dates"):
    dots_html = []
    for d, div in zip(m["last5_dates"], m["last5_div"]):
        if np.isnan(div):
            cls, label = "dot-gray", "·"
        elif div > 0.005:
            cls, label = "dot-green", "●"
        elif div < -0.005:
            cls, label = "dot-red", "●"
        else:
            cls, label = "dot-gray", "●"
        dots_html.append(f'<span class="{cls}" title="{d}: {div:.4f}">{label}</span> {d}')
    st.markdown(" &nbsp;&nbsp; ".join(dots_html), unsafe_allow_html=True)
    # Trend description
    divs = [x for x in m["last5_div"] if not np.isnan(x)]
    if len(divs) >= 2:
        if divs[-1] > divs[0] + 0.005:
            st.caption("Divergence is **rising** (ratio moving away from its 20-SMA) → stretch building.")
        elif divs[-1] < divs[0] - 0.005:
            st.caption("Divergence is **falling** (ratio moving back toward its 20-SMA) → mean-reversion in progress.")
        else:
            st.caption("Divergence is roughly **stable** over the last 5 sessions.")
    st.caption("Green = ratio above 20-SMA (positive divergence) · Red = ratio below 20-SMA · Gray = near zero")
else:
    st.caption("Not enough history for 5-day divergence dots.")

# ── Charts ──
if show_charts and HAS_PLOTLY and not selected_result["hist"].empty:
    hist = selected_result["hist"]
    ratio_s = m["ratio_series"]
    z_s = m["zscore_series"]

    fig = make_subplots(
        rows=3, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.07,
        subplot_titles=(
            f"Normalized Prices — {pair_info['s1']} vs {pair_info['s2']}",
            "Ratio (Price1 / Price2) + Reference Levels",
            "Z-Score of Ratio",
        ),
        row_heights=[0.35, 0.35, 0.30],
    )

    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100
    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=pair_info["s1"],
                             line=dict(color="#00E5FF", width=1.5)), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=pair_info["s2"],
                             line=dict(color="#FF5252", width=1.5)), row=1, col=1)

    # Ratio + levels
    fig.add_trace(go.Scatter(x=ratio_s.index, y=ratio_s, name="Ratio",
                             line=dict(color="#69F0AE", width=1.8)), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00",
                  annotation_text="60D Avg", row=2, col=1)
    fig.add_hline(y=m["ratio_max_30"], line_dash="dot", line_color="#FF8A80",
                  annotation_text="30D Max", row=2, col=1)
    fig.add_hline(y=m["ratio_min_30"], line_dash="dot", line_color="#80D8FF",
                  annotation_text="30D Min", row=2, col=1)
    fig.add_hline(y=m["ratio_max_60"], line_dash="dot", line_color="#FF1744",
                  annotation_text="60D Max", row=2, col=1)
    fig.add_hline(y=m["ratio_min_60"], line_dash="dot", line_color="#00B0FF",
                  annotation_text="60D Min", row=2, col=1)

    # Z-score
    fig.add_trace(go.Scatter(x=z_s.index, y=z_s, name="Z-Score",
                             line=dict(color="#E040FB", width=1.8)), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676", row=3, col=1)
    fig.add_hline(y=0, line_color="#78909C", row=3, col=1)

    fig.update_layout(
        height=780,
        template="plotly_dark",
        paper_bgcolor="#0E1117",
        plot_bgcolor="#1A1F2E",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color="#E0E0E0")),
        font=dict(color="#E0E0E0"),
    )
    st.plotly_chart(fig, use_container_width=True)

    # Z-score interpretation
    st.markdown("##### 📖 Z-Score Trend Interpretation")
    st.markdown(interpret_z_trend(m.get("last5_z", [])))

# ─────────────────────────────────────────────────────────────
# HOW EXCEL IS USED (transparency)
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📂 How Your Excel Sheet Is Used")
st.markdown("""
This dashboard is designed for **dual-confirmation** trading:

1. **Lot sizes** – Read from column A of the SIGNAL sheet on every refresh.  
   When NSE revises lots (usually month-end), update the Excel and the dashboard picks them up automatically.
2. **Excel signals** – Columns C/D (SELL/BUY per leg) and column AD (Buy / Do Not buy) are parsed.  
   If Excel says “Buy” / “PAIR ENTRY”, it is shown side-by-side with the independent engine.
3. **Excel Z / Ratio / Correlation** – Displayed for reference so you can compare the sheet’s numbers with live Yahoo Finance calculations.
4. **Independent engine** – Recomputes ratio, Z-score, correlation, half-life, support/resistance and mean-reversion signals from scratch using 1-year daily data.
5. **Dual confirmation badge** – Appears only when **both** Excel and the independent engine flash an entry.  
   That is the highest-conviction setup this tool can produce.

> Keep the Excel file named `PAIR TRADE -NEW (1).xlsx` in the same folder as the script (or set the `PAIR_EXCEL_PATH` environment variable).
""")

# ─────────────────────────────────────────────────────────────
# INDEPENDENT ANALYTICS NOTES
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("🧠 Independent Analytics — High-Conviction Entries")

notes = []
for r in results:
    if not r["metrics"]:
        continue
    s = r["signal"]
    m = r["metrics"]
    p = r["pair"]
    if "ENTER" in s["action"]:
        excel_tag = ""
        if p.get("excel", {}).get("excel_action") in ("BUY SIGNAL", "PAIR ENTRY"):
            excel_tag = " ✅ **DUAL CONFIRM with Excel**"
        notes.append(
            f"**{p['name']}** — {s['action']} (conf {s['confidence']}%). "
            f"Z={m['zscore']:.2f}, Corr30={m['corr_30']:.2f}. {s['reason']}{excel_tag}\n\n"
            f"&nbsp;&nbsp;&nbsp;*Deep rational:* {s.get('deep_reason', '')}"
        )

if notes:
    for n in notes:
        st.markdown(f"- {n}")
else:
    st.write("No high-conviction entry signals at the moment. Pairs are inside normal Z-score bands.")

# ─────────────────────────────────────────────────────────────
# HISTORICAL WIN / LOSS SECTION
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📊 Historical Pair Performance (Simulated — Last ~12 Months)")

st.caption(
    "Simple back-test: enter when |Z| ≥ entry threshold, exit when |Z| ≤ 0.5 or after 30 trading days. "
    "P&L measured on the ratio. Past results do not guarantee future performance."
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
    st.dataframe(pd.DataFrame(hist_rows), use_container_width=True, hide_index=True)

    # Detail for selected pair
    sel_hs = selected_result.get("hist_stats") or {}
    if sel_hs.get("details"):
        st.markdown(f"**Recent simulated trades for {selected}**")
        st.dataframe(pd.DataFrame(sel_hs["details"]), use_container_width=True, hide_index=True)
else:
    st.write("Not enough history to compute win-rate statistics.")

st.markdown("""
**Methodology (Independent of original Excel formulas):**
1. Fetch 1-year daily adjusted closes via Yahoo Finance.
2. Ratio = Stock1 / Stock2.
3. Z-score = (Ratio − RollingMean) / RollingStd (default 60-day).
4. Enter when |Z| > threshold **and** |correlation| ≥ 0.5.
5. Direction: positive Z → sell expensive leg / buy cheap leg.
6. Support / Resistance from 60-day high/low of each leg and of the ratio.
7. Half-life via AR(1) / Ornstein-Uhlenbeck regression.
8. Cash-neutrality: search integer lot multiples that minimise |Value1 − Value2|.
9. Historical stats: count mean-reversion opportunities and win rate over available history.
10. Excel SIGNAL sheet supplies lot sizes + its own Buy/Sell flags for dual confirmation.

> ⚠️ Educational & research use only. Past mean-reversion does not guarantee future results. Always manage risk and position size carefully.
""")

# ─────────────────────────────────────────────────────────────
# AUTO REFRESH
# ─────────────────────────────────────────────────────────────
if auto_refresh:
    time.sleep(60)
    st.rerun()
