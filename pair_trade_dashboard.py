"""
Pair Trading Dashboard - Streamlit App (Enhanced v3)
====================================================
Automated pair-trading signal dashboard with dual confirmation
(Excel SIGNAL sheet + Independent analytics).

How to run (local or Streamlit Cloud / GitHub):
    1. pip install streamlit pandas numpy yfinance plotly openpyxl scipy
    2. Place "PAIR TRADE -NEW (1).xlsx" in the same folder as this script
       (or set env PAIR_EXCEL_PATH). On GitHub / Streamlit Cloud the file
       must be committed in the repo – the dashboard reads it at runtime.
    3. streamlit run pair_trade_dashboard_v3.py

Data source confirmation
------------------------
• Live prices, ratio, Z-score, correlation, charts → always fetched fresh
  from Yahoo Finance (yfinance) on every refresh / page load.
• Lot sizes, Excel Buy/Sell flags, Excel Z/ratio → read from the Excel
  file that is present in the working directory (the file you committed
  to GitHub). There is NO dependence on a desktop path once the file
  lives in the repo. When you update the Excel and push to GitHub,
  Streamlit Cloud will pick up the new file on the next deploy / restart.
• The timestamp shown on the dashboard is the moment the Excel was last
  successfully parsed + the moment live prices were pulled.

v3 changes vs previous version
------------------------------
1. Full SIGNAL sheet rendered at the top (horizontal scroll only).
2. Mobile-friendly sidebar collapse arrows (dark high-contrast).
3. Live Signal Board shows only ENTER / EXIT-HOLD (no WATCH rows);
   whole-row colour highlight, smaller font.
4. Metric cards (Pairs Scanned etc.) reduced in size.
5. Deep-dive KPI boxes slightly tighter.
6. Richer Independent Deep Rational + extra analytics for selected pair.
7. Cash-neutrality section shows recommended multi-lot + explicit 1:1 example.
8. Profit potential uses the recommended cash-neutral position (₹).
9. Last-5-day divergence dots fixed + numeric divergence values shown.
10. Charts: mobile drag/zoom disabled, ratio-level colour legend,
    ratio chart legend moved to bottom, Z-score analytics + historical best Z.
11. Historical performance section more colourful.
12. Refresh button at top + optional auto-refresh every 60 s.
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
# PAGE CONFIG + DARK THEME (desktop + mobile)
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
        font-size: 2.1rem;
        font-weight: 800;
        background: linear-gradient(90deg, #00E5FF, #76FF03, #FFEA00);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 0.9rem;
        color: #90A4AE !important;
        margin-bottom: 0.8rem;
    }

    /* Compact metric cards */
    [data-testid="stMetric"] {
        background: #1A1F2E;
        border-radius: 8px;
        padding: 6px 10px !important;
        border-left: 3px solid #00E5FF;
        box-shadow: 0 1px 4px rgba(0,0,0,0.35);
        min-height: 60px;
    }
    [data-testid="stMetric"] label {
        color: #90CAF9 !important;
        font-size: 0.72rem !important;
    }
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 1.15rem !important;
    }
    [data-testid="stMetric"] [data-testid="stMetricDelta"] {
        font-size: 0.75rem !important;
    }

    /* Signal badges */
    .badge-enter {
        background: linear-gradient(135deg, #00C853, #1B5E20);
        color: #fff; padding: 5px 12px; border-radius: 7px;
        font-weight: 700; font-size: 0.9rem; display: inline-block;
    }
    .badge-watch {
        background: linear-gradient(135deg, #FFD600, #F57F17);
        color: #000; padding: 5px 12px; border-radius: 7px;
        font-weight: 700; font-size: 0.9rem; display: inline-block;
    }
    .badge-exit {
        background: linear-gradient(135deg, #546E7A, #263238);
        color: #fff; padding: 5px 12px; border-radius: 7px;
        font-weight: 700; font-size: 0.9rem; display: inline-block;
    }
    .badge-confirm {
        background: linear-gradient(135deg, #00B0FF, #01579B);
        color: #fff; padding: 3px 9px; border-radius: 5px;
        font-weight: 600; font-size: 0.8rem;
    }
    .badge-excel {
        background: linear-gradient(135deg, #AB47BC, #4A148C);
        color: #fff; padding: 3px 9px; border-radius: 5px;
        font-weight: 600; font-size: 0.8rem;
    }

    /* Ratio reference table */
    .ratio-box {
        background: #1A1F2E;
        border: 1px solid #37474F;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 6px 0 12px 0;
        font-size: 0.85rem;
    }
    .ratio-box table { width: 100%; border-collapse: collapse; color: #E0E0E0; }
    .ratio-box th {
        background: #263238; color: #80D8FF;
        padding: 6px 10px; text-align: center; font-size: 0.8rem;
    }
    .ratio-box td {
        padding: 5px 10px; text-align: center;
        border-bottom: 1px solid #37474F; font-size: 0.85rem;
    }

    /* Divergence dots */
    .dot-green { color: #00E676; font-size: 1.25rem; }
    .dot-red   { color: #FF5252; font-size: 1.25rem; }
    .dot-gray  { color: #78909C; font-size: 1.25rem; }
    .div-val   { font-size: 0.72rem; color: #B0BEC5; }

    /* Sidebar – dark + mobile-friendly collapse control */
    [data-testid="stSidebar"] {
        background-color: #0D1117 !important;
    }
    [data-testid="stSidebar"] * {
        color: #E0E0E0 !important;
    }
    /* Make the collapse / expand chevron high-contrast on mobile */
    [data-testid="stSidebarCollapsedControl"],
    button[kind="header"],
    [data-testid="collapsedControl"] {
        background-color: #1A237E !important;
        color: #FFFFFF !important;
        border: 2px solid #00E5FF !important;
        border-radius: 8px !important;
        padding: 4px 8px !important;
    }
    [data-testid="stSidebarCollapsedControl"] svg,
    button[kind="header"] svg {
        fill: #00E5FF !important;
        color: #00E5FF !important;
        width: 1.4rem !important;
        height: 1.4rem !important;
    }

    /* Dataframe */
    .stDataFrame { background: #1A1F2E; }
    .stAlert { background-color: #1A2332 !important; color: #E0E0E0 !important; }

    /* Selectbox / inputs */
    .stSelectbox label, .stSlider label, .stMultiSelect label {
        color: #90CAF9 !important;
    }
    hr { border-color: #37474F !important; }
    .stCaption, small { color: #78909C !important; }

    /* Compact deep-dive signal box */
    .signal-box {
        padding: 12px 16px !important;
        border-radius: 10px;
        margin: 8px 0;
    }

    /* Excel SIGNAL full-sheet container – horizontal scroll only */
    .excel-signal-wrap {
        overflow-x: auto;
        overflow-y: hidden;
        max-height: none;
        border: 1px solid #37474F;
        border-radius: 8px;
        background: #12161F;
        padding: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# PATH TO EXCEL (works on desktop, GitHub, Streamlit Cloud)
# ─────────────────────────────────────────────────────────────
EXCEL_PATH = os.environ.get(
    "PAIR_EXCEL_PATH",
    os.path.join(os.path.dirname(__file__) if "__file__" in dir() else ".", "PAIR TRADE -NEW (1).xlsx")
)
_CANDIDATES = [
    EXCEL_PATH,
    "/home/workdir/attachments/PAIR TRADE -NEW (1).xlsx",
    "PAIR TRADE -NEW (1).xlsx",
    os.path.join(os.getcwd(), "PAIR TRADE -NEW (1).xlsx"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "PAIR TRADE -NEW (1).xlsx"),
]

# ─────────────────────────────────────────────────────────────
# DEFAULT PAIRS (lot sizes overridden by Excel when available)
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
# LOAD FULL SIGNAL SHEET (for top display) + parse map
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=120, show_spinner=False)
def load_excel_full_and_map():
    """Return (full_signal_df for display, excel_map for lots/signals, status string)."""
    path = None
    for c in _CANDIDATES:
        if os.path.isfile(c):
            path = c
            break
    if not path or not HAS_OPENPYXL:
        return pd.DataFrame(), {}, "Excel not found – using built-in defaults"

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet_name = None
        for s in wb.sheetnames:
            if s.strip().upper() == "SIGNAL":
                sheet_name = s
                break
        if not sheet_name:
            return pd.DataFrame(), {}, "SIGNAL sheet not found"

        ws = wb[sheet_name]

        # ── Build a clean display table of the active pairs region ──
        # Headers (row 3 in 1-indexed sheet)
        headers = []
        for c in range(1, min(32, ws.max_column + 1)):
            v = ws.cell(row=3, column=c).value
            headers.append(str(v).strip() if v is not None else f"Col{c}")

        rows_data = []
        excel_map = {}
        row = 5
        while row < 80:
            a = ws.cell(row=row, column=1).value
            b = ws.cell(row=row, column=2).value
            if a is not None and b is not None and isinstance(a, (int, float)):
                # Leg 1 row
                leg1 = [ws.cell(row=row, column=c).value for c in range(1, len(headers) + 1)]
                rows_data.append(leg1)
                # Leg 2 row
                leg2 = [ws.cell(row=row + 1, column=c).value for c in range(1, len(headers) + 1)]
                rows_data.append(leg2)
                # Blank separator for readability
                rows_data.append([None] * len(headers))

                lot1 = int(a)
                raw1 = str(b)
                s1 = (raw1.replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip())
                a2 = ws.cell(row=row + 1, column=1).value
                b2 = ws.cell(row=row + 1, column=2).value
                if a2 is not None and b2 is not None:
                    lot2 = int(a2)
                    raw2 = str(b2)
                    s2 = (raw2.replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip())
                    z_val = ws.cell(row=row, column=12).value
                    ratio_val = ws.cell(row=row, column=15).value
                    corr30 = ws.cell(row=row, column=5).value
                    avg_r = ws.cell(row=row, column=9).value
                    max60 = ws.cell(row=row, column=10).value
                    min30 = ws.cell(row=row, column=11).value
                    excel_buy = ws.cell(row=row, column=30).value
                    sell1 = str(ws.cell(row=row, column=3).value or "")
                    buy1 = str(ws.cell(row=row, column=4).value or "")
                    sell2 = str(ws.cell(row=row + 1, column=3).value or "")
                    buy2 = str(ws.cell(row=row + 1, column=4).value or "")

                    excel_action = "NEUTRAL"
                    if excel_buy and "Buy" in str(excel_buy) and "Do Not" not in str(excel_buy):
                        excel_action = "BUY SIGNAL"
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

        # Build DataFrame – keep only useful columns, format nicely
        useful_cols = list(range(min(20, len(headers))))  # first ~20 cols
        display_headers = [headers[i] if i < len(headers) else f"C{i}" for i in useful_cols]
        clean_rows = []
        for r in rows_data:
            clean_rows.append([r[i] if i < len(r) else None for i in useful_cols])

        df_signal = pd.DataFrame(clean_rows, columns=display_headers)
        # Drop completely empty rows for cleaner look
        df_signal = df_signal.dropna(how="all")

        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        status = f"Loaded from Excel ({os.path.basename(path)}) @ {ts}"
        return df_signal, excel_map, status
    except Exception as e:
        return pd.DataFrame(), {}, f"Excel load error: {e}"


def merge_excel_lots(pairs, excel_map):
    updated = []
    for p in pairs:
        key = p.get("excel_key", p["name"].split("-")[0]).upper()
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

    # Divergence vs 20-SMA of ratio (more sensitive threshold)
    ratio_sma20 = ratio.rolling(20).mean()
    divergence = (ratio - ratio_sma20) / ratio_sma20.replace(0, np.nan)
    last5_div = divergence.tail(5).tolist() if len(divergence) >= 5 else []
    last5_dates = [d.strftime("%d-%b") for d in divergence.tail(5).index] if len(divergence) >= 5 else []
    last5_z = zscore.tail(5).tolist() if len(zscore) >= 5 else []

    # Historical extreme Z values (for analytics)
    z_valid = zscore.dropna()
    z_hist_max = float(z_valid.max()) if len(z_valid) else np.nan
    z_hist_min = float(z_valid.min()) if len(z_valid) else np.nan
    z_hist_abs_mean = float(z_valid.abs().mean()) if len(z_valid) else np.nan
    # Best historical entry Z (90th percentile of |Z| when |Z| > 1.5)
    extreme_z = z_valid[z_valid.abs() >= 1.5]
    best_entry_z = float(extreme_z.abs().quantile(0.75)) if len(extreme_z) > 5 else 2.0

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
        "z_hist_max": z_hist_max,
        "z_hist_min": z_hist_min,
        "z_hist_abs_mean": z_hist_abs_mean,
        "best_entry_z": best_entry_z,
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
                f"rolling mean. Historically, such stretches reverse toward the mean. "
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
    if not metrics:
        return {}
    p1, p2 = metrics["price1"], metrics["price2"]
    v1 = p1 * lot1
    v2 = p2 * lot2
    best = None
    for n1 in range(1, 8):
        for n2 in range(1, 8):
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
            pnl_a = sh1 * (p1 - p1_new_a)
            pnl_b = sh2 * (p2_new_b - p2)
        elif direction == "BUY":
            pnl_a = sh1 * (p1_new_a - p1)
            pnl_b = sh2 * (p2 - p2_new_b)
        else:
            pnl_a = pnl_b = 0

        pnl = (pnl_a + pnl_b) / 2
        move_pct = abs(tgt_r - cur_r) / cur_r * 100
        rows.append({
            "Target": label,
            "Target Ratio": round(tgt_r, 4),
            "Ratio Move %": round(move_pct, 2),
            "Est. P&L (₹)": f"₹{round(pnl):,}",
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


def extra_pair_analytics(metrics: dict, signal: dict, hist_stats: dict) -> str:
    """Richer independent analysis text for the selected pair."""
    if not metrics:
        return ""
    lines = []
    z = metrics["zscore"]
    corr30 = metrics.get("corr_30", np.nan)
    hl = metrics.get("half_life", np.nan)
    best_z = metrics.get("best_entry_z", 2.0)

    lines.append(f"**Current stretch:** |Z| = {abs(z):.2f}. "
                 f"Historical 75th-percentile |Z| for strong entries on this pair ≈ {best_z:.2f}.")
    if abs(z) >= best_z:
        lines.append("→ Current stretch is **at or above** the historical high-conviction zone.")
    elif abs(z) >= 1.5:
        lines.append("→ Stretch is tradeable but not extreme; position size more conservatively.")
    else:
        lines.append("→ Inside normal band; no statistical edge yet.")

    if not np.isnan(corr30):
        if corr30 >= 0.7:
            lines.append(f"**Correlation quality:** Strong ({corr30:.2f}) – pair co-movement is reliable.")
        elif corr30 >= 0.5:
            lines.append(f"**Correlation quality:** Acceptable ({corr30:.2f}).")
        else:
            lines.append(f"**Correlation quality:** Weak ({corr30:.2f}) – mean-reversion less trustworthy.")

    if not np.isnan(hl):
        if hl < 10:
            lines.append(f"**Half-life ≈ {hl:.1f} days** – fast mean-reversion; suitable for short holding periods.")
        elif hl < 25:
            lines.append(f"**Half-life ≈ {hl:.1f} days** – moderate reversion speed.")
        else:
            lines.append(f"**Half-life ≈ {hl:.1f} days** – slow reversion; wider stops / longer patience needed.")

    if hist_stats and hist_stats.get("trades", 0) > 0:
        lines.append(
            f"**Simulated 12-mo track record:** {hist_stats['wins']}/{hist_stats['trades']} wins "
            f"({hist_stats['win_rate']}%), avg ratio P&L {hist_stats['avg_pnl_pct']}%."
        )

    # Ratio position vs 30D range
    r = metrics["ratio"]
    rmin, rmax = metrics["ratio_min_30"], metrics["ratio_max_30"]
    if rmax > rmin:
        pct_in_range = (r - rmin) / (rmax - rmin) * 100
        lines.append(f"**Ratio vs 30-day range:** currently at the {pct_in_range:.0f}th percentile "
                     f"(0 = 30D low, 100 = 30D high).")

    return "\n\n".join(lines)


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

    auto_refresh = st.checkbox("Auto-refresh every 60 s", value=False)
    show_charts = st.checkbox("Show detailed charts", value=True)

    st.markdown("---")
    st.caption("Data: Yahoo Finance (live) + your Excel SIGNAL sheet (repo)")
    st.caption("Independent engine confirms / challenges Excel signals")
    st.caption(f"Clock: {datetime.now().strftime('%H:%M:%S')} IST")

    if st.button("🔄 Refresh Now", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


# ─────────────────────────────────────────────────────────────
# LOAD EXCEL + MERGE
# ─────────────────────────────────────────────────────────────
df_signal_full, excel_map, excel_status = load_excel_full_and_map()
ALL_PAIRS = merge_excel_lots(DEFAULT_PAIRS, excel_map)

# ─────────────────────────────────────────────────────────────
# HEADER + TOP REFRESH
# ─────────────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([5, 1])
with col_h1:
    st.markdown('<p class="main-header">📈 Automated Pair Trading Dashboard</p>', unsafe_allow_html=True)
    st.markdown(
        f'<p class="sub-header">Dual confirmation (Excel + Independent) • '
        f'Updated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} IST</p>',
        unsafe_allow_html=True,
    )
with col_h2:
    st.write("")
    if st.button("🔄 Refresh", key="top_refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.caption(f"📂 {excel_status}")

if not HAS_YF:
    st.error("yfinance is not installed. Run: `pip install yfinance`")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 1. FULL EXCEL SIGNAL SHEET (top section, horizontal scroll)
# ─────────────────────────────────────────────────────────────
st.subheader("📋 Full Excel SIGNAL Sheet (as in workbook)")
st.caption("Horizontal scroll only • Colour-coded Buy / Do Not buy columns preserved where possible")

if not df_signal_full.empty:
    # Style helper – highlight Buy / Do Not buy cells
    def highlight_signal_cells(val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return ""
        s = str(val)
        if "Buy" in s and "Do Not" not in s:
            return "background-color: #1B5E20; color: #FFFFFF; font-weight: 600"
        if "Do Not buy" in s or "Do Not BUY" in s:
            return "background-color: #37474F; color: #B0BEC5"
        if "Do Not SELL" in s or "Do Not BUY" in s:
            return "background-color: #263238; color: #90A4AE"
        return ""

    # Limit height so it doesn't dominate; force horizontal overflow via container
    styled_excel = df_signal_full.style.map(highlight_signal_cells)
    st.dataframe(
        styled_excel,
        use_container_width=True,
        height=320,          # fixed height → vertical space limited
        hide_index=True,
    )
    st.caption("Tip: scroll horizontally to see all columns (Z-score, Ratio, Cash Neutrality, Buy/Sell flag, etc.)")
else:
    st.info("Could not load SIGNAL sheet for full display. Pair-level signals below still work from defaults.")

st.markdown("---")

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
# 3. LIVE SIGNAL BOARD – only ENTER + EXIT/HOLD (no WATCH)
# ─────────────────────────────────────────────────────────────
st.subheader("🎯 Live Signal Board — Enter & Exit / Hold only")

rows = []
for r in results:
    m = r["metrics"]
    s = r["signal"]
    p = r["pair"]
    if not m:
        continue
    # Skip pure WATCH rows
    if "WATCH" in s["action"] and "ENTER" not in s["action"] and "EXIT" not in s["action"]:
        continue

    excel_info = p.get("excel", {})
    excel_act = excel_info.get("excel_action", "—")
    dual = ""
    if "ENTER" in s["action"] and excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
        dual = "✅ DUAL"
    elif "ENTER" in s["action"]:
        dual = "⚡ Indep."
    elif excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
        dual = "📋 Excel"
    else:
        dual = "—"

    rows.append({
        "Pair": p["name"],
        "Sector": p["sector"],
        "Lot1": p["lot1"],
        "Lot2": p["lot2"],
        "P1": round(m["price1"], 2),
        "P2": round(m["price2"], 2),
        "Ratio": round(m["ratio"], 4),
        "Z": round(m["zscore"], 2),
        "Corr30": round(m["corr_30"], 2) if not np.isnan(m["corr_30"]) else "—",
        "Action": s["action"],
        "L1": s["leg1"],
        "L2": s["leg2"],
        "Conf%": s["confidence"],
        "Excel": excel_act,
        "Dual": dual,
        "HL": round(m["half_life"], 1) if not np.isnan(m.get("half_life", np.nan)) else "—",
    })

if not rows:
    st.info("No ENTER or EXIT/HOLD signals right now. All pairs are inside the neutral Z-band (see full Excel sheet above).")
else:
    df_signals = pd.DataFrame(rows)

    def color_row(row):
        styles = [""] * len(row)
        action = str(row.get("Action", ""))
        if "ENTER" in action:
            # lighter green, dark text for readability
            styles = ["background-color: #E8F5E9; color: #1B5E20; font-weight: 600"] * len(row)
        elif "EXIT" in action or "HOLD" in action:
            styles = ["background-color: #ECEFF1; color: #37474F"] * len(row)
        return styles

    styled = (df_signals.style
              .apply(color_row, axis=1)
              .format({
                  "P1": "{:.2f}", "P2": "{:.2f}",
                  "Ratio": "{:.4f}", "Z": "{:.2f}",
              }))

    st.dataframe(
        styled,
        use_container_width=True,
        height=min(380, 40 + 32 * len(rows)),
        hide_index=True,
    )

# Compact metrics row
enter_count = sum(1 for r in rows if "ENTER" in r["Action"])
dual_count = sum(1 for r in rows if "DUAL" in r["Dual"])
exit_count = sum(1 for r in rows if "EXIT" in r["Action"] or "HOLD" in r["Action"])
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Pairs Scanned", len([r for r in results if r["metrics"]]))
col2.metric("🟢 Enter", enter_count)
col3.metric("✅ Dual", dual_count)
col4.metric("⬜ Exit/Hold", exit_count)
col5.metric("Avg |Z|", round(np.mean([abs(r["Z"]) for r in rows]), 2) if rows else 0)

# ─────────────────────────────────────────────────────────────
# DETAILED VIEW
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("🔍 Pair Deep-Dive & Trade Recommendations")

pair_names = [r["pair"]["name"] for r in results if r["metrics"]]
if not pair_names:
    st.warning("No pair data available.")
    st.stop()

selected = st.selectbox("Select pair for detailed analysis", pair_names)

selected_result = next(r for r in results if r["pair"]["name"] == selected)
m = selected_result["metrics"]
s = selected_result["signal"]
pr = selected_result["prices"]
pair_info = selected_result["pair"]
lots_info = selected_result["lots_info"]
hist_stats = selected_result["hist_stats"]
excel_info = pair_info.get("excel", {})

# Compact KPI row
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}",
          delta="Stretched" if abs(m["zscore"]) > z_entry else "Normal")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if not np.isnan(m["corr_60"]) else "—")
k5.metric("Half-Life", f"{m['half_life']:.1f}d" if not np.isnan(m.get("half_life", np.nan)) else "—")

# Signal box (slightly tighter)
if "ENTER" in s["action"]:
    action_color = "#1B5E20"
elif "WATCH" in s["action"]:
    action_color = "#F9A825"
else:
    action_color = "#37474F"

excel_line = ""
if excel_info:
    excel_line = (f'<p style="margin:4px 0 0 0;opacity:0.95;font-size:0.9rem;">'
                  f'<span class="badge-excel">Excel: {excel_info.get("excel_action", "—")}</span> '
                  f'&nbsp; Z(excel)={excel_info.get("excel_z")} &nbsp; '
                  f'Flag: {excel_info.get("excel_buy_flag", "")}</p>')

dual_note = ""
if "ENTER" in s["action"] and excel_info.get("excel_action") in ("BUY SIGNAL", "PAIR ENTRY"):
    dual_note = '<p style="margin:6px 0 0 0;"><span class="badge-confirm">✅ DUAL CONFIRMATION — Excel + Independent agree</span></p>'

st.markdown(
    f"""
    <div class="signal-box" style="background:{action_color};color:white;">
        <h4 style="margin:0;color:white;font-size:1.15rem;">{s['action']} — Conf {s['confidence']}%</h4>
        <p style="margin:6px 0 0 0;font-size:1.0rem;">
            <b>{pair_info['s1']}</b> → <b>{s['leg1']}</b> &nbsp;|&nbsp;
            <b>{pair_info['s2']}</b> → <b>{s['leg2']}</b>
        </p>
        <p style="margin:6px 0 0 0;opacity:0.95;font-size:0.9rem;">{s['reason']}</p>
        {excel_line}
        {dual_note}
    </div>
    """,
    unsafe_allow_html=True,
)

# 6. Independent Deep Rational + extra analytics
st.markdown("#### 🧠 Independent Deep Rational & Analytics")
if s.get("deep_reason"):
    st.info(s["deep_reason"])
extra = extra_pair_analytics(m, s, hist_stats)
if extra:
    st.markdown(extra)

# Ratio reference box
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

# 7. Lot sizes + Cash Neutrality (recommended + 1:1)
st.markdown("#### 📦 Lot Sizes & Cash-Neutrality Recommendation")
c1, c2 = st.columns(2)
with c1:
    st.markdown("**Exchange lot sizes (from Excel)**")
    st.markdown(
        f"| Leg | Ticker | Lot |\n|-----|--------|-----|\n"
        f"| 1 | {pair_info['s1']} | **{pair_info['lot1']}** |\n"
        f"| 2 | {pair_info['s2']} | **{pair_info['lot2']}** |"
    )
    st.caption("Update Excel on month-end lot changes; dashboard re-reads on refresh.")

with c2:
    if lots_info:
        st.markdown("**✅ Recommended (max cash-neutrality)**")
        st.markdown(
            f"**{lots_info['lots1']} × lot1** ({lots_info['shares1']} sh) of {pair_info['s1']}  \n"
            f"**{lots_info['lots2']} × lot2** ({lots_info['shares2']} sh) of {pair_info['s2']}  \n"
            f"Value1 ≈ ₹{lots_info['value1']:,.0f} | Value2 ≈ ₹{lots_info['value2']:,.0f}  \n"
            f"Imbalance: ₹{lots_info['imbalance']:,.0f} ({lots_info['imbalance_pct']}%)"
        )
        st.markdown("**— vs simple 1 : 1 lot —**")
        st.caption(
            f"1×lot1 vs 1×lot2 → imbalance ₹{lots_info['single_imbalance']:,.0f} "
            f"({lots_info['single_imbalance_pct']}%). Multi-lot reduces residual risk."
        )

# 8. Profit potential (₹, recommended lots)
st.markdown("#### 💰 Profit Potential (based on recommended cash-neutral lots)")
pnl_rows = profit_potential(m, s, lots_info)
if pnl_rows:
    st.dataframe(pd.DataFrame(pnl_rows), use_container_width=True, hide_index=True)
    st.caption("Estimates assume the recommended multi-lot position above. Not a guarantee.")
else:
    st.caption("Profit grid available when lot data is present (ENTER signals give directional P&L).")

# Entry / Target / Stop
if pr:
    st.markdown("#### 🎯 Suggested Entry / Target / Stop")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{pair_info['s1']}** (Lot {pair_info['lot1']})")
        st.markdown(
            f"| Entry | Target | Stop | Support | Resist |\n"
            f"|-------|--------|------|---------|--------|\n"
            f"| {pr['entry1']} | {pr['target1']} | {pr['stop1']} | {pr['support1']} | {pr['resist1']} |"
        )
    with c2:
        st.markdown(f"**{pair_info['s2']}** (Lot {pair_info['lot2']})")
        st.markdown(
            f"| Entry | Target | Stop | Support | Resist |\n"
            f"|-------|--------|------|---------|--------|\n"
            f"| {pr['entry2']} | {pr['target2']} | {pr['stop2']} | {pr['support2']} | {pr['resist2']} |"
        )
    st.info(
        "Targets assume ratio mean-reversion. Stops sit beyond recent ratio extremes. "
        "Educational only — not financial advice."
    )

# 9. Last 5 trading days – fixed dots + divergence values
st.markdown("#### 📅 Last 5 Trading Days — Divergence Trend")
if m.get("last5_div") and m.get("last5_dates"):
    # Lower threshold so small moves still colour
    dots_html = []
    vals_html = []
    for d, div in zip(m["last5_dates"], m["last5_div"]):
        if div is None or (isinstance(div, float) and np.isnan(div)):
            cls, label = "dot-gray", "●"
            val_str = "—"
        elif div > 0.002:
            cls, label = "dot-green", "●"
            val_str = f"+{div:.3f}"
        elif div < -0.002:
            cls, label = "dot-red", "●"
            val_str = f"{div:.3f}"
        else:
            cls, label = "dot-gray", "●"
            val_str = f"{div:.3f}"
        dots_html.append(f'<span class="{cls}" title="{d}: {div}">{label}</span> <small>{d}</small>')
        vals_html.append(f'<span class="div-val">{val_str}</span>')

    st.markdown(" &nbsp;&nbsp; ".join(dots_html), unsafe_allow_html=True)
    st.markdown(" &nbsp;&nbsp; ".join(vals_html), unsafe_allow_html=True)

    divs = [x for x in m["last5_div"] if x is not None and not (isinstance(x, float) and np.isnan(x))]
    if len(divs) >= 2:
        if divs[-1] > divs[0] + 0.003:
            st.caption("Divergence **rising** (ratio moving away from 20-SMA) → stretch building.")
        elif divs[-1] < divs[0] - 0.003:
            st.caption("Divergence **falling** (ratio moving back toward 20-SMA) → mean-reversion in progress.")
        else:
            st.caption("Divergence roughly **stable** over the last 5 sessions.")
    st.caption("Green = ratio above 20-SMA · Red = below · Gray = near zero · Values = (Ratio − SMA20) / SMA20")
else:
    st.caption("Not enough history for 5-day divergence dots.")

# 10 + 11. Charts – mobile-friendly, legends, descriptions, Z analytics
if show_charts and HAS_PLOTLY and not selected_result["hist"].empty:
    hist = selected_result["hist"]
    ratio_s = m["ratio_series"]
    z_s = m["zscore_series"]

    fig = make_subplots(
        rows=3, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        subplot_titles=(
            f"Normalized Prices — {pair_info['s1']} vs {pair_info['s2']}",
            "Ratio (Price1 / Price2) + Reference Levels",
            "Z-Score of Ratio",
        ),
        row_heights=[0.33, 0.34, 0.33],
    )

    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100
    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=pair_info["s1"],
                             line=dict(color="#00E5FF", width=1.5)), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=pair_info["s2"],
                             line=dict(color="#FF5252", width=1.5)), row=1, col=1)

    # Ratio
    fig.add_trace(go.Scatter(x=ratio_s.index, y=ratio_s, name="Ratio",
                             line=dict(color="#69F0AE", width=1.8)), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00",
                  annotation_text="60D Avg", annotation_position="bottom right", row=2, col=1)
    fig.add_hline(y=m["ratio_max_30"], line_dash="dot", line_color="#FF8A80",
                  annotation_text="30D Max", annotation_position="top right", row=2, col=1)
    fig.add_hline(y=m["ratio_min_30"], line_dash="dot", line_color="#80D8FF",
                  annotation_text="30D Min", annotation_position="bottom left", row=2, col=1)
    fig.add_hline(y=m["ratio_max_60"], line_dash="dot", line_color="#FF1744",
                  annotation_text="60D Max", annotation_position="top left", row=2, col=1)
    fig.add_hline(y=m["ratio_min_60"], line_dash="dot", line_color="#00B0FF",
                  annotation_text="60D Min", annotation_position="bottom right", row=2, col=1)

    # Z-score
    fig.add_trace(go.Scatter(x=z_s.index, y=z_s, name="Z-Score",
                             line=dict(color="#E040FB", width=1.8)), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252",
                  annotation_text=f"+{z_entry}", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676",
                  annotation_text=f"-{z_entry}", row=3, col=1)
    fig.add_hline(y=0, line_color="#78909C", row=3, col=1)
    # Historical extremes
    if not np.isnan(m.get("z_hist_max", np.nan)):
        fig.add_hline(y=m["z_hist_max"], line_dash="dot", line_color="#FF1744",
                      annotation_text="Hist Max Z", annotation_position="top right", row=3, col=1)
    if not np.isnan(m.get("z_hist_min", np.nan)):
        fig.add_hline(y=m["z_hist_min"], line_dash="dot", line_color="#00B0FF",
                      annotation_text="Hist Min Z", annotation_position="bottom right", row=3, col=1)

    fig.update_layout(
        height=820,
        template="plotly_dark",
        paper_bgcolor="#0E1117",
        plot_bgcolor="#1A1F2E",
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.12,          # bottom of the whole figure
            xanchor="center",
            x=0.5,
            font=dict(color="#E0E0E0", size=11),
        ),
        font=dict(color="#E0E0E0", size=11),
        margin=dict(l=40, r=40, t=50, b=80),
        # Mobile-friendly: disable drag-to-zoom that disfigures on scroll
        dragmode=False,
    )
    # Disable scroll-zoom and double-click zoom for touch devices
    config = {
        "scrollZoom": False,
        "displayModeBar": True,
        "modeBarButtonsToRemove": ["zoom2d", "pan2d", "select2d", "lasso2d", "autoScale2d"],
        "responsive": True,
    }
    st.plotly_chart(fig, use_container_width=True, config=config)

    # Colour legend for ratio lines
    st.markdown(
        """
        <p style="font-size:0.75rem;color:#90A4AE;margin-top:-8px;">
        <b>Ratio chart lines:</b>
        <span style="color:#FFEA00;">■</span> Yellow dashed = 60D average &nbsp;|&nbsp;
        <span style="color:#FF8A80;">■</span> Light red dotted = 30D max &nbsp;|&nbsp;
        <span style="color:#80D8FF;">■</span> Light blue dotted = 30D min &nbsp;|&nbsp;
        <span style="color:#FF1744;">■</span> Bright red dotted = 60D max &nbsp;|&nbsp;
        <span style="color:#00B0FF;">■</span> Bright blue dotted = 60D min
        </p>
        """,
        unsafe_allow_html=True,
    )

    # Z-score interpretation + historical best Z analytics
    st.markdown("##### 📖 Z-Score Trend Interpretation")
    st.markdown(interpret_z_trend(m.get("last5_z", [])))

    st.markdown("##### 📊 Z-Score Historical Analytics (this pair)")
    zmax = m.get("z_hist_max", np.nan)
    zmin = m.get("z_hist_min", np.nan)
    zabs = m.get("z_hist_abs_mean", np.nan)
    bestz = m.get("best_entry_z", np.nan)
    st.markdown(
        f"| Metric | Value |\n|--------|-------|\n"
        f"| Current Z | **{m['zscore']:.2f}** |\n"
        f"| Historical Max Z (1y) | {zmax:.2f} |\n"
        f"| Historical Min Z (1y) | {zmin:.2f} |\n"
        f"| Mean \\|Z\\| | {zabs:.2f} |\n"
        f"| Suggested high-conviction entry \\|Z\\| (75th pct of extremes) | **{bestz:.2f}** |"
    )
    st.caption(
        "Entering only when |Z| is near or above the historical high-conviction level "
        "has historically improved win rate on mean-reversion pairs. "
        "Combine with correlation ≥ 0.5 and dual confirmation when possible."
    )

# ─────────────────────────────────────────────────────────────
# HOW EXCEL IS USED
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📂 How Your Excel Sheet Is Used")
st.markdown("""
This dashboard is designed for **dual-confirmation** trading:

1. **Lot sizes** – Read from column A of the SIGNAL sheet on every refresh.  
   When NSE revises lots (usually month-end), update the Excel, commit/push to GitHub, and the dashboard picks them up.
2. **Excel signals** – Columns C/D (SELL/BUY per leg) and the Buy / Do Not buy flag are parsed.  
   If Excel says “Buy” / “PAIR ENTRY”, it is shown side-by-side with the independent engine.
3. **Excel Z / Ratio / Correlation** – Displayed for reference so you can compare the sheet’s numbers with live Yahoo Finance calculations.
4. **Independent engine** – Recomputes ratio, Z-score, correlation, half-life, support/resistance and mean-reversion signals from scratch using 1-year daily data **live from Yahoo Finance**.
5. **Dual confirmation badge** – Appears only when **both** Excel and the independent engine flash an entry.

> **Live-data confirmation:** Prices, charts, Z-scores and all independent analytics are fetched fresh from Yahoo Finance on every page load / refresh.  
> The Excel file supplies only the static SIGNAL snapshot (lots + your sheet’s own flags).  
> Once the Excel is committed in the GitHub repo that Streamlit Cloud (or your host) pulls from, there is **no dependence on a desktop path**.
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
# 13. HISTORICAL PAIR PERFORMANCE (more colourful)
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
    df_hist = pd.DataFrame(hist_rows)

    def color_winrate(val):
        try:
            v = float(val)
            if v >= 60:
                return "background-color: #1B5E20; color: #FFFFFF; font-weight: 700"
            if v >= 50:
                return "background-color: #33691E; color: #FFFFFF"
            if v >= 40:
                return "background-color: #F9A825; color: #000000"
            return "background-color: #B71C1C; color: #FFFFFF"
        except Exception:
            return ""

    def color_pnl(val):
        try:
            v = float(val)
            if v > 0:
                return "background-color: #E8F5E9; color: #1B5E20; font-weight: 600"
            return "background-color: #FFEBEE; color: #B71C1C; font-weight: 600"
        except Exception:
            return ""

    styled_hist = (df_hist.style
                   .map(color_winrate, subset=["Win Rate %"])
                   .map(color_pnl, subset=["Avg P&L % (ratio)"]))
    st.dataframe(styled_hist, use_container_width=True, hide_index=True)

    sel_hs = selected_result.get("hist_stats") or {}
    if sel_hs.get("details"):
        st.markdown(f"**Recent simulated trades for {selected}**")
        st.dataframe(pd.DataFrame(sel_hs["details"]), use_container_width=True, hide_index=True)
else:
    st.write("Not enough history to compute win-rate statistics.")

st.markdown("""
**Methodology (Independent of original Excel formulas):**
1. Fetch 1-year daily adjusted closes via Yahoo Finance (live).
2. Ratio = Stock1 / Stock2.
3. Z-score = (Ratio − RollingMean) / RollingStd (default 60-day).
4. Enter when |Z| > threshold **and** |correlation| ≥ 0.5.
5. Direction: positive Z → sell expensive leg / buy cheap leg.
6. Support / Resistance from 60-day high/low of each leg and of the ratio.
7. Half-life via AR(1) / Ornstein-Uhlenbeck regression.
8. Cash-neutrality: search integer lot multiples that minimise |Value1 − Value2|.
9. Historical stats: count mean-reversion opportunities and win rate over available history.
10. Excel SIGNAL sheet (from GitHub repo) supplies lot sizes + its own Buy/Sell flags for dual confirmation.

> ⚠️ Educational & research use only. Past mean-reversion does not guarantee future results. Always manage risk and position size carefully.
""")

# ─────────────────────────────────────────────────────────────
# AUTO REFRESH
# ─────────────────────────────────────────────────────────────
if auto_refresh:
    time.sleep(60)
    st.rerun()
