"""
Pair Trading Dashboard - Updated Streamlit App
================================================

Install:
    pip install streamlit pandas numpy yfinance plotly openpyxl scipy

Run:
    streamlit run pair_trade_dashboard_updated.py

Important Excel behavior:
    - A local workbook can be read only when Streamlit runs on the same computer.
    - Set DEFAULT_EXCEL_PATH or choose the workbook in the sidebar.
    - Formula values must be recalculated and saved by Excel before this app
      can read the updated values.
    - The app does not write to the workbook.

Expected workbook:
    - File name can be "Pair trade1.xlsx" or another .xlsx/.xlsm file.
    - Signal sheet name defaults to "SIGNAL".
    - Pair-definition columns are detected by common aliases such as:
      Pair, Stock1/Ticker1, Stock2/Ticker2, Lot1, Lot2, Sector.
"""

from pathlib import Path
import html
import math
import os
import time
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

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


# ---------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------

DEFAULT_EXCEL_PATH = r""  # Example: r"C:\Users\YourName\Documents\Pair trade1.xlsx"
DEFAULT_SIGNAL_SHEET = "SIGNAL"

st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_PAIRS = [
    {"name": "ICICIBANK-HDFCBANK", "s1": "ICICIBANK.NS", "s2": "HDFCBANK.NS", "lot1": 700, "lot2": 550, "sector": "Banking"},
    {"name": "AXISBANK-ICICIBANK", "s1": "AXISBANK.NS", "s2": "ICICIBANK.NS", "lot1": 625, "lot2": 700, "sector": "Banking"},
    {"name": "CIPLA-SUNPHARMA", "s1": "CIPLA.NS", "s2": "SUNPHARMA.NS", "lot1": 650, "lot2": 350, "sector": "Pharma"},
    {"name": "BPCL-HINDPETRO", "s1": "BPCL.NS", "s2": "HINDPETRO.NS", "lot1": 1975, "lot2": 2025, "sector": "Oil & Gas"},
    {"name": "ACC-AMBUJACEM", "s1": "ACC.NS", "s2": "AMBUJACEM.NS", "lot1": 300, "lot2": 900, "sector": "Cement"},
    {"name": "TCS-HCLTECH", "s1": "TCS.NS", "s2": "HCLTECH.NS", "lot1": 175, "lot2": 350, "sector": "IT"},
    {"name": "TATASTEEL-JSWSTEEL", "s1": "TATASTEEL.NS", "s2": "JSWSTEEL.NS", "lot1": 5500, "lot2": 675, "sector": "Metals"},
    {"name": "NESTLE-HINDUNILVR", "s1": "NESTLEIND.NS", "s2": "HINDUNILVR.NS", "lot1": 200, "lot2": 300, "sector": "FMCG"},
    {"name": "NIFTY-BANKNIFTY", "s1": "^NSEI", "s2": "^NSEBANK", "lot1": 25, "lot2": 15, "sector": "Index"},
    {"name": "BRITANNIA-HINDUNILVR", "s1": "BRITANNIA.NS", "s2": "HINDUNILVR.NS", "lot1": 200, "lot2": 300, "sector": "FMCG"},
    {"name": "BAJFINANCE-BAJAJFINSV", "s1": "BAJFINANCE.NS", "s2": "BAJAJFINSV.NS", "lot1": 100, "lot2": 500, "sector": "Finance"},
    {"name": "HDFCBANK-BANKNIFTY", "s1": "HDFCBANK.NS", "s2": "^NSEBANK", "lot1": 550, "lot2": 15, "sector": "Banking"},
    {"name": "RECLTD-PFC", "s1": "RECLTD.NS", "s2": "PFC.NS", "lot1": 2000, "lot2": 1300, "sector": "Finance"},
    {"name": "IGL-GAIL", "s1": "IGL.NS", "s2": "GAIL.NS", "lot1": 1375, "lot2": 4575, "sector": "Oil & Gas"},
    {"name": "ALKEM-AUROPHARMA", "s1": "ALKEM.NS", "s2": "AUROPHARMA.NS", "lot1": 100, "lot2": 550, "sector": "Pharma"},
    {"name": "BERGEPAINT-ASIANPAINT", "s1": "BERGEPAINT.NS", "s2": "ASIANPAINT.NS", "lot1": 1320, "lot2": 200, "sector": "Paints"},
    {"name": "BANKBARODA-CANBK", "s1": "BANKBARODA.NS", "s2": "CANBK.NS", "lot1": 2925, "lot2": 6750, "sector": "Banking"},
    {"name": "COALINDIA-GAIL", "s1": "COALINDIA.NS", "s2": "GAIL.NS", "lot1": 2925, "lot2": 6750, "sector": "Energy"},
]


# ---------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------

st.markdown(
    """
<style>
.main-header {
    font-size: 2rem;
    font-weight: 750;
    color: #1565C0;
    margin: 0 0 .1rem 0;
}
.sub-header {
    font-size: .84rem;
    color: #607D8B;
    margin-bottom: .7rem;
}
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #F8FBFF, #EEF5FC);
    border: 1px solid #D6E4F0;
    border-radius: 9px;
    padding: 5px 9px;
    min-height: 68px;
}
div[data-testid="stMetricLabel"] {
    font-size: .72rem;
}
div[data-testid="stMetricValue"] {
    font-size: 1.12rem;
}
div[data-testid="stDataFrame"] {
    font-size: .76rem;
}
.compact-box {
    padding: 10px 12px;
    border: 1px solid #DCE6EF;
    border-radius: 9px;
    background: #FAFCFE;
}
.signal-scroll {
    width: 100%;
    overflow-x: auto;
    overflow-y: visible;
    border: 1px solid #CFD8DC;
    border-radius: 8px;
    background: white;
    -webkit-overflow-scrolling: touch;
}
.signal-scroll table {
    border-collapse: collapse;
    min-width: max-content;
    width: 100%;
    font-size: .72rem;
}
.signal-scroll th {
    position: sticky;
    top: 0;
    z-index: 2;
    padding: 7px 9px;
    white-space: nowrap;
    background: #1565C0;
    color: white;
    border: 1px solid #BBDEFB;
}
.signal-scroll td {
    padding: 6px 9px;
    white-space: nowrap;
    border: 1px solid #E5E9EC;
    color: #17202A;
}
.legend-note {
    font-size: .69rem;
    color: #546E7A;
    line-height: 1.45;
    padding: 2px 5px 8px 5px;
}
.day-grid {
    display: grid;
    grid-template-columns: repeat(5, minmax(64px, 1fr));
    gap: 7px;
    margin: 5px 0 12px 0;
}
.day-cell {
    text-align: center;
    background: #F7F9FA;
    border: 1px solid #DEE5E8;
    border-radius: 8px;
    padding: 7px 3px;
    font-size: .70rem;
}
.dot-up, .dot-down, .dot-flat {
    height: 13px;
    width: 13px;
    border-radius: 50%;
    display: inline-block;
    margin: 4px;
}
.dot-up { background: #43A047; }
.dot-down { background: #E53935; }
.dot-flat { background: #90A4AE; }

@media (max-width: 768px) {
    .main-header { font-size: 1.45rem; }
    .sub-header { font-size: .72rem; }
    div[data-testid="stMetric"] {
        min-height: 58px;
        padding: 3px 7px;
    }
    div[data-testid="stMetricValue"] { font-size: .94rem; }
    div[data-testid="stDataFrame"] { font-size: .66rem; }

    button[data-testid="stSidebarCollapseButton"],
    button[data-testid="stSidebarCollapseButton"] *,
    button[kind="header"],
    button[kind="header"] * {
        color: #102A43 !important;
        fill: #102A43 !important;
        opacity: 1 !important;
    }
    section[data-testid="stSidebar"] button,
    section[data-testid="stSidebar"] button * {
        color: #102A43 !important;
        fill: #102A43 !important;
    }
    .day-grid {
        grid-template-columns: repeat(5, minmax(52px, 1fr));
        gap: 3px;
    }
    .day-cell {
        font-size: .61rem;
        padding: 5px 1px;
    }
}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------
# GENERAL HELPERS
# ---------------------------------------------------------------------

def clean_label(value) -> str:
    return " ".join(str(value).strip().lower().replace("_", " ").replace("-", " ").split())


def safe_float(value, default=np.nan):
    try:
        result = float(value)
        return result if np.isfinite(result) else default
    except (TypeError, ValueError):
        return default


def add_nse_suffix(ticker: str) -> str:
    ticker = str(ticker).strip().upper()
    if not ticker:
        return ticker
    if ticker.startswith("^") or ticker.endswith((".NS", ".BO")):
        return ticker
    return f"{ticker}.NS"


def inr(value) -> str:
    if value is None or not np.isfinite(value):
        return "₹—"
    return f"₹{value:,.2f}"


def nearest_multiple(value: float, lot: int) -> int:
    lot = max(1, int(lot or 1))
    return max(lot, int(round(value / lot)) * lot)


def format_excel_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, float):
        if value.is_integer():
            return f"{value:,.0f}"
        return f"{value:,.4f}".rstrip("0").rstrip(".")
    return str(value)


# ---------------------------------------------------------------------
# EXCEL INPUT
# ---------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def read_excel_signal(path: str, sheet_name: str, modified_ns: int):
    """Read the saved values of the requested workbook sheet."""
    del modified_ns
    return pd.read_excel(
        path,
        sheet_name=sheet_name,
        engine="openpyxl",
        dtype=object,
    )


def locate_default_workbook():
    candidates = [
        DEFAULT_EXCEL_PATH,
        "Pair trade1.xlsx",
        "Pair Trade1.xlsx",
        "pair trade1.xlsx",
        "Pair trade1.xlsm",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).expanduser().is_file():
            return str(Path(candidate).expanduser().resolve())
    return ""


def find_column(df: pd.DataFrame, aliases):
    normalized = {clean_label(col): col for col in df.columns}
    for alias in aliases:
        if clean_label(alias) in normalized:
            return normalized[clean_label(alias)]
    return None


def pairs_from_signal_sheet(signal_df: pd.DataFrame):
    """
    Build pair definitions from the SIGNAL sheet when recognizable columns exist.
    Otherwise the built-in pair list remains available.
    """
    if signal_df is None or signal_df.empty:
        return []

    pair_col = find_column(signal_df, ["pair", "pair name", "pairname", "symbol pair"])
    s1_col = find_column(signal_df, ["s1", "stock1", "stock 1", "ticker1", "ticker 1", "symbol1", "symbol 1", "leg1 symbol"])
    s2_col = find_column(signal_df, ["s2", "stock2", "stock 2", "ticker2", "ticker 2", "symbol2", "symbol 2", "leg2 symbol"])
    lot1_col = find_column(signal_df, ["lot1", "lot 1", "stock1 lot", "leg1 lot", "qty1"])
    lot2_col = find_column(signal_df, ["lot2", "lot 2", "stock2 lot", "leg2 lot", "qty2"])
    sector_col = find_column(signal_df, ["sector", "industry"])

    if not s1_col or not s2_col:
        return []

    pairs = []
    for _, row in signal_df.iterrows():
        raw_s1 = row.get(s1_col)
        raw_s2 = row.get(s2_col)
        if pd.isna(raw_s1) or pd.isna(raw_s2):
            continue

        s1 = add_nse_suffix(raw_s1)
        s2 = add_nse_suffix(raw_s2)
        if not s1 or not s2:
            continue

        default_name = f"{s1.replace('.NS', '')}-{s2.replace('.NS', '')}"
        pair_name = (
            str(row.get(pair_col)).strip()
            if pair_col and pd.notna(row.get(pair_col))
            else default_name
        )

        lot1 = safe_float(row.get(lot1_col), 1) if lot1_col else 1
        lot2 = safe_float(row.get(lot2_col), 1) if lot2_col else 1

        pairs.append(
            {
                "name": pair_name,
                "s1": s1,
                "s2": s2,
                "lot1": max(1, int(lot1)),
                "lot2": max(1, int(lot2)),
                "sector": (
                    str(row.get(sector_col)).strip()
                    if sector_col and pd.notna(row.get(sector_col))
                    else "Other"
                ),
            }
        )

    unique = {}
    for pair in pairs:
        unique[pair["name"]] = pair
    return list(unique.values())


def excel_cell_style(value):
    text = str(value).strip().upper()
    if "ENTER" in text or text in {"BUY", "LONG"}:
        return "background:#C8E6C9;color:#12351A;font-weight:700;"
    if "EXIT" in text or "HOLD" in text or "CLOSE" in text:
        return "background:#ECEFF1;color:#263238;font-weight:700;"
    if "SELL" in text or "SHORT" in text:
        return "background:#FFCDD2;color:#5D1117;font-weight:700;"
    if "WATCH" in text:
        return "background:#FFF3CD;color:#5D4A00;font-weight:700;"
    return ""


def dataframe_as_scroll_table(df: pd.DataFrame):
    """Render all Excel columns with horizontal scrolling only."""
    if df is None or df.empty:
        st.info("No rows were found in the selected Excel sheet.")
        return

    display = df.copy()
    display.columns = [str(c) for c in display.columns]

    html_rows = ["<div class='signal-scroll'><table><thead><tr>"]
    for col in display.columns:
        html_rows.append(f"<th>{html.escape(str(col))}</th>")
    html_rows.append("</tr></thead><tbody>")

    for _, row in display.iterrows():
        row_text = " ".join(str(v).upper() for v in row.values if pd.notna(v))
        if "ENTER" in row_text:
            row_style = "background:#E8F5E9;"
        elif "EXIT" in row_text or "HOLD" in row_text or "CLOSE" in row_text:
            row_style = "background:#F2F4F5;"
        elif "WATCH" in row_text:
            row_style = "background:#FFFBEA;"
        else:
            row_style = "background:#FFFFFF;"

        html_rows.append(f"<tr style='{row_style}'>")
        for value in row.values:
            formatted = format_excel_value(value)
            style = excel_cell_style(formatted)
            html_rows.append(
                f"<td style='{style}'>{html.escape(formatted)}</td>"
            )
        html_rows.append("</tr>")

    html_rows.append("</tbody></table></div>")
    st.markdown("".join(html_rows), unsafe_allow_html=True)


# ---------------------------------------------------------------------
# MARKET DATA AND PAIR METRICS
# ---------------------------------------------------------------------

@st.cache_data(ttl=55, show_spinner=False)
def fetch_pair_history(ticker1: str, ticker2: str, period: str = "1y"):
    if not HAS_YF:
        return pd.DataFrame()

    try:
        raw = yf.download(
            [ticker1, ticker2],
            period=period,
            auto_adjust=True,
            progress=False,
            threads=True,
            group_by="column",
        )

        if raw.empty:
            return pd.DataFrame()

        if isinstance(raw.columns, pd.MultiIndex):
            close = raw["Close"]
            if ticker1 not in close.columns or ticker2 not in close.columns:
                return pd.DataFrame()
            df = pd.DataFrame(
                {
                    "price1": close[ticker1],
                    "price2": close[ticker2],
                }
            )
        else:
            return pd.DataFrame()

        df = df.dropna()
        df.index = pd.to_datetime(df.index).tz_localize(None)
        return df

    except Exception:
        return pd.DataFrame()


def rolling_beta_hedge_ratio(df: pd.DataFrame, length=60):
    recent = df.tail(min(length, len(df))).dropna()
    if len(recent) < 20:
        return np.nan

    x = np.log(recent["price2"].clip(lower=1e-9))
    y = np.log(recent["price1"].clip(lower=1e-9))
    variance = x.var()
    if not np.isfinite(variance) or variance == 0:
        return np.nan
    return float(x.cov(y) / variance)


def count_zero_crossings(series: pd.Series):
    clean = series.dropna()
    if len(clean) < 2:
        return 0
    signs = np.sign(clean)
    return int((signs.diff().abs() == 2).sum())


def compute_pair_metrics(df: pd.DataFrame, lookback: int = 60):
    if df.empty or len(df) < 30:
        return {}

    ratio = df["price1"] / df["price2"]
    window = min(lookback, max(20, len(ratio) - 1))
    rolling_mean = ratio.rolling(window).mean()
    rolling_std = ratio.rolling(window).std(ddof=1).replace(0, np.nan)
    zscore = (ratio - rolling_mean) / rolling_std

    returns = df.pct_change().dropna()

    def corr_for(days):
        sample = returns.tail(days)
        return (
            sample["price1"].corr(sample["price2"])
            if len(sample) >= min(20, days)
            else np.nan
        )

    corr_30 = corr_for(30)
    corr_60 = corr_for(60)
    corr_90 = corr_for(90)

    current_ratio = float(ratio.iloc[-1])
    current_z = safe_float(zscore.iloc[-1], 0.0)
    p1 = float(df["price1"].iloc[-1])
    p2 = float(df["price2"].iloc[-1])

    recent60 = df.tail(60)
    ratio60 = ratio.tail(60)
    z_valid = zscore.dropna()

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

    spread_change_5d = (
        float(ratio.iloc[-1] / ratio.iloc[-6] - 1)
        if len(ratio) >= 6
        else np.nan
    )
    ratio_volatility = float(ratio.pct_change().tail(60).std() * np.sqrt(252))
    beta = rolling_beta_hedge_ratio(df, 60)

    historical_abs_z = z_valid.abs()
    current_percentile = (
        float((historical_abs_z <= abs(current_z)).mean() * 100)
        if not historical_abs_z.empty
        else np.nan
    )

    max_z = float(z_valid.max()) if not z_valid.empty else np.nan
    min_z = float(z_valid.min()) if not z_valid.empty else np.nan
    avg_z = float(z_valid.mean()) if not z_valid.empty else np.nan

    recent_z = z_valid.tail(252)
    extreme_count = int((recent_z.abs() >= 1.5).sum())
    zero_crossings = count_zero_crossings(recent_z)

    return {
        "ratio": current_ratio,
        "zscore": current_z,
        "corr_30": corr_30,
        "corr_60": corr_60,
        "corr_90": corr_90,
        "price1": p1,
        "price2": p2,
        "support1": float(recent60["price1"].min()),
        "resist1": float(recent60["price1"].max()),
        "support2": float(recent60["price2"].min()),
        "resist2": float(recent60["price2"].max()),
        "ratio_support": float(ratio60.min()),
        "ratio_resist": float(ratio60.max()),
        "ratio_mean": float(ratio60.mean()),
        "ratio_q25": float(ratio60.quantile(.25)),
        "ratio_q75": float(ratio60.quantile(.75)),
        "ratio_volatility": ratio_volatility,
        "spread_change_5d": spread_change_5d,
        "half_life": half_life,
        "hedge_beta": beta,
        "z_percentile": current_percentile,
        "max_z": max_z,
        "min_z": min_z,
        "avg_z": avg_z,
        "extreme_count": extreme_count,
        "zero_crossings": zero_crossings,
        "ratio_series": ratio,
        "ratio_mean_series": rolling_mean,
        "zscore_series": zscore,
        "df": df,
    }


def generate_signal(metrics: dict, z_entry=1.5, z_exit=.5):
    if not metrics:
        return {
            "action": "NO DATA",
            "leg1": "—",
            "leg2": "—",
            "confidence": 0,
            "reason": "Insufficient market history.",
        }

    z = metrics["zscore"]
    correlations = [
        metrics.get("corr_30"),
        metrics.get("corr_60"),
        metrics.get("corr_90"),
    ]
    corr = next((c for c in correlations if np.isfinite(c)), 0.0)
    corr_ok = abs(corr) >= .5

    if abs(z) >= z_entry and corr_ok:
        if z > 0:
            leg1, leg2 = "SELL", "BUY"
            reason = (
                f"Ratio is {z:.2f} standard deviations above its rolling mean. "
                "The setup expects the ratio to decline."
            )
        else:
            leg1, leg2 = "BUY", "SELL"
            reason = (
                f"Ratio is {z:.2f} standard deviations below its rolling mean. "
                "The setup expects the ratio to rise."
            )
        action = "ENTER PAIR"
        confidence = min(
            95,
            int(38 + abs(z) * 15 + abs(corr) * 20),
        )
    elif abs(z) <= z_exit:
        action = "EXIT / HOLD"
        leg1, leg2 = "CLOSE", "CLOSE"
        confidence = 30
        reason = (
            f"Z-score {z:.2f} is close to the rolling mean. "
            "This is normally an exit or no-entry zone."
        )
    else:
        action = "WATCH"
        leg1, leg2 = "—", "—"
        confidence = 20
        reason = f"Z-score {z:.2f} remains between entry and exit thresholds."

    return {
        "action": action,
        "leg1": leg1,
        "leg2": leg2,
        "confidence": confidence,
        "reason": reason,
        "z": z,
        "corr": corr,
    }


def recommend_prices(metrics: dict, signal: dict):
    if not metrics or signal["action"] != "ENTER PAIR":
        return {}

    z = metrics["zscore"]
    p1 = metrics["price1"]
    p2 = metrics["price2"]
    ratio = metrics["ratio"]
    target_ratio = metrics["ratio_mean"]

    if z > 0:
        stop_ratio = metrics["ratio_resist"] * 1.02
    else:
        stop_ratio = metrics["ratio_support"] * .98

    if signal["leg1"] == "SELL":
        target1 = p1 * target_ratio / ratio
        stop1 = p1 * stop_ratio / ratio
        target2 = p2 * ratio / target_ratio
        stop2 = p2 * ratio / stop_ratio
    else:
        target1 = p1 * target_ratio / ratio
        stop1 = p1 * stop_ratio / ratio
        target2 = p2 * ratio / target_ratio
        stop2 = p2 * ratio / stop_ratio

    return {
        "entry1": round(p1, 2),
        "target1": round(target1, 2),
        "stop1": round(stop1, 2),
        "entry2": round(p2, 2),
        "target2": round(target2, 2),
        "stop2": round(stop2, 2),
        "support1": round(metrics["support1"], 2),
        "resist1": round(metrics["resist1"], 2),
        "support2": round(metrics["support2"], 2),
        "resist2": round(metrics["resist2"], 2),
    }


# ---------------------------------------------------------------------
# CASH-NEUTRAL SIZING AND PROFIT ESTIMATE
# ---------------------------------------------------------------------

def optimize_cash_neutral_lots(pair, metrics, max_contracts=40):
    """
    Search whole-lot combinations that minimize cash-value mismatch.

    This is cash-neutral sizing, not beta-neutral sizing. If leg directions are
    opposite, the long and short market values are matched as closely as whole
    lots allow.
    """
    p1 = metrics["price1"]
    p2 = metrics["price2"]
    lot1 = max(1, int(pair.get("lot1", 1)))
    lot2 = max(1, int(pair.get("lot2", 1)))

    unit_value1 = p1 * lot1
    unit_value2 = p2 * lot2

    best = None
    for contracts1 in range(1, max_contracts + 1):
        ideal2 = unit_value1 * contracts1 / unit_value2
        candidates = {
            max(1, int(math.floor(ideal2))),
            max(1, int(round(ideal2))),
            max(1, int(math.ceil(ideal2))),
        }

        for contracts2 in candidates:
            if contracts2 > max_contracts:
                continue

            qty1 = contracts1 * lot1
            qty2 = contracts2 * lot2
            value1 = qty1 * p1
            value2 = qty2 * p2
            difference = value1 - value2
            mismatch_pct = abs(difference) / max(value1, value2) * 100

            size_penalty = (contracts1 + contracts2) * 0.0001
            score = mismatch_pct + size_penalty

            candidate = {
                "contracts1": contracts1,
                "contracts2": contracts2,
                "qty1": qty1,
                "qty2": qty2,
                "value1": value1,
                "value2": value2,
                "difference": difference,
                "mismatch_pct": mismatch_pct,
                "score": score,
            }

            if best is None or candidate["score"] < best["score"]:
                best = candidate

    one_to_one_value1 = lot1 * p1
    one_to_one_value2 = lot2 * p2
    one_to_one_difference = one_to_one_value1 - one_to_one_value2

    if best["difference"] > 0:
        additional_leg = 2
        additional_shares = max(0, int(round(best["difference"] / p2)))
    elif best["difference"] < 0:
        additional_leg = 1
        additional_shares = max(0, int(round(abs(best["difference"]) / p1)))
    else:
        additional_leg = 0
        additional_shares = 0

    best.update(
        {
            "lot1": lot1,
            "lot2": lot2,
            "one_to_one_value1": one_to_one_value1,
            "one_to_one_value2": one_to_one_value2,
            "one_to_one_difference": one_to_one_difference,
            "additional_leg": additional_leg,
            "additional_shares": additional_shares,
        }
    )
    return best


def estimate_profit(signal, prices, sizing):
    if not prices or signal["action"] != "ENTER PAIR":
        return {}

    q1 = sizing["qty1"]
    q2 = sizing["qty2"]

    if signal["leg1"] == "BUY":
        pnl1 = (prices["target1"] - prices["entry1"]) * q1
    else:
        pnl1 = (prices["entry1"] - prices["target1"]) * q1

    if signal["leg2"] == "BUY":
        pnl2 = (prices["target2"] - prices["entry2"]) * q2
    else:
        pnl2 = (prices["entry2"] - prices["target2"]) * q2

    gross_exposure = sizing["value1"] + sizing["value2"]
    total = pnl1 + pnl2

    return {
        "pnl1": pnl1,
        "pnl2": pnl2,
        "total": total,
        "gross_exposure": gross_exposure,
        "return_pct": total / gross_exposure * 100 if gross_exposure else np.nan,
    }


# ---------------------------------------------------------------------
# TABLE STYLING
# ---------------------------------------------------------------------

def active_signal_row_style(row):
    action = str(row.get("Action", "")).upper()
    if "ENTER" in action:
        return [
            "background-color:#DFF3E3;color:#17351E;font-size:11px;"
            for _ in row
        ]
    return [
        "background-color:#ECEFF1;color:#263238;font-size:11px;"
        for _ in row
    ]


# ---------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------

detected_workbook = locate_default_workbook()

with st.sidebar:
    st.title("⚙️ Controls")

    z_entry = st.slider(
        "Z-score entry threshold",
        1.0,
        3.0,
        1.5,
        .1,
    )
    z_exit = st.slider(
        "Z-score exit threshold",
        .1,
        1.5,
        .5,
        .1,
    )
    lookback = st.selectbox(
        "Z-score lookback",
        [30, 60, 90],
        index=1,
    )

    st.markdown("---")
    st.caption("Local Excel source")

    excel_path = st.text_input(
        "Workbook path",
        value=detected_workbook,
        placeholder=r"C:\Users\YourName\Documents\Pair trade1.xlsx",
    )
    signal_sheet = st.text_input(
        "Signal sheet",
        value=DEFAULT_SIGNAL_SHEET,
    )

    auto_refresh = st.checkbox("Auto-refresh every 60 seconds", value=True)
    show_charts = st.checkbox("Show detailed charts", value=True)

    if st.button("🔄 Refresh all data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


# ---------------------------------------------------------------------
# LOAD EXCEL
# ---------------------------------------------------------------------

signal_excel_df = pd.DataFrame()
excel_error = ""
excel_modified = None

if excel_path:
    workbook = Path(excel_path).expanduser()
    if workbook.is_file():
        try:
            stat = workbook.stat()
            excel_modified = datetime.fromtimestamp(stat.st_mtime)
            signal_excel_df = read_excel_signal(
                str(workbook.resolve()),
                signal_sheet,
                stat.st_mtime_ns,
            )
        except Exception as exc:
            excel_error = str(exc)
    else:
        excel_error = "The workbook path does not exist on the Streamlit computer."

excel_pairs = pairs_from_signal_sheet(signal_excel_df)
PAIR_SOURCE = excel_pairs if excel_pairs else DEFAULT_PAIRS

all_sectors = sorted({p["sector"] for p in PAIR_SOURCE})

with st.sidebar:
    sector_filter = st.multiselect(
        "Filter analytical scan by sector",
        options=all_sectors,
        default=all_sectors,
    )

    st.markdown("---")
    st.caption("Market source: Yahoo Finance")
    if excel_modified:
        st.caption(
            f"Saved Excel version: {excel_modified:%Y-%m-%d %H:%M:%S}"
        )
    elif excel_path:
        st.caption("Excel source unavailable")
    else:
        st.caption("Using built-in pair definitions")


# ---------------------------------------------------------------------
# HEADER AND TOP REFRESH CONTROL
# ---------------------------------------------------------------------

head1, head2 = st.columns([5, 1])
with head1:
    st.markdown(
        '<p class="main-header">📈 Automated Pair Trading Dashboard</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<p class="sub-header">Excel signal view + independent market analytics · '
        f'Updated {datetime.now():%Y-%m-%d %H:%M:%S} IST</p>',
        unsafe_allow_html=True,
    )

with head2:
    if st.button("🔄 Refresh", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

if not HAS_YF:
    st.error("Install yfinance with: pip install yfinance")
    st.stop()


# ---------------------------------------------------------------------
# TOP: COMPLETE EXCEL SIGNAL SHEET
# ---------------------------------------------------------------------

st.subheader("📋 Excel SIGNAL Dashboard")

if excel_error:
    st.warning(
        f"Excel SIGNAL sheet is not currently available: {excel_error} "
        "The independent scan below will use the built-in pair list."
    )
elif signal_excel_df.empty:
    st.info(
        "Enter the local workbook path in the sidebar. Until then, the "
        "independent dashboard uses the built-in pairs."
    )
else:
    source_note = (
        f"Loaded {len(signal_excel_df):,} rows from "
        f"`{Path(excel_path).name}` / `{signal_sheet}`. "
        f"Workbook saved: {excel_modified:%Y-%m-%d %H:%M:%S}."
    )
    st.caption(source_note)
    dataframe_as_scroll_table(signal_excel_df)


# ---------------------------------------------------------------------
# SCAN PAIRS
# ---------------------------------------------------------------------

pairs_to_scan = [
    p for p in PAIR_SOURCE
    if p["sector"] in sector_filter
]

if not pairs_to_scan:
    st.warning("Select at least one sector.")
    st.stop()

progress = st.progress(0, text="Fetching market data...")
results = []

for index, pair in enumerate(pairs_to_scan):
    progress.progress(
        (index + 1) / len(pairs_to_scan),
        text=f"Scanning {pair['name']}...",
    )
    history = fetch_pair_history(pair["s1"], pair["s2"], period="1y")
    metrics = compute_pair_metrics(history, lookback)
    signal = generate_signal(metrics, z_entry, z_exit)
    prices = recommend_prices(metrics, signal)

    results.append(
        {
            "pair": pair,
            "metrics": metrics,
            "signal": signal,
            "prices": prices,
            "hist": history,
        }
    )

progress.empty()


# ---------------------------------------------------------------------
# ACTIVE LIVE SIGNAL BOARD: ENTER AND EXIT/HOLD ONLY
# ---------------------------------------------------------------------

st.subheader("🎯 Active Live Signal Board")

active_rows = []
for result in results:
    m = result["metrics"]
    s = result["signal"]
    pair = result["pair"]

    if not m or s["action"] not in {"ENTER PAIR", "EXIT / HOLD"}:
        continue

    active_rows.append(
        {
            "Pair": pair["name"],
            "Sector": pair["sector"],
            "Price 1": m["price1"],
            "Price 2": m["price2"],
            "Ratio": m["ratio"],
            "Z-Score": m["zscore"],
            "Corr 30D": m["corr_30"],
            "Corr 60D": m["corr_60"],
            "Action": s["action"],
            "Leg 1": s["leg1"],
            "Leg 2": s["leg2"],
            "Confidence": s["confidence"],
            "Half-Life": m["half_life"],
        }
    )

if active_rows:
    active_df = pd.DataFrame(active_rows)
    styled_active = (
        active_df.style
        .apply(active_signal_row_style, axis=1)
        .format(
            {
                "Price 1": "{:,.2f}",
                "Price 2": "{:,.2f}",
                "Ratio": "{:.4f}",
                "Z-Score": "{:.2f}",
                "Corr 30D": "{:.2f}",
                "Corr 60D": "{:.2f}",
                "Confidence": "{:.0f}%",
                "Half-Life": "{:.1f}",
            },
            na_rep="—",
        )
    )
    st.dataframe(
        styled_active,
        use_container_width=True,
        hide_index=True,
        height=min(420, 42 + len(active_df) * 35),
    )
else:
    st.info("No ENTER PAIR or EXIT / HOLD signals are available.")


# ---------------------------------------------------------------------
# COMPACT SUMMARY METRICS
# ---------------------------------------------------------------------

valid_results = [r for r in results if r["metrics"]]
enter_count = sum(
    r["signal"]["action"] == "ENTER PAIR"
    for r in valid_results
)
exit_count = sum(
    r["signal"]["action"] == "EXIT / HOLD"
    for r in valid_results
)
average_z = (
    np.mean([abs(r["metrics"]["zscore"]) for r in valid_results])
    if valid_results
    else np.nan
)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Pairs Scanned", len(valid_results))
m2.metric("Enter Signals", enter_count)
m3.metric("Exit / Hold", exit_count)
m4.metric("Average |Z|", f"{average_z:.2f}" if np.isfinite(average_z) else "—")

if not valid_results:
    st.warning("No market data was returned. Check ticker symbols and internet access.")
    st.stop()


# ---------------------------------------------------------------------
# PAIR DEEP DIVE
# ---------------------------------------------------------------------

st.markdown("---")
st.subheader("🔍 Pair Deep-Dive")

pair_names = [r["pair"]["name"] for r in valid_results]
selected_name = st.selectbox(
    "Select pair",
    pair_names,
)

selected_result = next(
    r for r in valid_results
    if r["pair"]["name"] == selected_name
)
m = selected_result["metrics"]
s = selected_result["signal"]
pr = selected_result["prices"]
pair_info = selected_result["pair"]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Current Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if np.isfinite(m["corr_30"]) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if np.isfinite(m["corr_60"]) else "—")
k5.metric("Half-Life", f"{m['half_life']:.1f} d" if np.isfinite(m["half_life"]) else "—")

if s["action"] == "ENTER PAIR":
    signal_bg, signal_fg = "#DFF3E3", "#17351E"
elif s["action"] == "EXIT / HOLD":
    signal_bg, signal_fg = "#ECEFF1", "#263238"
else:
    signal_bg, signal_fg = "#FFF4CC", "#4E4200"

st.markdown(
    f"""
    <div style="
        background:{signal_bg};
        color:{signal_fg};
        padding:10px 13px;
        border-radius:9px;
        border:1px solid rgba(40,60,70,.14);
        margin:7px 0 10px 0;
        font-size:.88rem;">
        <b>{html.escape(s["action"])}</b> · Confidence {s["confidence"]}%<br>
        <b>{html.escape(pair_info["s1"])}</b>: {html.escape(s["leg1"])}
        &nbsp; | &nbsp;
        <b>{html.escape(pair_info["s2"])}</b>: {html.escape(s["leg2"])}
        <br><span style="font-size:.78rem;">{html.escape(s["reason"])}</span>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------
# ENTRY / TARGET / STOP
# ---------------------------------------------------------------------

if pr:
    st.markdown("### Suggested Entry, Target and Stop")

    price_df = pd.DataFrame(
        [
            {
                "Leg": pair_info["s1"],
                "Direction": s["leg1"],
                "Entry": pr["entry1"],
                "Target": pr["target1"],
                "Stop": pr["stop1"],
                "Support": pr["support1"],
                "Resistance": pr["resist1"],
            },
            {
                "Leg": pair_info["s2"],
                "Direction": s["leg2"],
                "Entry": pr["entry2"],
                "Target": pr["target2"],
                "Stop": pr["stop2"],
                "Support": pr["support2"],
                "Resistance": pr["resist2"],
            },
        ]
    )
    st.dataframe(
        price_df.style.format(
            {
                "Entry": "₹{:,.2f}",
                "Target": "₹{:,.2f}",
                "Stop": "₹{:,.2f}",
                "Support": "₹{:,.2f}",
                "Resistance": "₹{:,.2f}",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )


# ---------------------------------------------------------------------
# CASH NEUTRALITY, 1:1 DIFFERENCE AND EXTRA SHARES
# ---------------------------------------------------------------------

st.markdown("### ⚖️ Lot Size and Cash Neutrality")

sizing = optimize_cash_neutral_lots(pair_info, m)

cash_df = pd.DataFrame(
    [
        {
            "Leg": pair_info["s1"],
            "Exchange lot": sizing["lot1"],
            "Recommended contracts": sizing["contracts1"],
            "Recommended quantity": sizing["qty1"],
            "Cash value": sizing["value1"],
        },
        {
            "Leg": pair_info["s2"],
            "Exchange lot": sizing["lot2"],
            "Recommended contracts": sizing["contracts2"],
            "Recommended quantity": sizing["qty2"],
            "Cash value": sizing["value2"],
        },
    ]
)

st.dataframe(
    cash_df.style.format({"Cash value": "₹{:,.2f}"}),
    hide_index=True,
    use_container_width=True,
)

c1, c2, c3 = st.columns(3)
c1.metric(
    "Recommended mismatch",
    inr(abs(sizing["difference"])),
    f"{sizing['mismatch_pct']:.2f}%",
)
c2.metric(
    "1:1 lot mismatch",
    inr(abs(sizing["one_to_one_difference"])),
)
c3.metric(
    "Recommended gross value",
    inr(sizing["value1"] + sizing["value2"]),
)

if sizing["additional_leg"]:
    extra_ticker = pair_info[f"s{sizing['additional_leg']}"]
    st.caption(
        f"After the recommended whole-lot combination, approximately "
        f"{sizing['additional_shares']} additional shares of {extra_ticker} "
        f"would reduce the remaining cash-value difference. Verify whether "
        f"the extra cash-market position matches your intended long/short "
        f"direction, broker rules, margin and transaction costs."
    )
else:
    st.caption("The recommended whole-lot combination is already cash matched.")

st.caption(
    "Cash neutrality equalizes rupee market values; it does not guarantee "
    "beta, volatility or factor neutrality. The search is capped at 40 "
    "contracts per leg to prevent impractically large recommendations."
)


# ---------------------------------------------------------------------
# PROFIT POTENTIAL
# ---------------------------------------------------------------------

st.markdown("### 💰 Estimated Profit Potential")

profit = estimate_profit(s, pr, sizing)
if profit:
    p1c, p2c, p3c, p4c = st.columns(4)
    p1c.metric(f"{pair_info['s1']} target P&L", inr(profit["pnl1"]))
    p2c.metric(f"{pair_info['s2']} target P&L", inr(profit["pnl2"]))
    p3c.metric("Combined target P&L", inr(profit["total"]))
    p4c.metric("Return on gross value", f"{profit['return_pct']:.2f}%")

    st.caption(
        "The estimate uses the recommended cash-neutral quantities and the "
        "displayed target prices. It excludes brokerage, taxes, slippage, "
        "dividends, borrowing costs, margin interest and execution mismatch."
    )
else:
    st.caption("Profit potential is shown only for an active ENTER PAIR signal.")


# ---------------------------------------------------------------------
# LAST FIVE TRADING DAYS
# ---------------------------------------------------------------------

st.markdown("### 📅 Last 5 Trading Days - Ratio Divergence")

ratio5 = m["ratio_series"].dropna().tail(5)
mean5 = m["ratio_mean_series"].reindex(ratio5.index)

day_cells = []
for position, (date, ratio_value) in enumerate(ratio5.items()):
    mean_value = mean5.loc[date] if date in mean5.index else np.nan

    if np.isfinite(mean_value) and mean_value != 0:
        divergence = (ratio_value / mean_value - 1) * 100
    else:
        divergence = np.nan

    if position == 0:
        daily_change = np.nan
    else:
        previous = ratio5.iloc[position - 1]
        daily_change = (ratio_value / previous - 1) * 100

    if not np.isfinite(daily_change) or abs(daily_change) < .01:
        dot_class = "dot-flat"
        arrow = "•"
    elif daily_change > 0:
        dot_class = "dot-up"
        arrow = "↑"
    else:
        dot_class = "dot-down"
        arrow = "↓"

    divergence_text = (
        f"{divergence:+.2f}%"
        if np.isfinite(divergence)
        else "—"
    )

    day_cells.append(
        f"""
        <div class="day-cell">
            <b>{date:%d %b}</b><br>
            <span class="{dot_class}"></span><br>
            {arrow} Ratio {ratio_value:.4f}<br>
            <span style="color:#455A64;">Div {divergence_text}</span>
        </div>
        """
    )

st.markdown(
    f"<div class='day-grid'>{''.join(day_cells)}</div>",
    unsafe_allow_html=True,
)
st.caption(
    "Green means the ratio increased from the prior trading day; red means "
    "it decreased; grey means the move was effectively flat. Divergence is "
    "the percentage distance from the rolling ratio mean."
)


# ---------------------------------------------------------------------
# CHARTS: MOBILE-SAFE CONFIGURATION
# ---------------------------------------------------------------------

MOBILE_SAFE_CONFIG = {
    "displayModeBar": False,
    "scrollZoom": False,
    "doubleClick": False,
    "responsive": True,
    "showTips": False,
}

if show_charts and HAS_PLOTLY:
    history = selected_result["hist"]
    ratio_series = m["ratio_series"]
    z_series = m["zscore_series"]

    st.markdown("### 📊 Normalized Price Performance")

    normalized1 = history["price1"] / history["price1"].iloc[0] * 100
    normalized2 = history["price2"] / history["price2"].iloc[0] * 100

    price_fig = go.Figure()
    price_fig.add_trace(
        go.Scatter(
            x=history.index,
            y=normalized1,
            name=pair_info["s1"],
            line={"color": "#1976D2", "width": 2},
        )
    )
    price_fig.add_trace(
        go.Scatter(
            x=history.index,
            y=normalized2,
            name=pair_info["s2"],
            line={"color": "#E53935", "width": 2},
        )
    )
    price_fig.update_layout(
        height=360,
        margin={"l": 35, "r": 20, "t": 20, "b": 75},
        template="plotly_white",
        dragmode=False,
        hovermode="x unified",
        legend={
            "orientation": "h",
            "yanchor": "top",
            "y": -0.20,
            "xanchor": "center",
            "x": .5,
        },
    )
    st.plotly_chart(
        price_fig,
        use_container_width=True,
        config=MOBILE_SAFE_CONFIG,
    )

    st.markdown("### Ratio and Reference Levels")

    ratio_fig = go.Figure()
    ratio_fig.add_trace(
        go.Scatter(
            x=ratio_series.index,
            y=ratio_series,
            name="Ratio",
            line={"color": "#1565C0", "width": 2},
        )
    )
    ratio_fig.add_hline(
        y=m["ratio_resist"],
        line_color="#E53935",
        line_dash="dash",
    )
    ratio_fig.add_hline(
        y=m["ratio_q75"],
        line_color="#F9A825",
        line_dash="dot",
    )
    ratio_fig.add_hline(
        y=m["ratio_mean"],
        line_color="#43A047",
        line_dash="dash",
    )
    ratio_fig.add_hline(
        y=m["ratio_support"],
        line_color="#1E88E5",
        line_dash="dot",
    )
    ratio_fig.update_layout(
        height=390,
        margin={"l": 35, "r": 20, "t": 20, "b": 40},
        template="plotly_white",
        dragmode=False,
        hovermode="x unified",
        showlegend=False,
    )
    st.plotly_chart(
        ratio_fig,
        use_container_width=True,
        config=MOBILE_SAFE_CONFIG,
    )
    st.markdown(
        """
        <div class="legend-note">
        <b style="color:#E53935;">Red</b>: 60-day maximum/resistance ·
        <b style="color:#F9A825;">Yellow</b>: upper quartile ·
        <b style="color:#43A047;">Green</b>: 60-day average/mean ·
        <b style="color:#1E88E5;">Blue</b>: 60-day minimum/support.
        These are statistical reference levels, not guaranteed reversal prices.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### Z-Score History")

    z_fig = go.Figure()
    z_fig.add_trace(
        go.Scatter(
            x=z_series.index,
            y=z_series,
            name="Z-score",
            line={"color": "#7B1FA2", "width": 2},
        )
    )
    z_fig.add_hline(
        y=z_entry,
        line_color="#E53935",
        line_dash="dash",
    )
    z_fig.add_hline(
        y=-z_entry,
        line_color="#43A047",
        line_dash="dash",
    )
    z_fig.add_hline(
        y=m["max_z"],
        line_color="#F57C00",
        line_dash="dot",
    )
    z_fig.add_hline(
        y=m["avg_z"],
        line_color="#607D8B",
        line_dash="dot",
    )
    z_fig.add_hline(
        y=m["min_z"],
        line_color="#0288D1",
        line_dash="dot",
    )
    z_fig.add_hline(
        y=0,
        line_color="#263238",
        line_width=1,
    )
    z_fig.update_layout(
        height=410,
        margin={"l": 35, "r": 20, "t": 20, "b": 40},
        template="plotly_white",
        dragmode=False,
        hovermode="x unified",
        showlegend=False,
    )
    st.plotly_chart(
        z_fig,
        use_container_width=True,
        config=MOBILE_SAFE_CONFIG,
    )
    st.markdown(
        f"""
        <div class="legend-note">
        <b style="color:#E53935;">Red</b>: positive entry threshold +{z_entry:.1f} ·
        <b style="color:#43A047;">Green</b>: negative entry threshold -{z_entry:.1f} ·
        <b style="color:#F57C00;">Orange</b>: historical maximum {m["max_z"]:.2f} ·
        <b style="color:#607D8B;">Grey</b>: historical average {m["avg_z"]:.2f} ·
        <b style="color:#0288D1;">Blue</b>: historical minimum {m["min_z"]:.2f}.
        A larger absolute Z-score is a rarer deviation, but it can also indicate
        structural breakdown rather than a better trade.
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------
# SELECTED-PAIR ANALYTICS AND RATIONALE
# ---------------------------------------------------------------------

st.markdown("### 🧠 Selected-Pair Analytics and Rationale")

corr_quality = (
    "strong"
    if np.isfinite(m["corr_60"]) and abs(m["corr_60"]) >= .70
    else "moderate"
    if np.isfinite(m["corr_60"]) and abs(m["corr_60"]) >= .50
    else "weak"
)

half_life_text = (
    f"{m['half_life']:.1f} trading days"
    if np.isfinite(m["half_life"])
    else "not reliably estimable"
)

percentile_text = (
    f"{m['z_percentile']:.1f}th percentile"
    if np.isfinite(m["z_percentile"])
    else "unavailable"
)

analysis_points = [
    (
        f"**Signal state:** {s['action']} with Z-score {m['zscore']:.2f}. "
        f"{s['reason']}"
    ),
    (
        f"**Relationship quality:** 60-day return correlation is "
        f"{m['corr_60']:.2f} ({corr_quality})"
        if np.isfinite(m["corr_60"])
        else "**Relationship quality:** 60-day correlation is unavailable."
    ),
    (
        f"**Mean-reversion speed:** estimated half-life is {half_life_text}. "
        "A very long or unstable half-life can tie up capital and weaken the setup."
    ),
    (
        f"**Historical rarity:** current absolute Z-score is at approximately "
        f"the {percentile_text} of the available history."
    ),
    (
        f"**Recent divergence:** the ratio changed "
        f"{m['spread_change_5d'] * 100:+.2f}% over five trading days."
        if np.isfinite(m["spread_change_5d"])
        else "**Recent divergence:** insufficient data for a five-day comparison."
    ),
    (
        f"**Ratio volatility:** annualized 60-day ratio volatility is "
        f"{m['ratio_volatility'] * 100:.2f}%."
    ),
    (
        f"**Historical behavior:** the last available year contains "
        f"{m['extreme_count']} observations with |Z| at least 1.5 and "
        f"{m['zero_crossings']} zero crossings."
    ),
]

for point in analysis_points:
    st.markdown(f"- {point}")

st.info(
    "Historical extreme Z-scores should not be treated as automatic high-win-rate "
    "entries. A practical setup combines a sufficiently extreme Z-score with "
    "stable correlation, plausible half-life, liquidity, transaction costs and "
    "evidence that the economic relationship has not structurally changed."
)


# ---------------------------------------------------------------------
# ALL-PAIR INDEPENDENT NOTES
# ---------------------------------------------------------------------

st.markdown("---")
st.subheader("🌈 Independent Analytics Across All Pairs")

entry_results = [
    r for r in valid_results
    if r["signal"]["action"] == "ENTER PAIR"
]

if entry_results:
    cards = st.columns(min(3, len(entry_results)))
    for index, result in enumerate(entry_results):
        with cards[index % len(cards)]:
            rm = result["metrics"]
            rs = result["signal"]
            st.markdown(
                f"""
                <div style="
                    background:linear-gradient(135deg,#E8F5E9,#F3FFF5);
                    border-left:5px solid #43A047;
                    border-radius:9px;
                    padding:10px;
                    margin-bottom:8px;
                    min-height:125px;">
                    <b>{html.escape(result["pair"]["name"])}</b><br>
                    <span style="font-size:.78rem;">
                    Z: {rm["zscore"]:.2f}<br>
                    Corr 30D: {rm["corr_30"]:.2f}<br>
                    Confidence: {rs["confidence"]}%<br>
                    {html.escape(rs["leg1"])} / {html.escape(rs["leg2"])}
                    </span>
                </div>
                """,
                unsafe_allow_html=True,
            )
else:
    st.success(
        "No high-conviction entry signals are currently present. "
        "The scanned pairs are in exit, hold or watch zones."
    )


# ---------------------------------------------------------------------
# HISTORICAL PERFORMANCE SUMMARY
# ---------------------------------------------------------------------

st.subheader("📚 Historical Pair Performance Summary")

history_rows = []
for result in valid_results:
    rm = result["metrics"]
    history_rows.append(
        {
            "Pair": result["pair"]["name"],
            "Current Z": rm["zscore"],
            "Historical Max Z": rm["max_z"],
            "Historical Avg Z": rm["avg_z"],
            "Historical Min Z": rm["min_z"],
            "|Z| Percentile": rm["z_percentile"],
            "Extreme Days": rm["extreme_count"],
            "Zero Crossings": rm["zero_crossings"],
            "Half-Life": rm["half_life"],
        }
    )

history_df = pd.DataFrame(history_rows)

def history_color(value):
    if not isinstance(value, (int, float, np.number)) or not np.isfinite(value):
        return ""
    if value >= 2:
        return "background-color:#FFCDD2;color:#5D1117;"
    if value <= -2:
        return "background-color:#C8E6C9;color:#12351A;"
    return ""

styled_history = (
    history_df.style
    .map(
        history_color,
        subset=["Current Z", "Historical Max Z", "Historical Min Z"],
    )
    .format(
        {
            "Current Z": "{:.2f}",
            "Historical Max Z": "{:.2f}",
            "Historical Avg Z": "{:.2f}",
            "Historical Min Z": "{:.2f}",
            "|Z| Percentile": "{:.1f}%",
            "Half-Life": "{:.1f}",
        },
        na_rep="—",
    )
)

st.dataframe(
    styled_history,
    use_container_width=True,
    hide_index=True,
    height=min(460, 42 + len(history_df) * 35),
)


# ---------------------------------------------------------------------
# METHODOLOGY AND REFRESH
# ---------------------------------------------------------------------

with st.expander("Methodology and important limitations"):
    st.markdown(
        """
1. The Excel SIGNAL dashboard displays the workbook's **last saved values**.
2. Market analysis uses adjusted daily prices from Yahoo Finance.
3. The ratio is Stock 1 price divided by Stock 2 price.
4. The Z-score compares the ratio with its rolling mean and standard deviation.
5. An entry requires the configured Z-score threshold and correlation filter.
6. Cash-neutral sizing searches whole exchange-lot combinations with similar
   rupee market values.
7. Profit potential is a scenario estimate, not an executable guarantee.
8. Corporate actions, changing fundamentals, liquidity, slippage, short-sale
   constraints and taxes can invalidate a historical relationship.

This dashboard is for research and education, not individualized financial advice.
        """
    )

if auto_refresh:
    time.sleep(60)
    st.cache_data.clear()
    st.rerun()
