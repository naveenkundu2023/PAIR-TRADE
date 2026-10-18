"""
Pair Trading Dashboard - Streamlit App
======================================
Automated pair-trading signal dashboard with independent analytics.

How to run:
    1. pip install streamlit pandas numpy yfinance plotly openpyxl scipy
    2. streamlit run pair_trade_dashboard.py

Features:
    - Loads pair definitions from your Excel (or built-in list)
    - Live price fetch via yfinance (NSE stocks)
    - Correlation, Z-score, Ratio Spread, Mean-reversion signals
    - Independent Grok analytics + Support/Resistance levels
    - Auto-refresh every 60 seconds
    - Buy / Sell price recommendations
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import time
import warnings
warnings.filterwarnings("ignore")

# Optional heavy imports with graceful fallback
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
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E88E5;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #666;
        margin-bottom: 1.5rem;
    }
    .signal-buy {
        background-color: #1B5E20;
        color: white;
        padding: 4px 12px;
        border-radius: 6px;
        font-weight: 600;
    }
    .signal-sell {
        background-color: #B71C1C;
        color: white;
        padding: 4px 12px;
        border-radius: 6px;
        font-weight: 600;
    }
    .signal-hold {
        background-color: #455A64;
        color: white;
        padding: 4px 12px;
        border-radius: 6px;
        font-weight: 600;
    }
    .metric-card {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 12px;
        border-left: 4px solid #1E88E5;
    }
    .stMetric label { font-size: 0.85rem !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# PAIR DEFINITIONS (from your Excel SIGNAL sheet)
# ─────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────
# DATA HELPERS
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=55, show_spinner=False)
def fetch_pair_history(ticker1: str, ticker2: str, period: str = "6mo") -> pd.DataFrame:
    """Fetch adjusted close prices for a pair. Returns aligned DataFrame."""
    if not HAS_YF:
        return pd.DataFrame()
    try:
        t1 = yf.Ticker(ticker1)
        t2 = yf.Ticker(ticker2)
        h1 = t1.history(period=period, auto_adjust=True)
        h2 = t2.history(period=period, auto_adjust=True)
        if h1.empty or h2.empty:
            return pd.DataFrame()
        df = pd.DataFrame({
            "price1": h1["Close"],
            "price2": h2["Close"],
        }).dropna()
        df.index = pd.to_datetime(df.index).tz_localize(None)
        return df
    except Exception:
        return pd.DataFrame()


def compute_pair_metrics(df: pd.DataFrame, lookback: int = 60) -> dict:
    """
    Core pair-trading metrics:
      - Ratio = price1 / price2
      - Z-score of ratio (mean-reversion signal)
      - Rolling correlations (30/60/90)
      - Support / Resistance of ratio and of each leg
      - Divergence from SMA
    """
    if df.empty or len(df) < 30:
        return {}

    ratio = df["price1"] / df["price2"]
    window = min(lookback, len(ratio) - 1)

    # Rolling stats on ratio
    ratio_mean = ratio.rolling(window).mean()
    ratio_std = ratio.rolling(window).std()
    zscore = (ratio - ratio_mean) / ratio_std

    # Correlations
    corr_30 = df["price1"].tail(30).corr(df["price2"].tail(30)) if len(df) >= 30 else np.nan
    corr_60 = df["price1"].tail(60).corr(df["price2"].tail(60)) if len(df) >= 60 else np.nan
    corr_90 = df["price1"].tail(90).corr(df["price2"].tail(90)) if len(df) >= 90 else np.nan

    # Current values
    cur_ratio = float(ratio.iloc[-1])
    cur_z = float(zscore.iloc[-1]) if not np.isnan(zscore.iloc[-1]) else 0.0
    cur_p1 = float(df["price1"].iloc[-1])
    cur_p2 = float(df["price2"].iloc[-1])

    # SMA & bands for each leg
    sma20_1 = float(df["price1"].tail(20).mean())
    sma20_2 = float(df["price2"].tail(20).mean())
    std20_1 = float(df["price1"].tail(20).std())
    std20_2 = float(df["price2"].tail(20).std())

    # Support / Resistance (recent swing + Bollinger-style)
    recent = df.tail(60)
    support1 = float(recent["price1"].min())
    resist1 = float(recent["price1"].max())
    support2 = float(recent["price2"].min())
    resist2 = float(recent["price2"].max())

    # Ratio support/resistance
    ratio_support = float(ratio.tail(60).min())
    ratio_resist = float(ratio.tail(60).max())
    ratio_mean_60 = float(ratio.tail(60).mean())

    # Half-life of mean reversion (Ornstein-Uhlenbeck approximation)
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
        "ratio_support": ratio_support,
        "ratio_resist": ratio_resist,
        "ratio_mean": ratio_mean_60,
        "half_life": half_life,
        "ratio_series": ratio,
        "zscore_series": zscore,
        "df": df,
    }


def generate_signal(metrics: dict, z_entry: float = 1.5, z_exit: float = 0.5) -> dict:
    """
    Independent signal engine.
    Classic pair-trade logic:
      - High positive Z  → ratio too high → SHORT stock1 / LONG stock2
      - High negative Z  → ratio too low  → LONG stock1 / SHORT stock2
    Also requires reasonable correlation.
    """
    if not metrics:
        return {"action": "NO DATA", "leg1": "—", "leg2": "—", "confidence": 0, "reason": "Insufficient data"}

    z = metrics["zscore"]
    corr = metrics.get("corr_30") or metrics.get("corr_60") or 0
    abs_z = abs(z)

    # Correlation quality
    corr_ok = abs(corr) >= 0.5 if not np.isnan(corr) else False

    if abs_z >= z_entry and corr_ok:
        if z > 0:
            # Ratio elevated → expect mean reversion down
            action = "ENTER PAIR"
            leg1 = "SELL"
            leg2 = "BUY"
            reason = f"Z-score +{z:.2f} (ratio stretched high). Expect reversion → Sell leg1 / Buy leg2"
        else:
            action = "ENTER PAIR"
            leg1 = "BUY"
            leg2 = "SELL"
            reason = f"Z-score {z:.2f} (ratio stretched low). Expect reversion → Buy leg1 / Sell leg2"
        confidence = min(95, int(40 + abs_z * 15 + (abs(corr) * 20)))
    elif abs_z <= z_exit:
        action = "EXIT / HOLD"
        leg1 = "CLOSE"
        leg2 = "CLOSE"
        reason = f"Z-score near mean ({z:.2f}). No edge / take profit zone"
        confidence = 30
    else:
        action = "WATCH"
        leg1 = "—"
        leg2 = "—"
        reason = f"Z-score {z:.2f} inside neutral band. Waiting for stretch"
        confidence = 20

    if not corr_ok and action == "ENTER PAIR":
        action = "WATCH (Low Corr)"
        reason += f" | Correlation only {corr:.2f} – pair quality weak"
        confidence = max(10, confidence - 30)

    return {
        "action": action,
        "leg1": leg1,
        "leg2": leg2,
        "confidence": confidence,
        "reason": reason,
        "z": z,
        "corr": corr,
    }


def recommend_prices(metrics: dict, signal: dict) -> dict:
    """
    Suggest entry / target / stop based on Z-score targets and S/R.
    """
    if not metrics or signal["action"] not in ("ENTER PAIR",):
        return {}

    z = metrics["zscore"]
    p1, p2 = metrics["price1"], metrics["price2"]
    ratio = metrics["ratio"]
    mean_r = metrics["ratio_mean"]

    # Target: ratio returns to mean (Z≈0)
    # Stop: further stretch to |Z|≈2.5 or recent extreme
    target_ratio = mean_r
    if z > 0:
        # Selling leg1 / Buying leg2 → target lower ratio
        stop_ratio = metrics["ratio_resist"] * 1.02
    else:
        stop_ratio = metrics["ratio_support"] * 0.98

    # Approximate price targets keeping one leg roughly constant
    # (simplified – real sizing uses hedge ratio)
    if signal["leg1"] == "SELL":
        entry1 = p1
        target1 = p1 * (target_ratio / ratio) * 0.98   # mild buffer
        stop1 = p1 * (stop_ratio / ratio) * 1.01
        entry2 = p2
        target2 = p2 * (ratio / target_ratio) * 1.02
        stop2 = p2 * (ratio / stop_ratio) * 0.99
    else:
        entry1 = p1
        target1 = p1 * (target_ratio / ratio) * 1.02
        stop1 = p1 * (stop_ratio / ratio) * 0.99
        entry2 = p2
        target2 = p2 * (ratio / target_ratio) * 0.98
        stop2 = p2 * (ratio / stop_ratio) * 1.01

    return {
        "entry1": round(entry1, 2),
        "target1": round(target1, 2),
        "stop1": round(stop1, 2),
        "entry2": round(entry2, 2),
        "target2": round(target2, 2),
        "stop2": round(stop2, 2),
        "support1": round(metrics["support1"], 2),
        "resist1": round(metrics["resist1"], 2),
        "support2": round(metrics["support2"], 2),
        "resist2": round(metrics["resist2"], 2),
    }


# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Controls")
    st.markdown("---")

    z_entry = st.slider("Z-score Entry Threshold", 1.0, 3.0, 1.5, 0.1,
                        help="Enter pair trade when |Z| exceeds this")
    z_exit = st.slider("Z-score Exit Threshold", 0.1, 1.5, 0.5, 0.1)
    lookback = st.selectbox("Z-score Lookback (days)", [30, 60, 90], index=1)

    sector_filter = st.multiselect(
        "Filter by Sector",
        options=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
        default=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
    )

    auto_refresh = st.checkbox("Auto-refresh every 60s", value=True)
    show_charts = st.checkbox("Show detailed charts", value=True)

    st.markdown("---")
    st.caption("Data source: Yahoo Finance (yfinance)")
    st.caption("Signals are independent of the original Excel formulas.")
    st.caption(f"Last refresh: {datetime.now().strftime('%H:%M:%S')}")

    if st.button("🔄 Refresh Now"):
        st.cache_data.clear()
        st.rerun()


# ─────────────────────────────────────────────────────────────
# MAIN HEADER
# ─────────────────────────────────────────────────────────────
st.markdown('<p class="main-header">📈 Automated Pair Trading Dashboard</p>', unsafe_allow_html=True)
st.markdown(
    f'<p class="sub-header">Independent analytics • Mean-reversion Z-score engine • '
    f'Updated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} IST</p>',
    unsafe_allow_html=True,
)

if not HAS_YF:
    st.error("yfinance is not installed. Run: `pip install yfinance`")
    st.stop()

# ─────────────────────────────────────────────────────────────
# SCAN ALL PAIRS
# ─────────────────────────────────────────────────────────────
pairs_to_scan = [p for p in DEFAULT_PAIRS if p["sector"] in sector_filter]

progress = st.progress(0, text="Fetching live data…")
results = []

for i, pair in enumerate(pairs_to_scan):
    progress.progress((i + 1) / len(pairs_to_scan), text=f"Scanning {pair['name']}…")
    hist = fetch_pair_history(pair["s1"], pair["s2"], period="6mo")
    metrics = compute_pair_metrics(hist, lookback=lookback)
    signal = generate_signal(metrics, z_entry=z_entry, z_exit=z_exit)
    prices = recommend_prices(metrics, signal)

    results.append({
        "pair": pair,
        "metrics": metrics,
        "signal": signal,
        "prices": prices,
        "hist": hist,
    })

progress.empty()

# ─────────────────────────────────────────────────────────────
# SUMMARY TABLE
# ─────────────────────────────────────────────────────────────
st.subheader("🎯 Live Signal Board")

rows = []
for r in results:
    m = r["metrics"]
    s = r["signal"]
    p = r["pair"]
    if not m:
        continue
    rows.append({
        "Pair": p["name"],
        "Sector": p["sector"],
        "Price 1": round(m["price1"], 2),
        "Price 2": round(m["price2"], 2),
        "Ratio": round(m["ratio"], 4),
        "Z-Score": round(m["zscore"], 2),
        "Corr 30D": round(m["corr_30"], 2) if not np.isnan(m["corr_30"]) else "—",
        "Corr 60D": round(m["corr_60"], 2) if not np.isnan(m["corr_60"]) else "—",
        "Action": s["action"],
        "Leg1": s["leg1"],
        "Leg2": s["leg2"],
        "Confidence": s["confidence"],
        "Half-Life (d)": round(m["half_life"], 1) if not np.isnan(m.get("half_life", np.nan)) else "—",
    })

if not rows:
    st.warning("No data returned. Check internet / ticker symbols.")
    st.stop()

df_signals = pd.DataFrame(rows)

# Color the Action column
def color_action(val):
    if "ENTER" in str(val):
        return "background-color: #1B5E20; color: white; font-weight: bold"
    if "EXIT" in str(val) or "HOLD" in str(val):
        return "background-color: #37474F; color: white"
    if "WATCH" in str(val):
        return "background-color: #F9A825; color: black"
    return ""

styled = df_signals.style.map(color_action, subset=["Action"]).format({
    "Price 1": "{:.2f}",
    "Price 2": "{:.2f}",
    "Ratio": "{:.4f}",
    "Z-Score": "{:.2f}",
})

st.dataframe(styled, use_container_width=True, height=420)

# Quick stats
enter_count = sum(1 for r in rows if "ENTER" in r["Action"])
watch_count = sum(1 for r in rows if "WATCH" in r["Action"])
col1, col2, col3, col4 = st.columns(4)
col1.metric("Pairs Scanned", len(rows))
col2.metric("🟢 Enter Signals", enter_count)
col3.metric("🟡 Watchlist", watch_count)
col4.metric("Avg |Z|", round(np.mean([abs(r["Z-Score"]) for r in rows]), 2))

# ─────────────────────────────────────────────────────────────
# DETAILED VIEW – one pair at a time
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

# KPI row
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Current Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}",
          delta=f"{'Stretched' if abs(m['zscore']) > z_entry else 'Normal'}")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if not np.isnan(m['corr_30']) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if not np.isnan(m['corr_60']) else "—")
k5.metric("Half-Life (days)", f"{m['half_life']:.1f}" if not np.isnan(m.get('half_life', np.nan)) else "—")

# Signal box
action_color = "#1B5E20" if "ENTER" in s["action"] else ("#F9A825" if "WATCH" in s["action"] else "#37474F")
st.markdown(
    f"""
    <div style="background:{action_color};color:white;padding:16px;border-radius:10px;margin:12px 0;">
        <h3 style="margin:0;color:white;">{s['action']} — Confidence {s['confidence']}%</h3>
        <p style="margin:6px 0 0 0;font-size:1.05rem;">
            <b>{pair_info['s1']}</b> → <b>{s['leg1']}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
            <b>{pair_info['s2']}</b> → <b>{s['leg2']}</b>
        </p>
        <p style="margin:8px 0 0 0;opacity:0.9;">{s['reason']}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Price recommendations
if pr:
    st.markdown("#### 💰 Suggested Entry / Target / Stop")
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
        "Targets assume ratio mean-reversion. Stops are placed beyond recent ratio extremes. "
        "Always size positions for approximate cash neutrality (use lot sizes above as starting point). "
        "These are analytical suggestions, not financial advice."
    )
else:
    st.caption("No active entry recommendation for this pair right now.")

# Charts
if show_charts and HAS_PLOTLY and not selected_result["hist"].empty:
    hist = selected_result["hist"]
    ratio_s = m["ratio_series"]
    z_s = m["zscore_series"]

    fig = make_subplots(
        rows=3, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.06,
        subplot_titles=(
            f"Prices — {pair_info['s1']} vs {pair_info['s2']}",
            "Ratio (Price1 / Price2) + Mean",
            "Z-Score of Ratio",
        ),
        row_heights=[0.4, 0.3, 0.3],
    )

    # Normalized prices for comparison
    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100
    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=pair_info["s1"], line=dict(color="#1E88E5")), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=pair_info["s2"], line=dict(color="#E53935")), row=1, col=1)

    # Ratio
    fig.add_trace(go.Scatter(x=ratio_s.index, y=ratio_s, name="Ratio", line=dict(color="#43A047")), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="gray", row=2, col=1)

    # Z-score
    fig.add_trace(go.Scatter(x=z_s.index, y=z_s, name="Z-Score", line=dict(color="#8E24AA")), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="red", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="green", row=3, col=1)
    fig.add_hline(y=0, line_color="gray", row=3, col=1)

    fig.update_layout(height=720, template="plotly_white", showlegend=True,
                      legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# INDEPENDENT ANALYTICS SUMMARY
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("🧠 Independent Analytics Notes")

notes = []
for r in results:
    if not r["metrics"]:
        continue
    s = r["signal"]
    m = r["metrics"]
    if "ENTER" in s["action"]:
        notes.append(
            f"**{r['pair']['name']}** — {s['action']} (conf {s['confidence']}%). "
            f"Z={m['zscore']:.2f}, Corr30={m['corr_30']:.2f}. {s['reason']}"
        )

if notes:
    for n in notes:
        st.markdown(f"- {n}")
else:
    st.write("No high-conviction entry signals at the moment. Market pairs are within normal Z-score bands.")

st.markdown("""
**Methodology (Independent of original Excel):**
1. Fetch 6-month daily closes via Yahoo Finance.
2. Compute price ratio = Stock1 / Stock2.
3. Z-score = (Ratio − RollingMean) / RollingStd (default 60-day window).
4. Enter when |Z| > threshold **and** correlation ≥ 0.5.
5. Direction: positive Z → sell the expensive leg / buy the cheap leg (mean reversion).
6. Support/Resistance taken from 60-day high/low of each leg and of the ratio.
7. Half-life estimated via AR(1) / Ornstein-Uhlenbeck regression.

> ⚠️ This tool is for educational & research purposes only. Past mean-reversion does not guarantee future results. Always manage risk and position size carefully.
""")

# ─────────────────────────────────────────────────────────────
# AUTO REFRESH
# ─────────────────────────────────────────────────────────────
if auto_refresh:
    time.sleep(60)
    st.rerun()
