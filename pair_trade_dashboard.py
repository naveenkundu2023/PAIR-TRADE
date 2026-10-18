    df_signal = pd.DataFrame(display_rows, columns=headers).dropna(how="all")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return df_signal, excel_map, f"{source_status} • read at {ts}"

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
