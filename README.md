# Pair Trade Dashboard

Automated pair-trading signal dashboard built with Streamlit.

## Features

- **Live market data** via Yahoo Finance (`yfinance`) for NSE stocks & indices
- **18 pre-configured pairs** matching the original Excel SIGNAL sheet
- **Independent analytics engine**:
  - Price ratio & rolling Z-score (mean-reversion)
  - 30 / 60 / 90-day correlation
  - Half-life of mean reversion (Ornstein-Uhlenbeck)
  - Support / Resistance levels
- **Trade recommendations**: Entry, Target, Stop for both legs
- **Auto-refresh** every 60 seconds
- Interactive Plotly charts (normalized prices, ratio, Z-score)

## Quick Start

```bash
# 1. Clone this repo
git clone <your-repo-url>
cd <repo-folder>

# 2. Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the dashboard
streamlit run pair_trade_dashboard.py
```

Open the URL shown in the terminal (usually http://localhost:8501).

## How Signals Work

| Condition | Action |
|-----------|--------|
| \|Z-score\| ≥ Entry threshold **and** Corr ≥ 0.5 | **ENTER PAIR** |
| Positive Z (ratio high) | Sell Stock1 / Buy Stock2 |
| Negative Z (ratio low) | Buy Stock1 / Sell Stock2 |
| \|Z-score\| ≤ Exit threshold | EXIT / HOLD (near mean) |
| Otherwise | WATCH |

You can adjust Entry/Exit Z thresholds and lookback window in the sidebar.

## Pairs Included

| Pair | Sector |
|------|--------|
| ICICIBANK – HDFCBANK | Banking |
| AXISBANK – ICICIBANK | Banking |
| CIPLA – SUNPHARMA | Pharma |
| BPCL – HINDPETRO | Oil & Gas |
| ACC – AMBUJACEM | Cement |
| TCS – HCLTECH | IT |
| TATASTEEL – JSWSTEEL | Metals |
| NESTLE – HINDUNILVR | FMCG |
| NIFTY – BANKNIFTY | Index |
| BRITANNIA – HINDUNILVR | FMCG |
| BAJFINANCE – BAJAJFINSV | Finance |
| HDFCBANK – BANKNIFTY | Banking |
| RECLTD – PFC | Finance |
| IGL – GAIL | Oil & Gas |
| ALKEM – AUROPHARMA | Pharma |
| BERGEPAINT – ASIANPAINT | Paints |
| BANKBARODA – CANBK | Banking |
| COALINDIA – GAIL | Energy |

## Notes

- Data is delayed ~15 min for Indian markets (Yahoo Finance limitation).
- Lot sizes are taken from the original spreadsheet for cash-neutrality reference.
- This is **not financial advice**. Use for research only.

## License

MIT
