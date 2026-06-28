# Portfolio Link Market Analysis Backend

This is the Python engine that powers `Report Market Analysis`.

## What it does

- Reads the existing Google Sheets `Holdings` tab sent by Apps Script
- Excludes cash and money market positions
- Pulls market data with `yfinance`, with Stooq fallback
- Calculates 20 EMA, 50 SMA, 200 SMA, MACD, RSI, 5/20/60-day returns, volatility, drawdown, and relative strength versus SPY
- Scores sector rotation using the 11 SPDR sector ETFs
- Builds Raj's exact report sections in JSON for Apps Script to write back into Google Sheets

## Local run

```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn api:app --reload --port 8080
```

Health check:

```bash
curl http://127.0.0.1:8080/health
```

## Deploy

Deploy this `backend` folder to Render, Railway, Fly.io, or Google Cloud Run.

Set this environment variable on the backend host:

```text
BACKEND_API_TOKEN=your-private-token
```

Then set these Apps Script Properties:

```text
MARKET_BACKEND_URL=https://your-backend-domain.com
MARKET_BACKEND_TOKEN=your-private-token
```

Then run in Apps Script or from the sidebar:

```text
buildMarketAnalysisReport
```

## Important

This is a risk-managed research tool. It does not guarantee profits and does not trade automatically.
