from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from io import StringIO
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple
import math
import os
import re
import time

import numpy as np
import pandas as pd
import requests

try:
    import yfinance as yf
except Exception:
    yf = None

SECTOR_ETFS: List[Tuple[str, str]] = [
    ("XLC", "Communication Services"),
    ("XLY", "Consumer Discretionary"),
    ("XLP", "Consumer Staples"),
    ("XLE", "Energy"),
    ("XLF", "Financials"),
    ("XLV", "Healthcare"),
    ("XLI", "Industrials"),
    ("XLB", "Materials"),
    ("XLRE", "Real Estate"),
    ("XLK", "Technology"),
    ("XLU", "Utilities"),
]

MARKET_PROXIES: List[Tuple[str, str]] = [
    ("SPY", "S&P 500 / benchmark"),
    ("QQQ", "Growth / Nasdaq"),
    ("DIA", "Dow / blue chips"),
    ("IWM", "Small caps"),
    ("GLD", "Gold"),
    ("USO", "Oil"),
    ("TLT", "Long bonds / yield pressure"),
    ("UUP", "Dollar"),
    ("HYG", "High yield credit"),
    ("BIL", "T-bills / cash proxy"),
]

TICKER_MAP: Dict[str, Dict[str, str]] = {
    "SPYM": {"sleeve": "Core Equity", "sector_proxy": "SPY", "thesis": "Low-cost S&P 500 core exposure."},
    "DIA": {"sleeve": "Core Equity", "sector_proxy": "DIA", "thesis": "Dow blue-chip exposure."},
    "SCHG": {"sleeve": "Growth Equity", "sector_proxy": "XLK", "thesis": "Large-cap growth tilt; sensitive to rates and tech leadership."},
    "SPMO": {"sleeve": "Momentum Equity", "sector_proxy": "SPY", "thesis": "Momentum factor exposure tied to risk appetite."},
    "BAC": {"sleeve": "Financials", "sector_proxy": "XLF", "thesis": "Bank exposure; sensitive to rates, credit, and yield curve."},
    "MS": {"sleeve": "Financials", "sector_proxy": "XLF", "thesis": "Capital markets and wealth-management exposure."},
    "STT": {"sleeve": "Financials", "sector_proxy": "XLF", "thesis": "Custody bank / asset-servicing exposure."},
    "SONY": {"sleeve": "Consumer / ADR", "sector_proxy": "XLY", "thesis": "Consumer technology, gaming, media, and ADR exposure."},
    "LMT": {"sleeve": "Defense / Industrials", "sector_proxy": "XLI", "thesis": "Defense industrial; can act as geopolitical hedge."},
    "HTD": {"sleeve": "Income Equity", "sector_proxy": "XLU", "thesis": "Dividend-income sleeve with utility/financial income profile."},
    "SGOL": {"sleeve": "Gold / Alternative", "sector_proxy": "GLD", "thesis": "Gold hedge against real-rate, dollar, and geopolitical stress."},
    "JPST": {"sleeve": "Short Duration Safety", "sector_proxy": "BIL", "thesis": "Ultra-short income stabilizer."},
    "VRIG": {"sleeve": "Floating Rate Safety", "sector_proxy": "BIL", "thesis": "Floating-rate investment-grade income stabilizer."},
    "CLOZ": {"sleeve": "Credit Income", "sector_proxy": "HYG", "thesis": "CLO credit-income sleeve; sensitive to credit spreads."},
}

HOLDINGS_HEADERS = [
    "institution_name", "account_name", "account_official_name", "account_mask",
    "account_type", "account_subtype", "ticker_symbol", "security_name",
    "security_type", "security_subtype", "quantity", "cost_basis",
    "institution_price", "institution_value", "calculated_market_value",
    "unrealized_gain_loss", "unrealized_gain_loss_pct", "portfolio_weight",
    "account_weight", "close_price", "close_price_as_of", "pulled_at", "item_id", "account_id",
]

@dataclass
class Holding:
    ticker: str
    name: str
    security_type: str
    security_subtype: str
    quantity: float
    cost_basis: float
    price: float
    value: float
    unrealized: float
    pnl_pct: float
    portfolio_weight: float
    account_name: str = ""
    institution_name: str = ""

@dataclass
class TechnicalSnapshot:
    ticker: str
    price: float
    ema20: float
    sma50: float
    sma200: float
    macd: float
    macd_signal: float
    macd_hist: float
    rsi14: float
    ret5d: float
    ret20d: float
    ret60d: float
    vol20: float
    drawdown63d: float
    dist_52w_high: float
    trend_score: float
    momentum_score: float
    risk_score: float
    composite_score: float
    price_vs_20ema: str
    price_vs_50sma: str
    price_vs_200sma: str
    cross_20_50: str
    trend_50_200: str
    macd_vs_signal: str
    macd_zero: str
    read: str
    source: str

def money(x: Any) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def pct(x: Any) -> str:
    try:
        if x is None or str(x) == "":
            return "n/a"
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return "n/a"

def num(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return 0.0 if math.isnan(float(value)) or math.isinf(float(value)) else float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[$,%()\s]", "", s).replace(",", "")
    try:
        out = float(s)
        return -out if neg else out
    except Exception:
        return 0.0

def num_pct(value: Any) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        v = float(value)
        return v / 100.0 if abs(v) > 1 else v
    s = str(value).strip()
    v = num(s)
    return v / 100.0 if "%" in s else v

def normalize_header(h: Any) -> str:
    return re.sub(r"[^a-z0-9_]", "", re.sub(r"\s+", "_", str(h or "").strip().lower()))

def parse_holdings_from_rows(rows: Sequence[Sequence[Any]], headers: Optional[Sequence[str]] = None) -> List[Holding]:
    if not rows:
        return []
    norm_rows: List[List[Any]] = []
    for row in rows:
        if len(row) == 1 and "\t" in str(row[0]):
            norm_rows.append(str(row[0]).split("\t"))
        else:
            norm_rows.append(list(row))

    header_row_idx = None
    header_map: Dict[str, int] = {}
    if headers:
        header_map = {normalize_header(h): i for i, h in enumerate(headers) if normalize_header(h)}

    if not header_map:
        for i, row in enumerate(norm_rows[:75]):
            ix = {normalize_header(h): j for j, h in enumerate(row) if normalize_header(h)}
            score = sum(1 for k in ("ticker_symbol", "security_name", "quantity", "institution_value") if k in ix)
            if score >= 2:
                header_row_idx = i
                header_map = ix
                break

    start = header_row_idx + 1 if header_row_idx is not None else 0
    if not header_map:
        header_map = {h: i for i, h in enumerate(HOLDINGS_HEADERS)}

    def get(row: Sequence[Any], key: str) -> Any:
        idx = header_map.get(key)
        return "" if idx is None or idx >= len(row) else row[idx]

    holdings: List[Holding] = []
    for row in norm_rows[start:]:
        joined = "|".join(normalize_header(x) for x in row)
        if "ticker_symbol" in joined or "security_name" in joined:
            continue
        ticker = str(get(row, "ticker_symbol") or "").strip().upper()
        name = str(get(row, "security_name") or "").strip()
        security_type = str(get(row, "security_type") or "").strip().lower()
        if not ticker and not name:
            continue
        if is_excluded_holding(ticker, name, security_type):
            continue
        quantity = num(get(row, "quantity"))
        value = num(get(row, "institution_value")) or num(get(row, "calculated_market_value"))
        price = num(get(row, "institution_price"))
        if not price and value and quantity:
            price = value / quantity
        holdings.append(Holding(
            ticker=ticker or "(NO TICKER)",
            name=name,
            security_type=security_type,
            security_subtype=str(get(row, "security_subtype") or "").strip().lower(),
            quantity=quantity,
            cost_basis=num(get(row, "cost_basis")),
            price=price,
            value=value,
            unrealized=num(get(row, "unrealized_gain_loss")),
            pnl_pct=num_pct(get(row, "unrealized_gain_loss_pct")),
            portfolio_weight=num_pct(get(row, "portfolio_weight")),
            account_name=str(get(row, "account_name") or "").strip(),
            institution_name=str(get(row, "institution_name") or "").strip(),
        ))
    return holdings

def is_excluded_holding(ticker: str, name: str, security_type: str) -> bool:
    t = (ticker or "").upper()
    n = (name or "").lower()
    st = (security_type or "").lower()
    return st == "cash" or t.startswith("CUR:") or t in {"VMFXX", "SWVXX", "SPAXX", "FDRXX"} or "money market" in n or "sweep" in n

class MarketDataClient:
    def __init__(self, lookback: str = "2y", sleep_seconds: float = 0.05):
        self.lookback = lookback
        self.sleep_seconds = sleep_seconds
        self.cache: Dict[str, pd.DataFrame] = {}
        self.status: Dict[str, str] = {}
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept": "application/json,text/csv,text/plain,*/*",
        })

    def get_history(self, ticker: str) -> pd.DataFrame:
        ticker = ticker.upper().strip()
        if ticker in self.cache:
            return self.cache[ticker].copy()
        errors: List[str] = []
        for name, fn in [
            ("Yahoo chart API", self._fetch_yahoo_chart),
            ("yfinance", self._fetch_yfinance),
            ("Stooq CSV", self._fetch_stooq),
        ]:
            try:
                df = fn(ticker)
                if not df.empty and len(df) >= 50:
                    df = df.sort_index()
                    df = df[~df.index.duplicated(keep="last")]
                    df.attrs["source"] = name
                    self.cache[ticker] = df.copy()
                    self.status[ticker] = f"OK: {name}, rows={len(df)}"
                    if self.sleep_seconds:
                        time.sleep(self.sleep_seconds)
                    return df
                errors.append(f"{name}: empty/short rows={len(df) if df is not None else 0}")
            except Exception as exc:
                errors.append(f"{name}: {type(exc).__name__}: {exc}")
        self.status[ticker] = "FAILED: " + " | ".join(errors[:3])
        raise ValueError(self.status[ticker])

    def _fetch_yahoo_chart(self, ticker: str) -> pd.DataFrame:
        urls = [
            f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?range={self.lookback}&interval=1d&events=history&includeAdjustedClose=true",
            f"https://query2.finance.yahoo.com/v8/finance/chart/{ticker}?range={self.lookback}&interval=1d&events=history&includeAdjustedClose=true",
        ]
        last_error = None
        for url in urls:
            try:
                res = self.session.get(url, timeout=20)
                if res.status_code >= 400:
                    last_error = f"HTTP {res.status_code}"
                    continue
                data = res.json()
                result = (((data or {}).get("chart") or {}).get("result") or [None])[0]
                if not result:
                    last_error = "no chart result"
                    continue
                ts = result.get("timestamp") or []
                quote = (((result.get("indicators") or {}).get("quote") or [None])[0]) or {}
                adj = (((result.get("indicators") or {}).get("adjclose") or [None])[0]) or {}
                close = adj.get("adjclose") or quote.get("close") or []
                if not ts or not close:
                    last_error = "missing timestamp/close"
                    continue
                df = pd.DataFrame({"close": pd.to_numeric(close, errors="coerce")}, index=pd.to_datetime(ts, unit="s", utc=True).tz_convert(None))
                return df.dropna()
            except Exception as exc:
                last_error = str(exc)
        raise ValueError(last_error or "Yahoo chart failed")

    def _fetch_yfinance(self, ticker: str) -> pd.DataFrame:
        if yf is None:
            return pd.DataFrame()
        df = yf.download(ticker, period=self.lookback, interval="1d", auto_adjust=True, progress=False, threads=False)
        if df is None or df.empty:
            return pd.DataFrame()
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [c[0] for c in df.columns]
        close_col = "Close" if "Close" in df.columns else "Adj Close"
        out = pd.DataFrame({"close": pd.to_numeric(df[close_col], errors="coerce")})
        out.index = pd.to_datetime(out.index).tz_localize(None)
        return out.dropna()

    def _fetch_stooq(self, ticker: str) -> pd.DataFrame:
        symbols = [f"{ticker.lower()}.us", ticker.lower(), f"{ticker.upper()}.US"]
        last_error = None
        for symbol in symbols:
            url = f"https://stooq.com/q/d/l/?s={symbol}&i=d"
            try:
                res = self.session.get(url, timeout=20, verify=False)
                if res.status_code >= 400:
                    last_error = f"HTTP {res.status_code}"
                    continue
                text = res.text.strip()
                if not text or "Date" not in text or "Close" not in text:
                    last_error = text[:80] or "empty"
                    continue
                df = pd.read_csv(StringIO(text))
                if df.empty or "Close" not in df.columns:
                    last_error = "empty csv"
                    continue
                out = pd.DataFrame({"close": pd.to_numeric(df["Close"], errors="coerce")})
                out.index = pd.to_datetime(df["Date"]).tz_localize(None)
                return out.dropna()
            except Exception as exc:
                last_error = str(exc)
        raise ValueError(last_error or "Stooq failed")

def ema(series: pd.Series, span: int) -> pd.Series:
    return series.ewm(span=span, adjust=False).mean()

def sma(series: pd.Series, window: int) -> pd.Series:
    return series.rolling(window, min_periods=max(5, window // 3)).mean()

def rsi(series: pd.Series, window: int = 14) -> pd.Series:
    delta = series.diff()
    up = delta.clip(lower=0.0)
    down = -delta.clip(upper=0.0)
    avg_gain = up.ewm(alpha=1 / window, adjust=False).mean()
    avg_loss = down.ewm(alpha=1 / window, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    return (100 - (100 / (1 + rs))).fillna(50)

def macd_calc(series: pd.Series) -> Tuple[pd.Series, pd.Series, pd.Series]:
    line = ema(series, 12) - ema(series, 26)
    signal = ema(line, 9)
    return line, signal, line - signal

def max_drawdown(series: pd.Series, window: int = 63) -> float:
    s = series.tail(window)
    if s.empty:
        return 0.0
    return float((s / s.cummax() - 1.0).min())

def compute_technical(ticker: str, history: pd.DataFrame) -> TechnicalSnapshot:
    close = history["close"].astype(float).dropna()
    if len(close) < 50:
        raise ValueError(f"Not enough history for {ticker}: {len(close)} rows")
    ema20 = ema(close, 20)
    sma50 = sma(close, 50)
    sma200 = sma(close, 200)
    macd_line, signal, hist = macd_calc(close)
    rsi14 = rsi(close, 14)
    daily_ret = close.pct_change()
    price = float(close.iloc[-1])
    e20 = float(ema20.iloc[-1])
    s50 = float(sma50.iloc[-1])
    s200 = float(sma200.iloc[-1])
    m = float(macd_line.iloc[-1])
    sig = float(signal.iloc[-1])
    h = float(hist.iloc[-1])
    def ret(n: int) -> float:
        return float(close.iloc[-1] / close.iloc[-n - 1] - 1) if len(close) > n else 0.0
    ret5, ret20, ret60 = ret(5), ret(20), ret(60)
    vol20 = float(daily_ret.tail(20).std() * math.sqrt(252)) if len(daily_ret) >= 20 else 0.0
    dd63 = max_drawdown(close, 63)
    high252 = float(close.tail(252).max())
    dist_high = price / high252 - 1 if high252 else 0.0
    trend_points = [price >= e20, price >= s50, price >= s200, e20 >= s50, s50 >= s200]
    momentum_points = [m >= sig, m >= 0, ret20 > 0, ret60 > 0, float(rsi14.iloc[-1]) >= 50]
    trend_score = sum(trend_points) / len(trend_points)
    momentum_score = sum(momentum_points) / len(momentum_points)
    risk_score = max(0.0, min(1.0, 1.0 - (vol20 / 0.45) + (dd63 / 0.35)))
    composite = 0.45 * trend_score + 0.40 * momentum_score + 0.15 * risk_score
    read = "Bullish confirmed" if composite >= 0.78 else "Improving" if composite >= 0.58 else "Mixed" if composite >= 0.38 else "Weak"
    return TechnicalSnapshot(ticker, price, e20, s50, s200, m, sig, h, float(rsi14.iloc[-1]), ret5, ret20, ret60, vol20, dd63, dist_high, trend_score, momentum_score, risk_score, composite, "Above" if price >= e20 else "Below", "Above" if price >= s50 else "Below", "Above" if price >= s200 else "Below", "20 EMA > 50 SMA" if e20 >= s50 else "20 EMA < 50 SMA", "50 SMA > 200 SMA" if s50 >= s200 else "50 SMA < 200 SMA", "MACD > Signal" if m >= sig else "MACD < Signal", "Above 0" if m >= 0 else "Below 0", read, history.attrs.get("source", "market data"))

def blank_tech(symbol: str, reason: str) -> TechnicalSnapshot:
    return TechnicalSnapshot(symbol, 0, 0, 0, 0, 0, 0, 0, 50, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "n/a", "n/a", "n/a", "n/a", "n/a", "n/a", "n/a", "Needs data", reason)

def build_market_state(client: MarketDataClient) -> Dict[str, TechnicalSnapshot]:
    symbols = sorted(set([x[0] for x in SECTOR_ETFS] + [x[0] for x in MARKET_PROXIES] + ["SPY"]))
    out: Dict[str, TechnicalSnapshot] = {}
    for symbol in symbols:
        try:
            out[symbol] = compute_technical(symbol, client.get_history(symbol))
        except Exception as exc:
            client.status[symbol] = f"FAILED: {exc}"
            out[symbol] = blank_tech(symbol, str(exc))
    return out

def relative_strength_read(tech: TechnicalSnapshot, spy: TechnicalSnapshot) -> str:
    if tech.read == "Needs data" or spy.read == "Needs data":
        return "Needs data"
    diff = tech.ret20d - spy.ret20d
    if diff > 0.02:
        return "Outperforming SPY"
    if diff < -0.02:
        return "Underperforming SPY"
    return "In line with SPY"

def rotation_read(tech: TechnicalSnapshot, spy: TechnicalSnapshot) -> str:
    if tech.read == "Needs data":
        return "Needs data"
    if tech.composite_score >= 0.75 and tech.ret20d > spy.ret20d:
        return "Leadership"
    if tech.composite_score >= 0.60:
        return "Positive trend"
    if tech.composite_score >= 0.45 and tech.ret20d >= spy.ret20d:
        return "Improving"
    if tech.composite_score < 0.35:
        return "Lagging / weak"
    return "Mixed"

def flow_pressure(tech: TechnicalSnapshot, spy: TechnicalSnapshot) -> str:
    if tech.read == "Needs data" or spy.read == "Needs data":
        return "Needs data"
    if tech.ret20d > spy.ret20d + 0.02 and tech.composite_score >= 0.55:
        return "Positive pressure"
    if tech.ret20d < spy.ret20d - 0.02 and tech.composite_score <= 0.50:
        return "Negative pressure"
    return "Mixed pressure"

def action_bias(rot: str, tech: TechnicalSnapshot) -> str:
    if rot == "Leadership":
        return "Best add/hold candidates after macro confirmation"
    if rot == "Positive trend":
        return "Hold / selective add only"
    if rot == "Improving":
        return "Watchlist / starter only"
    if rot == "Lagging / weak":
        return "Avoid adds / review trims"
    return "Data needed / wait"

def meta_for_holding(h: Holding) -> Dict[str, str]:
    t = h.ticker.upper()
    if t in TICKER_MAP:
        return TICKER_MAP[t].copy()
    n = h.name.lower()
    if "fixed" in h.security_type or "fdic" in n or " cd " in f" {n} ":
        return {"sleeve": "Fixed Income Safety", "sector_proxy": "BIL", "thesis": "Principal/income stabilizer; watch safety overweight."}
    if "etf" in h.security_type:
        return {"sleeve": "ETF / Unmapped", "sector_proxy": "SPY", "thesis": "ETF exposure; map sleeve manually if material."}
    return {"sleeve": "Equity / Unmapped", "sector_proxy": "SPY", "thesis": "Single-stock exposure; validate thesis manually."}

def sleeve_weights(holdings: Sequence[Holding]) -> Dict[str, float]:
    weights: Dict[str, float] = {}
    for h in holdings:
        sleeve = meta_for_holding(h)["sleeve"]
        weights[sleeve] = weights.get(sleeve, 0.0) + h.portfolio_weight
    return weights

def tolerance_status(h: Holding, meta: Mapping[str, str], tech: TechnicalSnapshot, weights: Mapping[str, float]) -> str:
    sleeve = meta["sleeve"]
    sleeve_w = weights.get(sleeve, 0.0)
    safety_like = any(k in sleeve for k in ["Safety", "Income", "Credit"])
    if safety_like and sleeve_w > 0.45:
        return "Overweight safety"
    if safety_like and sleeve_w > 0.18:
        return "Safety overlap"
    if safety_like:
        return "Income tilt"
    if h.pnl_pct > 0.50:
        return "Profit outlier"
    if h.pnl_pct < -0.10 and tech.composite_score < 0.45:
        return "Weak"
    if h.portfolio_weight < 0.015:
        return "Too small"
    if h.portfolio_weight > 0.08 or sleeve_w > 0.18:
        return "Near limit"
    if tech.composite_score < 0.35 and tech.read != "Needs data":
        return "Out of tolerance"
    return "In tolerance"

def exact_action(h: Holding, meta: Mapping[str, str], tech: TechnicalSnapshot, tolerance: str) -> Tuple[str, int, float, str]:
    price = h.price or (h.value / h.quantity if h.quantity else 0.0)
    qty = max(0, int(math.floor(h.quantity)))
    safety_like = any(k in meta["sleeve"] for k in ["Safety", "Income", "Credit", "Fixed Income"])
    src = "E*TRADE institution price/value"
    if qty <= 0:
        return "Hold", 0, 0.0, f"No share quantity available. Price source: {src}."
    if tech.read == "Needs data":
        return "Hold", 0, 0.0, f"No market technical data yet. Price source: {src}."
    if tolerance == "Profit outlier" and tech.composite_score < 0.65:
        q = max(1, int(math.floor(qty * 0.10)))
        return f"Trim-QTY {q}", q, q * price, f"Profit outlier and mapped trend is not leadership. Estimate uses {src}."
    if tolerance in {"Weak", "Out of tolerance"} and tech.composite_score < 0.40 and not safety_like:
        q = max(1, int(math.floor(qty * 0.15)))
        return f"Sell-QTY {q}", q, q * price, f"Weak mapped trend plus tolerance pressure. Estimate uses {src}."
    if tolerance in {"Overweight safety", "Safety overlap"} and safety_like:
        q = max(1, int(math.floor(qty * 0.05)))
        return f"Trim-QTY {q}", q, q * price, f"Safety sleeve overlap; trim only if reallocating to confirmed leadership. Estimate uses {src}."
    if tolerance == "Too small" and tech.composite_score >= 0.75 and not safety_like:
        q = 1 if price >= 100 else max(1, int(math.floor(500 / max(price, 1))))
        return f"Add-QTY {q}", q, q * price, f"Small position with mapped technical leadership. Estimate uses {src}."
    return "Hold", 0, 0.0, f"Hold-no-add until macro/news confirms. Mapped trend: {tech.read}. Price source: {src}."

def build_data_quality_rows(client: MarketDataClient, market: Mapping[str, TechnicalSnapshot]) -> List[List[str]]:
    rows = []
    for symbol in sorted(market.keys()):
        t = market[symbol]
        rows.append([symbol, "OK" if t.read != "Needs data" else "Failed", t.source if t.read != "Needs data" else client.status.get(symbol, t.source), f"{t.price:.2f}" if t.price else "n/a"])
    return rows

def build_macro_rows(market: Mapping[str, TechnicalSnapshot]) -> List[List[str]]:
    spy = market["SPY"]
    uso = market.get("USO", spy)
    tlt = market.get("TLT", spy)
    uup = market.get("UUP", spy)
    gld = market.get("GLD", spy)
    hyg = market.get("HYG", spy)
    return [
        ["War / Geopolitics", f"GLD 20d {pct(gld.ret20d)}; USO 20d {pct(uso.ret20d)}", "Elevated watch" if gld.ret20d > 0.03 or uso.ret20d > 0.05 else "Normal watch", "Gold/defense matter more if risk shock rises."],
        ["Oil / Energy shock", f"USO 20d {pct(uso.ret20d)}", "Oil pressure rising" if uso.ret20d > 0.05 else "Oil easing" if uso.ret20d < -0.05 else "Neutral", "Affects XLE, inflation expectations, and consumer pressure."],
        ["Crisis / Liquidity", f"SPY {spy.price_vs_50sma} 50 SMA; HYG 20d {pct(hyg.ret20d)}", "Risk-off watch" if spy.price_vs_50sma == "Below" or hyg.ret20d < -0.03 else "Risk-on acceptable", "Controls how aggressive adds should be."],
        ["Fed / Yields", f"TLT 20d {pct(tlt.ret20d)}", "Yield pressure rising" if tlt.ret20d < -0.04 else "Yield pressure easing" if tlt.ret20d > 0.04 else "Neutral", "Affects growth stocks, banks, real estate, and bond sleeves."],
        ["Inflation / Dollar", f"UUP 20d {pct(uup.ret20d)}; USO 20d {pct(uso.ret20d)}", "Tighter impulse" if uup.ret20d > 0.03 or uso.ret20d > 0.05 else "Contained", "Affects ADRs, gold, and valuation multiples."],
        ["Jobs / Growth", f"SPY technical read: {spy.read}", "Growth trend constructive" if spy.composite_score >= 0.58 else "Growth trend mixed/weak", "Confirms whether rotation supports risk assets."],
    ]

def build_sector_rows(market: Mapping[str, TechnicalSnapshot]) -> List[List[str]]:
    spy = market["SPY"]
    return [[ticker, sector, rotation_read(market[ticker], spy), f"{flow_pressure(market[ticker], spy)} / {relative_strength_read(market[ticker], spy)}", action_bias(rotation_read(market[ticker], spy), market[ticker])] for ticker, sector in SECTOR_ETFS]

def build_technical_rows(market: Mapping[str, TechnicalSnapshot]) -> List[List[str]]:
    spy = market["SPY"]
    rows = []
    for ticker, sector in SECTOR_ETFS:
        t = market[ticker]
        rows.append([ticker, sector, t.price_vs_20ema, t.price_vs_50sma, t.price_vs_200sma, t.cross_20_50, t.trend_50_200, t.macd_vs_signal, t.macd_zero, relative_strength_read(t, spy)])
    return rows

def build_portfolio_rows(holdings: Sequence[Holding], market: Mapping[str, TechnicalSnapshot]) -> List[List[str]]:
    weights = sleeve_weights(holdings)
    rows = []
    for h in holdings:
        meta = meta_for_holding(h)
        tech = market.get(meta["sector_proxy"], market["SPY"])
        tol = tolerance_status(h, meta, tech, weights)
        action, action_qty, est, reason = exact_action(h, meta, tech, tol)
        rows.append([h.ticker, f"{h.quantity:g}", money(h.cost_basis), money(h.value), money(h.unrealized), pct(h.pnl_pct), tol, meta["thesis"], action, str(action_qty), money(est), f"{reason} Sector proxy {meta['sector_proxy']} = {tech.read} ({tech.composite_score:.2f}).", meta["sleeve"]])
    priority = {"Add-QTY": 1, "Trim-QTY": 2, "Sell-QTY": 3, "Avoid": 4, "Hold": 9}
    def key(row: List[str]):
        prefix = "Hold"
        for p in ["Add-QTY", "Trim-QTY", "Sell-QTY", "Avoid"]:
            if str(row[8]).startswith(p):
                prefix = p
        return (priority.get(prefix, 8), -num(row[3]))
    return sorted(rows, key=key)

def prior_week_text(market: Mapping[str, TechnicalSnapshot]) -> str:
    spy = market["SPY"]
    leaders, laggards = [], []
    for ticker, sector in SECTOR_ETFS:
        tech = market[ticker]
        rot = rotation_read(tech, spy)
        if rot == "Leadership":
            leaders.append(f"{ticker} {sector}")
        elif rot == "Lagging / weak":
            laggards.append(f"{ticker} {sector}")
    return f"Forward trend: SPY is {spy.read} with 20-day change of {pct(spy.ret20d)}. Confirmed technical leadership: {', '.join(leaders) if leaders else 'none'}. Weak/lagging sectors: {', '.join(laggards[:4]) if laggards else 'none'}. Use live macro/news and earnings checks before trading."

def total_by_action(rows: Sequence[Sequence[str]], prefix: str) -> float:
    return sum(num(row[10]) for row in rows if str(row[8]).startswith(prefix))

def build_report(holdings_rows: Sequence[Sequence[Any]], headers: Optional[Sequence[str]] = None) -> Dict[str, Any]:
    holdings = parse_holdings_from_rows(holdings_rows, headers=headers)
    if not holdings:
        raise ValueError("No usable holdings after exclusions.")
    client = MarketDataClient(lookback=os.getenv("MARKET_LOOKBACK", "2y"))
    market = build_market_state(client)
    portfolio_rows = build_portfolio_rows(holdings, market)
    return {
        "title": "Raj Market Rotation Report",
        "as_of": datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M:%S %Z"),
        "sections": [
            {"type": "paragraph", "title": "", "text": "Uses connected E*TRADE holdings excluding cash and money market. Technical engine uses daily market data, EMA/SMA/MACD/RS, and portfolio tolerance controls."},
            {"type": "table", "title": "Data Quality Check", "headers": ["Ticker", "Status", "Source / Error", "Last Price"], "rows": build_data_quality_rows(client, market)},
            {"type": "table", "title": "Macro Risk Dashboard", "headers": ["Risk Area", "Current Read", "Risk Level", "Portfolio Meaning"], "rows": build_macro_rows(market)},
            {"type": "paragraph", "title": "Macro Summary", "text": "Macro reads use market proxies for oil, yields/bonds, dollar, gold, high-yield credit, and SPY trend. Headlines, Fed calendar, CPI/jobs, and earnings still need live review."},
            {"type": "table", "title": "Sector Rotation - SPDR Map", "headers": ["ETF", "Sector", "Rotation Read", "Flow Pressure", "Action Bias"], "rows": build_sector_rows(market)},
            {"type": "paragraph", "title": "Sector Summary", "text": "Leadership requires both trend confirmation and relative strength versus SPY. Mixed or lagging sectors are not add candidates without a separate macro catalyst."},
            {"type": "table", "title": "Technical Confirmation Snapshot", "headers": ["ETF", "Sector", "Price vs 20 EMA", "Price vs 50 SMA", "Price vs 200 SMA", "20/50 Crossover", "50/200 Trend", "MACD vs Signal", "MACD Zero", "RS vs SPY"], "rows": build_technical_rows(market)},
            {"type": "paragraph", "title": "Prior Week Recap + Forward Trend", "text": prior_week_text(market)},
            {"type": "table", "title": "Raj Portfolio Impact - Exact Actions", "headers": ["Ticker", "Qty Held", "Cost Basis", "Current Value", "Unrealized $", "P&L %", "Tolerance Status", "Thesis", "Exact Action", "Action Qty", "Est. $ Value", "Reason / Price Source", "Sleeve"], "rows": portfolio_rows},
            {"type": "table", "title": "Total Suggested Trims and Primary Goal", "headers": ["Metric", "Value"], "rows": [["Total suggested trims", money(total_by_action(portfolio_rows, "Trim-QTY"))], ["Total suggested sells", money(total_by_action(portfolio_rows, "Sell-QTY"))], ["Total suggested adds", money(total_by_action(portfolio_rows, "Add-QTY"))], ["Primary goal", "Keep actions small, reduce overlap, protect profit outliers, and add only when macro/sector thesis confirms technical trend."]]},
            {"type": "table", "title": "Aggressive Growth Setup With Risk Controls", "headers": ["Setup", "Trigger", "Risk Control", "Action Bias"], "rows": [["Core growth add", "SPY/XLK above 20 EMA and 50 SMA with MACD > signal", "Starter size only; do not add into weak macro", "Add-QTY only when confirmed"], ["Profit protection", "Large P&L outlier or overweight sleeve", "Trim 5% to 15%, not full exit", "Trim-QTY"], ["Safety redeployment", "Safety sleeve overweight and growth leadership confirmed", "Keep liquidity buffer", "Gradual shift only"], ["Avoid weak trend", "Below 20 EMA/50 SMA with MACD < signal", "No averaging down without macro confirmation", "Avoid or Hold"]]},
            {"type": "table", "title": "Key Catalysts to Watch", "headers": ["Catalyst", "Why It Matters"], "rows": [["Fed/FOMC and Treasury yields", "Affects growth multiples, banks, real estate, and fixed income."], ["Oil and geopolitical headlines", "Affects energy, inflation pressure, defense, and gold."], ["Earnings guidance", "Can override sector trend."], ["Credit spreads / liquidity", "Important for CLOZ, banks, and risk appetite."], ["SPY breadth and sector relative strength", "Confirms whether rotation is broadening or narrowing."]]},
            {"type": "table", "title": "Chicago-Style Source List", "headers": ["Source", "Use"], "rows": [["Yahoo chart API", "Primary daily price source."], ["Yahoo Finance via yfinance", "Fallback daily price source."], ["Stooq CSV", "Second fallback daily price source."], ["Connected E*TRADE Holdings tab", "Portfolio quantities, prices, values, weights, and P&L."], ["SPDR sector ETF map", "11-sector rotation framework."], ["SPY benchmark", "Relative-strength benchmark."]]},
        ],
    }

def report_from_tsv(tsv: str) -> Dict[str, Any]:
    rows = [line.split("\t") for line in tsv.splitlines() if line.strip()]
    return build_report(rows)

if __name__ == "__main__":
    import argparse, json
    parser = argparse.ArgumentParser()
    parser.add_argument("--holdings-tsv", required=True, help="Path to Holdings tab pasted/exported as TSV")
    parser.add_argument("--out", default="market_report.json")
    args = parser.parse_args()
    report = report_from_tsv(open(args.holdings_tsv, "r", encoding="utf-8").read())
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)
    print(f"Wrote {args.out}")
