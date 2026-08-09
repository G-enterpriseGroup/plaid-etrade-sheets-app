from __future__ import annotations

import argparse
import json
import math
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import yfinance as yf

ET = ZoneInfo("America/New_York")
OUT = Path("/tmp/pmf_result.json")


def now_et() -> str:
    return datetime.now(ET).replace(microsecond=0).isoformat()


def log(msg: str) -> None:
    print(f"[{now_et()}] {msg}", flush=True)


def fnum(value):
    try:
        n = float(value)
        return n if math.isfinite(n) else None
    except Exception:
        return None


def retry(label, fn, attempts: int = 3):
    last = None
    for i in range(attempts):
        try:
            return fn()
        except Exception as exc:
            last = exc
            log(f"{label} failed attempt {i + 1}/{attempts}: {exc}")
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"{label} failed: {last}")


class Yahoo:
    def __init__(self):
        self._tickers = {}
        self._chains = {}

    def ticker(self, symbol: str):
        symbol = symbol.upper().strip()
        if symbol not in self._tickers:
            self._tickers[symbol] = yf.Ticker(symbol)
        return self._tickers[symbol]

    def expirations(self, symbol: str):
        return list(retry(f"{symbol} expirations", lambda: self.ticker(symbol).options or ()))

    def chain(self, symbol: str, expiration: str):
        key = (symbol.upper().strip(), expiration)
        if key not in self._chains:
            self._chains[key] = retry(
                f"{key[0]} {expiration} option chain",
                lambda: self.ticker(key[0]).option_chain(expiration),
            )
        return self._chains[key]


def normalized_type(value: str) -> str:
    value = str(value or "").upper()
    return "CALL" if "CALL" in value else "PUT"


def frame_for(chain, option_type: str):
    return chain.calls if normalized_type(option_type) == "CALL" else chain.puts


def find_contract(df: pd.DataFrame, strike: float):
    if df is None or df.empty or "strike" not in df.columns:
        return None
    strikes = pd.to_numeric(df["strike"], errors="coerce")
    hit = df[(strikes - float(strike)).abs() < 1e-6]
    if hit.empty:
        return None
    return hit.iloc[0]


def quote_from_row(row) -> dict:
    bid = fnum(row.get("bid"))
    ask = fnum(row.get("ask"))
    last = fnum(row.get("lastPrice"))

    mark = None
    if bid is not None and ask is not None and bid > 0 and ask > 0 and ask >= bid:
        mark = (bid + ask) / 2
    elif last is not None and last > 0:
        mark = last
    elif bid is not None and bid > 0:
        mark = bid
    elif ask is not None and ask > 0:
        mark = ask

    return {
        "bid": bid,
        "ask": ask,
        "last": last,
        "mark": mark,
        "contract_symbol": str(row.get("contractSymbol") or ""),
        "open_interest": fnum(row.get("openInterest")),
        "volume": fnum(row.get("volume")),
        "implied_volatility": fnum(row.get("impliedVolatility")),
    }


def make_result(request_id: str, mode: str, **payload) -> dict:
    out = {
        "request_id": request_id,
        "mode": mode,
        "source": "Yahoo Finance via yfinance",
        "updated_at": now_et(),
    }
    out.update(payload)
    return out


def run_expirations(y: Yahoo, args) -> dict:
    ticker = args.ticker.upper().strip()
    expirations = y.expirations(ticker)
    return make_result(
        args.request_id,
        "expirations",
        row=args.row,
        ticker=ticker,
        option_type=normalized_type(args.option_type),
        expirations=expirations,
    )


def run_strikes(y: Yahoo, args) -> dict:
    ticker = args.ticker.upper().strip()
    chain = y.chain(ticker, args.expiration)
    df = frame_for(chain, args.option_type)
    strikes = []
    if df is not None and not df.empty and "strike" in df.columns:
        strikes = sorted(
            {
                float(x)
                for x in pd.to_numeric(df["strike"], errors="coerce").dropna().tolist()
                if math.isfinite(float(x))
            }
        )
    return make_result(
        args.request_id,
        "strikes",
        row=args.row,
        ticker=ticker,
        option_type=normalized_type(args.option_type),
        expiration=args.expiration,
        strikes=strikes,
    )


def quote_contract(y: Yahoo, contract: dict) -> dict:
    ticker = str(contract.get("ticker") or "").upper().strip()
    expiration = str(contract.get("expiration") or "").strip()
    option_type = normalized_type(contract.get("type"))
    strike = fnum(contract.get("strike"))
    key = str(contract.get("key") or "").strip()
    row_number = int(contract.get("row") or 0)

    result = {
        "row": row_number,
        "key": key,
        "ticker": ticker,
        "expiration": expiration,
        "type": option_type,
        "strike": strike,
        "status": "YFINANCE ERROR",
        "bid": None,
        "ask": None,
        "last": None,
        "mark": None,
        "contract_symbol": "",
        "open_interest": None,
        "volume": None,
        "implied_volatility": None,
    }

    if not ticker or not expiration or strike is None:
        result["status"] = "INCOMPLETE CONTRACT"
        return result

    try:
        chain = y.chain(ticker, expiration)
        row = find_contract(frame_for(chain, option_type), strike)
        if row is None:
            result["status"] = "NO YFINANCE CONTRACT"
            return result
        result.update(quote_from_row(row))
        result["status"] = "YFINANCE" if result["mark"] is not None else "NO YFINANCE QUOTE"
        return result
    except Exception as exc:
        result["error"] = str(exc)
        return result


def run_quotes(y: Yahoo, args) -> dict:
    contracts = json.loads(args.contracts_json or "[]")
    if not isinstance(contracts, list):
        raise ValueError("contracts_json must be a JSON array")
    quotes = [quote_contract(y, c) for c in contracts]
    return make_result(args.request_id, "quotes", quotes=quotes)


def parse_args():
    p = argparse.ArgumentParser(description="Put Money Flow Yahoo/yfinance worker for GitHub Actions.")
    p.add_argument("--request-id", required=True)
    p.add_argument("--mode", required=True, choices=["expirations", "strikes", "quotes"])
    p.add_argument("--row", default="")
    p.add_argument("--ticker", default="")
    p.add_argument("--option-type", default="")
    p.add_argument("--expiration", default="")
    p.add_argument("--contracts-json", default="[]")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    yahoo = Yahoo()

    if args.mode == "expirations":
        result = run_expirations(yahoo, args)
    elif args.mode == "strikes":
        result = run_strikes(yahoo, args)
    else:
        result = run_quotes(yahoo, args)

    OUT.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    log(f"wrote {OUT} for {args.mode} request {args.request_id}")


if __name__ == "__main__":
    main()
