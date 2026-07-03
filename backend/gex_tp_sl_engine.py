"""GEX Take Profit / Stop Limit engine.

Reads a holdings-style CSV/JSON list of tickers, pulls delayed Cboe option chains,
estimates dealer gamma exposure, and returns Call Wall, Put Wall, Gamma Flip,
Take Profit, Stop Limit, cost basis, and TP/SL percent guide levels.
"""

from __future__ import annotations

import csv
import json
import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

CBOE_BASE = "https://cdn.cboe.com/api/global/delayed_quotes/options/"
MAX_DTE = 45
MIN_OI = 1
EXCLUDE = {"VMFXX", "SWVXX", "SPAXX", "FDRXX"}


@dataclass
class Holding:
    ticker: str
    qty: float = 0.0
    price: float = 0.0
    cost_basis: float = 0.0


def norm_pdf(x: float) -> float:
    return math.exp(-0.5 * x * x) / math.sqrt(2 * math.pi)


def bs_gamma(spot: float, strike: float, dte: int, iv: float = 0.35) -> float:
    if not spot or not strike or not dte:
        return 0.0
    t = max(dte / 365.0, 1 / 365.0)
    v = max(float(iv or 0.35), 0.05)
    d1 = (math.log(spot / strike) + 0.5 * v * v * t) / (v * math.sqrt(t))
    return norm_pdf(d1) / (spot * v * math.sqrt(t))


def dte(expiration: str) -> int:
    try:
        d = datetime.fromisoformat(str(expiration)[:10]).date()
        return (d - date.today()).days
    except Exception:
        return 9999


def fetch_cboe_chain(ticker: str) -> Dict[str, Any]:
    symbol = re.sub(r"[^A-Z0-9.]", "", ticker.upper())
    last_error = ""
    for sym in [symbol, symbol.replace(".", "-")]:
        url = f"{CBOE_BASE}{sym}.json"
        try:
            r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
            if r.status_code != 200:
                last_error = f"Cboe HTTP {r.status_code}"
                continue
            data = r.json()
            options = data.get("options") or data.get("data", {}).get("options") or []
            spot = data.get("current_price") or data.get("currentPrice") or data.get("data", {}).get("current_price") or 0
            if not options:
                last_error = "no options in Cboe response"
                continue
            return {"spot": float(spot or 0), "options": options, "source": "Cboe delayed quotes"}
        except Exception as exc:
            last_error = str(exc)
    raise RuntimeError(last_error or "Cboe fetch failed")


def parse_option(row: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    option = str(row.get("option") or row.get("option_symbol") or row.get("symbol") or "").upper()
    opt_type = str(row.get("option_type") or row.get("type") or "").upper()[:1]
    strike = float(row.get("strike") or row.get("strike_price") or 0)
    exp = row.get("expiration_date") or row.get("expiration") or row.get("expiry") or ""

    if (not opt_type or not strike or not exp) and option:
        m = re.search(r"(\d{6})([CP])(\d{8})$", option)
        if m:
            opt_type = m.group(2)
            strike = int(m.group(3)) / 1000.0
            exp = f"20{m.group(1)[:2]}-{m.group(1)[2:4]}-{m.group(1)[4:6]}"

    if opt_type not in {"C", "P"} or not strike:
        return None
    iv = float(row.get("iv") or row.get("implied_volatility") or row.get("greeks", {}).get("iv") or 0)
    if iv > 3:
        iv /= 100.0
    return {
        "type": opt_type,
        "strike": strike,
        "dte": dte(str(exp)),
        "oi": float(row.get("open_interest") or row.get("openInterest") or row.get("oi") or 0),
        "gamma": float(row.get("gamma") or row.get("greeks", {}).get("gamma") or 0),
        "iv": iv,
    }


def max_oi_strike(oi_by_strike: Dict[float, float]) -> Optional[float]:
    return max(oi_by_strike.items(), key=lambda kv: kv[1])[0] if oi_by_strike else None


def gamma_flip(gex_by_strike: Dict[float, float]) -> Optional[float]:
    strikes = sorted(gex_by_strike)
    if not strikes:
        return None
    cum = 0.0
    prev = None
    best = strikes[0]
    best_abs = float("inf")
    for k in strikes:
        cum += gex_by_strike[k]
        if abs(cum) < best_abs:
            best_abs = abs(cum)
            best = k
        if prev is not None and ((prev <= 0 <= cum) or (prev >= 0 >= cum)):
            return k
        prev = cum
    return best


def analyze_ticker(holding: Holding) -> Dict[str, Any]:
    chain = fetch_cboe_chain(holding.ticker)
    spot = float(chain.get("spot") or holding.price or 0)
    if not spot:
        raise RuntimeError("missing spot")

    gex_by_strike: Dict[float, float] = {}
    call_oi: Dict[float, float] = {}
    put_oi: Dict[float, float] = {}

    for raw in chain["options"]:
        opt = parse_option(raw)
        if not opt or opt["dte"] < 0 or opt["dte"] > MAX_DTE or opt["oi"] < MIN_OI:
            continue
        gamma = opt["gamma"] or bs_gamma(spot, opt["strike"], opt["dte"], opt["iv"] or 0.35)
        gex = gamma * opt["oi"] * 100 * spot * spot * 0.01
        if opt["type"] == "P":
            gex = -abs(gex)
            put_oi[opt["strike"]] = put_oi.get(opt["strike"], 0) + opt["oi"]
        else:
            gex = abs(gex)
            call_oi[opt["strike"]] = call_oi.get(opt["strike"], 0) + opt["oi"]
        gex_by_strike[opt["strike"]] = gex_by_strike.get(opt["strike"], 0) + gex

    if not gex_by_strike:
        raise RuntimeError(f"no option rows inside {MAX_DTE} DTE")

    call_wall = max_oi_strike(call_oi)
    put_wall = max_oi_strike(put_oi)
    flip = gamma_flip(gex_by_strike)
    net_gex = sum(gex_by_strike.values())
    regime = "Positive GEX" if net_gex >= 0 else "Negative GEX"
    tp = call_wall if call_wall and call_wall > spot else next((x for x in sorted([call_wall, flip]) if x and x > spot), None)
    stops = [x for x in [put_wall, flip] if x and x < spot]
    sl = max(stops) if stops else put_wall or flip
    return {
        "ticker": holding.ticker,
        "qty": holding.qty,
        "cost_basis": holding.cost_basis,
        "spot": spot,
        "call_wall": call_wall,
        "put_wall": put_wall,
        "gamma_flip": flip,
        "take_profit": tp,
        "take_profit_pct_from_spot": (tp / spot - 1) if tp and spot else None,
        "stop_limit": sl,
        "stop_limit_pct_from_spot": (sl / spot - 1) if sl and spot else None,
        "dealer_regime": regime,
        "net_gex": net_gex,
        "source": chain["source"],
    }


def to_float(value: Any) -> float:
    if value is None:
        return 0.0
    text = str(value)
    neg = "(" in text and ")" in text
    text = re.sub(r"[$,%\s()]", "", text)
    try:
        n = float(text)
        return -n if neg else n
    except ValueError:
        return 0.0


def read_holdings(path: Path) -> List[Holding]:
    if path.suffix.lower() == ".json":
        data = json.loads(path.read_text())
        rows = data.get("holdings", data if isinstance(data, list) else [])
    else:
        rows = list(csv.DictReader(path.open(newline="")))
    out: List[Holding] = []
    seen = set()
    for r in rows:
        ticker = str(r.get("ticker_symbol") or r.get("ticker") or r.get("symbol") or "").upper().strip()
        if not ticker or ticker in seen or ticker in EXCLUDE or ticker.startswith("CUR:"):
            continue
        seen.add(ticker)
        out.append(
            Holding(
                ticker=ticker,
                qty=to_float(r.get("quantity") or r.get("qty") or 0),
                price=to_float(r.get("institution_price") or r.get("price") or 0),
                cost_basis=to_float(r.get("cost_basis") or r.get("cost") or 0),
            )
        )
    return out


def run(input_path: str) -> List[Dict[str, Any]]:
    return [analyze_ticker(h) for h in read_holdings(Path(input_path))]


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("input_path")
    parser.add_argument("--output", default="gex_tp_sl_report.json")
    args = parser.parse_args()
    result = run(args.input_path)
    Path(args.output).write_text(json.dumps(result, indent=2))
    print(f"Wrote {len(result)} rows to {args.output}")
