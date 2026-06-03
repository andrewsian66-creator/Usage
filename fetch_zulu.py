#!/usr/bin/env python3
"""
fetch_zulu.py  —  Zulu Pre-Trade Checklist Auto-Filler
Usage:  python fetch_zulu.py TICKER
Example: python fetch_zulu.py BARC.L
"""

import sys
import os
from datetime import date
import warnings

import pandas as pd
import yfinance as yf
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore")

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "zulu_pretrade_checklist.html")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _fmt_cap(mc):
    if not mc:
        return "N/A"
    if mc >= 1_000_000_000:
        return f"£{mc / 1_000_000_000:.2f}B"
    if mc >= 1_000_000:
        return f"£{mc / 1_000_000:.1f}M"
    return f"£{mc:,.0f}"


def _sym(currency):
    return {"GBp": "p", "GBX": "p", "GBP": "£", "USD": "$", "EUR": "€"}.get(currency, currency + " ")


def _pct(value):
    if value is None:
        return "N/A"
    return f"{value * 100:+.1f}%"


# ── Data Fetching ─────────────────────────────────────────────────────────────

def fetch_stock_data(ticker: str) -> dict:
    print("  → Connecting to Yahoo Finance …")
    stock = yf.Ticker(ticker)
    info  = stock.info

    print("  → Fetching 1-year historical price / volume data …")
    hist  = stock.history(period="1y")

    if hist.empty:
        raise ValueError(f"No data returned for '{ticker}'. Check the ticker symbol and try again.")

    data = {}
    currency = info.get("currency", "GBp" if ticker.upper().endswith(".L") else "USD")
    sym      = _sym(currency)

    # ── Candidate Details ────────────────────────────────────────────────────
    data["ticker"]       = ticker.upper()
    data["company-name"] = info.get("longName") or info.get("shortName", ticker)
    data["date"]         = date.today().strftime("%d/%m/%Y")
    data["currency"]     = currency

    raw_price = (info.get("currentPrice")
                 or info.get("regularMarketPrice")
                 or info.get("previousClose")
                 or float(hist["Close"].iloc[-1]))
    data["current-price"] = f"{sym}{raw_price:.2f}"

    data["market-cap"] = _fmt_cap(info.get("marketCap"))
    data["sector"]     = info.get("sector",   "N/A")
    data["industry"]   = info.get("industry", "N/A")

    # ── Market Check ─────────────────────────────────────────────────────────
    exchange      = info.get("exchange",          "").upper()
    full_exchange = info.get("fullExchangeName",  "").upper()

    if "AIM" in exchange or "AIM" in full_exchange:
        data["market"] = "aim"
    elif ticker.upper().endswith(".L") or exchange in ("LSE", "LON", "XLON"):
        data["market"] = "main-market"
    else:
        data["market"] = "other"

    hi52 = info.get("fiftyTwoWeekHigh") or float(hist["High"].max())
    lo52 = info.get("fiftyTwoWeekLow")  or float(hist["Low"].min())
    data["week-52-high"] = f"{sym}{hi52:.2f}"
    data["week-52-low"]  = f"{sym}{lo52:.2f}"
    if hi52 and hi52 > 0:
        data["price-vs-high-pct"] = f"{(raw_price - hi52) / hi52 * 100:+.1f}%"
    else:
        data["price-vs-high-pct"] = "N/A"

    # ── Chart Leg — Moving Averages ───────────────────────────────────────────
    print("  → Calculating 50-day and 200-day SMAs …")
    closes = hist["Close"]
    sma50 = sma200 = None

    if len(closes) >= 50:
        hist = hist.copy()
        hist["SMA50"] = closes.rolling(50).mean()
        sma50 = float(hist["SMA50"].iloc[-1])
        data["sma-50"]        = f"{sym}{sma50:.2f}"
        data["price-vs-50ma"] = "above" if raw_price > sma50 else "below"
    else:
        data["sma-50"]        = "N/A (< 50 days data)"
        data["price-vs-50ma"] = ""

    if len(closes) >= 200:
        hist["SMA200"] = closes.rolling(200).mean()
        sma200 = float(hist["SMA200"].iloc[-1])
        data["sma-200"]        = f"{sym}{sma200:.2f}"
        data["price-vs-200ma"] = "above" if raw_price > sma200 else "below"
    else:
        data["sma-200"]        = "N/A (< 200 days data)"
        data["price-vs-200ma"] = ""

    if sma50 is not None and sma200 is not None:
        data["golden-cross"] = "yes" if sma50 > sma200 else "no"

        # 10-day trend in the gap between the two MAs
        if len(hist) >= 211:
            gap_now = float(hist["SMA50"].iloc[-1])  - float(hist["SMA200"].iloc[-1])
            gap_10d = float(hist["SMA50"].iloc[-11]) - float(hist["SMA200"].iloc[-11])
            if sma50 > sma200:
                data["ma-trend"] = "widening" if gap_now > gap_10d * 1.015 else (
                                   "approaching" if gap_now < gap_10d * 0.985 else "widening")
            else:
                data["ma-trend"] = "approaching" if abs(gap_now) < abs(gap_10d) * 0.985 else "declining"
        else:
            data["ma-trend"] = ""

        # Crossover age — scan backwards for the most recent SMA50/SMA200 cross
        valid = hist[["SMA50", "SMA200"]].dropna()
        diff  = valid["SMA50"] - valid["SMA200"]
        crossover_ts = None
        for i in range(len(diff) - 1, 0, -1):
            if (diff.iloc[i - 1] <= 0 < diff.iloc[i]) or (diff.iloc[i - 1] >= 0 > diff.iloc[i]):
                crossover_ts = diff.index[i]
                break

        if crossover_ts is not None:
            now      = pd.Timestamp.now(tz=crossover_ts.tzinfo)
            age_days = (now - crossover_ts).days
            kind     = "golden" if sma50 > sma200 else "death"
            data["cross-age"] = f"{age_days} days ({kind} cross)"
        else:
            data["cross-age"] = "> 1 year ago (no cross in available data)"
    else:
        data["golden-cross"] = ""
        data["ma-trend"]     = ""
        data["cross-age"]    = "N/A"

    # ── Volume Profile ────────────────────────────────────────────────────────
    print("  → Analysing 30-day volume profile …")
    recent = hist.tail(30).copy()
    recent["green"] = recent["Close"] >= recent["Open"]

    avg_green = recent.loc[recent["green"],  "Volume"].mean()
    avg_red   = recent.loc[~recent["green"], "Volume"].mean()

    data["avg-green-volume"] = f"{avg_green:,.0f}" if pd.notna(avg_green) and avg_green > 0 else "N/A"
    data["avg-red-volume"]   = f"{avg_red:,.0f}"   if pd.notna(avg_red)   and avg_red   > 0 else "N/A"

    if pd.notna(avg_green) and pd.notna(avg_red) and avg_red > 0:
        ratio = avg_green / avg_red
        data["volume-signal"] = "accumulation" if ratio > 1.10 else ("distribution" if ratio < 0.91 else "neutral")
    else:
        data["volume-signal"] = ""

    # ── Quants Leg ────────────────────────────────────────────────────────────
    print("  → Pulling fundamental data …")

    def _get(key):
        return info.get(key)

    peg     = _get("pegRatio")
    pe      = _get("trailingPE")
    fwd_pe  = _get("forwardPE")
    eps_t   = _get("trailingEps")
    eps_f   = _get("forwardEps")
    rev_g   = _get("revenueGrowth")
    roe     = _get("returnOnEquity")
    d2e     = _get("debtToEquity")
    div_y   = _get("dividendYield")

    data["peg-ratio"]     = f"{peg:.2f}"    if peg    is not None else "N/A"
    data["pe-ratio"]      = f"{pe:.2f}"     if pe     is not None else "N/A"
    data["forward-pe"]    = f"{fwd_pe:.2f}" if fwd_pe is not None else "N/A"
    data["eps-trailing"]  = f"{sym}{eps_t:.4f}" if eps_t is not None else "N/A"
    data["eps-forward"]   = f"{sym}{eps_f:.4f}" if eps_f is not None else "N/A"
    data["revenue-growth"]= _pct(rev_g)
    data["roe"]           = f"{roe  * 100:.1f}%" if roe  is not None else "N/A"
    data["debt-to-equity"]= f"{d2e:.2f}"          if d2e  is not None else "N/A"
    data["dividend-yield"]= f"{div_y * 100:.2f}%" if div_y is not None else "N/A"

    if eps_t is not None and eps_f is not None and eps_t != 0:
        data["eps-growth"] = f"{(eps_f - eps_t) / abs(eps_t) * 100:+.1f}%"
    else:
        data["eps-growth"] = "N/A"

    return data


# ── HTML Injection ────────────────────────────────────────────────────────────

TEXT_IDS = [
    "ticker", "company-name", "date", "current-price", "currency",
    "market-cap", "sector", "industry",
    "week-52-high", "week-52-low", "price-vs-high-pct",
    "sma-50", "sma-200", "cross-age",
    "avg-green-volume", "avg-red-volume",
    "peg-ratio", "pe-ratio", "forward-pe",
    "eps-trailing", "eps-forward", "eps-growth",
    "revenue-growth", "roe", "debt-to-equity", "dividend-yield",
]

SELECT_IDS = [
    "market", "golden-cross", "ma-trend",
    "price-vs-50ma", "price-vs-200ma", "volume-signal",
]


def inject_into_html(data: dict, output_path: str):
    with open(TEMPLATE, "r", encoding="utf-8") as fh:
        soup = BeautifulSoup(fh.read(), "html.parser")

    for fid in TEXT_IDS:
        el = soup.find(id=fid)
        if el and fid in data:
            el["value"] = data[fid]

    for fid in SELECT_IDS:
        el = soup.find(id=fid)
        if el and data.get(fid):
            target = data[fid]
            for opt in el.find_all("option"):
                if opt.get("value") == target:
                    opt["selected"] = "selected"
                elif "selected" in opt.attrs:
                    del opt["selected"]

    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(str(soup))


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("\nUsage:   python fetch_zulu.py TICKER")
        print("Example: python fetch_zulu.py BARC.L   (UK stock)")
        print("         python fetch_zulu.py AAPL      (US stock)")
        sys.exit(1)

    ticker = sys.argv[1].strip().upper()

    if not os.path.exists(TEMPLATE):
        print(f"\nERROR: Template not found at {TEMPLATE}")
        sys.exit(1)

    print(f"\n{'═' * 52}")
    print(f"  Zulu Pre-Trade Checklist  ·  {ticker}")
    print(f"{'═' * 52}")

    try:
        data = fetch_stock_data(ticker)
    except Exception as exc:
        print(f"\n  FAILED: {exc}")
        sys.exit(1)

    safe_name   = ticker.replace(".", "_").replace("/", "_")
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"checklist_{safe_name}.html")

    print("  → Injecting data into checklist template …")
    inject_into_html(data, output_path)

    print(f"\n{'─' * 52}")
    print(f"  Results for {data.get('company-name', ticker)}")
    print(f"{'─' * 52}")
    for k, v in sorted(data.items()):
        print(f"  {k:<22}  {v}")

    print(f"\n  ✓  Saved → {output_path}")
    print("  Open that file in your browser to view the pre-filled checklist.\n")


if __name__ == "__main__":
    main()
