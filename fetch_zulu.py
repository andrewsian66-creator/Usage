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

# Injected at end of <body> so pre-filled dropdowns/inputs trigger the JS
# scoring engine on page load — otherwise signals stay "Pending".
INIT_SCRIPT = """
document.addEventListener('DOMContentLoaded', function () {
  [evalPeg1, evalPeg2, evalEps, evalBuyback, evalMarket,
   evalMoat, evalDirectors, evalSkeletons, evalSmoothness, evalGoodwill,
   evalRegime, evalMADirection, evalCrossAge, evalPriceVsMA, evalVolume,
   updateTitle].forEach(function (fn) { try { fn(); } catch (e) {} });
});
"""


# ── Helpers ───────────────────────────────────────────────────────────────────

def _sym(currency: str) -> str:
    return {"GBp": "p", "GBX": "p", "GBP": "£", "USD": "$", "EUR": "€"}.get(currency, currency + " ")


# ── Data Fetching ─────────────────────────────────────────────────────────────

def fetch_stock_data(ticker: str) -> dict:
    print("  → Connecting to Yahoo Finance …")
    stock = yf.Ticker(ticker)
    info  = stock.info

    print("  → Fetching 1-year historical price / volume data …")
    hist = stock.history(period="1y")

    if hist.empty:
        raise ValueError(f"No data returned for '{ticker}'. Check the ticker symbol and try again.")

    data     = {}
    currency = info.get("currency", "GBp" if ticker.upper().endswith(".L") else "USD")
    sym      = _sym(currency)

    # ── Candidate Details ─────────────────────────────────────────────────────
    data["ticker"]    = ticker.upper()
    data["company"]   = info.get("longName") or info.get("shortName", ticker)
    data["tradedate"] = date.today().strftime("%d %b %Y")

    raw_price = (info.get("currentPrice")
                 or info.get("regularMarketPrice")
                 or info.get("previousClose")
                 or float(hist["Close"].iloc[-1]))
    # Price field label says "(p)" — store raw number; for non-pence stocks
    # the user can see the currency in the company name / their broker
    data["price"] = f"{raw_price:.2f}"

    # ── LEG 1 — QUANTS ───────────────────────────────────────────────────────
    print("  → Pulling fundamental data …")

    # PEG (primary proxy — Yahoo Finance; Stockopedia cross-reference is manual)
    peg = info.get("pegRatio")
    data["peg1"] = f"{peg:.2f}" if peg is not None else ""

    # EPS Growth → maps to select#eps
    eps_t = info.get("trailingEps")
    eps_f = info.get("forwardEps")
    if eps_t is not None and eps_f is not None and eps_t != 0:
        growth = (eps_f - eps_t) / abs(eps_t) * 100
        if growth > 15:
            data["eps"] = "strong"
        elif growth > 0:
            data["eps"] = "positive"
        elif growth > -5:
            data["eps"] = "flat"
        else:
            data["eps"] = "negative"
    else:
        data["eps"] = ""

    # Market → select#market (values: main | aim_exception | aim)
    exchange      = info.get("exchange",         "").upper()
    full_exchange = info.get("fullExchangeName", "").upper()
    if "AIM" in exchange or "AIM" in full_exchange:
        data["market"] = "aim"
    elif ticker.upper().endswith(".L") or exchange in ("LSE", "LON", "XLON"):
        data["market"] = "main"
    else:
        data["market"] = "main"   # default; user can correct for non-UK

    # Goodwill → select#goodwill (values: low | moderate | high)
    # Calculated from balance sheet vs operating profit
    data["goodwill"] = ""
    try:
        bs = stock.balance_sheet
        fs = stock.financials
        gw = None
        for lbl in ("Goodwill", "GoodwillAndOtherIntangibleAssets",
                    "Goodwill And Intangible Assets"):
            if bs is not None and not bs.empty and lbl in bs.index:
                gw = float(bs.loc[lbl].iloc[0])
                break
        op = None
        for lbl in ("Operating Income", "EBIT", "Ebit"):
            if fs is not None and not fs.empty and lbl in fs.index:
                op = float(fs.loc[lbl].iloc[0])
                break
        if gw is not None and gw == 0:
            data["goodwill"] = "low"
        elif gw is not None and op is not None and op > 0:
            ratio = gw / op
            data["goodwill"] = "low" if ratio < 1 else ("moderate" if ratio < 3 else "high")
    except Exception:
        pass

    # ── LEG 3 — CHART ────────────────────────────────────────────────────────
    print("  → Calculating 50-day and 200-day SMAs …")
    hist  = hist.copy()
    closes = hist["Close"]
    sma50 = sma200 = None

    if len(closes) >= 50:
        hist["SMA50"] = closes.rolling(50).mean()
        sma50 = float(hist["SMA50"].iloc[-1])

    if len(closes) >= 200:
        hist["SMA200"] = closes.rolling(200).mean()
        sma200 = float(hist["SMA200"].iloc[-1])

    if sma50 is not None and sma200 is not None:
        # MA Regime → select#regime
        if sma50 > sma200:
            data["regime"] = "golden"
        elif abs(sma50 - sma200) / sma200 < 0.02:
            data["regime"] = "near_golden"
        else:
            data["regime"] = "death"

        # MA Direction → select#madirection
        if len(hist) >= 211:
            gap_now = float(hist["SMA50"].iloc[-1])  - float(hist["SMA200"].iloc[-1])
            gap_10d = float(hist["SMA50"].iloc[-11]) - float(hist["SMA200"].iloc[-11])
            if sma50 > sma200:
                # Golden cross — is it widening or the cross is decaying?
                if gap_now > gap_10d * 1.015:
                    data["madirection"] = "widening"
                elif gap_now < gap_10d * 0.97:
                    data["madirection"] = "declining"
                else:
                    data["madirection"] = "neutral"
            else:
                # Death cross — is 50-day moving toward 200-day from below?
                if abs(gap_now) < abs(gap_10d) * 0.97:
                    data["madirection"] = "approaching"
                elif abs(gap_now) > abs(gap_10d) * 1.015:
                    data["madirection"] = "declining"
                else:
                    data["madirection"] = "neutral"
        else:
            data["madirection"] = "neutral"

        # Cross Age → select#crossage (only meaningful for a live golden cross)
        crossover_ts = None
        valid = hist[["SMA50", "SMA200"]].dropna()
        diff  = valid["SMA50"] - valid["SMA200"]
        for i in range(len(diff) - 1, 0, -1):
            prev, curr = diff.iloc[i - 1], diff.iloc[i]
            if (prev <= 0 < curr) or (prev >= 0 > curr):
                crossover_ts = diff.index[i]
                break

        if crossover_ts is not None and sma50 > sma200:
            now_ts   = pd.Timestamp.now(tz=crossover_ts.tzinfo)
            age_days = (now_ts - crossover_ts).days
            if age_days < 30:
                data["crossage"] = "fresh"
            elif age_days <= 90:
                data["crossage"] = "aged"
            else:
                data["crossage"] = "old"
        else:
            data["crossage"] = "na"

        # Price vs MAs → select#pricevsma
        if raw_price > sma50 and raw_price > sma200:
            data["pricevsma"] = "above_both"
        elif raw_price < sma50 and raw_price < sma200:
            data["pricevsma"] = "below_both"
        elif raw_price < sma50 and raw_price >= sma200:
            data["pricevsma"] = "below_50"
        else:
            data["pricevsma"] = "between"
    else:
        data["regime"]      = "no_cross"
        data["madirection"] = ""
        data["crossage"]    = "na"
        data["pricevsma"]   = ""

    # Volume Profile → select#volume
    print("  → Analysing 30-day volume profile …")
    recent = hist.tail(30).copy()
    recent["green"] = recent["Close"] >= recent["Open"]
    avg_green = recent.loc[ recent["green"], "Volume"].mean()
    avg_red   = recent.loc[~recent["green"], "Volume"].mean()

    if pd.notna(avg_green) and pd.notna(avg_red) and avg_red > 0:
        ratio = avg_green / avg_red
        data["volume"] = "accumulation" if ratio > 1.10 else (
                         "distribution" if ratio < 0.91 else "neutral")
    else:
        data["volume"] = "neutral"

    return data


# ── HTML Injection ────────────────────────────────────────────────────────────

# Plain text inputs in the dark "Candidate Details" strip
TEXT_INPUTS = ["ticker", "company", "tradedate", "price"]

# Numeric input (PEG primary source)
NUMBER_INPUTS = ["peg1"]

# Select dropdowns we can auto-populate
SELECT_AUTO = ["eps", "market", "goodwill",
               "regime", "madirection", "crossage", "pricevsma", "volume"]


def inject_into_html(data: dict, output_path: str) -> None:
    with open(TEMPLATE, "r", encoding="utf-8") as fh:
        soup = BeautifulSoup(fh.read(), "html.parser")

    # Text inputs
    for fid in TEXT_INPUTS:
        el = soup.find(id=fid)
        if el and data.get(fid):
            el["value"] = data[fid]

    # Number inputs — only set if we have a real number
    for fid in NUMBER_INPUTS:
        el = soup.find(id=fid)
        val = data.get(fid, "")
        if el and val and val not in ("N/A", ""):
            el["value"] = val

    # Select dropdowns — mark the right <option> as selected
    for fid in SELECT_AUTO:
        el = soup.find(id=fid)
        if not el or not data.get(fid):
            continue
        target = data[fid]
        for opt in el.find_all("option"):
            if opt.get("value") == target:
                opt["selected"] = "selected"
            elif "selected" in opt.attrs:
                del opt["selected"]

    # Append init script so the JS scoring engine fires on page load
    tag = soup.new_tag("script")
    tag.string = INIT_SCRIPT
    soup.body.append(tag)

    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(str(soup))


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("\nUsage:   python fetch_zulu.py TICKER")
        print("Example: python fetch_zulu.py BARC.L   (UK — Barclays)")
        print("         python fetch_zulu.py TSCO.L   (UK — Tesco)")
        print("         python fetch_zulu.py AAPL     (US — Apple)")
        sys.exit(1)

    ticker = sys.argv[1].strip().upper()

    if not os.path.exists(TEMPLATE):
        print(f"\nERROR: Template not found → {TEMPLATE}")
        sys.exit(1)

    print(f"\n{'═' * 56}")
    print(f"  Zulu Pre-Trade Checklist  ·  {ticker}")
    print(f"{'═' * 56}")

    try:
        data = fetch_stock_data(ticker)
    except Exception as exc:
        print(f"\n  FAILED: {exc}")
        sys.exit(1)

    safe      = ticker.replace(".", "_").replace("/", "_")
    out_path  = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             f"checklist_{safe}.html")

    print("  → Injecting data into checklist template …")
    inject_into_html(data, out_path)

    print(f"\n{'─' * 56}")
    print(f"  Auto-filled fields for {data.get('company', ticker)}")
    print(f"{'─' * 56}")
    labels = {
        "ticker":      "Ticker",
        "company":     "Company",
        "tradedate":   "Date",
        "price":       "Price",
        "peg1":        "PEG ratio (Yahoo proxy)",
        "eps":         "EPS growth signal",
        "market":      "Market",
        "goodwill":    "Goodwill signal",
        "regime":      "MA Regime",
        "madirection": "MA Direction (FCH rule)",
        "crossage":    "Cross age",
        "pricevsma":   "Price vs MAs",
        "volume":      "Volume signal",
    }
    for k, label in labels.items():
        v = data.get(k, "")
        print(f"  {label:<30}  {v or '(blank — fill manually)'}")

    print(f"\n  ✓  Saved → {out_path}")
    print("  Open that file in your browser — signals will auto-calculate.\n")

    # Remind user which fields need manual completion
    print("  Manual fields still needed:")
    print("    Leg 1 · PEG secondary source (Simply Wall St)")
    print("    Leg 1 · Buyback programme type")
    print("    Leg 2 · Moat, Director dealings, Skeletons, Revenue smoothness")
    print("    Leg 2 · Goodwill (if not auto-detected)")
    print()


if __name__ == "__main__":
    main()
