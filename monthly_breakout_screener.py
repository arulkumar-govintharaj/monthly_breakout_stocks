#!/usr/bin/env python3
"""
monthly_breakout_screener.py

Screens the NSE F&O stock universe for symbols whose price has broken
the PREVIOUS CALENDAR MONTH's HIGH and/or LOW during the current
session — output as a single Excel file, with an optional email
(stock names in the body, spreadsheet attached).

52-WEEK LOW SCAN (NIFTY 500)
----------------------------
In the same run, the script also scans every NIFTY 500 constituent
and flags stocks whose session LOW is at or below their lowest low of
the previous 52 weeks (i.e. a fresh 52-week low today). These appear
in their own section of the email and on a second sheet of the Excel
file. Stocks that were NOT on the list at the previous run are marked
NEW (tracked in state/52w_low_state.csv).

The NIFTY 500 constituent list is downloaded from NSE on each run and
cached to data/nifty500_list.csv; if NSE can't be reached, the cached
copy is used. Disable the scan with --no-52w-low.

STATE / NEW-BREAKOUT TRACKING
------------------------------
Each run writes the current set of breakout symbols to a small CSV
state file (default: state/breakout_state.csv). On the NEXT run, the
script loads that file and compares it to today's breakouts:

  - If no state file exists yet (first run, or it wasn't committed),
    the email lists ALL current breakouts and the CSV is attached.
  - If a state file exists, the email lists ONLY the symbols that are
    breaking out today but were NOT in the previous state (i.e. new
    breakouts since the last run).

The state file is meant to be committed back to the repo by the
GitHub Actions workflow after each run so the next run can diff
against it.

Usage:
    python monthly_breakout_screener.py
    python monthly_breakout_screener.py --symbols RELIANCE,SBIN,SONACOMS
    python monthly_breakout_screener.py --no-email
    python monthly_breakout_screener.py --only-breakouts
    python monthly_breakout_screener.py --state-file state/breakout_state.csv

Email is sent via Gmail SMTP using credentials from environment
variables (see README for GitHub Actions secrets setup):
    EMAIL_USERNAME     - sender Gmail address
    EMAIL_APP_PASSWORD - Gmail App Password (not your normal password)
    EMAIL_TO           - comma-separated recipient address(es)
"""

import argparse
import datetime as dt
import io
import os
import smtplib
import sys
import time
import urllib.request
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
import yfinance as yf

# --------------------------------------------------------------------------
# Default NSE F&O universe (add/remove symbols as NSE revises the F&O list)
# --------------------------------------------------------------------------
DEFAULT_FNO_UNIVERSE = [
    "RELIANCE", "TCS", "HDFCBANK", "ICICIBANK", "INFY", "SBIN", "AXISBANK",
    "KOTAKBANK", "BAJFINANCE", "BHARTIARTL", "ITC", "LT", "HINDUNILVR",
    "MARUTI", "TITAN", "SUNPHARMA", "ULTRACEMCO", "M&M", "TATASTEEL",
    "TATAMOTORS", "ADANIENT", "ADANIPORTS", "ASIANPAINT", "BAJAJFINSV",
    "NTPC", "POWERGRID", "HCLTECH", "WIPRO", "TECHM", "GRASIM",
    "HINDALCO", "JSWSTEEL", "COALINDIA", "ONGC", "BPCL", "IOC",
    "DIVISLAB", "DRREDDY", "CIPLA", "APOLLOHOSP", "NESTLEIND", "BRITANNIA",
    "TATACONSUM", "EICHERMOT", "HEROMOTOCO", "BAJAJ-AUTO", "SHRIRAMFIN",
    "SBILIFE", "HDFCLIFE", "INDUSINDBK", "PIDILITIND", "DLF", "GODREJCP",
    "HAVELLS", "SIEMENS", "AMBUJACEM", "ACC", "BANDHANBNK", "BEL",
    "BOSCHLTD", "CANBK", "CHOLAFIN", "COLPAL", "CONCOR", "CUMMINSIND",
    "DABUR", "DELHIVERY", "GAIL", "GLENMARK", "HINDPETRO", "IDEA",
    "IDFCFIRSTB", "IGL", "INDHOTEL", "INDIGO", "IRCTC", "JINDALSTEL",
    "JUBLFOOD", "LICHSGFIN", "LTF", "LTIM", "LUPIN", "MFSL", "MOTHERSON",
    "MPHASIS", "MRF", "MUTHOOTFIN", "NAUKRI", "NMDC", "OBEROIRLTY",
    "OFSS", "PAGEIND", "PATANJALI", "PEL", "PERSISTENT", "PETRONET",
    "PFC", "PIIND", "PNB", "POLYCAB", "RECLTD", "SAIL", "SONACOMS",
    "SRF", "TATACOMM", "TATAPOWER", "TORNTPHARM", "TRENT", "TVSMOTOR",
    "UBL", "UPL", "VEDL", "VOLTAS", "ZYDUSLIFE", "ZOMATO", "PAYTM",
]

# --------------------------------------------------------------------------
# NIFTY 500 universe for the 52-week-low scan
# --------------------------------------------------------------------------
NIFTY500_URLS = [
    "https://archives.nseindia.com/content/indices/ind_nifty500list.csv",
    "https://www.niftyindices.com/IndexConstituent/ind_nifty500list.csv",
]
NIFTY500_MIN_ROWS = 400  # sanity check that we got the real list, not an error page


def _cache_stamp_path(cache_path: Path) -> Path:
    return cache_path.with_name(cache_path.stem + "_updated.txt")


def _cache_age_days(cache_path: Path):
    """Days since the cached list was downloaded, from the sidecar date
    file (git checkout resets file mtimes, so mtime can't be trusted).
    Returns None if unknown."""
    stamp = _cache_stamp_path(cache_path)
    try:
        updated = dt.date.fromisoformat(stamp.read_text().strip())
        return (dt.date.today() - updated).days
    except Exception:  # noqa: BLE001
        return None


_HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/csv,*/*",
}


def load_cached_list(name, cache_path: Path, urls, parse, max_age_days: int = 30):
    """
    Generic "keep it in the repo, refresh it every N days" loader.
    `parse(text) -> DataFrame` must raise if the text isn't the real list.
    Order of preference:
      1. Repo copy at `cache_path`, if downloaded < `max_age_days` ago
      2. Fresh download from `urls` (raw text saved to `cache_path` + date stamp)
      3. Stale repo copy, if every download fails
    Returns None if none of these work.
    """
    age = _cache_age_days(cache_path)
    if cache_path.exists() and age is not None and age < max_age_days:
        try:
            df = parse(cache_path.read_text(encoding="utf-8"))
            print(f"[INFO] Using {name} from repo ({cache_path}, {age} days old; "
                  f"refreshes after {max_age_days} days)")
            return df
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] Repo copy of {name} unreadable ({exc}) — re-downloading",
                  file=sys.stderr)
    if not cache_path.exists():
        reason = "not in repo yet"
    elif age is None:
        reason = "has no download-date stamp"
    else:
        reason = f"is {age} days old"
    print(f"[INFO] {name} {reason} — downloading fresh copy from NSE")

    for url in urls:
        try:
            req = urllib.request.Request(url, headers=_HTTP_HEADERS)
            with urllib.request.urlopen(req, timeout=30) as resp:
                text = resp.read().decode("utf-8", errors="replace")
            df = parse(text)
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            cache_path.write_text(text, encoding="utf-8")
            _cache_stamp_path(cache_path).write_text(dt.date.today().isoformat() + "\n")
            print(f"[INFO] {name} downloaded from {url} ({len(df)} symbols), saved to {cache_path}")
            return df
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] Could not download {name} from {url}: {exc}", file=sys.stderr)

    if cache_path.exists():
        try:
            df = parse(cache_path.read_text(encoding="utf-8"))
            print(f"[WARN] Download failed — using existing (stale) {name} at {cache_path}",
                  file=sys.stderr)
            return df
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] Stale {name} unreadable too: {exc}", file=sys.stderr)
    return None


def _parse_nifty500(text: str) -> pd.DataFrame:
    df = pd.read_csv(io.StringIO(text))
    df.columns = [str(c).strip() for c in df.columns]
    if "Symbol" not in df.columns:
        df = df.rename(columns={df.columns[0]: "Symbol"})
    out = pd.DataFrame({"Symbol": df["Symbol"].astype(str).str.strip().str.upper()})
    out["Company"] = df["Company Name"].astype(str).str.strip() if "Company Name" in df else ""
    out["Industry"] = df["Industry"].astype(str).str.strip() if "Industry" in df else ""
    out = out[out["Symbol"].ne("") & out["Symbol"].ne("NAN")].drop_duplicates("Symbol")
    return out.reset_index(drop=True)


def load_nifty500(cache_path: Path, override_path: str = None,
                  max_age_days: int = 30) -> pd.DataFrame:
    """
    Returns a DataFrame with columns Symbol, Company, Industry for the
    NIFTY 500 — from the --nifty500-file override if given, otherwise the
    repo copy (refreshed from NSE every `max_age_days`). Empty if unavailable.
    """
    if override_path:
        p = Path(override_path)
        if not p.exists():
            print(f"[ERROR] NIFTY 500 file not found: {p}", file=sys.stderr)
            sys.exit(1)
        return _parse_nifty500(p.read_text(encoding="utf-8"))

    def parse(text):
        df = _parse_nifty500(text)
        if len(df) < NIFTY500_MIN_ROWS:
            raise ValueError(f"unexpected content ({len(df)} rows)")
        return df

    df = load_cached_list("NIFTY 500 list", cache_path, NIFTY500_URLS, parse, max_age_days)
    if df is None:
        print("[WARN] No NIFTY 500 list available — 52-week-low scan skipped.", file=sys.stderr)
        return pd.DataFrame(columns=["Symbol", "Company", "Industry"])
    return df


# --------------------------------------------------------------------------
# Official NSE F&O stock list (with lot sizes) — decides which 52W-low
# stocks go in the "has futures" table.
# --------------------------------------------------------------------------
FNO_LOTS_URLS = [
    "https://archives.nseindia.com/content/fo/fo_mktlots.csv",
    "https://nsearchives.nseindia.com/content/fo/fo_mktlots.csv",
]
FNO_MIN_ROWS = 100
_INDEX_UNDERLYINGS = {"NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "NIFTYNXT50"}


def _parse_fno_lots(text: str) -> pd.DataFrame:
    """
    Parses NSE's fo_mktlots.csv (UNDERLYING, SYMBOL, <month columns>...;
    padded with spaces, index rows first). Returns Symbol, Lot_Size.
    """
    df = pd.read_csv(io.StringIO(text), dtype=str, skipinitialspace=True)
    df.columns = [str(c).strip() for c in df.columns]
    sym_col = next((c for c in df.columns if c.upper() == "SYMBOL"), None)
    if sym_col is None:
        raise ValueError("no SYMBOL column")
    month_cols = df.columns[df.columns.get_loc(sym_col) + 1:]

    def first_lot(row):
        for c in month_cols:
            v = str(row[c]).strip().replace(",", "")
            if v.isdigit():
                return int(v)
        return None

    out = pd.DataFrame({
        "Symbol": df[sym_col].astype(str).str.strip().str.upper(),
        "Lot_Size": df.apply(first_lot, axis=1),
    })
    out = out[out["Symbol"].str.match(r"^[A-Z0-9&\-]+$") & ~out["Symbol"].isin(_INDEX_UNDERLYINGS)
              & out["Symbol"].ne("SYMBOL")]
    out = out.drop_duplicates("Symbol").reset_index(drop=True)
    if len(out) < FNO_MIN_ROWS:
        raise ValueError(f"unexpected content ({len(out)} stocks)")
    return out


def load_fno_lots(cache_path: Path, max_age_days: int = 30) -> dict:
    """
    Returns {symbol: lot_size_or_None} for every stock in NSE's F&O segment.
    Falls back to the built-in DEFAULT_FNO_UNIVERSE (no lot sizes) if the
    official list can't be loaded.
    """
    df = load_cached_list("NSE F&O list", cache_path, FNO_LOTS_URLS, _parse_fno_lots, max_age_days)
    if df is None:
        print("[WARN] Official F&O list unavailable — using built-in F&O list for tagging.",
              file=sys.stderr)
        return {s: None for s in DEFAULT_FNO_UNIVERSE}
    return {s: (int(l) if pd.notna(l) else None) for s, l in zip(df["Symbol"], df["Lot_Size"])}


def _extract_ticker_frame(data: pd.DataFrame, ticker: str) -> pd.DataFrame:
    """Pulls one ticker's OHLC frame out of a yf.download() result."""
    if isinstance(data.columns, pd.MultiIndex):
        if ticker in data.columns.get_level_values(0):
            return data[ticker]
        if ticker in data.columns.get_level_values(1):
            return data.xs(ticker, axis=1, level=1)
        return pd.DataFrame()
    return data


def find_52w_low(df: pd.DataFrame, min_history_weeks: int = 50):
    """
    Given one stock's daily OHLC (oldest -> newest), decides whether the
    latest session made a new 52-week low. Returns a dict of metrics, or
    None if there isn't enough clean history to judge.

    New 52W low  := latest session's Low <= lowest Low of the sessions in
                    the 52 calendar weeks BEFORE the latest session.
    """
    df = df.dropna(subset=["High", "Low", "Close"])
    if len(df) < 2:
        return None

    last_date = df.index[-1]
    window_start = last_date - pd.Timedelta(weeks=52)
    prior = df[(df.index >= window_start) & (df.index < last_date)]
    # Skip recent listings — a "52-week low" needs ~a year of history
    if prior.empty or prior.index[0] > last_date - pd.Timedelta(weeks=min_history_weeks):
        return None

    today = df.iloc[-1]
    prior_low = float(prior["Low"].min())
    prior_low_date = prior["Low"].idxmin()
    high_52w = float(max(prior["High"].max(), today["High"]))
    today_low = float(today["Low"])
    close = float(today["Close"])

    return {
        "Hit_52W_Low": today_low <= prior_low,
        "LTP": round(close, 2),
        "Today_Low": round(today_low, 2),
        "Prev_52W_Low": round(prior_low, 2),
        "Prev_52W_Low_Date": pd.Timestamp(prior_low_date).strftime("%Y-%m-%d"),
        "52W_High": round(high_52w, 2),
        "%_Low_vs_Prev_52W_Low": round((today_low - prior_low) / prior_low * 100, 2),
        "%_LTP_From_52W_High": round((close - high_52w) / high_52w * 100, 2),
        "Closed_Below_Prev_52W_Low": "YES" if close < prior_low else "",
        "Session_Date": pd.Timestamp(last_date).strftime("%Y-%m-%d"),
    }


def scan_52w_lows(universe: pd.DataFrame, fno_symbols, chunk_size: int = 100) -> pd.DataFrame:
    """
    Batch-downloads ~13 months of daily data for every symbol in
    `universe` and returns ONLY the rows that made a fresh 52-week low in
    the latest session, sorted deepest break first.
    `fno_symbols` is a {symbol: lot_size} dict (or any iterable of symbols).
    """
    meta = universe.set_index("Symbol")
    symbols = list(meta.index)
    fno = fno_symbols if isinstance(fno_symbols, dict) else {s: None for s in fno_symbols}
    rows = []

    for start in range(0, len(symbols), chunk_size):
        chunk = symbols[start:start + chunk_size]
        tickers = [f"{s}.NS" for s in chunk]
        print(f"[52W] Downloading {start + 1}-{start + len(chunk)} of {len(symbols)}...")
        try:
            data = yf.download(
                tickers, period="13mo", interval="1d", auto_adjust=False,
                group_by="ticker", threads=True, progress=False,
            )
        except Exception as exc:  # noqa: BLE001
            print(f"  [WARN] 52W batch download failed: {exc}", file=sys.stderr)
            continue

        for sym, ticker in zip(chunk, tickers):
            try:
                result = find_52w_low(_extract_ticker_frame(data, ticker))
            except Exception as exc:  # noqa: BLE001
                print(f"  [WARN] {sym}: 52W calc failed ({exc})", file=sys.stderr)
                continue
            if result is None:
                continue
            rows.append({
                "Symbol": sym,
                "Company": meta.at[sym, "Company"],
                "Industry": meta.at[sym, "Industry"],
                "In_FnO": "YES" if sym in fno else "",
                "Lot_Size": fno.get(sym),
                **result,
            })
        time.sleep(1)  # be polite between batches

    all_df = pd.DataFrame(rows)
    if all_df.empty:
        return all_df

    # Drop stale symbols (suspended / no trade today) — keep only the
    # most common latest session date across the universe.
    latest_session = all_df["Session_Date"].mode().iloc[0]
    stale = all_df[all_df["Session_Date"] != latest_session]
    if not stale.empty:
        print(f"[52W] Ignoring {len(stale)} symbols with stale data (not {latest_session})")
    all_df = all_df[all_df["Session_Date"] == latest_session]

    lows = all_df[all_df["Hit_52W_Low"]].drop(columns="Hit_52W_Low")
    print(f"[52W] {len(all_df)} symbols evaluated, {len(lows)} at a fresh 52-week low")
    return lows.sort_values("%_Low_vs_Prev_52W_Low").reset_index(drop=True)


def load_previous_52w(state_path: Path):
    """Returns (symbol_set, state_existed)."""
    if not state_path.exists():
        return set(), False
    try:
        prev = pd.read_csv(state_path)
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] Could not read 52W state {state_path}: {exc}", file=sys.stderr)
        return set(), False
    return (set(prev["Symbol"]) if "Symbol" in prev else set()), True


def save_52w_state(lows_df: pd.DataFrame, state_path: Path):
    state_path.parent.mkdir(parents=True, exist_ok=True)
    cols = ["Symbol", "Today_Low", "Prev_52W_Low", "Session_Date"]
    (lows_df[cols] if not lows_df.empty else pd.DataFrame(columns=cols)).to_csv(
        state_path, index=False
    )
    print(f"[INFO] 52W state saved to {state_path.resolve()} ({len(lows_df)} rows)")


def fetch_monthly_and_today(symbol: str, retries: int = 2, pause: float = 0.5):
    """
    Returns a dict with the previous COMPLETED calendar month's high/low
    and the latest available session's high/low/close for `symbol`
    (NSE), or None if data could not be fetched.
    """
    ticker = f"{symbol}.NS"
    for attempt in range(retries + 1):
        try:
            daily = yf.Ticker(ticker).history(
                period="6mo", interval="1d", auto_adjust=False
            )
            daily = daily.dropna(how="all")
            if daily.empty:
                return None

            monthly = daily.resample("ME").agg({"High": "max", "Low": "min"})
            monthly.index = monthly.index.to_period("M")

            current_period = pd.Timestamp.now().to_period("M")
            completed_months = monthly[monthly.index < current_period]
            if completed_months.empty:
                return None
            prev_month = completed_months.iloc[-1]
            prev_month_label = str(completed_months.index[-1])

            latest_day = daily.iloc[-1]
            latest_date = daily.index[-1]

            return {
                "Symbol": symbol,
                "Prev_Month": prev_month_label,
                "Prev_Month_High": round(float(prev_month["High"]), 2),
                "Prev_Month_Low": round(float(prev_month["Low"]), 2),
                "Latest_Date": latest_date.strftime("%Y-%m-%d"),
                "Latest_High": round(float(latest_day["High"]), 2),
                "Latest_Low": round(float(latest_day["Low"]), 2),
                "LTP": round(float(latest_day["Close"]), 2),
            }
        except Exception as exc:  # noqa: BLE001
            if attempt < retries:
                time.sleep(pause)
                continue
            print(f"  [WARN] {symbol}: failed to fetch data ({exc})", file=sys.stderr)
            return None


def build_screener_table(symbols):
    rows = []
    total = len(symbols)
    for i, symbol in enumerate(symbols, 1):
        print(f"[{i}/{total}] Fetching {symbol}...")
        data = fetch_monthly_and_today(symbol)
        if data is None:
            continue

        broke_high = data["Latest_High"] > data["Prev_Month_High"]
        broke_low = data["Latest_Low"] < data["Prev_Month_Low"]

        pct_from_high = round(
            (data["LTP"] - data["Prev_Month_High"]) / data["Prev_Month_High"] * 100, 2
        )
        pct_from_low = round(
            (data["LTP"] - data["Prev_Month_Low"]) / data["Prev_Month_Low"] * 100, 2
        )

        rows.append(
            {
                "Symbol": data["Symbol"],
                "LTP": data["LTP"],
                "Prev_Month": data["Prev_Month"],
                "Prev_Month_High": data["Prev_Month_High"],
                "Prev_Month_Low": data["Prev_Month_Low"],
                "Today_High": data["Latest_High"],
                "Today_Low": data["Latest_Low"],
                "Broke_Month_High": "YES" if broke_high else "",
                "Broke_Month_Low": "YES" if broke_low else "",
                "%_From_Month_High": pct_from_high,
                "%_From_Month_Low": pct_from_low,
                "Session_Date": data["Latest_Date"],
            }
        )
        time.sleep(0.2)  # be polite to the data source

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["_sort_key"] = (
        (df["Broke_Month_High"] == "YES").astype(int) * 2
        + (df["Broke_Month_Low"] == "YES").astype(int)
    )
    df = df.sort_values(["_sort_key", "Symbol"], ascending=[False, True]).drop(
        columns="_sort_key"
    )
    return df.reset_index(drop=True)


def load_universe(universe_arg, symbols_arg):
    if symbols_arg:
        return [s.strip().upper() for s in symbols_arg.split(",") if s.strip()]
    if universe_arg:
        path = Path(universe_arg)
        if not path.exists():
            print(f"[ERROR] Universe file not found: {path}", file=sys.stderr)
            sys.exit(1)
        col = pd.read_csv(path)
        first_col = col.columns[0]
        return [str(s).strip().upper() for s in col[first_col].dropna().tolist()]
    return DEFAULT_FNO_UNIVERSE


def load_previous_breakouts(state_path: Path):
    """
    Returns (prev_high_set, prev_low_set, state_existed).
    An empty/missing file means no previous state (first run).
    """
    if not state_path.exists():
        return set(), set(), False

    try:
        prev_df = pd.read_csv(state_path)
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] Could not read state file {state_path}: {exc}", file=sys.stderr)
        return set(), set(), False

    if prev_df.empty:
        return set(), set(), True

    prev_high = set(prev_df.loc[prev_df["Broke_Month_High"] == "YES", "Symbol"])
    prev_low = set(prev_df.loc[prev_df["Broke_Month_Low"] == "YES", "Symbol"])
    return prev_high, prev_low, True


def save_state(full_df: pd.DataFrame, state_path: Path):
    """Persists today's breakout rows (only) as the new state file."""
    state_path.parent.mkdir(parents=True, exist_ok=True)
    breakout_df = full_df[
        (full_df["Broke_Month_High"] == "YES") | (full_df["Broke_Month_Low"] == "YES")
    ].copy()
    breakout_df.to_csv(state_path, index=False)
    print(f"[INFO] State saved to {state_path.resolve()} ({len(breakout_df)} rows)")


def format_symbol_list(symbols):
    return ", ".join(sorted(symbols)) if symbols else "None"


def format_52w_section(lows_df, new_52w, status_note=None):
    """Plain-text email block listing NIFTY 500 stocks at a 52-week low."""
    lines = ["=" * 60]
    if status_note:
        lines += ["52-WEEK LOW STOCKS (NIFTY 500)", status_note]
        return lines
    count = 0 if lows_df is None else len(lows_df)
    lines.append(f"52-WEEK LOW STOCKS — NIFTY 500 ({count}, {len(new_52w)} new)")
    if not count:
        lines.append("None today.")
        return lines
    lines.append("[NEW] = not on the list at the previous run")
    fno_lows, cash_lows = split_52w_by_fno(lows_df)
    for title, part in (("F&O STOCKS (futures available)", fno_lows),
                        ("NON-F&O STOCKS (cash only)", cash_lows)):
        lines += ["", f"-- {title} ({len(part)}) --"]
        if part.empty:
            lines.append("None")
        for _, r in part.iterrows():
            lot = r.get("Lot_Size")
            lot_txt = f" | lot {int(lot)}" if lot is not None and pd.notna(lot) else ""
            lines.append(
                f"{r['Symbol']:<12} {'[NEW]' if r['Symbol'] in new_52w else ''}".rstrip()
                + f"\n    LTP {r['LTP']:,.2f}{lot_txt} | Low {r['Today_Low']:,.2f} vs prev 52W "
                  f"low {r['Prev_52W_Low']:,.2f} on {r['Prev_52W_Low_Date']} "
                  f"({r['%_Low_vs_Prev_52W_Low']:+.2f}%) | "
                  f"{r['%_LTP_From_52W_High']:+.1f}% from 52W high"
                + (" | closed below" if r["Closed_Below_Prev_52W_Low"] == "YES" else "")
            )
    return lines


# --------------------------------------------------------------------------
# HTML email (inline styles only — Gmail ignores most <style> blocks).
# Table cells use HTML attributes (align/nowrap/bgcolor) instead of
# per-cell inline CSS to keep the email under Gmail's ~102 KB clip limit.
# --------------------------------------------------------------------------
_C = {
    "ink": "#111827", "muted": "#6b7280", "line": "#e5e7eb", "zebra": "#f3f4f6",
    "head": "#1f2937", "green": "#15803d", "green_bg": "#dcfce7",
    "red": "#b91c1c", "red_bg": "#fee2e2", "amber": "#92400e", "amber_bg": "#fef3c7",
    "blue": "#1e40af", "blue_bg": "#dbeafe",
}
_FONT = "font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;"
_TH = f"background:{_C['head']};color:#fff;font-size:12px;font-weight:600"
MAX_EMAIL_ROWS = 60  # per table; the rest are in the attached spreadsheet


def _esc(v):
    import html
    return html.escape("" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))


def _num(v, dp=2):
    try:
        return f"{float(v):,.{dp}f}"
    except (TypeError, ValueError):
        return _esc(v)


def _pct(v, colour=True):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return _esc(v)
    c = _C["green"] if f > 0 else _C["red"] if f < 0 else _C["muted"]
    return f'<b style="color:{c if colour else _C["ink"]}">{f:+.2f}%</b>'


def _badge(text, fg, bg):
    return (f' <b style="padding:1px 5px;border-radius:3px;background:{bg};color:{fg};'
            f'font-size:10px">{text}</b>')


def _sub(text):
    """Small grey second line inside a cell."""
    return f'<br><small style="color:{_C["muted"]}">{text}</small>'


def _stat(label, value, fg, bg):
    return (f'<div style="display:inline-block;vertical-align:top;background:{bg};border-radius:8px;'
            f'padding:10px 14px;margin:0 6px 6px 0;min-width:100px">'
            f'<div style="font-size:22px;font-weight:700;color:{fg}">{value}</div>'
            f'<div style="font-size:11px;color:{fg};text-transform:uppercase;letter-spacing:.4px">'
            f'{label}</div></div>')


def _section_title(text, count, fg):
    return (f'<h2 style="margin:28px 0 8px;font-size:16px;color:{_C["ink"]}">'
            f'<span style="display:inline-block;width:4px;height:16px;background:{fg};'
            f'vertical-align:-2px;margin-right:8px;border-radius:2px"></span>{_esc(text)} '
            f'<span style="color:{_C["muted"]};font-weight:400">({count})</span></h2>')


def _table(headers, rows, total=None):
    """headers: list of (label, align); rows: list of lists of HTML cell strings.
    Shows at most MAX_EMAIL_ROWS rows; `total` is the full count for the note."""
    total = len(rows) if total is None else total
    rows = rows[:MAX_EMAIL_ROWS]
    th = "".join(f'<th align="{a}" nowrap style="{_TH}">{_esc(h)}</th>' for h, a in headers)
    aligns = [a for _, a in headers]
    body = []
    for i, r in enumerate(rows):
        bg = f' bgcolor="{_C["zebra"]}"' if i % 2 else ""
        tds = "".join(f'<td align="{a}" nowrap>{c}</td>' for a, c in zip(aligns, r))
        body.append(f"<tr{bg}>{tds}</tr>")
    more = ""
    if total > len(rows):
        more = (f'<div style="font-size:12px;color:{_C["muted"]};margin-top:4px">'
                f'Showing top {len(rows)} of {total} — see the attached spreadsheet for all.</div>')
    return (f'<div style="overflow-x:auto"><table cellpadding="7" cellspacing="0" width="100%" '
            f'style="border-collapse:collapse;border:1px solid {_C["line"]};font-size:13px">'
            f"<tr>{th}</tr>{''.join(body)}</table></div>{more}")


def _empty(text):
    return (f'<div style="padding:12px 14px;border:1px dashed {_C["line"]};border-radius:6px;'
            f'color:{_C["muted"]};font-size:13px">{_esc(text)}</div>')


def _breakout_table(full_df, symbols, side):
    """Table of monthly breakout rows. side = 'high' | 'low'."""
    if full_df is None or not symbols:
        return _empty("None")
    df = full_df[full_df["Symbol"].isin(symbols)].copy()
    lvl, day, pct = (("Prev_Month_High", "Today_High", "%_From_Month_High") if side == "high"
                     else ("Prev_Month_Low", "Today_Low", "%_From_Month_Low"))
    df = df.sort_values(pct, ascending=(side == "low"))
    held_cell = (f'<b style="color:{_C["green"] if side == "high" else _C["red"]}">'
                 f'✔ {"Above" if side == "high" else "Below"}</b>')
    back_cell = f'<span style="color:{_C["muted"]}">Back inside</span>'
    rows = []
    for _, r in df.iterrows():
        held = (r["LTP"] > r[lvl]) if side == "high" else (r["LTP"] < r[lvl])
        rows.append([
            f'<b>{_esc(r["Symbol"])}</b>',
            _num(r["LTP"]),
            _num(r[lvl]),
            _num(r[day]),
            _pct(r[pct]),
            held_cell if held else back_cell,
        ])
    label = "High" if side == "high" else "Low"
    return _table(
        [("Symbol", "left"), ("LTP", "right"), (f"Prev Month {label}", "right"),
         (f"Today {label}", "right"), ("LTP vs Level", "right"), ("Close", "left")],
        rows,
    )


def split_52w_by_fno(lows_df):
    """Returns (fno_df, cash_only_df), each keeping the deepest-break-first order."""
    if lows_df is None or lows_df.empty or "In_FnO" not in lows_df:
        empty = lows_df if lows_df is not None else pd.DataFrame()
        return empty, empty
    is_fno = lows_df["In_FnO"] == "YES"
    return lows_df[is_fno].reset_index(drop=True), lows_df[~is_fno].reset_index(drop=True)


def _w52_table(lows_df, new_52w, show_lot=False):
    yes_cell = (f'<b style="background:{_C["red_bg"]};color:{_C["red"]};padding:2px 7px;'
                f'border-radius:3px;font-size:11px">YES</b>')
    wick_cell = f'<span style="color:{_C["muted"]}">Wick only</span>'
    new_b = _badge("NEW", _C["amber"], _C["amber_bg"])
    rows = []
    for i, (_, r) in enumerate(lows_df.head(MAX_EMAIL_ROWS).iterrows(), 1):
        company = _esc(r.get("Company", ""))
        row = [
            i,
            f'<b>{_esc(r["Symbol"])}</b>'
            + (new_b if r["Symbol"] in new_52w else "")
            + (_sub(company[:28]) if company else ""),
            _num(r["LTP"]),
            _num(r["Today_Low"]),
            _num(r["Prev_52W_Low"]) + _sub(_esc(r["Prev_52W_Low_Date"])),
            _pct(r["%_Low_vs_Prev_52W_Low"]),
            _pct(r["%_LTP_From_52W_High"]),
            yes_cell if r.get("Closed_Below_Prev_52W_Low") == "YES" else wick_cell,
        ]
        if show_lot:
            lot = r.get("Lot_Size")
            row.insert(3, _num(lot, 0) if lot is not None and pd.notna(lot) else
                       f'<span style="color:{_C["muted"]}">—</span>')
        rows.append(row)
    headers = [("#", "right"), ("Symbol", "left"), ("LTP", "right"), ("Day Low", "right"),
               ("Prev 52W Low", "right"), ("Break %", "right"), ("From 52W High", "right"),
               ("Closed Below", "left")]
    if show_lot:
        headers.insert(3, ("Lot Size", "right"))
    return _table(headers, rows, total=len(lows_df))


def _sub_title(text, count, fg, bg):
    return (f'<div style="margin:18px 0 6px;font-size:14px;font-weight:700;color:{_C["ink"]}">'
            f'<b style="padding:2px 8px;border-radius:4px;background:{bg};color:{fg};'
            f'font-size:12px">{_esc(text)}</b> '
            f'<span style="color:{_C["muted"]};font-weight:400">({count})</span></div>')


def _ha_section(ha_signals, watch_n, note, min_red, recent_sessions):
    """Heikin Ashi reversal block: first doji/green HA candle after a red
    streak, on stocks whose 52W low was made in the last N sessions."""
    from ha_reversal_screener import _signal_table  # lazy: avoids circular import
    n = 0 if ha_signals is None else len(ha_signals)
    out = [_section_title("Heikin Ashi Reversals — 52-Week-Low Stocks", n, _C["green"])]
    if note:
        out.append(_empty(note))
        return "".join(out)
    out.append(
        f'<div style="font-size:12px;color:{_C["muted"]};margin-bottom:6px">'
        f'Watching {watch_n} NIFTY 500 stocks whose 52W low was made in the last '
        f'{recent_sessions} sessions. Alert = first <b>DOJI</b> or <b>GREEN</b> daily HA candle '
        f'after at least {min_red} red ones; <b>STRONG</b> = green with no lower wick. '
        f'A heads-up to confirm next session, not an entry.</div>')
    if n == 0:
        out.append(_empty("No 52-week-low stock printed a doji or green HA candle after a "
                          "red streak today."))
        return "".join(out)
    fno_sig, cash_sig = split_52w_by_fno(ha_signals)
    out.append(_sub_title("F&O stocks — futures available", len(fno_sig),
                          _C["blue"], _C["blue_bg"]))
    out.append(_signal_table(fno_sig) if len(fno_sig) else _empty("No F&O reversal today."))
    out.append(_sub_title("Non-F&O stocks — cash market only", len(cash_sig),
                          _C["amber"], _C["amber_bg"]))
    out.append(_signal_table(cash_sig) if len(cash_sig) else _empty("No non-F&O reversal today."))
    return "".join(out)


def build_email_html(today_str, is_first_run, current_high, current_low, new_high, new_low,
                     full_df, lows_df, new_52w, w52_status_note,
                     ha_signals=None, ha_watch_n=0, ha_note=None, ha_min_red=3,
                     ha_recent_sessions=20):
    show_high = current_high if is_first_run else new_high
    show_low = current_low if is_first_run else new_low
    prefix = "" if is_first_run else "New "
    run_note = ("Baseline run — no previous state, showing ALL current breakouts."
                if is_first_run else
                "Monthly section shows only NEW breakouts since the last run.")

    stats = [
        _stat(f"{prefix}Month-High breaks", len(show_high), _C["green"], _C["green_bg"]),
        _stat(f"{prefix}Month-Low breaks", len(show_low), _C["red"], _C["red_bg"]),
    ]
    fno_lows, cash_lows = split_52w_by_fno(lows_df)
    if lows_df is not None:
        stats.append(_stat("52W lows — F&O", len(fno_lows), _C["blue"], _C["blue_bg"]))
        stats.append(_stat("52W lows — non-F&O", len(cash_lows), _C["amber"], _C["amber_bg"]))
    if ha_signals is not None:
        stats.append(_stat("HA reversals", len(ha_signals), _C["green"], _C["green_bg"]))

    parts = [
        f'<div style="{_FONT}color:{_C["ink"]};max-width:900px;margin:0 auto">',
        f'<div style="background:{_C["head"]};color:#fff;padding:18px 20px;border-radius:8px 8px 0 0">'
        f'<div style="font-size:18px;font-weight:700">NSE Breakout, 52-Week Low &amp; '
        f'HA Reversal Screener</div>'
        f'<div style="font-size:13px;opacity:.8;margin-top:2px">{_esc(today_str)}</div></div>',
        f'<div style="border:1px solid {_C["line"]};border-top:0;border-radius:0 0 8px 8px;'
        f'padding:16px 20px">',
        f'<div>{"".join(stats)}</div>',
        f'<div style="font-size:12px;color:{_C["muted"]};margin-top:6px">{_esc(run_note)}</div>',

        _section_title(f"{prefix}Breakouts above previous month HIGH (F&O)", len(show_high),
                       _C["green"]),
        _breakout_table(full_df, show_high, "high"),
        _section_title(f"{prefix}Breakdowns below previous month LOW (F&O)", len(show_low),
                       _C["red"]),
        _breakout_table(full_df, show_low, "low"),
    ]

    if lows_df is not None or w52_status_note:
        n = 0 if lows_df is None else len(lows_df)
        parts.append(_section_title("52-Week Lows — NIFTY 500", n, _C["amber"]))
        if w52_status_note:
            parts.append(_empty(w52_status_note))
        elif n == 0:
            parts.append(_empty("No NIFTY 500 stock made a fresh 52-week low today."))
        else:
            parts.append(
                f'<div style="font-size:12px;color:{_C["muted"]};margin-bottom:6px">'
                f'Sorted by deepest break. <b>Break %</b> = day low vs previous 52W low. '
                f'<b>Closed Below</b> = YES means the break held at close; '
                f'"Wick only" means it recovered above the old low. '
                f'<b>NEW</b> = not on the list at the previous run.</div>'
            )
            parts.append(_sub_title("F&O stocks — futures available", len(fno_lows),
                                    _C["blue"], _C["blue_bg"]))
            parts.append(_w52_table(fno_lows, new_52w, show_lot=True) if len(fno_lows)
                         else _empty("No F&O stock made a fresh 52-week low today."))
            parts.append(_sub_title("Non-F&O stocks — cash market only", len(cash_lows),
                                    _C["amber"], _C["amber_bg"]))
            parts.append(_w52_table(cash_lows, new_52w) if len(cash_lows)
                         else _empty("No non-F&O stock made a fresh 52-week low today."))

    if ha_signals is not None or ha_note:
        parts.append(_ha_section(ha_signals, ha_watch_n, ha_note, ha_min_red, ha_recent_sessions))

    parts.append(
        f'<div style="margin-top:24px;padding-top:12px;border-top:1px solid {_C["line"]};'
        f'font-size:11px;color:{_C["muted"]}">Full spreadsheet (breakouts, 52W lows and '
        f'HA reversal sheets) '
        f'and breakout-state CSV attached. Data: Yahoo Finance EOD — cross-check with your '
        f'broker before trading.</div>'
    )
    parts.append("</div></div>")
    return "".join(parts)


def send_email(
    current_high,
    current_low,
    new_high,
    new_low,
    is_first_run: bool,
    excel_path: Path,
    state_path: Path,
    lows_df=None,
    new_52w=frozenset(),
    w52_status_note=None,
    full_df=None,
    ha_signals=None,
    ha_watch_n=0,
    ha_note=None,
    ha_min_red=3,
    ha_recent_sessions=20,
):
    """Sends the screener result via Gmail SMTP as an HTML email (with a
    plain-text fallback). No-op (with a warning) if the required
    environment variables aren't set."""
    user = os.environ.get("EMAIL_USERNAME")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    recipients = os.environ.get("EMAIL_TO")

    if not (user and app_password and recipients):
        print(
            "[INFO] EMAIL_USERNAME / EMAIL_APP_PASSWORD / EMAIL_TO not all set — "
            "skipping email, spreadsheet is still saved locally.",
            file=sys.stderr,
        )
        return

    today_str = dt.datetime.now().strftime("%d-%b-%Y")
    lines = [f"NSE F&O Monthly High/Low Breakout Screener — {today_str}", ""]

    if is_first_run:
        lines.append(
            "No previous state file found — this is a baseline run. "
            "Showing ALL current breakouts."
        )
        lines.append("")
        lines.append(f"Broke Previous Month HIGH ({len(current_high)}):")
        lines.append(format_symbol_list(current_high))
        lines.append("")
        lines.append(f"Broke Previous Month LOW ({len(current_low)}):")
        lines.append(format_symbol_list(current_low))
        subject_suffix = "Baseline"
    else:
        lines.append(
            "Comparing against the last committed state — showing only "
            "NEW breakouts since the last run."
        )
        lines.append("")
        lines.append(f"NEW breakouts above previous month HIGH ({len(new_high)}):")
        lines.append(format_symbol_list(new_high))
        lines.append("")
        lines.append(f"NEW breakouts below previous month LOW ({len(new_low)}):")
        lines.append(format_symbol_list(new_low))
        subject_suffix = "New Breakouts"

    lines.append("")
    if lows_df is not None or w52_status_note:
        lines += format_52w_section(lows_df, new_52w, w52_status_note)
        lines.append("")
        n_52w = 0 if lows_df is None else len(lows_df)
        subject_suffix += f" | 52W Lows: {n_52w}"
    if ha_signals is not None or ha_note:
        lines.append(f"HEIKIN ASHI REVERSALS ON 52-WEEK-LOW STOCKS "
                     f"({0 if ha_signals is None else len(ha_signals)})")
        if ha_note:
            lines.append(ha_note)
        elif ha_signals is None or ha_signals.empty:
            lines.append("None today.")
        else:
            for _, r in ha_signals.iterrows():
                lines.append(f"{r['Symbol']:<12} {'[F&O] ' if r['In_FnO'] == 'YES' else ''}"
                             f"{r['Signal']} after {r['Red_Streak_Before']} red | Close "
                             f"{r['Close']:,.2f} | 52W low {r['52W_Low']:,.2f} "
                             f"({r['%_Above_52W_Low']:+.2f}%)")
        lines.append("")
        subject_suffix += f" | HA Reversals: {0 if ha_signals is None else len(ha_signals)}"
    lines.append("Full spreadsheet (breakouts, 52W lows, HA reversal sheets) and "
                 "breakout-state CSV attached.")
    body = "\n".join(lines)
    html_body = build_email_html(
        today_str, is_first_run, current_high, current_low, new_high, new_low,
        full_df, lows_df, new_52w, w52_status_note,
        ha_signals, ha_watch_n, ha_note, ha_min_red, ha_recent_sessions,
    )

    msg = MIMEMultipart("mixed")
    msg["From"] = user
    msg["To"] = recipients
    msg["Subject"] = f"NSE Monthly Breakout Screener ({subject_suffix}) — {today_str}"
    alt = MIMEMultipart("alternative")  # mail apps show HTML, fall back to text
    alt.attach(MIMEText(body, "plain", "utf-8"))
    alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt)

    for attach_path in (excel_path, state_path):
        if attach_path and Path(attach_path).exists():
            with open(attach_path, "rb") as f:
                part = MIMEApplication(f.read(), Name=Path(attach_path).name)
            part["Content-Disposition"] = f'attachment; filename="{Path(attach_path).name}"'
            msg.attach(part)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(user, app_password)
        server.sendmail(user, [r.strip() for r in recipients.split(",")], msg.as_string())

    print(f"[INFO] Email sent to {recipients}")


def main():
    parser = argparse.ArgumentParser(
        description="Screen NSE F&O stocks for previous-month high/low breakouts."
    )
    parser.add_argument(
        "--universe", default=None,
        help="Path to a CSV with a column of NSE symbols (no .NS suffix). "
             "Overrides the built-in F&O list.",
    )
    parser.add_argument(
        "--symbols", default=None,
        help="Comma-separated list of symbols to scan instead of the full universe.",
    )
    parser.add_argument("--out", default="output", help="Output directory (default: ./output)")
    parser.add_argument(
        "--state-file", default="state/breakout_state.csv",
        help="Path to the committed state CSV used to detect new breakouts "
             "(default: state/breakout_state.csv)",
    )
    parser.add_argument(
        "--only-breakouts", action="store_true",
        help="Only include rows where the month's high and/or low was broken "
             "in the SPREADSHEET output (email content logic is unaffected).",
    )
    parser.add_argument(
        "--no-email", action="store_true",
        help="Skip sending email even if credentials are configured.",
    )
    parser.add_argument(
        "--no-52w-low", action="store_true",
        help="Skip the NIFTY 500 52-week-low scan.",
    )
    parser.add_argument(
        "--nifty500-file", default=None,
        help="CSV of NIFTY 500 constituents (NSE format, 'Symbol' column) to use "
             "instead of downloading from NSE.",
    )
    parser.add_argument(
        "--nifty500-cache", default="data/nifty500_list.csv",
        help="Where the NIFTY 500 list is kept in the repo (default: data/nifty500_list.csv)",
    )
    parser.add_argument(
        "--nifty500-max-age-days", type=int, default=30,
        help="Re-download the NIFTY 500 list only when the repo copy is at least "
             "this many days old (default: 30)",
    )
    parser.add_argument(
        "--fno-list-cache", default="data/fno_lot_sizes.csv",
        help="Where NSE's official F&O stock list (with lot sizes) is kept in the repo; "
             "refreshed on the same schedule as the NIFTY 500 list "
             "(default: data/fno_lot_sizes.csv)",
    )
    parser.add_argument(
        "--w52-state-file", default="state/52w_low_state.csv",
        help="State CSV used to mark NEW 52-week lows (default: state/52w_low_state.csv)",
    )
    parser.add_argument(
        "--no-ha-reversal", action="store_true",
        help="Skip the Heikin Ashi reversal scan on 52-week-low stocks.",
    )
    parser.add_argument("--ha-recent-sessions", type=int, default=20,
                        help="HA watchlist: 52W low made within this many sessions (default: 20)")
    parser.add_argument("--ha-min-red", type=int, default=3,
                        help="Red HA candles required before the signal candle (default: 3)")
    parser.add_argument("--ha-doji-ratio", type=float, default=0.10,
                        help="Max HA body/range for a doji (default: 0.10)")
    args = parser.parse_args()

    symbols = load_universe(args.universe, args.symbols)
    print(f"Scanning {len(symbols)} symbols for previous-month high/low breakouts...\n")

    full_df = build_screener_table(symbols)
    if full_df.empty:
        print("No data could be fetched for any symbol.", file=sys.stderr)
        sys.exit(1)

    state_path = Path(args.state_file)
    prev_high, prev_low, had_state = load_previous_breakouts(state_path)

    current_high = set(full_df.loc[full_df["Broke_Month_High"] == "YES", "Symbol"])
    current_low = set(full_df.loc[full_df["Broke_Month_Low"] == "YES", "Symbol"])
    new_high = current_high - prev_high
    new_low = current_low - prev_low
    is_first_run = not had_state

    # Spreadsheet output (unaffected by state logic, kept as before)
    output_df = full_df
    if args.only_breakouts:
        output_df = full_df[
            (full_df["Broke_Month_High"] == "YES") | (full_df["Broke_Month_Low"] == "YES")
        ].reset_index(drop=True)

    # ---- NIFTY 500 52-week-low scan (failures here never block the breakout email)
    lows_df, new_52w, w52_note = None, set(), None
    if not args.no_52w_low:
        try:
            n500 = load_nifty500(Path(args.nifty500_cache), args.nifty500_file,
                                 args.nifty500_max_age_days)
            if n500.empty:
                w52_note = "Scan skipped — NIFTY 500 constituent list could not be loaded."
            else:
                print(f"\nScanning {len(n500)} NIFTY 500 symbols for fresh 52-week lows...")
                fno_lots = load_fno_lots(Path(args.fno_list_cache), args.nifty500_max_age_days)
                lows_df = scan_52w_lows(n500, fno_lots)
                if lows_df.empty and not {"Symbol"}.issubset(lows_df.columns):
                    lows_df = pd.DataFrame(columns=["Symbol", "Today_Low", "Prev_52W_Low",
                                                    "Session_Date"])
                w52_state = Path(args.w52_state_file)
                prev_52w, had_52w_state = load_previous_52w(w52_state)
                today_52w = set(lows_df["Symbol"])
                new_52w = today_52w - prev_52w if had_52w_state else today_52w
                save_52w_state(lows_df, w52_state)
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] 52-week-low scan failed: {exc}", file=sys.stderr)
            lows_df, new_52w = None, set()
            w52_note = f"Scan failed this run ({type(exc).__name__}) — see Actions log."

    # ---- Heikin Ashi reversals on 52-week-low stocks (never blocks the email)
    ha_watch, ha_signals, ha_note = None, None, None
    if not args.no_ha_reversal:
        try:
            from ha_reversal_screener import scan as ha_scan  # lazy: avoids circular import
            n500 = load_nifty500(Path(args.nifty500_cache), args.nifty500_file,
                                 args.nifty500_max_age_days)
            if n500.empty:
                ha_note = "Scan skipped — NIFTY 500 constituent list could not be loaded."
            else:
                print(f"\nScanning {len(n500)} NIFTY 500 symbols for Heikin Ashi reversals...")
                fno_lots = load_fno_lots(Path(args.fno_list_cache), args.nifty500_max_age_days)
                ha_watch = ha_scan(n500, fno_lots, args.ha_recent_sessions,
                                   args.ha_min_red, args.ha_doji_ratio)
                if ha_watch.empty:
                    ha_watch = pd.DataFrame(columns=["Symbol", "In_FnO", "Signal"])
                ha_signals = ha_watch[ha_watch["Signal"] != ""].reset_index(drop=True)
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] Heikin Ashi scan failed: {exc}", file=sys.stderr)
            ha_watch, ha_signals = None, None
            ha_note = f"Scan failed this run ({type(exc).__name__}) — see Actions log."

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    out_path = out_dir / f"monthly_breakout_screener_{timestamp}.xlsx"
    with pd.ExcelWriter(out_path) as writer:
        output_df.to_excel(writer, index=False, sheet_name="Monthly_Breakout_Screener")
        if lows_df is not None:
            fno_lows, cash_lows = split_52w_by_fno(lows_df)
            for part, name in ((fno_lows, "52W_Lows_FnO"), (cash_lows, "52W_Lows_NonFnO")):
                sheet = part.copy()
                if "Symbol" in sheet:
                    sheet.insert(1, "New",
                                 sheet["Symbol"].map(lambda s: "NEW" if s in new_52w else ""))
                    sheet = sheet.drop(columns=["In_FnO"], errors="ignore")
                    if name == "52W_Lows_NonFnO":
                        sheet = sheet.drop(columns=["Lot_Size"], errors="ignore")
                sheet.to_excel(writer, index=False, sheet_name=name)
        if ha_watch is not None:
            ha_signals.to_excel(writer, index=False, sheet_name="HA_Reversal_Signals")
            ha_watch.to_excel(writer, index=False, sheet_name="HA_Watchlist_52W_Low")

    print(f"\nDone. {len(output_df)} symbols in spreadsheet output.")
    print(f"  Total breaking Previous-Month High today : {len(current_high)}")
    print(f"  Total breaking Previous-Month Low today  : {len(current_low)}")
    print(f"  NEW High breakouts vs last state         : {len(new_high)}")
    print(f"  NEW Low breakouts vs last state          : {len(new_low)}")
    if lows_df is not None:
        print(f"  NIFTY 500 stocks at fresh 52-week low    : {len(lows_df)} ({len(new_52w)} new)")
    if ha_signals is not None:
        print(f"  HA reversals on 52W-low stocks           : {len(ha_signals)} "
              f"(watchlist {len(ha_watch)})")
    print(f"Saved spreadsheet to: {out_path.resolve()}")

    # Update state file for next run (always overwrite with today's breakouts)
    save_state(full_df, state_path)

    if not args.no_email:
        send_email(
            current_high, current_low, new_high, new_low,
            is_first_run, out_path, state_path,
            lows_df=lows_df, new_52w=new_52w, w52_status_note=w52_note,
            full_df=full_df,
            ha_signals=ha_signals, ha_watch_n=0 if ha_watch is None else len(ha_watch),
            ha_note=ha_note, ha_min_red=args.ha_min_red,
            ha_recent_sessions=args.ha_recent_sessions,
        )


if __name__ == "__main__":
    main()
