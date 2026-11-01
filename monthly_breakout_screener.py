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


def load_nifty500(cache_path: Path, override_path: str = None,
                  max_age_days: int = 30) -> pd.DataFrame:
    """
    Returns a DataFrame with columns Symbol, Company, Industry for the
    NIFTY 500. Order of preference:
      1. --nifty500-file override (if given)
      2. Cached list in the repo at `cache_path`, if it was downloaded
         less than `max_age_days` ago (no network call)
      3. Fresh download from NSE (saved to `cache_path` + date stamp)
      4. Stale cached copy, if the download fails
    Returns an empty DataFrame if none of these work.
    """
    def _normalise(df: pd.DataFrame) -> pd.DataFrame:
        df.columns = [c.strip() for c in df.columns]
        if "Symbol" not in df.columns:
            df = df.rename(columns={df.columns[0]: "Symbol"})
        out = pd.DataFrame({"Symbol": df["Symbol"].astype(str).str.strip().str.upper()})
        out["Company"] = df["Company Name"].astype(str).str.strip() if "Company Name" in df else ""
        out["Industry"] = df["Industry"].astype(str).str.strip() if "Industry" in df else ""
        return out[out["Symbol"].ne("") & out["Symbol"].ne("NAN")].drop_duplicates("Symbol")

    if override_path:
        p = Path(override_path)
        if not p.exists():
            print(f"[ERROR] NIFTY 500 file not found: {p}", file=sys.stderr)
            sys.exit(1)
        return _normalise(pd.read_csv(p))

    age = _cache_age_days(cache_path)
    if cache_path.exists() and age is not None and age < max_age_days:
        print(f"[INFO] Using NIFTY 500 list from repo ({cache_path}, {age} days old; "
              f"refreshes after {max_age_days} days)")
        return _normalise(pd.read_csv(cache_path))
    if not cache_path.exists():
        reason = "not in repo yet"
    elif age is None:
        reason = "has no download-date stamp"
    else:
        reason = f"is {age} days old"
    print(f"[INFO] NIFTY 500 list {reason} — downloading fresh copy from NSE")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Accept": "text/csv,*/*",
    }
    for url in NIFTY500_URLS:
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as resp:
                text = resp.read().decode("utf-8", errors="replace")
            df = pd.read_csv(io.StringIO(text))
            if len(df) < NIFTY500_MIN_ROWS or "Symbol" not in [c.strip() for c in df.columns]:
                raise ValueError(f"unexpected content ({len(df)} rows)")
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(cache_path, index=False)
            _cache_stamp_path(cache_path).write_text(dt.date.today().isoformat() + "\n")
            print(f"[INFO] NIFTY 500 list downloaded from {url} ({len(df)} symbols), "
                  f"saved to {cache_path}")
            return _normalise(df)
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] Could not download NIFTY 500 list from {url}: {exc}", file=sys.stderr)

    if cache_path.exists():
        print(f"[WARN] Download failed — using existing (stale) list at {cache_path}",
              file=sys.stderr)
        return _normalise(pd.read_csv(cache_path))

    print("[WARN] No NIFTY 500 list available — 52-week-low scan skipped.", file=sys.stderr)
    return pd.DataFrame(columns=["Symbol", "Company", "Industry"])


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
    """
    meta = universe.set_index("Symbol")
    symbols = list(meta.index)
    fno = set(fno_symbols)
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
    lines.append("[NEW] = not on the list at the previous run, [F&O] = in F&O segment")
    lines.append("")
    for _, r in lows_df.iterrows():
        tags = " ".join(t for t, on in (("[NEW]", r["Symbol"] in new_52w),
                                         ("[F&O]", r["In_FnO"] == "YES")) if on)
        lines.append(
            f"{r['Symbol']:<12} {tags}".rstrip()
            + f"\n    LTP {r['LTP']:,.2f} | Low {r['Today_Low']:,.2f} vs prev 52W low "
              f"{r['Prev_52W_Low']:,.2f} on {r['Prev_52W_Low_Date']} "
              f"({r['%_Low_vs_Prev_52W_Low']:+.2f}%) | "
              f"{r['%_LTP_From_52W_High']:+.1f}% from 52W high"
            + (" | closed below" if r["Closed_Below_Prev_52W_Low"] == "YES" else "")
        )
    return lines


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
):
    """Sends the screener result via Gmail SMTP. No-op (with a warning)
    if the required environment variables aren't set."""
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
    lines.append("Full spreadsheet (breakouts + 52W lows sheets) and breakout-state CSV attached.")
    body = "\n".join(lines)

    msg = MIMEMultipart()
    msg["From"] = user
    msg["To"] = recipients
    msg["Subject"] = f"NSE Monthly Breakout Screener ({subject_suffix}) — {today_str}"
    msg.attach(MIMEText(body, "plain"))

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
        "--w52-state-file", default="state/52w_low_state.csv",
        help="State CSV used to mark NEW 52-week lows (default: state/52w_low_state.csv)",
    )
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
                lows_df = scan_52w_lows(n500, DEFAULT_FNO_UNIVERSE)
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

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    out_path = out_dir / f"monthly_breakout_screener_{timestamp}.xlsx"
    with pd.ExcelWriter(out_path) as writer:
        output_df.to_excel(writer, index=False, sheet_name="Monthly_Breakout_Screener")
        if lows_df is not None:
            sheet = lows_df.copy()
            if "Symbol" in sheet:
                sheet.insert(1, "New", sheet["Symbol"].map(lambda s: "NEW" if s in new_52w else ""))
            sheet.to_excel(writer, index=False, sheet_name="52W_Lows_NIFTY500")

    print(f"\nDone. {len(output_df)} symbols in spreadsheet output.")
    print(f"  Total breaking Previous-Month High today : {len(current_high)}")
    print(f"  Total breaking Previous-Month Low today  : {len(current_low)}")
    print(f"  NEW High breakouts vs last state         : {len(new_high)}")
    print(f"  NEW Low breakouts vs last state          : {len(new_low)}")
    if lows_df is not None:
        print(f"  NIFTY 500 stocks at fresh 52-week low    : {len(lows_df)} ({len(new_52w)} new)")
    print(f"Saved spreadsheet to: {out_path.resolve()}")

    # Update state file for next run (always overwrite with today's breakouts)
    save_state(full_df, state_path)

    if not args.no_email:
        send_email(
            current_high, current_low, new_high, new_low,
            is_first_run, out_path, state_path,
            lows_df=lows_df, new_52w=new_52w, w52_status_note=w52_note,
        )


if __name__ == "__main__":
    main()
