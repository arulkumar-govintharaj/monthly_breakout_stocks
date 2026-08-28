#!/usr/bin/env python3
"""
monthly_breakout_screener.py

Screens the NSE F&O stock universe for symbols whose price has broken
the PREVIOUS CALENDAR MONTH's HIGH and/or LOW during the current
session — output as a single Excel file, with an optional email
(stock names in the body, spreadsheet attached).

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
import os
import smtplib
import sys
import time
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


def send_email(
    current_high,
    current_low,
    new_high,
    new_low,
    is_first_run: bool,
    excel_path: Path,
    state_path: Path,
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
    lines.append("Full spreadsheet and current breakout-state CSV attached.")
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

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    out_path = out_dir / f"monthly_breakout_screener_{timestamp}.xlsx"
    output_df.to_excel(out_path, index=False, sheet_name="Monthly_Breakout_Screener")

    print(f"\nDone. {len(output_df)} symbols in spreadsheet output.")
    print(f"  Total breaking Previous-Month High today : {len(current_high)}")
    print(f"  Total breaking Previous-Month Low today  : {len(current_low)}")
    print(f"  NEW High breakouts vs last state         : {len(new_high)}")
    print(f"  NEW Low breakouts vs last state          : {len(new_low)}")
    print(f"Saved spreadsheet to: {out_path.resolve()}")

    # Update state file for next run (always overwrite with today's breakouts)
    save_state(full_df, state_path)

    if not args.no_email:
        send_email(
            current_high, current_low, new_high, new_low,
            is_first_run, out_path, state_path,
        )


if __name__ == "__main__":
    main()
