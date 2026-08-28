#!/usr/bin/env python3
"""
monthly_breakout_screener.py

Screens the NSE F&O stock universe for symbols whose price has broken
the PREVIOUS CALENDAR MONTH's HIGH and/or LOW during the current
session — output as a single Excel file, with an optional email
(stock names in the body, spreadsheet attached).

Usage:
    python monthly_breakout_screener.py
    python monthly_breakout_screener.py --symbols RELIANCE,SBIN,SONACOMS
    python monthly_breakout_screener.py --no-email
    python monthly_breakout_screener.py --only-breakouts

Email is sent via Gmail SMTP using credentials from environment
variables (see README for GitHub Actions secrets setup):
    EMAIL_USERNAME   - sender Gmail address
    EMAIL_APP_PASSWORD - Gmail App Password (not your normal password)
    EMAIL_TO         - comma-separated recipient address(es)
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


def send_email(df: pd.DataFrame, excel_path: Path):
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

    high_breaks = df[df["Broke_Month_High"] == "YES"]["Symbol"].tolist()
    low_breaks = df[df["Broke_Month_Low"] == "YES"]["Symbol"].tolist()
    today_str = dt.datetime.now().strftime("%d-%b-%Y")

    lines = [f"NSE F&O Monthly High/Low Breakout Screener — {today_str}", ""]
    lines.append(f"Broke Previous Month HIGH ({len(high_breaks)}):")
    lines.append(", ".join(high_breaks) if high_breaks else "  None")
    lines.append("")
    lines.append(f"Broke Previous Month LOW ({len(low_breaks)}):")
    lines.append(", ".join(low_breaks) if low_breaks else "  None")
    lines.append("")
    lines.append("Full spreadsheet attached.")
    body = "\n".join(lines)

    msg = MIMEMultipart()
    msg["From"] = user
    msg["To"] = recipients
    msg["Subject"] = f"NSE Monthly Breakout Screener — {today_str}"
    msg.attach(MIMEText(body, "plain"))

    with open(excel_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=excel_path.name)
    part["Content-Disposition"] = f'attachment; filename="{excel_path.name}"'
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
        "--only-breakouts", action="store_true",
        help="Only include rows where the month's high and/or low was broken "
             "(in the spreadsheet too, not just the email summary).",
    )
    parser.add_argument(
        "--no-email", action="store_true",
        help="Skip sending email even if credentials are configured.",
    )
    args = parser.parse_args()

    symbols = load_universe(args.universe, args.symbols)
    print(f"Scanning {len(symbols)} symbols for previous-month high/low breakouts...\n")

    df = build_screener_table(symbols)
    if df.empty:
        print("No data could be fetched for any symbol.", file=sys.stderr)
        sys.exit(1)

    if args.only_breakouts:
        df = df[
            (df["Broke_Month_High"] == "YES") | (df["Broke_Month_Low"] == "YES")
        ].reset_index(drop=True)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    out_path = out_dir / f"monthly_breakout_screener_{timestamp}.xlsx"
    df.to_excel(out_path, index=False, sheet_name="Monthly_Breakout_Screener")

    n_high = (df["Broke_Month_High"] == "YES").sum()
    n_low = (df["Broke_Month_Low"] == "YES").sum()
    print(f"\nDone. {len(df)} symbols in output.")
    print(f"  Broke Previous-Month High : {n_high}")
    print(f"  Broke Previous-Month Low  : {n_low}")
    print(f"Saved to: {out_path.resolve()}")

    if not args.no_email:
        send_email(df, out_path)


if __name__ == "__main__":
    main()
