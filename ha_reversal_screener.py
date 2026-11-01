#!/usr/bin/env python3
"""
ha_reversal_screener.py

Daily Heikin Ashi reversal alert for NIFTY 500 stocks sitting at / near
their 52-week low.

WATCHLIST
---------
A NIFTY 500 stock is on the watchlist when its lowest Low of the last
52 weeks was made within the last --recent-sessions sessions
(default 20, i.e. roughly the last month). These are the "52-week low"
stocks we want to catch turning.

SIGNAL (alert on change, not on state)
--------------------------------------
Daily OHLC is converted to Heikin Ashi:
    HA Close = (O + H + L + C) / 4
    HA Open  = (prev HA Open + prev HA Close) / 2
    HA High  = max(H, HA Open, HA Close)
    HA Low   = min(L, HA Open, HA Close)

The latest HA candle triggers an alert only if the --min-red candles
immediately before it were ALL red (HA Close < HA Open, not doji), and
the latest candle is one of:

    DOJI             body |HA Close - HA Open| <= --doji-ratio x range
                     (default 10% of HA High - HA Low)
    GREEN            HA Close > HA Open
    GREEN (STRONG)   green with (almost) no lower wick — the classic
                     Heikin Ashi trend-turn candle

Because the red streak must be unbroken right up to the latest candle,
a stock alerts once, on the first doji/green after the fall — it does
not keep alerting every day while it stays green.

Output: one Excel file (Signals + full Watchlist sheets) and, when there
is at least one signal, an HTML email via the same Gmail SMTP secrets as
monthly_breakout_screener.py (EMAIL_USERNAME, EMAIL_APP_PASSWORD,
EMAIL_TO). Use --always-email to also get a "no signals today" email.

Usage:
    python ha_reversal_screener.py
    python ha_reversal_screener.py --no-email
    python ha_reversal_screener.py --symbols SUZLON,INTELLECT --recent-sessions 60
    python ha_reversal_screener.py --min-red 4 --doji-ratio 0.08
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

import numpy as np
import pandas as pd
import yfinance as yf

from monthly_breakout_screener import (
    _C, _FONT, _badge, _empty, _esc, _extract_ticker_frame, _num, _pct, _section_title,
    _stat, _sub, _table, load_fno_lots, load_nifty500,
)

# Lower wick counts as "none" if it's within this fraction of the HA range
STRONG_WICK_TOLERANCE = 0.02


# --------------------------------------------------------------------------
# Heikin Ashi
# --------------------------------------------------------------------------
def heikin_ashi(df: pd.DataFrame) -> pd.DataFrame:
    """Daily OHLC (oldest -> newest) -> DataFrame of HA_Open/High/Low/Close."""
    o, h, l, c = (df[k].to_numpy(dtype=float) for k in ("Open", "High", "Low", "Close"))
    ha_close = (o + h + l + c) / 4.0
    ha_open = np.empty_like(ha_close)
    ha_open[0] = (o[0] + c[0]) / 2.0
    for i in range(1, len(ha_close)):
        ha_open[i] = (ha_open[i - 1] + ha_close[i - 1]) / 2.0
    ha_high = np.maximum.reduce([h, ha_open, ha_close])
    ha_low = np.minimum.reduce([l, ha_open, ha_close])
    return pd.DataFrame(
        {"HA_Open": ha_open, "HA_High": ha_high, "HA_Low": ha_low, "HA_Close": ha_close},
        index=df.index,
    )


def classify_candle(ha_row, doji_ratio: float) -> str:
    """'DOJI', 'GREEN', 'GREEN (STRONG)' or 'RED' for one HA candle."""
    body = ha_row["HA_Close"] - ha_row["HA_Open"]
    rng = ha_row["HA_High"] - ha_row["HA_Low"]
    if rng <= 0 or abs(body) <= doji_ratio * rng:
        return "DOJI"
    if body > 0:
        lower_wick = min(ha_row["HA_Open"], ha_row["HA_Close"]) - ha_row["HA_Low"]
        return "GREEN (STRONG)" if lower_wick <= STRONG_WICK_TOLERANCE * rng else "GREEN"
    return "RED"


def analyse_stock(df: pd.DataFrame, recent_sessions: int, min_red: int,
                  doji_ratio: float, min_history_weeks: int = 50):
    """
    One stock's daily OHLC -> dict of watchlist metrics + signal, or None
    if it isn't near its 52-week low (or lacks history).
    """
    if df is None or df.empty:
        return None
    df = df.dropna(subset=["Open", "High", "Low", "Close"]).sort_index()
    if len(df) < 60:
        return None

    last_date = df.index[-1]
    window = df[df.index > last_date - pd.Timedelta(weeks=52)]
    if df.index[0] > last_date - pd.Timedelta(weeks=min_history_weeks):
        return None  # recent listing — no meaningful 52-week low

    low_date = window["Low"].idxmin()
    low_52w = float(window["Low"].min())
    sessions_since = int((window.index > low_date).sum())
    if sessions_since > recent_sessions:
        return None  # 52W low is too old — not a "52-week low stock" any more

    ha = heikin_ashi(df)  # full history so HA Open has settled
    kinds = [classify_candle(r, doji_ratio) for _, r in ha.iloc[-(min_red + 120):].iterrows()]
    latest = kinds[-1]
    red_streak = 0
    for k in reversed(kinds[:-1]):
        if k != "RED":
            break
        red_streak += 1

    last, ha_last = df.iloc[-1], ha.iloc[-1]
    close = float(last["Close"])
    signal = latest if latest != "RED" and red_streak >= min_red else ""
    return {
        "Signal": signal,
        "HA_Candle": latest,
        "Red_Streak_Before": red_streak,
        "Close": round(close, 2),
        "Day_Chg_%": round((close / float(df["Close"].iloc[-2]) - 1) * 100, 2),
        "HA_Open": round(float(ha_last["HA_Open"]), 2),
        "HA_Close": round(float(ha_last["HA_Close"]), 2),
        "HA_High": round(float(ha_last["HA_High"]), 2),
        "HA_Low": round(float(ha_last["HA_Low"]), 2),
        "52W_Low": round(low_52w, 2),
        "52W_Low_Date": pd.Timestamp(low_date).strftime("%Y-%m-%d"),
        "Sessions_Since_52W_Low": sessions_since,
        "%_Above_52W_Low": round((close - low_52w) / low_52w * 100, 2),
        "Session_Date": pd.Timestamp(last_date).strftime("%Y-%m-%d"),
    }


# --------------------------------------------------------------------------
# Scan
# --------------------------------------------------------------------------
def scan(universe: pd.DataFrame, fno: dict, recent_sessions: int, min_red: int,
         doji_ratio: float, chunk_size: int = 100) -> pd.DataFrame:
    """Returns the full watchlist (one row per stock near its 52W low)."""
    meta = universe.set_index("Symbol")
    symbols = list(meta.index)
    rows = []
    for start in range(0, len(symbols), chunk_size):
        chunk = symbols[start:start + chunk_size]
        tickers = [f"{s}.NS" for s in chunk]
        print(f"[HA] Downloading {start + 1}-{start + len(chunk)} of {len(symbols)}...")
        try:
            data = yf.download(tickers, period="15mo", interval="1d", auto_adjust=False,
                               group_by="ticker", threads=True, progress=False)
        except Exception as exc:  # noqa: BLE001
            print(f"  [WARN] batch download failed: {exc}", file=sys.stderr)
            continue
        for sym, ticker in zip(chunk, tickers):
            try:
                res = analyse_stock(_extract_ticker_frame(data, ticker),
                                    recent_sessions, min_red, doji_ratio)
            except Exception as exc:  # noqa: BLE001
                print(f"  [WARN] {sym}: HA calc failed ({exc})", file=sys.stderr)
                continue
            if res is None:
                continue
            rows.append({
                "Symbol": sym,
                "Company": meta.at[sym, "Company"] if "Company" in meta else "",
                "Industry": meta.at[sym, "Industry"] if "Industry" in meta else "",
                "In_FnO": "YES" if sym in fno else "",
                "Lot_Size": fno.get(sym),
                **res,
            })
        time.sleep(1)

    wl = pd.DataFrame(rows)
    if wl.empty:
        return wl
    latest_session = wl["Session_Date"].mode().iloc[0]
    stale = wl["Session_Date"] != latest_session
    if stale.any():
        print(f"[HA] Ignoring {int(stale.sum())} symbols with stale data (not {latest_session})")
    wl = wl[~stale]
    order = {"GREEN (STRONG)": 0, "GREEN": 1, "DOJI": 2, "": 3}
    wl = wl.assign(_o=wl["Signal"].map(order)).sort_values(
        ["_o", "Red_Streak_Before", "%_Above_52W_Low"], ascending=[True, False, True]
    ).drop(columns="_o").reset_index(drop=True)
    return wl


# --------------------------------------------------------------------------
# Email
# --------------------------------------------------------------------------
_SIG_STYLE = {
    "GREEN (STRONG)": (_C["green"], _C["green_bg"]),
    "GREEN": (_C["green"], _C["green_bg"]),
    "DOJI": (_C["amber"], _C["amber_bg"]),
}


def _signal_table(df):
    rows = []
    for _, r in df.iterrows():
        fg, bg = _SIG_STYLE[r["Signal"]]
        lot = f"Lot {int(r['Lot_Size'])}" if pd.notna(r.get("Lot_Size")) else ""
        rows.append([
            f"<b>{_esc(r['Symbol'])}</b>" + (_badge("F&amp;O", _C["blue"], _C["blue_bg"])
                                             if r["In_FnO"] == "YES" else "")
            + _sub(_esc(r["Company"]) + (f" · {lot}" if lot else "")),
            _badge(_esc(r["Signal"]), fg, bg).strip(),
            f"{int(r['Red_Streak_Before'])} red",
            _num(r["Close"]) + _sub(_pct(r["Day_Chg_%"])),
            _num(r["52W_Low"]) + _sub(_esc(r["52W_Low_Date"])),
            _pct(r["%_Above_52W_Low"], colour=False),
        ])
    return _table([("Stock", "left"), ("HA Signal", "left"), ("Before", "right"),
                   ("Close", "right"), ("52W Low", "right"), ("Above Low", "right")], rows)


def build_html(today_str, signals, watchlist, args):
    n_strong = int((signals["Signal"] == "GREEN (STRONG)").sum()) if len(signals) else 0
    n_green = int((signals["Signal"] == "GREEN").sum()) if len(signals) else 0
    n_doji = int((signals["Signal"] == "DOJI").sum()) if len(signals) else 0
    parts = [
        f'<div style="{_FONT}color:{_C["ink"]};max-width:820px">',
        f'<div style="background:{_C["head"]};color:#fff;padding:16px 18px;border-radius:8px 8px 0 0">'
        f'<div style="font-size:18px;font-weight:700">Heikin Ashi Reversal — 52-Week Low Stocks</div>'
        f'<div style="font-size:12px;opacity:.8">{_esc(today_str)} · daily candles · NIFTY 500</div>'
        f'</div><div style="padding:14px 4px">',
        _stat("Watchlist", len(watchlist), _C["ink"], _C["zebra"]),
        _stat("Green strong", n_strong, _C["green"], _C["green_bg"]),
        _stat("Green", n_green, _C["green"], _C["green_bg"]),
        _stat("Doji", n_doji, _C["amber"], _C["amber_bg"]),
        _section_title("Reversal signals today", len(signals), _C["green"]),
    ]
    if len(signals):
        parts.append(
            f'<div style="font-size:12px;color:{_C["muted"]};margin-bottom:6px">'
            f'First doji / green HA candle after at least {args.min_red} red ones. '
            f'<b>STRONG</b> = green with no lower wick. '
            f'Watchlist = 52W low made in the last {args.recent_sessions} sessions.</div>')
        parts.append(_signal_table(signals))
    else:
        parts.append(_empty("No 52-week-low stock printed a doji or green HA candle "
                            "after a red streak today."))
    parts.append(
        f'<div style="margin-top:22px;font-size:11px;color:{_C["muted"]}">Full watchlist with each '
        f'stock\'s latest HA candle is in the attached spreadsheet. A reversal candle is a heads-up, '
        f'not an entry — confirm with the next session and your own levels.</div></div></div>')
    return "".join(parts)


def send_email(today_str, signals, watchlist, excel_path, args):
    user = os.environ.get("EMAIL_USERNAME")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    recipients = os.environ.get("EMAIL_TO")
    if not (user and app_password and recipients):
        print("[INFO] Email credentials not set — skipping email.", file=sys.stderr)
        return

    text = [f"Heikin Ashi Reversal — 52-Week Low Stocks — {today_str}", ""]
    if len(signals):
        for _, r in signals.iterrows():
            text.append(f"{r['Symbol']:<12} {r['Signal']:<15} after {r['Red_Streak_Before']} red | "
                        f"Close {r['Close']:,.2f} | 52W low {r['52W_Low']:,.2f} "
                        f"({r['%_Above_52W_Low']:+.2f}%)")
    else:
        text.append("No signals today.")
    text += ["", f"Watchlist size: {len(watchlist)} (full list in attached spreadsheet)."]

    msg = MIMEMultipart("mixed")
    msg["From"] = user
    msg["To"] = recipients
    msg["Subject"] = f"HA Reversal on 52W Lows ({len(signals)} signals) — {today_str}"
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText("\n".join(text), "plain", "utf-8"))
    alt.attach(MIMEText(build_html(today_str, signals, watchlist, args), "html", "utf-8"))
    msg.attach(alt)
    with open(excel_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=excel_path.name)
    part["Content-Disposition"] = f'attachment; filename="{excel_path.name}"'
    msg.attach(part)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(user, app_password)
        server.sendmail(user, [r.strip() for r in recipients.split(",")], msg.as_string())
    print(f"[INFO] Email sent to {recipients}")


# --------------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="Heikin Ashi reversal alerts for 52-week-low stocks.")
    p.add_argument("--symbols", default=None,
                   help="Comma-separated symbols to scan instead of the NIFTY 500.")
    p.add_argument("--recent-sessions", type=int, default=20,
                   help="Stock is on the watchlist if its 52W low was made within this many "
                        "sessions (default: 20)")
    p.add_argument("--min-red", type=int, default=3,
                   help="Red HA candles required right before the signal candle (default: 3)")
    p.add_argument("--doji-ratio", type=float, default=0.10,
                   help="Max body/range for a doji (default: 0.10)")
    p.add_argument("--out", default="output", help="Output directory (default: ./output)")
    p.add_argument("--nifty500-cache", default="data/nifty500_list.csv")
    p.add_argument("--fno-list-cache", default="data/fno_lot_sizes.csv")
    p.add_argument("--no-email", action="store_true", help="Don't send email.")
    p.add_argument("--always-email", action="store_true",
                   help="Send the email even when there are no signals.")
    args = p.parse_args()

    if args.symbols:
        universe = pd.DataFrame({"Symbol": [s.strip().upper() for s in args.symbols.split(",")
                                            if s.strip()]})
        universe["Company"], universe["Industry"] = "", ""
    else:
        universe = load_nifty500(Path(args.nifty500_cache))
        if universe.empty:
            print("[ERROR] NIFTY 500 list unavailable.", file=sys.stderr)
            sys.exit(1)
    fno = load_fno_lots(Path(args.fno_list_cache))

    print(f"Scanning {len(universe)} symbols for Heikin Ashi reversals near 52-week lows...")
    watchlist = scan(universe, fno, args.recent_sessions, args.min_red, args.doji_ratio)
    if watchlist.empty:
        watchlist = pd.DataFrame(columns=["Symbol", "Signal"])
    signals = watchlist[watchlist["Signal"] != ""].reset_index(drop=True)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"ha_reversal_52w_low_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
    with pd.ExcelWriter(out_path) as w:
        signals.to_excel(w, index=False, sheet_name="Signals")
        watchlist.to_excel(w, index=False, sheet_name="Watchlist")

    print(f"\nWatchlist (52W low in last {args.recent_sessions} sessions): {len(watchlist)}")
    print(f"Reversal signals today: {len(signals)}")
    for _, r in signals.iterrows():
        print(f"  {r['Symbol']:<12} {r['Signal']:<15} after {r['Red_Streak_Before']} red")
    print(f"Saved spreadsheet to: {out_path.resolve()}")

    if not args.no_email and (len(signals) or args.always_email):
        send_email(dt.datetime.now().strftime("%d-%b-%Y"), signals, watchlist, out_path, args)
    elif not args.no_email:
        print("[INFO] No signals — email not sent (use --always-email to send anyway).")


if __name__ == "__main__":
    main()
