#!/usr/bin/env python3
"""
ato_trader.py — places Zerodha ATO (Alert Triggers Order) trades for the F&O
stocks in the "Heikin Ashi Reversals — 52-Week-Low Stocks" section.

Runs on a machine with a STATIC IP registered on your Kite Connect app
(required for API orders since 1 April 2026). The 5 pm GitHub Actions scan
commits state/ha_fno_signals.csv; this script picks it up from there.

Commands (see trader/crontab.example for the schedule, all times IST):

  place    17:30 Mon-Fri  git pull, read today's F&O HA signals and create one
                          entry ATO per stock:
                            alert : NSE:<SYM> LastTradedPrice >= previous day's
                                    close + 0.5% (normal candle, not HA). "Previous
                                    day" = last completed daily candle: today's
                                    close after 15:30, else the prior session's
                            order : BUY 1 lot of the futures contract, NRML, MARKET
                                    (next month; if the order can fill on the expiry
                                    day or the day before, the month after that)
  manage   every 5 min    once an entry ATO has fired and the futures position
           09:00-15:55    has grown by the order quantity, create two exit ATOs:
                            stop-loss : NSE:<SYM> LTP <= signal-day HA Low
                            target    : NSE:<SYM> LTP >= last swing high before
                                        the 52-week low (from the scan)
                            order     : SELL the same futures quantity, NRML, MARKET
                          When the position is closed (SL, target or by you) the
                          other exit ATO is deleted.
  place --dry-run creates nothing but still emails the preview ([DRY RUN] subject).
  place skips a stock if you already hold any of its futures, if the signal has
  no target, if the target is closer than the stop (reward < risk), or if the
  price is already at/above the entry trigger (Kite rejects such alerts).
  Manual next-morning rerun after a failed 5:30 pm run:
      trader/run.sh place --session <signal date YYYY-MM-DD> --force
  expire   manual only    delete entry ATOs created before today that have not
                          triggered (entry ATOs otherwise stay open until they fire).
  status   any time       print the trades this script is tracking.
  login    any time       just log in (checks credentials / TOTP).

Add --dry-run to any command to print what it would do without creating or
deleting anything. Configuration comes from environment variables (see
trader/.env.example). The ATO alerts are placed on the stock's cash price
because the HA candles are built from it; orders go to the futures contract.
"""

import argparse
import datetime as dt
import html
import json
import math
import os
import re
import smtplib
import subprocess
import sys
import urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests

IST = ZoneInfo("Asia/Kolkata")
API = "https://api.kite.trade"
TAG = "HA52"                      # prefix on every alert this script creates
ENTRY_BUFFER = 0.005              # entry trigger = previous day close + 0.5%

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
DEFAULT_SIGNALS = REPO / "state" / "ha_fno_signals.csv"
DEFAULT_HOME = Path(os.environ.get("HA_TRADER_HOME", Path.home() / ".ha_trader"))


def now_ist() -> dt.datetime:
    return dt.datetime.now(IST)


def log(msg: str):
    print(f"{now_ist():%Y-%m-%d %H:%M:%S} {msg}", flush=True)


# --------------------------------------------------------------------------
# Config / local state
# --------------------------------------------------------------------------
def env(name: str, required: bool = True, default: str = "") -> str:
    val = os.environ.get(name, default).strip()
    if required and not val:
        sys.exit(f"[ERROR] Environment variable {name} is not set (see trader/.env.example).")
    return val


class State:
    """trades.json in HA_TRADER_HOME — what this script has placed and its status."""

    def __init__(self, home: Path):
        self.path = home / "trades.json"
        home.mkdir(parents=True, exist_ok=True)
        self.data = (json.loads(self.path.read_text()) if self.path.exists()
                     else {"processed_sessions": [], "trades": []})

    def save(self):
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.data, indent=2, default=str))
        tmp.replace(self.path)

    @property
    def trades(self):
        return self.data["trades"]

    def open_symbols(self):
        return {t["symbol"] for t in self.trades
                if t["status"] in ("entry_open", "entry_triggered", "sl_open")}


# --------------------------------------------------------------------------
# Login (automated TOTP) — token cached for the day
# --------------------------------------------------------------------------
def _request_token(api_key: str, user_id: str, password: str, totp_secret: str) -> str:
    import pyotp

    s = requests.Session()
    r = s.post("https://kite.zerodha.com/api/login",
               data={"user_id": user_id, "password": password}, timeout=20).json()
    if r.get("status") != "success":
        raise RuntimeError(f"Kite login failed: {r.get('message')}")
    request_id = r["data"]["request_id"]
    r = s.post("https://kite.zerodha.com/api/twofa",
               data={"user_id": user_id, "request_id": request_id,
                     "twofa_value": pyotp.TOTP(totp_secret).now(), "twofa_type": "totp"},
               timeout=20).json()
    if r.get("status") != "success":
        raise RuntimeError(f"Kite TOTP step failed: {r.get('message')}")

    url = f"https://kite.trade/connect/login?api_key={api_key}&v=3"
    for _ in range(10):  # follow redirects by hand until request_token shows up
        resp = s.get(url, allow_redirects=False, timeout=20)
        loc = resp.headers.get("Location", "")
        if "request_token=" in loc:
            return urllib.parse.parse_qs(urllib.parse.urlparse(loc).query)["request_token"][0]
        if not loc:
            break
        url = urllib.parse.urljoin(url, loc)
    raise RuntimeError("Could not obtain request_token — has the Kite Connect app been "
                       "authorised once manually, and is its redirect URL set?")


def get_kite(home: Path):
    from kiteconnect import KiteConnect

    api_key, api_secret = env("KITE_API_KEY"), env("KITE_API_SECRET")
    kite = KiteConnect(api_key=api_key)
    tok_file = home / "access_token.json"
    today = now_ist().date().isoformat()
    # Kite tokens expire around 06:00 IST the next morning
    if tok_file.exists():
        cached = json.loads(tok_file.read_text())
        if cached.get("date") == today and now_ist().hour >= 6:
            kite.set_access_token(cached["access_token"])
            try:
                kite.profile()
                return kite
            except Exception:  # noqa: BLE001
                log("[INFO] Cached token rejected, logging in again.")
    rt = _request_token(api_key, env("KITE_USER_ID"), env("KITE_PASSWORD"),
                        env("KITE_TOTP_SECRET"))
    sess = kite.generate_session(rt, api_secret=api_secret)
    kite.set_access_token(sess["access_token"])
    tok_file.write_text(json.dumps({"date": today, "access_token": sess["access_token"]}))
    os.chmod(tok_file, 0o600)
    log(f"[INFO] Logged in as {sess.get('user_id')}")
    return kite


# --------------------------------------------------------------------------
# Alerts API (not wrapped by pykiteconnect)
# --------------------------------------------------------------------------
class Alerts:
    def __init__(self, kite, dry_run: bool):
        self.dry = dry_run
        self.h = {"X-Kite-Version": "3",
                  "Authorization": f"token {kite.api_key}:{kite.access_token}"}

    def _call(self, method, path, **kw):
        r = requests.request(method, API + path, headers=self.h, timeout=20, **kw)
        body = r.json()
        if body.get("status") != "success":
            raise RuntimeError(f"Alerts API {method} {path}: {body.get('message', body)}")
        return body.get("data")

    def create_ato(self, name, symbol, operator, level, basket_items):
        payload = {
            "name": name, "type": "ato",
            "lhs_exchange": "NSE", "lhs_tradingsymbol": symbol,
            "lhs_attribute": "LastTradedPrice", "operator": operator,
            "rhs_type": "constant", "rhs_constant": level,
            "basket": json.dumps({"name": name, "type": "alert", "tags": [],
                                  "items": basket_items}),
        }
        if self.dry:
            log(f"[DRY-RUN] would create ATO: {json.dumps(payload, indent=1)}")
            return {"uuid": f"dry-{name}"}
        return self._call("POST", "/alerts", data=payload)

    def get(self, uuid):
        return self._call("GET", f"/alerts/{uuid}")

    def delete(self, uuid):
        if self.dry:
            log(f"[DRY-RUN] would delete alert {uuid}")
            return
        self._call("DELETE", "/alerts", params={"uuid": uuid})


def order_item(fut, transaction_type, quantity):
    """One basket item: futures MARKET order (Zerodha adds market protection)."""
    return {"type": "insert", "tradingsymbol": fut, "exchange": "NFO", "weight": 10000,
            "params": {"transaction_type": transaction_type, "product": "NRML",
                       "order_type": "MARKET", "validity": "DAY", "validity_ttl": 1,
                       "quantity": int(quantity), "price": 0, "trigger_price": 0,
                       "disclosed_quantity": 0, "last_price": 0, "variety": "regular",
                       "tags": [TAG], "squareoff": 0, "stoploss": 0,
                       "trailing_stoploss": 0, "iceberg_legs": 0,
                       "market_protection": -1}}


# --------------------------------------------------------------------------
# Contract selection
# --------------------------------------------------------------------------
def next_weekday(d: dt.date) -> dt.date:
    """The next trading day after d (weekends skipped; exchange holidays not)."""
    d += dt.timedelta(days=1)
    while d.weekday() >= 5:
        d += dt.timedelta(days=1)
    return d


def pick_future(nfo: pd.DataFrame, symbol: str, trade_date: dt.date, last_days: int = 2):
    """Next-month future for an order that can fill on trade_date. If
    trade_date is one of the last `last_days` trading days of the current
    month's contract (expiry day or the day before), the month after that
    instead — e.g. trading on Mon 23 / Tue 24 Nov 2026 -> January.
    Returns (tradingsymbol, lot_size, expiry) or None."""
    futs = nfo[(nfo["name"] == symbol) & (nfo["instrument_type"] == "FUT")].copy()
    futs["expiry"] = pd.to_datetime(futs["expiry"]).dt.date
    futs = futs[futs["expiry"] >= trade_date].sort_values("expiry").reset_index(drop=True)
    if futs.empty:
        return None
    days_left = int(np.busday_count(trade_date, futs.loc[0, "expiry"]))  # Mon->Tue = 1
    idx = 2 if days_left < last_days else 1
    if idx >= len(futs):
        return None
    row = futs.loc[idx]
    return row["tradingsymbol"], int(row["lot_size"]), row["expiry"]


def is_trading_day(kite) -> bool:
    """True if NSE traded today (handles exchange holidays without a calendar).
    Falls back to a weekday check if the Kite plan has no market-quote access."""
    try:
        q = kite.quote(["NSE:RELIANCE"])["NSE:RELIANCE"]
        ltt = q.get("last_trade_time") or q.get("timestamp")
        return bool(ltt) and pd.Timestamp(ltt).date() == now_ist().date()
    except Exception as exc:  # noqa: BLE001
        log(f"[WARN] quote unavailable ({exc}); assuming weekdays are trading days")
        return now_ist().weekday() < 5


def net_qty(kite, tradingsymbol: str) -> int:
    for p in kite.positions()["net"]:
        if p["exchange"] == "NFO" and p["tradingsymbol"] == tradingsymbol:
            return int(p["quantity"])
    return 0


def held_futures(kite, nfo: pd.DataFrame) -> dict:
    """{stock symbol: [futures tradingsymbols with a non-zero position]} —
    any month, so stocks you already trade are left alone."""
    fut_names = dict(zip(nfo.loc[nfo["instrument_type"] == "FUT", "tradingsymbol"],
                         nfo.loc[nfo["instrument_type"] == "FUT", "name"]))
    held = {}
    for p in kite.positions()["net"]:
        if p["exchange"] == "NFO" and int(p["quantity"]) != 0 and p["tradingsymbol"] in fut_names:
            held.setdefault(fut_names[p["tradingsymbol"]], []).append(p["tradingsymbol"])
    return held


def fired(alert: dict) -> bool:
    return int(alert.get("alert_count") or 0) > 0 or alert.get("status") != "enabled"


def round_up(x: float, step: float = 0.05) -> float:
    return round(math.ceil(round(x / step, 6)) * step, 2)


def entry_trigger(kite, sym: str, sig_row, session: str):
    """Entry level = close of the last completed normal (non-HA) daily candle
    + ENTRY_BUFFER, i.e. "previous day's close" for the next trading session.

    Last completed candle:
      * after 15:30 on a trading day (the 5:30 pm run) -> today's close
      * during market hours                           -> previous session's close
      * before the open / on a holiday                -> last trading day's close
    The signal file's Close (NSE official close from the 5 pm scan) is used when
    that day is the signal day; otherwise the Kite quote: ohlc.close during market
    hours (the previous close), or the last traded price after the close.
    Returns (trigger, ref_close, ref_desc, ltp); ltp is None if no quote access."""
    now = now_ist()
    today = now.date()
    q = None
    try:
        q = kite.quote([f"NSE:{sym}"])[f"NSE:{sym}"]
    except Exception as exc:  # noqa: BLE001
        log(f"[WARN] {sym}: quote unavailable ({exc}); using the signal-day Close")
    ltp = float(q["last_price"]) if q and q.get("last_price") else None
    ltt = (q.get("last_trade_time") or q.get("timestamp")) if q else None
    ltt_day = pd.Timestamp(ltt).date() if ltt else None
    file_close = sig_row.get("Close")
    file_close = None if file_close is None or pd.isna(file_close) else float(file_close)

    if ltt_day == today and now.time() < dt.time(15, 30):
        prev = (q.get("ohlc") or {}).get("close")
        if not prev:
            raise RuntimeError("Kite quote has no previous close")
        ref_close, ref_desc = float(prev), "prev session close"
    elif ltt_day is None or ltt_day.isoformat() == session:
        if file_close is None:
            raise RuntimeError("no Close in the signal file and no quote access")
        ref_close, ref_desc = file_close, f"close {session}"
    else:
        if ltp is None:
            raise RuntimeError("Kite quote has no last price")
        ref_close, ref_desc = ltp, f"close {ltt_day} (last traded price)"
    return round_up(ref_close * (1 + ENTRY_BUFFER)), round(ref_close, 2), ref_desc, ltp


# --------------------------------------------------------------------------
# Email
# --------------------------------------------------------------------------
def notify(subject: str, lines, html_body: str = "", dry_run: bool = False):
    """Plain-text email; with html_body, a multipart email (HTML + text fallback)."""
    user, pwd, to = (os.environ.get(k, "") for k in
                     ("EMAIL_USERNAME", "EMAIL_APP_PASSWORD", "EMAIL_TO"))
    body = "\n".join(lines)
    log(f"[NOTIFY] {subject}\n{body}")
    if not (user and pwd and to):
        return
    if html_body:
        msg = MIMEMultipart("alternative")
        msg.attach(MIMEText(body, "plain", "utf-8"))
        msg.attach(MIMEText(html_body, "html", "utf-8"))
    else:
        msg = MIMEText(body, "plain", "utf-8")
    msg["From"], msg["To"], msg["Subject"] = (user, to, ("[DRY RUN] " if dry_run else "")
                                            + f"[HA52 ATO] {subject}")
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
            s.starttls()
            s.login(user, pwd)
            s.sendmail(user, [r.strip() for r in to.split(",")], msg.as_string())
    except Exception as exc:  # noqa: BLE001
        log(f"[WARN] email failed: {exc}")


def _short_dates(text: str) -> str:
    """2026-11-23 -> 23-Nov (for the compact notes in the email table)."""
    return re.sub(r"\b(\d{4})-(\d{2})-(\d{2})\b",
                  lambda m: dt.date(int(m[1]), int(m[2]), int(m[3])).strftime("%d-%b"), text)


_SIGNAL_BADGE = {"DOJI": ("#fef3c7", "#92400e"), "GREEN": ("#dcfce7", "#166534"),
                 "GREEN (STRONG)": ("#bbf7d0", "#14532d")}
_TD = "padding:9px 10px;"
_NOTE = "font-size:11px;color:#6b7280;"


def place_email_html(session: str, placed_rows, skipped, dry_run: bool = False) -> str:
    """Colour-coded HTML for the 5:30 pm entry email (inline CSS for Gmail)."""
    e = html.escape

    def tile(label, n, bg, fg, dark):
        return (f'<td style="padding:14px 20px;background:{bg};width:50%;">'
                f'<div style="font-size:12px;color:{fg};text-transform:uppercase;'
                f'letter-spacing:.5px;">{label}</div>'
                f'<div style="font-size:24px;font-weight:700;color:{dark};">{n}</div></td>')

    def price(v):
        return f"{v:,.2f}"

    out = [
        '<div style="margin:0;padding:16px;background:#f3f5f9;font-family:Segoe UI,Arial,'
        'sans-serif;color:#1f2937;"><div style="max-width:960px;margin:0 auto;'
        'background:#ffffff;border-radius:8px;overflow:hidden;border:1px solid #e5e7eb;">',
        '<div style="background:#1e3a8a;color:#ffffff;padding:16px 20px;">'
        f'<div style="font-size:20px;font-weight:600;">HA52 ATO — Entry orders for '
        f'{e(session)} signals</div><div style="font-size:13px;opacity:.85;margin-top:4px;">'
        f'{"Previewed" if dry_run else "Placed"} {now_ist():%a %d %b %Y, %H:%M} IST · '
        'Heikin Ashi reversals at 52-week lows</div></div>',
        ('<div style="background:#fef9c3;color:#854d0e;padding:10px 20px;font-size:14px;'
         'font-weight:600;border-bottom:1px solid #fde68a;">🧪 DRY RUN — preview only, '
         'nothing was placed in Kite</div>') if dry_run else "",
        '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" '
        'style="border-collapse:collapse;"><tr>'
        + tile("Would place" if dry_run else "Placed", len(placed_rows), "#ecfdf5", "#047857", "#065f46")
        + tile("Skipped", len(skipped), "#fff7ed", "#c2410c", "#9a3412")
        + '</tr></table>',
        '<div style="padding:18px 20px 6px;font-size:16px;font-weight:600;color:#065f46;">'
        '✅ Entry ATOs — open until they trigger</div>'
        '<div style="padding:0 20px 18px;overflow-x:auto;">',
    ]
    if placed_rows:
        heads = [("Stock", "left"), ("Signal", "left"), ("Buy", "left"),
                 ("Entry (LTP ≥)", "right"), ("Stop-loss (LTP ≤)", "right"),
                 ("Target (LTP ≥)", "right"), ("Risk ₹", "right"), ("Reward ₹", "right"),
                 ("R:R", "center")]
        out.append('<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;'
                   'width:100%;font-size:13px;"><tr style="background:#1e3a8a;color:#ffffff;">'
                   + "".join(f'<th style="{_TD}text-align:{a};">{h}</th>' for h, a in heads)
                   + "</tr>")
        for i, r in enumerate(placed_rows):
            bg, fg = _SIGNAL_BADGE.get(r["signal"], ("#e5e7eb", "#374151"))
            rr_bg, rr_fg = (("#d1fae5", "#065f46") if r["rr"] >= 2 else ("#fef3c7", "#92400e"))
            out.append(
                f'<tr style="background:{"#ffffff" if i % 2 == 0 else "#f9fafb"};'
                'border-bottom:1px solid #e5e7eb;">'
                f'<td style="{_TD}font-weight:700;">{e(r["symbol"])}</td>'
                f'<td style="{_TD}"><span style="background:{bg};color:{fg};padding:2px 8px;'
                f'border-radius:10px;font-size:12px;font-weight:600;">{e(r["signal"])}</span></td>'
                f'<td style="{_TD}">{r["qty"]:,} × {e(r["future"])}</td>'
                f'<td style="{_TD}text-align:right;"><b style="color:#1d4ed8;">{price(r["entry"])}'
                f'</b><div style="{_NOTE}">{e(_short_dates(r["entry_note"]))}</div></td>'
                f'<td style="{_TD}text-align:right;"><b style="color:#b91c1c;">{price(r["sl"])}'
                f'</b><div style="{_NOTE}">HA Low</div></td>'
                f'<td style="{_TD}text-align:right;"><b style="color:#047857;">'
                f'{price(r["target"])}</b><div style="{_NOTE}">swing high '
                f'{e(_short_dates(r["target_date"]))}</div></td>'
                f'<td style="{_TD}text-align:right;color:#b91c1c;">{r["risk"]:,.0f}</td>'
                f'<td style="{_TD}text-align:right;color:#047857;">{r["reward"]:,.0f}</td>'
                f'<td style="{_TD}text-align:center;"><span style="background:{rr_bg};'
                f'color:{rr_fg};padding:2px 8px;border-radius:10px;font-weight:700;">'
                f'1 : {r["rr"]:.1f}</span></td></tr>')
        out.append("</table>")
    else:
        out.append('<div style="font-size:13px;color:#6b7280;">None</div>')
    out.append("</div>")

    out.append('<div style="padding:6px 20px 6px;font-size:16px;font-weight:600;'
               'color:#9a3412;">⚠️ Skipped</div><div style="padding:0 20px 18px;">')
    if skipped:
        out.append('<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;'
                   'width:100%;font-size:13px;"><tr style="background:#fed7aa;color:#7c2d12;">'
                   f'<th style="padding:8px 10px;text-align:left;width:140px;">Stock</th>'
                   f'<th style="padding:8px 10px;text-align:left;">Reason</th></tr>')
        for line in skipped:
            sym, _, reason = line.partition(": ")
            out.append('<tr style="background:#fff7ed;border-bottom:1px solid #fde3c8;">'
                       f'<td style="padding:8px 10px;font-weight:700;">{e(sym)}</td>'
                       f'<td style="padding:8px 10px;">{e(_short_dates(reason))}</td></tr>')
        out.append("</table>")
    else:
        out.append('<div style="font-size:13px;color:#6b7280;">None</div>')
    out.append('</div><div style="padding:12px 20px;background:#f9fafb;border-top:1px solid '
               '#e5e7eb;font-size:12px;color:#6b7280;">Orders: BUY futures, NRML, MARKET. '
               'After a fill, stop-loss and target ATOs are created automatically; whichever '
               'fires first closes the trade.</div></div></div>')
    return "".join(out)


# --------------------------------------------------------------------------
# Commands
# --------------------------------------------------------------------------
def cmd_place(args, state: State):
    if not args.no_pull:
        subprocess.run(["git", "-C", str(REPO), "pull", "-q", "--ff-only"], check=False)
    sig_path = Path(args.signals)
    if not sig_path.exists():
        notify("No signal file", [f"{sig_path} not found — did the 5 pm scan run?"])
        return
    sig = pd.read_csv(sig_path)
    today = now_ist().date()
    session = args.session or today.isoformat()
    if session in state.data["processed_sessions"] and not args.force:
        log(f"[INFO] Session {session} already processed — nothing to do.")
        return
    if sig.empty:
        log("[INFO] No F&O HA reversal signals today.")
        return
    if (sig["Session_Date"].astype(str) != session).all():
        notify("Signal file is stale",
               [f"{sig_path} is for {sig['Session_Date'].iloc[0]}, expected {session}.",
                "The 5 pm scan probably failed or has not committed yet. No ATOs placed."])
        return
    sig = sig[sig["Session_Date"].astype(str) == session]

    kite = get_kite(args.home)
    if not args.session and not is_trading_day(kite):
        log("[INFO] Not a trading day — skipping.")
        return
    alerts = Alerts(kite, args.dry_run)
    nfo = pd.DataFrame(kite.instruments("NFO"))
    open_syms = state.open_symbols()
    held = held_futures(kite, nfo)
    placed, placed_rows, skipped = [], [], []
    for _, r in sig.iterrows():
        sym = r["Symbol"]
        if sym in open_syms and not args.dry_run:  # dry run previews every signal
            skipped.append(f"{sym}: already has an open trade/ATO from this script")
            continue
        if sym in held:
            skipped.append(f"{sym}: you already hold futures ({', '.join(held[sym])})")
            continue
        # An order placed before the close can fill today; after it, next session
        fill_day = today if now_ist().time() < dt.time(15, 30) else next_weekday(today)
        fut = pick_future(nfo, sym, fill_day)
        if fut is None:
            skipped.append(f"{sym}: no suitable futures contract found")
            continue
        fut_sym, lot, expiry = fut
        qty = lot * args.lots
        try:
            level, ref_close, ref_desc, ltp = entry_trigger(kite, sym, r, session)
        except Exception as exc:  # noqa: BLE001
            skipped.append(f"{sym}: could not work out the entry trigger — {exc}")
            continue
        op, desc = ">=", f"{ref_desc} {ref_close} + 0.5%"
        if ltp is not None and ltp >= level:
            skipped.append(f"{sym}: LTP {ltp} already >= entry trigger {level} ({desc}) "
                           "— not chasing")
            continue
        sl = round(float(r["HA_Low"]), 2)
        target = r.get("Target")
        if target is None or pd.isna(target):
            skipped.append(f"{sym}: no swing-high target in the signal file")
            continue
        target = round(float(target), 2)
        risk, reward = max(abs(level - sl), 0.01), target - level
        if reward < risk:
            skipped.append(f"{sym}: target {target} too close — reward {reward:.2f} < "
                           f"risk {risk:.2f} (entry {level}, SL {sl})")
            continue
        name = f"{TAG} ENTRY {sym} {session}"
        try:
            res = alerts.create_ato(name, sym, op, level, [order_item(fut_sym, "BUY", qty)])
        except Exception as exc:  # noqa: BLE001
            if "already met" in str(exc).lower():
                skipped.append(f"{sym}: price already at/above entry trigger {level} "
                               f"({desc}) — not chasing")
            else:
                skipped.append(f"{sym}: ATO create failed — {exc}")
            continue
        trade = {
            "symbol": sym, "signal": r["Signal"], "session": session,
            "future": fut_sym, "expiry": str(expiry), "qty": qty,
            "entry_op": op, "entry_level": level, "entry_ref": desc, "sl_level": sl,
            "target": target, "target_date": str(r.get("Target_Date", "")),
            "baseline_qty": net_qty(kite, fut_sym) if not args.dry_run else 0,
            "entry_uuid": res["uuid"], "sl_uuid": None, "tgt_uuid": None,
            "status": "entry_open",
            "created_at": now_ist().isoformat(timespec="seconds"),
        }
        if not args.dry_run:
            state.trades.append(trade)
        placed_rows.append({
            "symbol": sym, "signal": str(r["Signal"]), "qty": qty, "future": fut_sym,
            "entry": level, "entry_note": desc, "sl": sl, "target": target,
            "target_date": str(r.get("Target_Date", "")), "risk": risk * qty,
            "reward": reward * qty, "rr": reward / risk})
        placed.append(f"{sym:<12} {r['Signal']:<15} BUY {qty} {fut_sym} when LTP {op} {level} "
                      f"({desc}) | after fill: SL LTP <= {sl} (HA Low), "
                      f"target LTP >= {target} (swing high {r.get('Target_Date', '')}) | "
                      f"risk ~Rs {risk * qty:,.0f}, reward ~Rs {reward * qty:,.0f} "
                      f"(1:{reward / risk:.1f})")
    if args.dry_run:
        for line in placed:
            log(f"[DRY-RUN] {line}")
        notify(f"{len(placed)} entry ATOs would be placed for {session}",
               ["DRY RUN — nothing was placed.", "",
                "Entry ATOs:", *(placed or ["None"]), "", "Skipped:", *(skipped or ["None"])],
               html_body=place_email_html(session, placed_rows, skipped, dry_run=True),
               dry_run=True)
        return
    state.data["processed_sessions"].append(session)
    state.save()
    notify(f"{len(placed)} entry ATOs placed for {session}",
           ["Entry ATOs (open until they trigger):", *(placed or ["None"]), "",
            "Skipped:", *(skipped or ["None"])],
           html_body=place_email_html(session, placed_rows, skipped))


def cmd_manage(args, state: State):
    active = [t for t in state.trades
              if t["status"] in ("entry_open", "entry_triggered", "sl_open")]
    if not active:
        return
    kite = get_kite(args.home)
    alerts = Alerts(kite, args.dry_run)
    changed = False
    for t in active:
        try:
            qty_now = net_qty(kite, t["future"])
            base = int(t.get("baseline_qty") or 0)
            if t["status"] in ("entry_open", "entry_triggered"):
                entry_fired = fired(alerts.get(t["entry_uuid"]))
                if entry_fired and qty_now >= base + t["qty"]:
                    _place_exits(kite, alerts, t)
                    changed = True
                elif entry_fired and t["status"] == "entry_open":
                    t["status"], t["triggered_at"] = "entry_triggered", now_ist().isoformat()
                    changed = True
                elif t["status"] == "entry_triggered":
                    since = dt.datetime.fromisoformat(t["triggered_at"])
                    if now_ist() - since > dt.timedelta(minutes=15):
                        t["status"] = "entry_failed"
                        notify(f"Entry for {t['symbol']} triggered but no position",
                               [f"The ATO fired but no new {t['future']} position appeared — "
                                "the order was probably rejected (margin, price band). "
                                "Check the Kite order book."])
                        changed = True
            elif t["status"] == "sl_open" and qty_now <= base:
                # Position closed: by the stop-loss, the target, or by you.
                legs = {"stop-loss": t.get("sl_uuid"), "target": t.get("tgt_uuid")}
                hit, removed = [], []
                for leg, uuid in legs.items():
                    if not uuid:
                        continue
                    a = alerts.get(uuid)
                    if fired(a):
                        hit.append(leg)
                    else:
                        alerts.delete(uuid)
                        removed.append(leg)
                how = (f"{' and '.join(hit)} hit" if hit else "closed manually")
                t.update(status="closed", closed_by=how,
                         closed_at=now_ist().isoformat(timespec="seconds"))
                notify(f"{t['symbol']} position closed ({how})",
                       [f"{t['future']}: {how}.",
                        f"Deleted the open {' and '.join(removed)} ATO." if removed
                        else "No open exit ATOs left."])
                changed = True
        except Exception as exc:  # noqa: BLE001
            log(f"[WARN] {t['symbol']}: {exc}")
    if changed:
        state.save()


def _place_exits(kite, alerts, t):
    """After the buy fills: stop-loss ATO (HA Low) and target ATO (swing high)."""
    sym = t["symbol"]
    res = alerts.create_ato(f"{TAG} SL {sym} {t['session']}", sym, "<=", t["sl_level"],
                            [order_item(t["future"], "SELL", t["qty"])])
    t.update(sl_uuid=res["uuid"], status="sl_open",
             filled_at=now_ist().isoformat(timespec="seconds"))
    lines = [f"Long {t['qty']} {t['future']}.",
             f"Stop-loss ATO: sell when {sym} LTP <= {t['sl_level']} (signal-day HA Low)."]
    if t.get("target"):
        try:
            res = alerts.create_ato(f"{TAG} TGT {sym} {t['session']}", sym, ">=",
                                    t["target"], [order_item(t["future"], "SELL", t["qty"])])
            t["tgt_uuid"] = res["uuid"]
            lines.append(f"Target ATO: sell when {sym} LTP >= {t['target']} "
                         f"(swing high {t.get('target_date', '')}).")
        except Exception as exc:  # noqa: BLE001
            lines.append(f"WARNING: target ATO failed ({exc}) — set a target manually.")
    if t.get("tgt_uuid"):
        lines.append("Whichever fires first closes the trade; the other is deleted within "
                     "5 minutes.")
    try:
        ltp = kite.ltp([f"NSE:{sym}"])[f"NSE:{sym}"]["last_price"]
        lines.append(f"Stock LTP now {ltp}.")
        if ltp <= t["sl_level"]:
            lines.append("WARNING: price is already below the stop — it will trigger now.")
    except Exception:  # noqa: BLE001
        pass
    notify(f"BUY filled: {sym} — stop-loss" + (" and target" if t.get("tgt_uuid") else "")
           + " placed", lines)


def cmd_expire(args, state: State):
    today = now_ist().date().isoformat()
    stale = [t for t in state.trades if t["status"] == "entry_open"
             and t["created_at"][:10] < today]
    if not stale:
        return
    kite = get_kite(args.home)
    if not is_trading_day(kite):
        log("[INFO] Not a trading day — entry ATOs stay open until the next session.")
        return
    alerts = Alerts(kite, args.dry_run)
    done = []
    for t in stale:
        try:
            a = alerts.get(t["entry_uuid"])
            if a.get("status") == "enabled" and int(a.get("alert_count") or 0) == 0:
                alerts.delete(t["entry_uuid"])
                t["status"] = "expired"
                done.append(t["symbol"])
        except Exception as exc:  # noqa: BLE001
            log(f"[WARN] {t['symbol']}: {exc}")
    if done and not args.dry_run:
        state.save()
        notify(f"{len(done)} untriggered entry ATOs expired", done)


def cmd_status(args, state: State):
    if not state.trades:
        print("No trades tracked yet.")
        return
    df = pd.DataFrame(state.trades)
    cols = ["session", "symbol", "signal", "future", "qty", "entry_op", "entry_level",
            "sl_level", "status"]
    print(df[[c for c in cols if c in df]].to_string(index=False))


def main():
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("command", choices=["place", "manage", "expire", "status", "login"])
    p.add_argument("--dry-run", action="store_true", help="Print actions; change nothing.")
    p.add_argument("--signals", default=str(DEFAULT_SIGNALS))
    p.add_argument("--home", type=Path, default=DEFAULT_HOME,
                   help="Folder for trades.json and the cached token (default: ~/.ha_trader)")
    p.add_argument("--lots", type=int, default=int(os.environ.get("HA_LOTS", "1")))
    p.add_argument("--session", help="Signal date YYYY-MM-DD to process (testing/back-fill)")
    p.add_argument("--force", action="store_true", help="Re-process an already done session")
    p.add_argument("--no-pull", action="store_true", help="Don't git pull before placing")
    args = p.parse_args()
    if os.environ.get("HA_ENTRY"):
        log("[INFO] HA_ENTRY is no longer used (entry is always LTP >= prev day close + 0.5%); "
            "remove it from ~/.ha_trader/.env")
    state = State(args.home)
    if args.command == "login":
        k = get_kite(args.home)
        print("Logged in:", k.profile()["user_name"])
        return
    try:
        {"place": cmd_place, "manage": cmd_manage, "expire": cmd_expire,
         "status": cmd_status}[args.command](args, state)
    except Exception as exc:  # noqa: BLE001
        alert_failure(args, exc)
        raise


def alert_failure(args, exc: Exception):
    """Email when a scheduled run crashes (e.g. the TOTP login fails).
    At most one email per command per day, so the 5-minute `manage` runs
    don't flood your inbox; the log keeps every failure."""
    if args.dry_run or args.command == "status":
        return
    flag = args.home / f"failed-{args.command}-{now_ist().date().isoformat()}"
    if flag.exists():
        return
    flag.touch()
    what = {"place": "no entry ATOs were placed",
            "manage": "fills are NOT being checked and no stop-loss/target ATOs are "
                      "being placed or cleaned up",
            "expire": "untriggered entry ATOs were not cleared"}[args.command]
    notify(f"{args.command} FAILED at {now_ist():%H:%M} — action needed",
           [f"The '{args.command}' run failed, so {what}.",
            f"Error: {type(exc).__name__}: {exc}", "",
            "Check Kite and any open HA52 positions yourself until this is fixed.",
            "Log: ~/.ha_trader/logs/ on the server. Test with: ato_trader.py login",
            "(Only the first failure per command per day is emailed.)"])


if __name__ == "__main__":
    main()
