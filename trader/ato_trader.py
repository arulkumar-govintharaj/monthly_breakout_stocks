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
                            alert : NSE:<SYM> LastTradedPrice <= signal-day HA Close
                                    (--entry breakout: >= signal-day HA High)
                            order : BUY 1 lot of the futures contract, NRML, MARKET
                                    (next month; in expiry week the one after that)
  manage   every 2 min    once an entry ATO has fired and the futures position
           09:15-15:30    exists, create the stop-loss ATO:
                            alert : NSE:<SYM> LastTradedPrice <= signal-day HA Low
                            order : SELL the same futures quantity, NRML, MARKET
                          and tidy up when a position is closed (SL hit or manual).
  expire   15:35 Mon-Fri  delete entry ATOs that did not trigger during their one
                          trading day of validity.
  status   any time       print the trades this script is tracking.
  login    any time       just log in (checks credentials / TOTP).

Add --dry-run to any command to print what it would do without creating or
deleting anything. Configuration comes from environment variables (see
trader/.env.example). The ATO alerts are placed on the stock's cash price
because the HA candles are built from it; orders go to the futures contract.
"""

import argparse
import datetime as dt
import json
import os
import smtplib
import subprocess
import sys
import urllib.parse
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import requests

IST = ZoneInfo("Asia/Kolkata")
API = "https://api.kite.trade"
TAG = "HA52"                      # prefix on every alert this script creates

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
def pick_future(nfo: pd.DataFrame, symbol: str, today: dt.date, expiry_week_days: int = 7):
    """Next-month future; if today is in the last week before the current
    month's expiry, the contract after that (e.g. 23 Nov -> January).
    Returns (tradingsymbol, lot_size, expiry) or None."""
    futs = nfo[(nfo["name"] == symbol) & (nfo["instrument_type"] == "FUT")].copy()
    futs["expiry"] = pd.to_datetime(futs["expiry"]).dt.date
    futs = futs[futs["expiry"] >= today].sort_values("expiry").reset_index(drop=True)
    if futs.empty:
        return None
    idx = 2 if (futs.loc[0, "expiry"] - today).days < expiry_week_days else 1
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


# --------------------------------------------------------------------------
# Email
# --------------------------------------------------------------------------
def notify(subject: str, lines):
    user, pwd, to = (os.environ.get(k, "") for k in
                     ("EMAIL_USERNAME", "EMAIL_APP_PASSWORD", "EMAIL_TO"))
    body = "\n".join(lines)
    log(f"[NOTIFY] {subject}\n{body}")
    if not (user and pwd and to):
        return
    msg = MIMEText(body, "plain", "utf-8")
    msg["From"], msg["To"], msg["Subject"] = user, to, f"[HA52 ATO] {subject}"
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
            s.starttls()
            s.login(user, pwd)
            s.sendmail(user, [r.strip() for r in to.split(",")], msg.as_string())
    except Exception as exc:  # noqa: BLE001
        log(f"[WARN] email failed: {exc}")


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
    placed, skipped = [], []
    for _, r in sig.iterrows():
        sym = r["Symbol"]
        if sym in open_syms:
            skipped.append(f"{sym}: already has an open trade/ATO from this script")
            continue
        fut = pick_future(nfo, sym, today)
        if fut is None:
            skipped.append(f"{sym}: no suitable futures contract found")
            continue
        fut_sym, lot, expiry = fut
        qty = lot * args.lots
        if args.entry == "breakout":
            op, level, desc = ">=", round(float(r["HA_High"]), 2), "above HA High"
        else:
            op, level, desc = "<=", round(float(r["HA_Close"]), 2), "at/below HA Close"
        sl = round(float(r["HA_Low"]), 2)
        name = f"{TAG} ENTRY {sym} {session}"
        try:
            res = alerts.create_ato(name, sym, op, level, [order_item(fut_sym, "BUY", qty)])
        except Exception as exc:  # noqa: BLE001
            skipped.append(f"{sym}: ATO create failed — {exc}")
            continue
        trade = {
            "symbol": sym, "signal": r["Signal"], "session": session,
            "future": fut_sym, "expiry": str(expiry), "qty": qty,
            "entry_op": op, "entry_level": level, "sl_level": sl,
            "entry_uuid": res["uuid"], "sl_uuid": None, "status": "entry_open",
            "created_at": now_ist().isoformat(timespec="seconds"),
        }
        if not args.dry_run:
            state.trades.append(trade)
        placed.append(f"{sym:<12} {r['Signal']:<15} BUY {qty} {fut_sym} when LTP {op} {level} "
                      f"({desc}) | SL after fill: LTP <= {sl} (HA Low) | "
                      f"risk ~Rs {abs(level - sl) * qty:,.0f}")
    if args.dry_run:
        for line in placed:
            log(f"[DRY-RUN] {line}")
        return
    state.data["processed_sessions"].append(session)
    state.save()
    notify(f"{len(placed)} entry ATOs placed for {session}",
           ["Entry ATOs (valid for the next trading day):", *(placed or ["None"]), "",
            "Skipped:", *(skipped or ["None"])])


def cmd_manage(args, state: State):
    active = [t for t in state.trades
              if t["status"] in ("entry_open", "entry_triggered", "sl_open")
              ]
    if not active:
        return
    kite = get_kite(args.home)
    alerts = Alerts(kite, args.dry_run)
    changed = False
    for t in active:
        try:
            qty_now = net_qty(kite, t["future"])
            if t["status"] in ("entry_open", "entry_triggered"):
                a = alerts.get(t["entry_uuid"])
                fired = int(a.get("alert_count") or 0) > 0 or a.get("status") != "enabled"
                if qty_now >= t["qty"]:
                    res = alerts.create_ato(
                        f"{TAG} SL {t['symbol']} {t['session']}", t["symbol"], "<=",
                        t["sl_level"], [order_item(t["future"], "SELL", t["qty"])])
                    t.update(sl_uuid=res["uuid"], status="sl_open",
                             filled_at=now_ist().isoformat(timespec="seconds"))
                    try:
                        ltp = kite.ltp([f"NSE:{t['symbol']}"])[f"NSE:{t['symbol']}"]["last_price"]
                    except Exception:  # noqa: BLE001
                        ltp = None
                    notify(f"BUY filled: {t['symbol']} — stop-loss placed",
                           [f"Long {t['qty']} {t['future']}.",
                            f"Stop-loss ATO: sell when {t['symbol']} LTP <= {t['sl_level']} "
                            f"(signal-day HA Low). Stock LTP now {ltp}."]
                           + (["WARNING: price is already below the stop — it will "
                               "trigger immediately."]
                              if ltp is not None and ltp <= t["sl_level"] else []))
                    changed = True
                elif fired and t["status"] == "entry_open":
                    t["status"], t["triggered_at"] = "entry_triggered", now_ist().isoformat()
                    changed = True
                elif t["status"] == "entry_triggered":
                    since = dt.datetime.fromisoformat(t["triggered_at"])
                    if now_ist() - since > dt.timedelta(minutes=10):
                        t["status"] = "entry_failed"
                        notify(f"Entry for {t['symbol']} triggered but no position",
                               [f"The ATO fired but no {t['future']} position exists after "
                                "10 minutes — the order was probably rejected (margin, "
                                "price band). Check the Kite order book."])
                        changed = True
            elif t["status"] == "sl_open":
                if qty_now <= 0:
                    a = alerts.get(t["sl_uuid"])
                    if a.get("status") == "enabled":
                        alerts.delete(t["sl_uuid"])   # exited manually — remove the SL
                        how = "closed manually — stop-loss ATO deleted"
                    else:
                        how = "stop-loss hit"
                    t.update(status="closed", closed_at=now_ist().isoformat(timespec="seconds"))
                    notify(f"{t['symbol']} position closed ({how})", [f"{t['future']}: {how}."])
                    changed = True
        except Exception as exc:  # noqa: BLE001
            log(f"[WARN] {t['symbol']}: {exc}")
    if changed:
        state.save()


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
    p.add_argument("--entry", choices=["dip", "breakout"],
                   default=os.environ.get("HA_ENTRY", "dip"),
                   help="dip: buy when LTP <= signal-day HA Close (default); "
                        "breakout: buy when LTP >= signal-day HA High")
    p.add_argument("--session", help="Signal date YYYY-MM-DD to process (testing/back-fill)")
    p.add_argument("--force", action="store_true", help="Re-process an already done session")
    p.add_argument("--no-pull", action="store_true", help="Don't git pull before placing")
    args = p.parse_args()
    state = State(args.home)
    if args.command == "login":
        k = get_kite(args.home)
        print("Logged in:", k.profile()["user_name"])
        return
    {"place": cmd_place, "manage": cmd_manage, "expire": cmd_expire,
     "status": cmd_status}[args.command](args, state)


if __name__ == "__main__":
    main()
