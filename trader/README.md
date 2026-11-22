# HA52 ATO trader — auto-orders for 52-week-low Heikin Ashi reversals

Places Zerodha **ATO (Alert Triggers Order)** trades for every stock in the
**F&O stocks — futures available** table of the email's *Heikin Ashi
Reversals — 52-Week-Low Stocks* section.

## How it works

| Time (IST) | Where | What |
|---|---|---|
| ~17:00 | GitHub Actions | Screener runs, emails you, commits `state/ha_fno_signals.csv` (the F&O HA signals) |
| 17:30 | Your server | `place`: pulls the repo and creates one **entry ATO** per stock |
| next day, every 2 min | Your server | `manage`: when an entry fills, creates the **stop-loss ATO** |
| next day 15:35 | Your server | `expire`: deletes entry ATOs that did not trigger |

**Entry ATO** — alert on the stock's NSE price, order on futures:
- trigger: stock LTP **≤ signal-day HA Close** (`HA_ENTRY=dip`, default) or
  **≥ signal-day HA High** (`HA_ENTRY=breakout`)
- order: **BUY 1 lot** (`HA_LOTS`) futures, NRML, MARKET (Zerodha adds market protection)
- contract: **next month**; if the order can fill on the **last 2 trading days**
  of the current month's contract (expiry day or the day before), the month
  after that. Expiry 24 Nov 2026: ATOs placed Fri 20 Nov or Mon 23 Nov buy
  **January**; placed Thu 19 Nov they buy December. (Weekends are skipped;
  exchange holidays are not.)
- valid for the next trading day only

**Stop-loss ATO** — created only after the futures position exists (so it can
never open a short):
- trigger: stock LTP **≤ signal-day HA Low**
- order: SELL the same quantity, NRML, MARKET
- if you close the position yourself, the stop-loss ATO is deleted automatically

Example from 20 Nov 2026 (dry run):

```
TMPV        GREEN (STRONG)  BUY 1600 TMPV27JANFUT when LTP <= 264.46 | SL: LTP <= 260.74 | risk ~Rs 5,952
ASTRAL      DOJI            BUY 425 ASTRAL27JANFUT when LTP <= 1247.42 | SL: LTP <= 1236.5 | risk ~Rs 4,641
SUPREMEIND  DOJI            BUY 175 SUPREMEIND27JANFUT when LTP <= 3207.88 | SL: LTP <= 3182.1 | risk ~Rs 4,512
```

There is **no position cap**: every F&O signal gets an ATO. If margin runs
out, Zerodha rejects that order and you get an email.

## Why a separate server?

Since **1 April 2026** every API order must come from a **static IP registered
on your Kite Connect app**. GitHub Actions has no fixed IP, so the scan stays
on GitHub and the order steps run on a small server you control.

## One-time setup

1. **Kite Connect app** — sign up at <https://developers.kite.trade>, create an
   app, note the API key and secret. Set any redirect URL (e.g.
   `https://127.0.0.1`). Log in through the app once in a browser to authorise it.
2. **Static-IP server** — a small Ubuntu VPS (a Mumbai/India region is best).
   Note its public IP and add it under **IP Whitelist** on your
   developers.kite.trade profile (only one change per week is allowed).
3. **External TOTP** — in Kite, enable external TOTP and keep the setup key
   (the base32 secret). This is what makes the daily login automatic.
4. **On the server:**
   ```bash
   sudo timedatectl set-timezone Asia/Kolkata
   sudo apt update && sudo apt install -y git python3-venv
   git clone https://github.com/arulkumar-govintharaj/monthly_breakout_stocks.git
   cd monthly_breakout_stocks
   python3 -m venv .venv && .venv/bin/pip install -r trader/requirements.txt
   mkdir -p ~/.ha_trader && cp trader/.env.example ~/.ha_trader/.env
   chmod 600 ~/.ha_trader/.env && nano ~/.ha_trader/.env    # fill in
   chmod +x trader/run.sh
   ```
5. **Test** (nothing is placed with `--dry-run`):
   ```bash
   set -a; . ~/.ha_trader/.env; set +a
   .venv/bin/python trader/ato_trader.py login
   .venv/bin/python trader/ato_trader.py place --dry-run --session 2026-11-20 --no-pull
   ```
6. **Schedule** — `crontab -e` and paste `trader/crontab.example` (fix the path).

## Day to day

- `trader/run.sh status` — trades the script is tracking.
- Logs: `~/.ha_trader/logs/YYYY-MM.log`. Emails for: ATOs placed, fills +
  stop-loss, rejections, expiries, position closed.
- To stop trading: comment out the cron lines (and delete open `HA52` alerts
  in Kite → Orders → Alerts if you want them gone).

## Things to know

- The login uses your password + TOTP secret stored in `~/.ha_trader/.env`
  on the server. Keep that server locked down (SSH keys only).
- The automated login uses Kite's web login flow, which Zerodha does not
  document; if Zerodha changes it, `login` will fail and you'll see it in the
  log/email before any order is missed.
- The stop-loss is a price **alert** checked by Zerodha, then a MARKET order —
  in a gap-down open it fills at the market price, below the HA Low.
- Positions are matched by futures contract; if you trade the same contract
  manually, the script treats it as the same position.
