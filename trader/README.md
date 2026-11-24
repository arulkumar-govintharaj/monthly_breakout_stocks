# HA52 ATO trader — auto-orders for 52-week-low Heikin Ashi reversals

Places Zerodha **ATO (Alert Triggers Order)** trades for every stock in the
**F&O stocks — futures available** table of the email's *Heikin Ashi
Reversals — 52-Week-Low Stocks* section.

## How it works

| Time (IST) | Where | What |
|---|---|---|
| ~17:00 | GitHub Actions | Screener runs, emails you, commits `state/ha_fno_signals.csv` (the F&O HA signals) |
| 17:30 | Your server | `place`: pulls the repo and creates one **entry ATO** per stock |
| every 5 min, 09:00–15:55 | Your server | `manage`: when an entry fills, creates the **stop-loss** and **target** ATOs; when one fires, deletes the other |

**Entry ATO** — alert on the stock's NSE price, order on futures:
- trigger: stock LTP **≥ previous day's close + 0.5%** (normal candle, not
  HA). "Previous day" is the last completed daily candle: **today's close**
  when run after 15:30 (the 5:30 pm run), otherwise the **previous session's
  close** (manual rerun before the open, during market hours, or on a holiday).
  The signal day's close comes from the 5 pm scan (NSE official close); a later
  day's comes from the Kite quote.
- if the price is already at or above the trigger, the stock is skipped
  ("not chasing") — Kite refuses alerts whose condition is already met.
- order: **BUY 1 lot** (`HA_LOTS`) futures, NRML, MARKET (Zerodha adds market protection)
- contract: **next month**; if the order can fill on the **last 2 trading days**
  of the current month's contract (expiry day or the day before), the month
  after that. Expiry 24 Nov 2026: ATOs placed Fri 20 Nov or Mon 23 Nov buy
  **January**; placed Thu 19 Nov they buy December. (Weekends are skipped;
  exchange holidays are not.)
- stays open until it triggers (not auto-deleted). While a stock has an open
  entry ATO or position from this script, new signals for it are skipped.
  `trader/run.sh expire` clears untriggered entry ATOs manually.

**Exit ATOs** — created only after the trader's own buy has filled (so they
can never open a short):
- **stop-loss:** stock LTP **≤ signal-day HA Low**
- **target:** stock LTP **≥ last swing high before the 52-week low** — the most
  recent day whose High is above the 5 sessions on either side, within 60
  sessions before the low (fallback: highest High of those 60 sessions).
  Calculated by the 5 pm scan and shown in the email.
- order: SELL the same quantity, NRML, MARKET
- whichever fires first closes the trade; the other is deleted at the next
  5-minute check. If you close the position yourself, both are deleted.

**Skipped (listed in the 5:30 pm email):**
- you already hold futures in that stock (any month) — your own positions are
  never touched;
- the target is closer than the stop-loss (reward < risk);
- the signal has no target;
- the price is already at/above the entry trigger.

**Rerunning manually the next morning** (e.g. after a failed 5:30 pm run):

```bash
trader/run.sh place --session 2026-11-23 --force    # the signal date
```

Stocks that already got an entry ATO are skipped automatically.

Example email line (dry run with test data):

```
NMDC  DOJI  BUY 6750 NMDC27JANFUT when LTP >= 67.75 (close 2026-11-24 67.4 + 0.5%) | after fill: SL LTP <= 66.91 (HA Low), target LTP >= 82.32 (swing high 2026-09-23) | risk ~Rs 5,670, reward ~Rs 98,347 (1:17.3)
KEI: LTP 3900.0 already >= entry trigger 3819.65 (close 2026-11-24 3800.6 + 0.5%) — not chasing
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
- Zerodha allows at most 200 ATOs per account; untriggered entries count
  towards it, so clear old ones now and then with `expire`.
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
