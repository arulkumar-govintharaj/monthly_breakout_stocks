# NSE Monthly Breakout Screener

Screens NSE F&O stocks for symbols whose price has broken the
**previous completed calendar month's HIGH and/or LOW**, writes the
result to a single Excel file, and emails you the breakout list
(with the spreadsheet attached) via Gmail SMTP.

Runs automatically **Monday–Friday at 5:00 PM IST** via GitHub Actions.

## Output columns

| Column | Meaning |
|---|---|
| `Symbol` | NSE symbol |
| `LTP` | Latest close/price |
| `Prev_Month` | The calendar month used as reference, e.g. `2026-07` |
| `Prev_Month_High` / `Prev_Month_Low` | That month's high/low |
| `Today_High` / `Today_Low` | Today's session high/low |
| `Broke_Month_High` | `YES` if today's high > previous month's high |
| `Broke_Month_Low` | `YES` if today's low < previous month's low |
| `%_From_Month_High` / `%_From_Month_Low` | How far LTP sits from each level |
| `Session_Date` | Date of the session used |

Rows are sorted with dual breakouts (both high and low broken) first.

## Quick start (local)

```bash
pip install -r requirements.txt
python monthly_breakout_screener.py --no-email
```

Writes `output/monthly_breakout_screener_<timestamp>.xlsx`.

## Options

```bash
# Scan a custom list
python monthly_breakout_screener.py --symbols RELIANCE,SBIN,SONACOMS --no-email

# Use your own universe file (single column of symbols)
python monthly_breakout_screener.py --universe my_universe.csv --no-email

# Only keep rows that actually broke out (spreadsheet + email both)
python monthly_breakout_screener.py --only-breakouts

# Custom output folder
python monthly_breakout_screener.py --out results/ --no-email
```

## Email setup (GitHub Actions)

The workflow sends email via Gmail SMTP using three **repository
secrets** — add these under **Settings → Secrets and variables →
Actions → New repository secret**:

| Secret | Value |
|---|---|
| `EMAIL_USERNAME` | Your Gmail address (the sender) |
| `EMAIL_APP_PASSWORD` | A Gmail **App Password** — not your normal Gmail password. Generate one at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) (requires 2-Step Verification enabled on the account) |
| `EMAIL_TO` | Recipient address, or comma-separated list for multiple |

Without these three secrets set, the script still runs and saves the
Excel file — it just skips the email step (you'll see a note in the
Actions log) and the file is still downloadable from the workflow's
**Artifacts**.

## Schedule

`.github/workflows/monthly_breakout.yml` runs at `30 11 * * 1-5` in
cron (UTC), which is **5:00 PM IST, Monday–Friday**. GitHub Actions
cron can run a few minutes late during high load — that's normal.

You can also trigger it manually anytime from the **Actions** tab
(`workflow_dispatch`), optionally passing a custom symbol list or
restricting the spreadsheet to only breakout rows.

## Notes

- Data source is Yahoo Finance via `yfinance` (`SYMBOL.NS`) — free,
  fine for EOD screening, not tick-accurate. Cross-check against your
  broker before acting on it.
- "Broke month high/low" means *at any point during today's session*
  (via the day's High/Low), not "currently trading above/below" — check
  `LTP` against the levels if you want to know whether the break has held.
- The default universe is a manually curated list of liquid F&O names.
  NSE revises the F&O list periodically — update `DEFAULT_FNO_UNIVERSE`
  in `monthly_breakout_screener.py`, or pass `--universe` with a fresh CSV.
- "Previous month" always means the last **fully completed** calendar
  month — e.g. if today is in September, it compares against August's
  high/low, not September-to-date.
