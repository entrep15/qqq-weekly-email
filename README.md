# QQQ Top 10 Weekly Email

A server-side cron (GitHub Actions) that emails the **top 10 Invesco QQQ holdings**
— company name, ticker, and % allocation within QQQ — every **Friday at 6:05 AM
Pacific**. Runs entirely on GitHub's infrastructure; nothing local needs to be on.

## How it works

- `qqq_email.py` fetches the actual QQQ ETF holdings from stockanalysis.com
  (two host endpoints for resilience), renders an HTML + plain-text table, and
  sends it via Gmail SMTP.
- `.github/workflows/qqq-weekly.yml` triggers the script.

### DST handling

GitHub cron is UTC and DST-unaware, so the workflow fires at **two** UTC times
every Friday (13:05 and 14:05 UTC). The script only proceeds when it is actually
the 6 AM hour in `America/Los_Angeles`, so exactly one run sends each week —
correct in both PDT and PST. Manual runs (`workflow_dispatch`) set `FORCE_SEND=1`
and bypass the guard.

## Configuration (GitHub repo secrets)

| Secret | Value |
| --- | --- |
| `GMAIL_USER` | the Gmail address mail is sent **from** |
| `GMAIL_APP_PASSWORD` | a Google [App Password](https://myaccount.google.com/apppasswords) (16 chars) |
| `RECIPIENTS` | comma-separated recipient addresses |

Set them with:

```bash
gh secret set GMAIL_USER --body "you@gmail.com"
gh secret set GMAIL_APP_PASSWORD --body "xxxx xxxx xxxx xxxx"
gh secret set RECIPIENTS --body "you@gmail.com,wife@example.com"
```

## Test it now

Trigger a one-off run from the Actions tab → "QQQ Top 10 Weekly Email" →
"Run workflow", or:

```bash
gh workflow run qqq-weekly.yml
```

## Notes

- Data source is stockanalysis.com (reports the real QQQ ETF holdings). If both
  host endpoints fail, the run errors loudly (visible in the Actions tab) rather
  than sending stale data.
- GitHub scheduled jobs can be delayed several minutes under load; the email may
  arrive a bit after 6:05 AM.
