#!/usr/bin/env python3
"""
Fetch the top 10 holdings of the Invesco QQQ ETF and email them.

Data source: stockanalysis.com's holdings API, which reports the ACTUAL QQQ
ETF holdings (name, ticker, % of fund). Two host endpoints are tried for
resilience; whichever responds first wins.

Email is sent via Gmail SMTP using an app password.

Required environment variables:
  GMAIL_USER          - the Gmail address the mail is sent FROM
  GMAIL_APP_PASSWORD  - a Google app password (16 chars, spaces ok)
  RECIPIENTS          - comma-separated list of recipient addresses
Optional:
  FORCE_SEND=1        - bypass the 6 AM Pacific time guard (manual runs)
"""

import os
import smtplib
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import requests

PACIFIC = ZoneInfo("America/Los_Angeles")
UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
HEADERS = {"User-Agent": UA, "Accept": "text/html,application/xhtml+xml,*/*"}
TIMEOUT = 30


# --------------------------------------------------------------------------- #
# Data sources                                                                 #
# --------------------------------------------------------------------------- #
def _clean_weight(value):
    """Parse a weight string like '8.93%' or '8.93' into a float percent."""
    s = str(value).strip().replace("%", "").replace(",", "")
    return float(s)


STOCKANALYSIS_HOSTS = (
    "https://stockanalysis.com/api/symbol/e/QQQ/holdings",
    "https://api.stockanalysis.com/api/symbol/e/QQQ/holdings",
)


def _parse_stockanalysis(url):
    """Read QQQ holdings JSON from a stockanalysis host endpoint.

    Response shape:
      {"status":200,"data":{"holdings":[
          {"no":1,"n":"NVIDIA Corporation","s":"$NVDA","as":"8.14%","sh":"..."}, ...
      ]}}
    """
    r = requests.get(url, headers={**HEADERS, "Accept": "application/json"},
                     timeout=TIMEOUT)
    r.raise_for_status()
    payload = r.json()
    holdings = payload.get("data", {}).get("holdings", [])
    rows = []
    for h in holdings:
        name = h.get("n") or h.get("name")
        ticker = (h.get("s") or h.get("symbol") or "").lstrip("$")
        weight = h.get("as") if h.get("as") is not None else h.get("weight")
        if name is None or weight is None:
            continue
        try:
            w = _clean_weight(weight)
        except (ValueError, TypeError):
            continue
        rows.append({"name": name, "ticker": ticker, "weight": w})
    rows.sort(key=lambda x: x["weight"], reverse=True)
    if len(rows) < 10:
        raise ValueError(f"only {len(rows)} holdings parsed")
    return rows[:10], "stockanalysis.com — Invesco QQQ Trust holdings"


def fetch_top10():
    """Return (rows, source, errors). Tries each host; first valid wins."""
    errors = []
    for url in STOCKANALYSIS_HOSTS:
        try:
            rows, source = _parse_stockanalysis(url)
            total = sum(r["weight"] for r in rows)
            # Sanity band: QQQ's top 10 typically sum to ~45-60%.
            if not (30 <= total <= 75):
                raise ValueError(f"weights look off (top-10 sum={total:.1f}%)")
            return rows, source, errors
        except Exception as e:  # noqa: BLE001 - try the next host
            host = url.split("/")[2]
            errors.append(f"{host}: {e}")
    raise RuntimeError("All data sources failed:\n  " + "\n  ".join(errors))


# --------------------------------------------------------------------------- #
# Email rendering                                                              #
# --------------------------------------------------------------------------- #
def render(rows, source, now):
    date_str = now.strftime("%b %d, %Y")
    ts = now.strftime("%b %d, %Y %I:%M %p %Z")
    total = sum(r["weight"] for r in rows)

    # plain text
    lines = [f"QQQ Top 10 Holdings — Week of {date_str}", ""]
    lines.append(f"{'#':>2}  {'Company':<28} {'Ticker':<8} {'% of QQQ':>8}")
    lines.append("-" * 52)
    for i, r in enumerate(rows, 1):
        lines.append(f"{i:>2}  {r['name'][:28]:<28} {r['ticker']:<8} {r['weight']:>7.2f}%")
    lines.append("-" * 52)
    lines.append(f"{'':>2}  {'Top 10 combined':<28} {'':<8} {total:>7.2f}%")
    lines.append("")
    lines.append(f"Source: {source}")
    lines.append(f"Retrieved: {ts}")
    text = "\n".join(lines)

    # html
    trs = ""
    for i, r in enumerate(rows, 1):
        bg = "#ffffff" if i % 2 else "#f5f7fa"
        trs += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:8px 12px;text-align:right;color:#888">{i}</td>'
            f'<td style="padding:8px 12px">{r["name"]}</td>'
            f'<td style="padding:8px 12px;font-weight:600">{r["ticker"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{r["weight"]:.2f}%</td>'
            f"</tr>"
        )
    html = f"""\
<div style="font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1a1a1a;max-width:560px">
  <h2 style="margin:0 0 4px">QQQ Top 10 Holdings</h2>
  <div style="color:#666;margin-bottom:16px">Week of {date_str}</div>
  <table style="border-collapse:collapse;width:100%;font-size:14px;border:1px solid #e2e6ea">
    <thead>
      <tr style="background:#0b3d91;color:#fff;text-align:left">
        <th style="padding:8px 12px;text-align:right">#</th>
        <th style="padding:8px 12px">Company</th>
        <th style="padding:8px 12px">Ticker</th>
        <th style="padding:8px 12px;text-align:right">% of QQQ</th>
      </tr>
    </thead>
    <tbody>
      {trs}
      <tr style="background:#eef1f5;font-weight:700">
        <td style="padding:8px 12px"></td>
        <td style="padding:8px 12px">Top 10 combined</td>
        <td style="padding:8px 12px"></td>
        <td style="padding:8px 12px;text-align:right">{total:.2f}%</td>
      </tr>
    </tbody>
  </table>
  <div style="color:#888;font-size:12px;margin-top:12px">
    Source: {source}<br>Retrieved: {ts}
  </div>
</div>"""
    subject = f"QQQ Top 10 Holdings — Week of {date_str}"
    return subject, text, html


def send_email(subject, text, html):
    user = os.environ["GMAIL_USER"]
    password = os.environ["GMAIL_APP_PASSWORD"]
    recipients = [a.strip() for a in os.environ["RECIPIENTS"].split(",") if a.strip()]
    if not recipients:
        raise RuntimeError("RECIPIENTS is empty")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(text, "plain"))
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(user, password.replace(" ", ""))
        server.sendmail(user, recipients, msg.as_string())
    print(f"Sent to: {', '.join(recipients)}")


# --------------------------------------------------------------------------- #
# Main                                                                         #
# --------------------------------------------------------------------------- #
def main():
    now = datetime.now(PACIFIC)
    force = (
        os.environ.get("FORCE_SEND") == "1"
        or os.environ.get("EVENT_NAME") == "workflow_dispatch"
    )

    # Time guard: GitHub cron is UTC and DST-unaware, so the workflow fires at
    # two UTC times year-round; only the one that lands on 6 AM Pacific proceeds.
    if not force and now.hour != 6:
        print(f"Skipping: Pacific time is {now:%H:%M %Z}, not the 6 AM window.")
        return 0

    rows, source, errors = fetch_top10()
    if errors:
        print("Note: some sources failed before success:\n  " + "\n  ".join(errors))

    subject, text, html = render(rows, source, now)
    print(text)
    send_email(subject, text, html)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # noqa: BLE001
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
