#!/usr/bin/env python3
"""
Fetch the top 10 holdings of the Invesco QQQ ETF and email them.

Data source: stockanalysis.com's holdings API, which reports the ACTUAL QQQ
ETF holdings (name, ticker, % of fund). Two host endpoints are tried for
resilience; whichever responds first wins.

Email is sent via Gmail SMTP using an app password.

Week-over-week change: the previous send's top-10 tickers are stored in
state/last_week.json. Each run diffs the current top 10 against it to fill a
"Change from last week" column ("No change" / "New. Replaces <TICKER>"). The
GitHub workflow commits the refreshed state file back to the repo after a send.

Required environment variables:
  GMAIL_USER          - the Gmail address the mail is sent FROM
  GMAIL_APP_PASSWORD  - a Google app password (16 chars, spaces ok)
  RECIPIENTS          - comma-separated list of recipient addresses
Optional:
  FORCE_SEND=1        - bypass the 6 AM Pacific time guard (manual runs)
"""

import json
import os
import smtplib
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import requests

PACIFIC = ZoneInfo("America/Los_Angeles")
STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "state", "last_week.json")
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
# Week-over-week state                                                         #
# --------------------------------------------------------------------------- #
def load_prev_tickers():
    """Return the previous send's ordered top-10 tickers, or None if first run."""
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            data = json.load(f)
        tickers = data.get("tickers") or []
        return tickers or None
    except (FileNotFoundError, ValueError):
        return None


def save_state(rows, now):
    """Persist this week's top-10 tickers so next week can diff against them."""
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    payload = {
        "date": now.strftime("%Y-%m-%d"),
        "tickers": [r["ticker"] for r in rows],
    }
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)
        f.write("\n")


def compute_changes(rows, prev_tickers):
    """Map each current ticker to its 'Change from last week' label.

    - First run (no prior data): every row is "No change" (baseline).
    - In both weeks: "No change".
    - New this week: "New. Replaces <TICKER>", pairing each new entrant with a
      ticker that dropped out of the top 10 (ordered, for the rare multi-change
      week). Top-10 size is fixed, so #new == #dropped.
    """
    current = [r["ticker"] for r in rows]
    if not prev_tickers:
        return {t: "No change" for t in current}, True

    prev_set, cur_set = set(prev_tickers), set(current)
    dropped = [t for t in prev_tickers if t not in cur_set]
    new = [t for t in current if t not in prev_set]
    pairing = {nt: dropped[i] if i < len(dropped) else "—"
               for i, nt in enumerate(new)}

    changes = {}
    for t in current:
        if t in prev_set:
            changes[t] = "No change"
        else:
            changes[t] = f"New. Replaces {pairing[t]}"
    return changes, False


# --------------------------------------------------------------------------- #
# Email rendering                                                              #
# --------------------------------------------------------------------------- #
def render(rows, source, now, changes, is_first):
    date_str = now.strftime("%b %d, %Y")
    ts = now.strftime("%b %d, %Y %I:%M %p %Z")
    total = sum(r["weight"] for r in rows)

    # plain text
    lines = [f"QQQ Top 10 Holdings — Week of {date_str}", ""]
    lines.append(f"{'#':>2}  {'Company':<26} {'Ticker':<7} {'% of QQQ':>8}  "
                 f"{'Change from last week':<26}")
    lines.append("-" * 78)
    for i, r in enumerate(rows, 1):
        lines.append(f"{i:>2}  {r['name'][:26]:<26} {r['ticker']:<7} "
                     f"{r['weight']:>7.2f}%  {changes.get(r['ticker'], ''):<26}")
    lines.append("-" * 78)
    lines.append(f"{'':>2}  {'Top 10 combined':<26} {'':<7} {total:>7.2f}%")
    lines.append("")
    if is_first:
        lines.append("Note: first send — this week is the baseline; real "
                     "week-over-week changes start next Friday.")
    lines.append(f"Source: {source}")
    lines.append(f"Retrieved: {ts}")
    text = "\n".join(lines)

    # html
    trs = ""
    for i, r in enumerate(rows, 1):
        bg = "#ffffff" if i % 2 else "#f5f7fa"
        change = changes.get(r["ticker"], "")
        # Green-highlight new entrants so they jump out; grey for "No change".
        is_new = change.startswith("New")
        change_style = ("color:#0a7d33;font-weight:600" if is_new
                        else "color:#999")
        trs += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:8px 12px;text-align:right;color:#888">{i}</td>'
            f'<td style="padding:8px 12px">{r["name"]}</td>'
            f'<td style="padding:8px 12px;font-weight:600">{r["ticker"]}</td>'
            f'<td style="padding:8px 12px;text-align:right">{r["weight"]:.2f}%</td>'
            f'<td style="padding:8px 12px;{change_style}">{change}</td>'
            f"</tr>"
        )
    first_note = (
        '<div style="color:#666;font-size:12px;margin-top:10px">'
        "Note: first send — this week is the baseline; real week-over-week "
        "changes start next Friday.</div>" if is_first else "")
    html = f"""\
<div style="font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1a1a1a;max-width:680px">
  <h2 style="margin:0 0 4px">QQQ Top 10 Holdings</h2>
  <div style="color:#666;margin-bottom:16px">Week of {date_str}</div>
  <table style="border-collapse:collapse;width:100%;font-size:14px;border:1px solid #e2e6ea">
    <thead>
      <tr style="background:#0b3d91;color:#fff;text-align:left">
        <th style="padding:8px 12px;text-align:right">#</th>
        <th style="padding:8px 12px">Company</th>
        <th style="padding:8px 12px">Ticker</th>
        <th style="padding:8px 12px;text-align:right">% of QQQ</th>
        <th style="padding:8px 12px">Change from last week</th>
      </tr>
    </thead>
    <tbody>
      {trs}
      <tr style="background:#eef1f5;font-weight:700">
        <td style="padding:8px 12px"></td>
        <td style="padding:8px 12px">Top 10 combined</td>
        <td style="padding:8px 12px"></td>
        <td style="padding:8px 12px;text-align:right">{total:.2f}%</td>
        <td style="padding:8px 12px"></td>
      </tr>
    </tbody>
  </table>
  {first_note}
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

    prev_tickers = load_prev_tickers()
    changes, is_first = compute_changes(rows, prev_tickers)

    subject, text, html = render(rows, source, now, changes, is_first)
    print(text)
    send_email(subject, text, html)

    # Persist this week's holdings only after a successful send, so next week
    # has a baseline to diff against. The workflow commits the file back.
    save_state(rows, now)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # noqa: BLE001
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
