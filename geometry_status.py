#!/usr/bin/env python3
"""
Daily Geometry Sprint status email.

Fetches the Geometry Sprint Plan Google Sheet (exported as .xlsx), reads the
Dashboard + Daily Checklist tabs, and emails a plain-language status summary
to the family via Gmail SMTP. Sent every day at 6 PM Pacific by GitHub Actions.

Environment variables:
  GMAIL_USER              - the Gmail address that sends the mail
  GMAIL_APP_PASSWORD      - a Gmail app password for that account
  GEOMETRY_RECIPIENTS     - comma-separated recipient addresses
  GEOMETRY_SHEET_URL      - full URL that returns the workbook as .xlsx
                            (e.g. https://docs.google.com/spreadsheets/d/<id>/export?format=xlsx)
  DRY_RUN=1               - parse and print the email, but do NOT send
  FORCE_SEND=1            - bypass the 6 PM Pacific time guard
  EVENT_NAME              - GitHub event name; "workflow_dispatch" also bypasses the guard
"""

import io
import os
import re
import smtplib
import sys
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import requests
import openpyxl

PACIFIC = ZoneInfo("America/Los_Angeles")
CATEGORIES = ("Quizzes", "Tests", "Assignments", "Final Exam")

# Fallbacks used only if the corresponding text can't be found in the sheet.
DEFAULT_DEADLINE = date(2026, 8, 2)
DEFAULT_TARGET = date(2026, 7, 10)


# --------------------------------------------------------------------------- #
# Fetch + low-level cell access                                               #
# --------------------------------------------------------------------------- #
def fetch_workbook(url):
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    ctype = resp.headers.get("Content-Type", "")
    if "html" in ctype.lower():
        raise RuntimeError(
            "The sheet URL returned an HTML page, not an .xlsx file. This usually "
            "means the sheet is not readable without signing in. Make it viewable "
            "via 'Anyone with the link' (or publish it) so the job can read it.\n"
            f"URL: {url}"
        )
    return openpyxl.load_workbook(io.BytesIO(resp.content), data_only=True)


def sheet_rows(ws):
    """Return the worksheet as a list of rows, each a list of trimmed strings."""
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(["" if c is None else str(c).strip() for c in row])
    return rows


def cell_text(rows):
    """Flatten every cell of a sheet into one lowercase blob for keyword tests."""
    return " ".join(c for row in rows for c in row).lower()


# --------------------------------------------------------------------------- #
# Parsing helpers                                                             #
# --------------------------------------------------------------------------- #
def as_int(value):
    m = re.search(r"-?\d+", value.replace(",", ""))
    return int(m.group()) if m else None


def as_percent(value):
    """Normalise a progress cell to an int percentage (0-100), or None."""
    value = value.strip()
    if not value:
        return None
    if value.endswith("%"):
        n = re.search(r"-?\d+(?:\.\d+)?", value)
        return round(float(n.group())) if n else None
    try:
        f = float(value)
    except ValueError:
        return None
    # Percent-formatted cells come back as a fraction (1.0, 0.09).
    return round(f * 100) if f <= 1.0 else round(f)


def find_dashboard(workbook):
    for ws in workbook.worksheets:
        blob = cell_text(sheet_rows(ws))
        if "progress" in blob and ("part 1" in blob or "part 2" in blob):
            return ws
    return None


def find_checklist(workbook):
    for ws in workbook.worksheets:
        blob = cell_text(sheet_rows(ws))
        if "task" in blob and "done" in blob and ("day" in blob and "date" in blob):
            return ws
    return None


def parse_dashboard(ws):
    """
    Returns dict: {
      'parts': {1: {'progress': int, 'cats': [(name, total, completed, remaining)]},
                2: {...}},
      'deadline': date|None, 'target': date|None
    }
    """
    rows = sheet_rows(ws)
    parts = {1: {"progress": None, "cats": []}, 2: {"progress": None, "cats": []}}
    current = None
    deadline = target = None

    for row in rows:
        joined = " ".join(row)
        low = joined.lower()

        if deadline is None:
            deadline = _find_date(joined, after=r"deadline")
        if target is None:
            target = _find_date(joined, after=r"done by")

        if re.search(r"\bpart\s*1\b", low):
            current = 1
        elif re.search(r"\bpart\s*2\b", low):
            current = 2

        # Progress rows, e.g. "P1 Progress:" | "100%"
        pm = re.search(r"\bp([12])\s*progress", low)
        if pm:
            pnum = int(pm.group(1))
            for cell in row:
                pct = as_percent(cell)
                if pct is not None and "progress" not in cell.lower():
                    parts[pnum]["progress"] = pct
                    break
            continue

        # Category rows: name in col, followed by total / completed / remaining ints.
        if current and row and row[0] in CATEGORIES:
            nums = [as_int(c) for c in row[1:] if as_int(c) is not None]
            if len(nums) >= 3:
                total, completed, remaining = nums[0], nums[1], nums[2]
                parts[current]["cats"].append(
                    (row[0], total, completed, remaining)
                )

    return {"parts": parts, "deadline": deadline, "target": target}


MONTHS = {m.lower(): i for i, m in enumerate(
    ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}


def _find_date(text, after):
    """Find a 'Mon DD, YYYY' / 'July 10th, 2026' style date appearing after a keyword."""
    m = re.search(
        after + r".*?([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})",
        text, re.IGNORECASE)
    if not m:
        return None
    mon = MONTHS.get(m.group(1)[:3].lower())
    if not mon:
        return None
    try:
        return date(int(m.group(3)), mon, int(m.group(2)))
    except ValueError:
        return None


def parse_checklist_today(ws, today):
    """Return (day_label, [(task, done_bool)]) for today, or (None, []) if none."""
    rows = sheet_rows(ws)
    # Locate header to find the Task and Done column indexes.
    task_col = done_col = date_col = day_col = None
    header_idx = None
    for i, row in enumerate(rows):
        low = [c.lower() for c in row]
        if "task" in low and any("done" in c for c in low):
            header_idx = i
            for j, c in enumerate(low):
                if c == "task":
                    task_col = j
                elif "done" in c:
                    done_col = j
                elif c == "date":
                    date_col = j
                elif c == "day":
                    day_col = j
            break
    if header_idx is None or task_col is None:
        return None, []

    today_keys = {
        today.strftime("%a %b %d"),
        today.strftime("%a %b %-d") if hasattr(today, "strftime") else "",
        f"{today.strftime('%a')} {today.strftime('%b')} {today.day}",
    }
    today_keys = {k for k in today_keys if k}

    cur_date = ""
    day_label = None
    tasks = []
    matching = False
    for row in rows[header_idx + 1:]:
        raw_date = row[date_col].strip() if date_col is not None and date_col < len(row) else ""
        if raw_date:
            cur_date = raw_date
            matching = cur_date in today_keys
            if matching and day_col is not None and day_col < len(row):
                day_label = f"Day {row[day_col]} — {cur_date}" if row[day_col] else cur_date
            elif matching:
                day_label = cur_date
        if matching:
            task = row[task_col].strip() if task_col < len(row) else ""
            if task:
                done_raw = ""
                if done_col is not None and done_col < len(row):
                    done_raw = row[done_col].strip().lower()
                done = done_raw in ("true", "1", "yes", "x", "✓", "done")
                tasks.append((task, done))
    return day_label, tasks


# --------------------------------------------------------------------------- #
# Render                                                                       #
# --------------------------------------------------------------------------- #
def _part_line(cats):
    done = sum(c for _, _, c, _ in cats)
    total = sum(t for _, t, _, _ in cats)
    remaining = sum(r for _, _, _, r in cats)
    return done, total, remaining


def render(dash, day_label, tasks, today):
    parts = dash["parts"]
    deadline = dash["deadline"] or DEFAULT_DEADLINE
    target = dash["target"] or DEFAULT_TARGET
    days_left = (deadline - today).days

    p1, p2 = parts[1], parts[2]
    p1_prog = p1["progress"]
    p2_prog = p2["progress"]

    subject = (
        f"Geometry Sprint — {today.strftime('%a %b %d')}: "
        f"Part 1 {fmt_pct(p1_prog)}, Part 2 {fmt_pct(p2_prog)}"
    )

    # ---- plain text ----
    lines = [f"Geometry Sprint Plan — status for {today.strftime('%A, %B %d, %Y')}", ""]
    lines.append(f"Deadline: {deadline.strftime('%b %d, %Y')}  ({days_left} days left)")
    lines.append(f"Target to finish: {target.strftime('%b %d, %Y')}")
    lines.append("")
    for pnum, label in ((1, "Part 1"), (2, "Part 2")):
        p = parts[pnum]
        lines.append(f"{label}: {fmt_pct(p['progress'])} complete")
        for name, total, completed, remaining in p["cats"]:
            flag = "" if remaining == 0 else f"  ({remaining} left)"
            lines.append(f"   - {name}: {completed}/{total}{flag}")
        lines.append("")

    lines.append("Today's checklist:")
    if tasks:
        if day_label:
            lines.append(f"  {day_label}")
        for task, done in tasks:
            lines.append(f"    [{'x' if done else ' '}] {task}")
    else:
        lines.append("  No tasks scheduled for today on the checklist.")
    lines.append("")
    lines.append("(Automated daily update, 6 PM PT.)")
    text = "\n".join(lines)

    # ---- html ----
    def cat_rows(cats):
        out = []
        for name, total, completed, remaining in cats:
            color = "#137333" if remaining == 0 else "#b06000"
            left = "done" if remaining == 0 else f"{remaining} left"
            out.append(
                f"<tr><td style='padding:2px 10px 2px 0'>{name}</td>"
                f"<td style='padding:2px 10px 2px 0'>{completed}/{total}</td>"
                f"<td style='padding:2px 0;color:{color}'>{left}</td></tr>"
            )
        return "".join(out)

    def part_block(pnum, label):
        p = parts[pnum]
        return (
            f"<h3 style='margin:16px 0 4px'>{label}: "
            f"<span style='color:#1a73e8'>{fmt_pct(p['progress'])}</span> complete</h3>"
            f"<table style='border-collapse:collapse;font-size:14px'>{cat_rows(p['cats'])}</table>"
        )

    if tasks:
        task_items = "".join(
            f"<li style='{'color:#137333' if done else ''}'>"
            f"{'✅' if done else '⬜'} {task}</li>"
            for task, done in tasks
        )
        today_html = (
            (f"<div style='color:#555;font-size:13px'>{day_label}</div>" if day_label else "")
            + f"<ul style='margin:4px 0'>{task_items}</ul>"
        )
    else:
        today_html = "<p style='color:#777'>No tasks scheduled for today on the checklist.</p>"

    html = f"""\
<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:560px">
  <h2 style="margin:0 0 2px">📐 Geometry Sprint — {today.strftime('%A, %b %d')}</h2>
  <div style="color:#555;font-size:14px">
    Deadline <b>{deadline.strftime('%b %d, %Y')}</b> ({days_left} days left)
    · target to finish <b>{target.strftime('%b %d, %Y')}</b>
  </div>
  {part_block(1, "Part 1")}
  {part_block(2, "Part 2")}
  <h3 style="margin:16px 0 4px">Today's checklist</h3>
  {today_html}
  <p style="color:#999;font-size:12px;margin-top:18px">Automated daily update · 6 PM PT</p>
</div>"""
    return subject, text, html


def fmt_pct(p):
    return "—" if p is None else f"{p}%"


# --------------------------------------------------------------------------- #
# Send                                                                         #
# --------------------------------------------------------------------------- #
def send_email(subject, text, html):
    user = os.environ["GMAIL_USER"]
    password = os.environ["GMAIL_APP_PASSWORD"]
    recipients = [a.strip() for a in os.environ["GEOMETRY_RECIPIENTS"].split(",") if a.strip()]
    if not recipients:
        raise RuntimeError("GEOMETRY_RECIPIENTS is empty")

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
    today = now.date()
    dry_run = os.environ.get("DRY_RUN") == "1"
    force = (
        os.environ.get("FORCE_SEND") == "1"
        or os.environ.get("EVENT_NAME") == "workflow_dispatch"
        or dry_run
    )

    # GitHub cron is UTC and DST-unaware, so the workflow fires at two UTC times
    # year-round; only the one that lands in the 6 PM Pacific hour proceeds.
    if not force and now.hour != 18:
        print(f"Skipping: Pacific time is {now:%H:%M %Z}, not the 6 PM window.")
        return 0

    url = os.environ["GEOMETRY_SHEET_URL"]
    wb = fetch_workbook(url)

    dash_ws = find_dashboard(wb)
    if dash_ws is None:
        raise RuntimeError("Could not find the Dashboard tab in the workbook.")
    dash = parse_dashboard(dash_ws)

    chk_ws = find_checklist(wb)
    day_label, tasks = parse_checklist_today(chk_ws, today) if chk_ws else (None, [])

    subject, text, html = render(dash, day_label, tasks, today)
    print("SUBJECT:", subject)
    print(text)

    if dry_run:
        print("\n[DRY_RUN] Not sending.")
        return 0

    send_email(subject, text, html)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
