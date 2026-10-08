#!/usr/bin/env python3
"""Read one client's monthly claims from the timesheet workbook.

Layout this expects (based on "mock data.xlsx"):
  * One date row near the top: the first real date (e.g. E2) starts the month,
    and each column to the right is the next day, up to month end.
  * One section per client, each starting at a "Client Name" label with the
    name in the cell to its right. Inside the section:
      - "rate $/hr" label with the hourly rate to its right
      - one row per service ("Homemaking hours", "Personal Care Hours",
        "Respite hours") with hours under each date column.
  A section runs until the next "Client Name" label.

Prints the claim lines the online form needs (service date YYYY-MM-DD,
service, number of hours, total cost). With --url it can also enter them in the
online form using Playwright. It clicks "Add claim" for each line but never
submits the whole form; you review and submit yourself.

After each client the script returns to the client menu. Type "exit" to quit.

Usage:
  python read_claims.py FILE.xlsx --list
  python read_claims.py FILE.xlsx --client 2           # by number from --list
  python read_claims.py FILE.xlsx --client "John Doe"  # or by name
  python read_claims.py FILE.xlsx                      # menu; "exit" quits; asks for the URL after a client is picked
  python read_claims.py FILE.xlsx --client 1 --json out.json
  python read_claims.py FILE.xlsx --start 20260810 --end 20260820   # only that date range
  python read_claims.py FILE.xlsx --url http://localhost:3000   # fill the form
  ... --url http://localhost:3000 --first-only   # enter just the first claim (testing)
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import urllib.request
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

# Row label in the sheet (lowercased) -> option text in the form's Service dropdown.
SERVICES = {
    "homemaking hours": "Homemaking",
    "personal care hours": "Personal Care",
    "respite hours": "Respite",
}
CLIENT_LABEL = "client name"
RATE_LABEL = "rate $/hr"
MAX_HOURS_PER_DAY = 24
CENT = Decimal("0.01")
MAX_CLAIMS_PER_GROUP = 10  # the website accepts at most this many claims at a time
EXIT_WORDS = {"exit", "quit", "q"}

# How the form fields are found on the page (labels from the "Claim details" form).
FIELD_SERVICE_DATE = "Service date"
FIELD_SERVICE = "Service"
FIELD_HOURS = "Number of hours"
FIELD_COST = "Total cost"
BUTTON_ADD = "Add claim"
BUTTON_PREDETERMINE = "Predetermine Claim"
# How the website's running totals are read from the page text ("Total hours: 4.00", "Total cost: $136.12").
TOTAL_HOURS_PATTERN = r"Total hours\s*:\s*([\d,]*\.?\d+)"
TOTAL_COST_PATTERN = r"Total cost\s*:\s*\$?\s*([\d,]*\.?\d+)"
TYPING_DELAY_MS = 30      # pause between keystrokes when typing hours/cost (config key typing_delay_ms)
FIELD_TRIES = 3           # times a single form box is re-filled if what it holds is not what was typed
MAX_ATTEMPTS = 3          # tries per claim before giving up for good
TOTALS_WAIT_SECONDS = 5   # how long to wait for the website totals to update after Add claim


def apply_config(config):
    """Let config.json override the form wording, so a site text change needs no rebuild."""
    global TYPING_DELAY_MS, FIELD_SERVICE_DATE, FIELD_SERVICE, FIELD_HOURS, FIELD_COST, BUTTON_ADD, BUTTON_PREDETERMINE
    FIELD_SERVICE_DATE = config.get("field_service_date", FIELD_SERVICE_DATE)
    FIELD_SERVICE = config.get("field_service", FIELD_SERVICE)
    FIELD_HOURS = config.get("field_hours", FIELD_HOURS)
    FIELD_COST = config.get("field_cost", FIELD_COST)
    BUTTON_ADD = config.get("button_add", BUTTON_ADD)
    TYPING_DELAY_MS = config.get("typing_delay_ms", TYPING_DELAY_MS)
    BUTTON_PREDETERMINE = config.get("button_predetermine", BUTTON_PREDETERMINE)


def norm(value):
    return str(value).strip().lower() if isinstance(value, str) else None


def as_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return None


def find_date_columns(ws):
    """Return [(column, date)] for each day of the month in the date row."""
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20)):
        for cell in row:
            start = as_date(cell.value)
            if start is None:
                continue
            # Following cells are formulas; use their cached value when the file
            # was saved by Excel, otherwise count up one day per column.
            columns = []
            day = start
            col = cell.column
            while day.month == start.month:
                cached = as_date(ws.cell(cell.row, col).value)
                columns.append((col, cached or day))
                col += 1
                day = day + dt.timedelta(days=1)
            return cell.row, columns
    sys.exit("Could not find the date row (no date cell in the first 20 rows).")


def find_clients(ws):
    """Return a list of client sections: name, first row, last row."""
    starts = []
    for row in ws.iter_rows():
        for cell in row:
            if norm(cell.value) == CLIENT_LABEL:
                name = ws.cell(cell.row, cell.column + 1).value
                starts.append((cell.row, str(name).strip() if name else f"(unnamed, row {cell.row})"))
    clients = []
    for i, (row, name) in enumerate(starts):
        end = starts[i + 1][0] - 1 if i + 1 < len(starts) else ws.max_row
        clients.append({"name": name, "first_row": row, "last_row": end})
    return clients


def to_money(value):
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def parse_day(text):
    """Turn YYYYMMDD (or YYYY-MM-DD) into a date; blank means no limit. Raises ValueError."""
    text = (text or "").strip()
    if not text:
        return None
    digits = text.replace("-", "")
    try:
        if len(digits) != 8:  # strptime would accept short forms like 2026081
            raise ValueError
        return dt.datetime.strptime(digits, "%Y%m%d").date()
    except ValueError:
        raise ValueError(f"'{text}' is not a valid date. Use YYYYMMDD, e.g. 20260815.")


def read_client(ws, client, date_columns, start=None, end=None):
    """start/end are optional dates; only days inside that range (inclusive) are read."""
    rate = None
    service_rows = {}
    for row in ws.iter_rows(min_row=client["first_row"], max_row=client["last_row"]):
        for cell in row:
            label = norm(cell.value)
            if label == RATE_LABEL and rate is None:
                rate = ws.cell(cell.row, cell.column + 1).value
            elif label in SERVICES and cell.row not in service_rows:
                service_rows[cell.row] = SERVICES[label]

    problems, warnings = [], []
    if not isinstance(rate, (int, float)) or rate <= 0:
        problems.append(f"No valid hourly rate next to '{RATE_LABEL}' (found {rate!r}).")
        rate = None
    if not service_rows:
        problems.append("No service rows found (Homemaking / Personal Care / Respite hours).")

    claims = []
    for row, service in sorted(service_rows.items()):
        for col, day in date_columns:
            if (start and day < start) or (end and day > end):
                continue
            hours = ws.cell(row, col).value
            if hours in (None, "", 0):
                continue
            where = ws.cell(row, col).coordinate
            if not isinstance(hours, (int, float)):
                problems.append(f"{where}: hours is not a number ({hours!r}).")
                continue
            if hours < 0 or hours > MAX_HOURS_PER_DAY:
                problems.append(f"{where}: {hours} hours is outside 0-{MAX_HOURS_PER_DAY}.")
                continue
            claims.append({
                "service_date": day.strftime("%Y-%m-%d"),
                "service": service,
                "hours": hours,
                "total_cost": to_money(hours * rate) if rate else None,
                "cell": where,
            })
    claims.sort(key=lambda c: (c["service_date"], c["service"]))
    return {"client": client["name"], "rate": rate, "claims": claims,
            "problems": problems, "warnings": warnings}


def group_claims(claims, size=MAX_CLAIMS_PER_GROUP):
    """Split claims into consecutive groups of at most `size` (15 -> 10 + 5)."""
    return [claims[i:i + size] for i in range(0, len(claims), size)]


def group_label(groups, n):
    """Human text for group n (1-based), e.g. 'Group 2 of 2 (claims 11-15)'."""
    start = sum(len(g) for g in groups[:n - 1]) + 1
    return f"Group {n} of {len(groups)} (claims {start}-{start + len(groups[n - 1]) - 1})"


def choose_client(clients, choice):
    """Return the matching client, or None if the choice is not usable."""
    if choice.isdigit() and 1 <= int(choice) <= len(clients):
        return clients[int(choice) - 1]
    matches = [c for c in clients if c["name"].lower() == choice.lower()]
    if not matches:
        matches = [c for c in clients if choice.lower() in c["name"].lower()]
    if len(matches) == 1:
        return matches[0]
    if matches:
        print(f"'{choice}' matches more than one client: " + ", ".join(c["name"] for c in matches))
    else:
        print(f"No client matches '{choice}'.")
    return None


def print_client_list(clients):
    for i, c in enumerate(clients, 1):
        print(f"  {i}. {c['name']}  (rows {c['first_row']}-{c['last_row']})")


def print_result(result):
    print(f"\nClient: {result['client']}    Rate: ${result['rate']}/hr")
    print(f"{'Service date':<14}{'Service':<16}{'Hours':>7}{'Total cost ($)':>16}   Cell")
    total_hours = total_cost = 0
    for c in result["claims"]:
        cost = "" if c["total_cost"] is None else f"{c['total_cost']:.2f}"
        print(f"{c['service_date']:<14}{c['service']:<16}{c['hours']:>7g}{cost:>16}   {c['cell']}")
        total_hours += c["hours"]
        total_cost += c["total_cost"] or 0
    print(f"{len(result['claims'])} claim lines, {total_hours:g} hours, ${total_cost:,.2f}")
    for w in result["warnings"]:
        print("WARNING:", w)
    for p in result["problems"]:
        print("PROBLEM:", p)


def find_chrome():
    """Path of the installed Google Chrome, or a RuntimeError telling the user to install it."""
    if sys.platform == "darwin":
        candidates = ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                      os.path.expanduser("~/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")]
    elif sys.platform.startswith("win"):
        candidates = [os.path.join(os.environ.get(var, ""), "Google", "Chrome", "Application", "chrome.exe")
                      for var in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA")]
    else:
        candidates = [shutil.which(n) or "" for n in ("google-chrome", "google-chrome-stable", "chrome")]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    raise RuntimeError("Google Chrome was not found. Please install Google Chrome and try again.")


def profile_dir():
    """Where this program keeps its own Chrome profile (logins, passed security checks)."""
    if sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    elif sys.platform.startswith("win"):
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    else:
        base = os.path.expanduser("~/.local/share")
    return os.path.join(base, "ClaimFiller", "chrome-profile")


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class FormFiller:
    """Drives one browser window for the whole session, so the login is kept
    while you move from client to client. Call open() first, let the user log in
    and reach the Claim details form, then call fill()."""

    def __init__(self, url, log=print):
        self.url = url
        self.log = log
        self.pw = self.browser = self.page = self.chrome_proc = None
        self.port = None
        self.screenshot_dir = os.getcwd()  # where failure screenshots go (the app points this next to the program)

    def open(self):
        """Start the user's own Google Chrome on the form address, with no automation attached.

        Chrome is started like a normal browser (just with a debugging port open), so
        Cloudflare-style checks see an ordinary visit and the person can pass them and
        log in by hand. The script only connects to the window later, in attach(),
        once fill() is called. A separate profile folder is kept between runs, so
        logins and passed checks are usually remembered.
        """
        chrome = find_chrome()
        self.port = free_port()
        os.makedirs(profile_dir(), exist_ok=True)
        self.chrome_proc = subprocess.Popen(
            [chrome, f"--remote-debugging-port={self.port}", f"--user-data-dir={profile_dir()}",
             "--no-first-run", "--no-default-browser-check", self.url],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def attach(self):
        """Connect Playwright to the running Chrome and pick the tab showing the claim form."""
        if self.page:
            return
        from playwright.sync_api import sync_playwright
        endpoint = f"http://127.0.0.1:{self.port}"
        deadline = time.time() + 15
        while True:
            try:
                urllib.request.urlopen(endpoint + "/json/version", timeout=1).close()
                break
            except Exception:
                if self.chrome_proc.poll() is not None or time.time() > deadline:
                    raise RuntimeError(
                        "Could not connect to Chrome. Close any Chrome window opened by this program "
                        "(and any other window using the Claim Filler profile), then open the website again.")
                time.sleep(0.3)
        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.connect_over_cdp(endpoint)
        pages = [p for c in self.browser.contexts for p in c.pages]
        if not pages:
            raise RuntimeError("Chrome has no open tab.")
        # Use the tab that has the form; otherwise the last tab (the one most recently opened).
        self.page = next((p for p in pages if p.get_by_role("button", name=BUTTON_ADD).count()), pages[-1])
        self.page.set_default_timeout(10000)  # fail a stuck step quickly so the retry logic can react

    def save_screenshot(self, name):
        """Keep a picture of the page when an attempt fails, to see what the form looked like."""
        if not self.screenshot_dir:
            return
        try:
            path = os.path.join(self.screenshot_dir, f"{name}.png")
            self.page.screenshot(path=path)
            self.log(f"    screenshot saved: {path}")
        except Exception:
            pass

    def read_totals(self):
        """The website's (total hours, total cost) as Decimals, or None if it shows no totals."""
        text = self.page.inner_text("body")
        hours = re.search(TOTAL_HOURS_PATTERN, text, re.I)
        cost = re.search(TOTAL_COST_PATTERN, text, re.I)
        if not (hours and cost):
            return None
        return (Decimal(hours.group(1).replace(",", "")).quantize(CENT),
                Decimal(cost.group(1).replace(",", "")).quantize(CENT))

    def wait_for_totals(self, expected):
        """Poll until the website totals equal `expected` or the wait runs out; return the last totals seen."""
        deadline = time.time() + TOTALS_WAIT_SECONDS
        seen = self.read_totals()
        while seen != expected and time.time() < deadline:
            self.page.wait_for_timeout(250)
            seen = self.read_totals()
        return seen

    # --- form boxes ----------------------------------------------------------
    def boxes(self):
        page = self.page
        return {
            "date": page.get_by_label(FIELD_SERVICE_DATE),  # label reads "Service date (YYYY-MM-DD)"
            "service": page.get_by_label(FIELD_SERVICE, exact=True),
            "hours": page.get_by_label(FIELD_HOURS, exact=True),
            "cost": page.get_by_label(FIELD_COST),  # label reads "Total cost ($)"
        }

    def values(self):
        """What the four boxes hold right now (the Service box as its visible text)."""
        b = self.boxes()
        return {
            "date": b["date"].input_value(),
            "service": b["service"].evaluate("e => e.options[e.selectedIndex].text"),
            "hours": b["hours"].input_value(),
            "cost": b["cost"].input_value(),
        }

    def describe_boxes(self):
        try:
            v = self.values()
            return f"form held date={v['date']!r} service={v['service']!r} hours={v['hours']!r} cost={v['cost']!r}"
        except Exception as e:
            return f"could not read the form ({str(e).splitlines()[0]})"

    @staticmethod
    def same(wanted, actual):
        """Equal as numbers when both are numbers (so 4 and 4.00 match), else as trimmed text."""
        try:
            return Decimal(wanted) == Decimal(actual.replace(",", "").replace("$", "").strip())
        except Exception:
            return wanted.strip().lower() == actual.strip().lower()

    def settle(self):
        """Wait until the form has stopped changing, so a late reset from the previous add
        can't land in the middle of typing the next claim."""
        try:
            self.page.wait_for_load_state("networkidle", timeout=2000)
        except Exception:
            pass
        last = None
        for _ in range(12):  # at most ~3.6 s
            now = self.values()
            if now == last:
                return
            last = now
            self.page.wait_for_timeout(300)

    def set_box(self, name, text, typed):
        """Put text in a box and check it stuck; re-fill up to FIELD_TRIES times. True if it holds the text."""
        box = self.boxes()[name]
        for _ in range(FIELD_TRIES):
            if typed:
                # Real keystrokes + Tab fire keydown/keyup/change/blur like a person typing.
                box.fill("")
                box.press_sequentially(text, delay=TYPING_DELAY_MS)
                box.press("Tab")
            else:
                box.fill(text)
                if name == "date":
                    box.press("Escape")  # close the calendar picker if it opened
            self.page.wait_for_timeout(100)
            if self.same(text, self.values()[name]):
                return True
        return False

    def enter_claim(self, c):
        self.settle()
        wanted = {"date": c["service_date"], "hours": f"{c['hours']:g}"}
        if c["total_cost"] is not None:
            wanted["cost"] = f"{c['total_cost']:.2f}"
        if not self.set_box("date", wanted["date"], typed=False):
            raise RuntimeError("the service date box would not keep its value")
        self.select_service(c["service"])
        for name in ("hours", "cost"):
            if name in wanted and not self.set_box(name, wanted[name], typed=True):
                raise RuntimeError(f"the {name} box would not keep its value")
        # Last look before clicking: a late reset or re-render may have wiped something.
        for _ in range(FIELD_TRIES):
            now = self.values()
            wrong = [n for n, w in wanted.items() if not self.same(w, now[n])]
            if now["service"].strip().lower() != c["service"].strip().lower():
                wrong.append("service")
            if not wrong:
                break
            self.log(f"    form changed while filling ({self.describe_boxes()}); filling again")
            for n in wrong:
                if n == "service":
                    self.select_service(c["service"])
                else:
                    self.set_box(n, wanted[n], typed=n in ("hours", "cost"))
        else:
            raise RuntimeError(f"form is not right and will not be submitted: {self.describe_boxes()}")
        self.page.get_by_role("button", name=BUTTON_ADD).click()

    def fill(self, claims, predetermine=True):
        """Add each claim and check it landed by comparing the website's totals with a rolling sum.

        A claim counts as added only when the website's total hours and total cost equal
        (starting totals + everything added so far). Each claim gets MAX_ATTEMPTS tries; a
        third failure stops the whole run. If the totals changed to something unexpected
        (for example a duplicate), the run stops at once, because retrying could add more
        wrong entries. When every claim is verified and predetermine is True, the
        "Predetermine Claim" button is clicked.

        Returns {"ok": bool, "added": n, "message": str}.
        """
        self.attach()
        start = self.read_totals() or (Decimal("0.00"), Decimal("0.00"))
        expected = start
        self.log(f"  website totals before adding: {start[0]} hours, ${start[1]}")
        for i, c in enumerate(claims, 1):
            label = f"{c['service_date']} {c['service']} {c['hours']:g}h"
            before = expected
            expected = (expected[0] + Decimal(str(c["hours"])).quantize(CENT),
                        expected[1] + Decimal(str(c["total_cost"] or 0)).quantize(CENT))
            for attempt in range(1, MAX_ATTEMPTS + 1):
                problem = None
                try:
                    self.enter_claim(c)
                except Exception as e:
                    problem = f"could not use the form ({str(e).splitlines()[0]})"
                seen = self.wait_for_totals(expected)
                if seen == expected:
                    self.log(f"  added {i}/{len(claims)}: {label}  (website now {seen[0]} hours, ${seen[1]})")
                    break
                if seen is not None and seen != before:
                    msg = (f"Stopped at claim {i} ({label}): the website shows {seen[0]} hours / ${seen[1]} "
                           f"but {expected[0]} hours / ${expected[1]} was expected. Check the website before continuing.")
                    self.log(msg)
                    return {"ok": False, "added": i - 1, "message": msg}
                self.log(f"  claim {i} ({label}) attempt {attempt}/{MAX_ATTEMPTS} failed: "
                         f"{problem or 'the website totals did not change'}; {self.describe_boxes()}")
                self.save_screenshot(f"failed-claim-{i}-attempt-{attempt}")
            else:
                msg = f"Stopped: claim {i} ({label}) failed {MAX_ATTEMPTS} times. {i - 1} of {len(claims)} claims were added."
                self.log(msg)
                return {"ok": False, "added": i - 1, "message": msg}
        msg = f"All {len(claims)} claims added; totals match ({expected[0]} hours, ${expected[1]})."
        if not predetermine:
            self.log("  " + msg)
        else:
            try:
                self.page.get_by_role("button", name=BUTTON_PREDETERMINE).click()
            except Exception as e:
                msg += f" But the '{BUTTON_PREDETERMINE}' button could not be clicked: {str(e).splitlines()[0]}"
                self.log("  " + msg)
                return {"ok": False, "added": len(claims), "message": msg}
            msg += f" Clicked '{BUTTON_PREDETERMINE}'."
            self.log("  " + msg)
        return {"ok": True, "added": len(claims), "message": msg}

    def select_service(self, service):
        dropdown = self.page.get_by_label(FIELD_SERVICE, exact=True)
        options = dropdown.locator("option").all_text_contents()
        # Match ignoring case, so "Personal Care" finds "Personal care".
        for text in options:
            if text.strip().lower() == service.lower():
                dropdown.select_option(label=text)
                return
        raise RuntimeError(f"Service '{service}' is not in the dropdown (options: {options}).")

    def close(self):
        """Disconnect and close the Chrome window this program opened."""
        try:
            if self.pw:
                self.pw.stop()
        except Exception:
            pass
        if self.chrome_proc and self.chrome_proc.poll() is None:
            self.chrome_proc.terminate()
        self.pw = self.browser = self.page = self.chrome_proc = None


class Session:
    """Holds the form URL and the browser, opened only when first needed and
    reused for every client after that."""

    def __init__(self, url=None):
        self.url = url
        self.filler = None

    def get_filler(self):
        """Return the open FormFiller, asking for the URL if not known yet.
        Returns None if the user skips (blank URL) or the browser can't open."""
        if self.filler:
            return self.filler
        if not self.url:
            self.url = input("Form URL (Enter to skip filling): ").strip() or None
            if not self.url:
                return None
        filler = FormFiller(self.url)
        try:
            filler.open()
        except Exception as e:
            print(f"Could not open {self.url}: {e}")
            self.url = None  # ask again next time
            return None
        input("Log in if needed and open the Claim details form in the browser, then press Enter here... ")
        self.filler = filler
        return filler

    def close(self):
        if self.filler:
            self.filler.close()


def process_client(ws, client, date_columns, args, session):
    result = read_client(ws, client, date_columns, args.start, args.end)
    print_result(result)
    if args.json:
        with open(args.json, "w") as f:
            json.dump(result, f, indent=2)
        print(f"Saved {args.json}")
    if result["problems"]:
        print("Fix the problems above before filling the form.")
        return
    if not result["claims"]:
        return
    filler = session.get_filler()
    if not filler:
        return  # no URL: back to the client menu
    groups = group_claims(result["claims"])
    default = "f" if args.first_only else "a"
    while True:
        if len(groups) > 1:
            print(f"The website takes {MAX_CLAIMS_PER_GROUP} claims at a time, so these are split into groups:")
            for n in range(1, len(groups) + 1):
                print(f"  {n}. {group_label(groups, n)}")
            pick = input(f"Which group? [1-{len(groups)}, Enter = 1, b = back to the client list] ").strip().lower() or "1"
            if pick in ("b", "back"):
                return
            if not (pick.isdigit() and 1 <= int(pick) <= len(groups)):
                print("Please type one of the group numbers.")
                continue
            group = groups[int(pick) - 1]
        else:
            group = groups[0]
        answer = input(f"Enter [a]ll {len(group)} claims (then click '{BUTTON_PREDETERMINE}'), "
                       f"[f]irst only (test, no '{BUTTON_PREDETERMINE}'), or [b]ack? [{default}] ").strip().lower() or default
        if answer in ("a", "f"):
            claims = group[:1] if answer == "f" else group
            try:
                outcome = filler.fill(claims, predetermine=(answer == "a"))
            except Exception as e:  # keep the menu alive so the other clients aren't lost
                print(f"Form filling stopped: {e}")
                return
            print(outcome["message"])
            if not outcome["ok"]:
                return  # a failed add stops everything for this client
        if len(groups) == 1:
            return  # with several groups we loop so the next group can be chosen


def main():
    parser = argparse.ArgumentParser(description="Read one client's claims from the timesheet workbook.")
    parser.add_argument("workbook", help="path to the .xlsx file")
    parser.add_argument("--sheet", help="sheet name (default: first sheet)")
    parser.add_argument("--list", action="store_true", help="list the clients found and exit")
    parser.add_argument("--client", help="client number from --list, or client name")
    parser.add_argument("--json", metavar="FILE", help="also save the parsed claims as JSON (overwritten for each client)")
    parser.add_argument("--start", metavar="YYYYMMDD", help="only entries on or after this date (default: all)")
    parser.add_argument("--end", metavar="YYYYMMDD", help="only entries on or before this date (default: all)")
    parser.add_argument("--first-only", action="store_true",
                        help="make \"first claim only\" the default answer when entering claims (to test what the form does after Add claim)")
    parser.add_argument("--url", help="online claim form URL (if omitted, asked for after you pick a client)")
    args = parser.parse_args()
    try:
        args.start, args.end = parse_day(args.start), parse_day(args.end)
    except ValueError as e:
        parser.error(str(e))
    if args.start and args.end and args.start > args.end:
        parser.error("--start is after --end.")

    # data_only=True gives the values Excel last calculated instead of formula text.
    wb = openpyxl.load_workbook(args.workbook, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb.worksheets[0]

    clients = find_clients(ws)
    if not clients:
        sys.exit(f"No '{CLIENT_LABEL}' labels found on sheet '{ws.title}'.")
    if args.list:
        print_client_list(clients)
        return

    date_row, date_columns = find_date_columns(ws)
    session = Session(args.url)
    choice = args.client
    try:
        while True:
            if choice is None:
                print()
                print_client_list(clients)
                choice = input('Pick a client number or name ("exit" to quit): ').strip()
            if choice.lower() in EXIT_WORDS:
                break
            client = choose_client(clients, choice) if choice else None
            choice = None
            if client:
                process_client(ws, client, date_columns, args, session)
    except (EOFError, KeyboardInterrupt):
        print()
    finally:
        session.close()


if __name__ == "__main__":
    main()
