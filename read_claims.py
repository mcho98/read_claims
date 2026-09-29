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
import sys
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
EXIT_WORDS = {"exit", "quit", "q"}

# How the form fields are found on the page (labels from the "Claim details" form).
FIELD_SERVICE_DATE = "Service date"
FIELD_SERVICE = "Service"
FIELD_HOURS = "Number of hours"
FIELD_COST = "Total cost"
BUTTON_ADD = "Add claim"


def apply_config(config):
    """Let config.json override the form wording, so a site text change needs no rebuild."""
    global FIELD_SERVICE_DATE, FIELD_SERVICE, FIELD_HOURS, FIELD_COST, BUTTON_ADD
    FIELD_SERVICE_DATE = config.get("field_service_date", FIELD_SERVICE_DATE)
    FIELD_SERVICE = config.get("field_service", FIELD_SERVICE)
    FIELD_HOURS = config.get("field_hours", FIELD_HOURS)
    FIELD_COST = config.get("field_cost", FIELD_COST)
    BUTTON_ADD = config.get("button_add", BUTTON_ADD)


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


class FormFiller:
    """Drives one browser window for the whole session, so the login is kept
    while you move from client to client. Call open() first, let the user log in
    and reach the Claim details form, then call fill()."""

    def __init__(self, url, log=print):
        self.url = url
        self.log = log
        self.pw = self.browser = self.page = None

    def open(self):
        from playwright.sync_api import sync_playwright
        self.pw = sync_playwright().start()
        # Prefer the Edge/Chrome already installed (no browser download needed);
        # fall back to Playwright's own Chromium if neither is present.
        for channel in ("msedge", "chrome", None):
            try:
                self.browser = self.pw.chromium.launch(headless=False, channel=channel)
                break
            except Exception:
                if channel is None:
                    self.close()
                    raise
        self.page = self.browser.new_page()
        self.page.goto(self.url)

    def fill(self, claims):
        page = self.page
        for i, c in enumerate(claims, 1):
            date_box = page.get_by_label(FIELD_SERVICE_DATE)  # label reads "Service date (YYYY-MM-DD)"
            date_box.fill(c["service_date"])
            date_box.press("Escape")  # close the calendar picker if it opened
            self.select_service(c["service"])
            page.get_by_label(FIELD_HOURS, exact=True).fill(f"{c['hours']:g}")
            if c["total_cost"] is not None:
                page.get_by_label(FIELD_COST).fill(f"{c['total_cost']:.2f}")  # label reads "Total cost ($)"
            page.get_by_role("button", name=BUTTON_ADD).click()
            self.log(f"  added {i}/{len(claims)}: {c['service_date']} {c['service']} {c['hours']:g}h")

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
        if self.browser:
            self.browser.close()
        if self.pw:
            self.pw.stop()
        self.pw = self.browser = self.page = None


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
    default = "f" if args.first_only else "a"
    n = len(result["claims"])
    answer = input(f"Enter [a]ll {n} claims, [f]irst only, or [b]ack to the client list? [{default}] ").strip().lower() or default
    if answer in ("a", "f"):
        claims = result["claims"][:1] if answer == "f" else result["claims"]
        try:
            filler.fill(claims)
        except Exception as e:  # keep the menu alive so the other clients aren't lost
            print(f"Form filling stopped: {e}")


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
