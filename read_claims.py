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

Read-only: prints the claim lines the online form needs
(service date YYYY-MM-DD, service, number of hours, total cost).

Usage:
  python read_claims.py FILE.xlsx --list
  python read_claims.py FILE.xlsx --client 2           # by number from --list
  python read_claims.py FILE.xlsx --client "John Doe"  # or by name
  python read_claims.py FILE.xlsx                      # asks which client
  python read_claims.py FILE.xlsx --client 1 --json out.json
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
    "personal care hours": "Personal care",
    "respite hours": "Respite",
}
CLIENT_LABEL = "client name"
RATE_LABEL = "rate $/hr"
MAX_HOURS_PER_DAY = 24


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


def read_client(ws, client, date_columns):
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
            if (hours * 4) % 1:
                warnings.append(f"{where}: {hours} hours is not a quarter-hour amount, check it.")
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
    if choice is None:
        print_client_list(clients)
        choice = input("Pick a client number: ").strip()
    if choice.isdigit() and 1 <= int(choice) <= len(clients):
        return clients[int(choice) - 1]
    matches = [c for c in clients if c["name"].lower() == choice.lower()]
    if not matches:
        matches = [c for c in clients if choice.lower() in c["name"].lower()]
    if len(matches) == 1:
        return matches[0]
    if matches:
        sys.exit(f"'{choice}' matches more than one client: " + ", ".join(c["name"] for c in matches))
    sys.exit(f"No client matches '{choice}'. Use --list to see the choices.")


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


def main():
    parser = argparse.ArgumentParser(description="Read one client's claims from the timesheet workbook.")
    parser.add_argument("workbook", help="path to the .xlsx file")
    parser.add_argument("--sheet", help="sheet name (default: first sheet)")
    parser.add_argument("--list", action="store_true", help="list the clients found and exit")
    parser.add_argument("--client", help="client number from --list, or client name")
    parser.add_argument("--json", metavar="FILE", help="also save the parsed claims as JSON")
    args = parser.parse_args()

    # data_only=True gives the values Excel last calculated instead of formula text.
    wb = openpyxl.load_workbook(args.workbook, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb.worksheets[0]

    clients = find_clients(ws)
    if not clients:
        sys.exit(f"No '{CLIENT_LABEL}' labels found on sheet '{ws.title}'.")
    if args.list:
        print_client_list(clients)
        return

    client = choose_client(clients, args.client)
    date_row, date_columns = find_date_columns(ws)
    result = read_client(ws, client, date_columns)
    print_result(result)
    if args.json:
        with open(args.json, "w") as f:
            json.dump(result, f, indent=2)
        print(f"Saved {args.json}")
    if result["problems"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
