# Excel timesheet to online claim form

Goal: read one client's monthly hours from the timesheet workbook and enter each day as a claim in an online form. Build it in steps. Step 1 (reading) is done. Step 2 (filling the form) has not started.

## Decisions so far
- The user does not copy/paste each client into a new file. The script finds every client section and the user picks one with `--list`, `--client N|"name"`, or an interactive prompt.
- Reading only for now. Nothing fills or submits the form yet.
- Python with openpyxl, kept simple.

## Workbook layout (from the mock file)
- One date row near the top. The first date cell (E2 = 2026-08-01) starts the month. Columns to the right are formulas adding one day, up to month end (E:AI).
- Each client section starts at a "Client Name" label with the name in the next cell (B5/C5). A section ends where the next "Client Name" starts. The real file has one date row, and this structure repeats for each client.
- Inside a section: "rate $/hr" with the rate in the next cell (C7/D7 = 34.03), then the rows "Homemaking hours", "Personal Care Hours" and "Respite hours" (column B), with hours under each date column.
- The rows below (totals, $ per day, half-month reconciliation, several #REF! formulas) are ignored.

## Target form ("Claim details")
- Service date (YYYY-MM-DD): text box with a calendar picker
- Service: dropdown (the screenshot shows "Personal care")
- Number of hours
- Total cost ($)
- "Add claim" button, one claim per day and service

## read_claims.py
- `python read_claims.py FILE.xlsx --list`
- `python read_claims.py FILE.xlsx --client 2` or `--client "John Doe"`; with no `--client` it asks
- `--json out.json` saves the parsed lines for the form-filling step
- Cost = hours x rate, rounded to cents per line. Warns on hours that aren't quarter hours and stops on invalid hours or a missing rate.
- Mock result for John Doe: 18 lines, 55.53 hours, $1,889.71.

## Open questions for the user
- The exact dropdown wording for Homemaking and Respite. The script assumes "Homemaking" and "Respite".
- Whether AB9 = 1.53 hours in the mock is a typo. (user context: it is not a typo, non-quarter amounts for hours is allowed)
- Whether cost should be rounded per line (current behavior) or on the total. Per-line rounding comes out 2 cents higher than the sheet's total. (user context: do not round. allow the final online form to handle rounding)
