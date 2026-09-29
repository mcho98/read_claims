# Excel timesheet to online claim form

Goal: read one client's monthly hours from the timesheet workbook and enter each day as a claim in an online form. Build it in steps. Step 1 (reading) is done. Step 2 (filling the form with Playwright) is written but untested against the real form.

## Decisions so far
- The user does not copy/paste each client into a new file. The script finds every client section and the user picks one with `--list`, `--client N|"name"`, or an interactive prompt.
- Step 2: `--url` opens a headed Chromium via Playwright; the user logs in and opens the form, then the script fills each claim and clicks "Add claim". It never submits the whole form and asks y/N before filling.
- The script loops back to the client menu after each client; "exit" quits. The workbook is read once, and the browser stays open for the whole session.
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
- `python read_claims.py FILE.xlsx --client 2` or `--client "John Doe"`; with no `--client` it shows the menu
- `--url FORM_URL` enables form filling (mock site: http://localhost:3000, a Next.js page with labelled fields); `--first-only` enters just the first claim per client for testing post-claim behavior. The real form's date label is "Service date (YYYY-MM-DD)". Field labels and button name are constants at the top of the script (Service date, Service, Number of hours, Total cost, Add claim); the Service dropdown match ignores case
- `--json out.json` saves the parsed lines for the form-filling step
- Cost = hours x rate, rounded to cents per line. Warns on hours that aren't quarter hours and stops on invalid hours or a missing rate.
- Mock result for John Doe: 18 lines, 55.53 hours, $1,889.71.

## Open questions for the user
- The form URL, and whether the field labels match the constants (the form has not been seen live).
- Whether "Add claim" resets the form or needs a wait between claims.
- The exact dropdown wording for Homemaking and Respite. The script assumes "Homemaking" and "Respite".
- Whether AB9 = 1.53 hours in the mock is a typo.
- Whether cost should be rounded per line (current behavior) or on the total. Per-line rounding comes out 2 cents higher than the sheet's total.

## Packaging for the end user (Windows, non-technical)
- Chosen approach: tkinter GUI (`app.py`) + PyInstaller one-file exe (`build.bat`, run on Windows). Docker was rejected: it needs Docker Desktop/WSL2 and makes the visible login browser harder.
- `app.py` steps: choose Excel file, pick client (preview table), Open claim website, Fill form; "Test: only the first claim" checkbox. Playwright runs on a worker thread; UI updates go through a queue.
- `FormFiller` (in `read_claims.py`) prefers installed Edge, then Chrome, then Playwright's Chromium, so the exe needs no browser download.
- Form URL lives in `config.json` next to the exe/script.
- Not yet done: building and testing the exe on Windows.

## Releasing and updating the exe
- Version lives in `version.py` (`VERSION`, `REPO = mcho98/read_claims`). Shown in the window title.
- Release: edit code, test against the mock site, bump `VERSION`, commit, `git tag vX.Y.Z && git push --tags`. `.github/workflows/build.yml` builds `ClaimFiller.exe` on windows-latest and attaches it to a GitHub Release. The repo must be public for the app to read releases without a token.
- `updater.py` (stdlib only): the app checks the latest release on start and via "Check for updates"; it downloads `ClaimFiller.new.exe`, and a generated `update.bat` swaps it in after exit, keeps `ClaimFiller.old.exe` as rollback, and relaunches. `config.json` is never replaced.
- `config.json` may also override form wording without a rebuild: `field_service_date`, `field_service`, `field_hours`, `field_cost`, `button_add` (`read_claims.apply_config`).
- Untested: the actual swap on Windows and the GitHub Actions build (needs a first tag push).
