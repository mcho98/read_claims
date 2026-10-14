# Claim Filler — Chrome extension

JavaScript version of `../python-app`. Everything happens inside Chrome: read the timesheet
workbook, pick a client, then fill the claim form in your own browser tab (no Playwright, no
debugging port, no separate profile).

## Install
From a release: download `ClaimFiller-extension.zip` from the repo's Releases page, unzip it, then follow steps 1-3 below with the unzipped folder. From a clone: use this `extension` folder.

## Install steps
1. Chrome → `chrome://extensions` → turn on **Developer mode**.
2. **Load unpacked** → choose this `extension` folder.
3. Click the Claim Filler toolbar icon: the side panel opens and stays open while you work.

## Use
1. **Choose Excel file…** (or **open an exported JSON…** made by the Python app / this extension).
2. Pick a client; optionally set From / To as `YYYYMMDD`. Claims are shown in groups of 10.
3. Paste the claim website address and click **Open claim website** (Chrome asks once for permission to use that site). Log in and open the **Claim details** form.
4. Choose the group, then **Fill form**. Each claim is verified against the website's running total cost; a claim is retried up to 3 times, then everything stops. When the whole group matches, **Predetermine** is clicked. The "Test: only the first claim" box enters one claim and never clicks Predetermine.

## Files
- `src/core.js` — workbook reading, grouping, JSON export (port of `read_claims.py`; pure, testable).
- `src/content.js` — the form-filling engine, injected into the claim page (port of `FormFiller`).
- `panel.html/.css/.js` — side panel UI. `background.js` — opens the panel.
- `vendor/xlsx.full.min.js` — SheetJS 0.18.5 (Apache-2.0, see `vendor/xlsx-LICENSE.txt`); Chrome extensions cannot load it from a CDN.
- `shared/` (repo root) — `mock_data.xlsx` and the `out.json` the Python app made from it.

## Log
The panel log has time stamps for every step (tab found, form values, each box typed, how long the page refresh took, the total after each Add, any message the site shows). Use **Copy log** / **Save log…** to send it when something goes wrong.

## Tests
- `npm test` — the JS reader must reproduce `../shared/out.json` exactly from `../shared/mock_data.xlsx`, plus date range, grouping, rounding and date parsing.
- `python3 test/panel_flow_test.py [--first-only]` — loads the extension into Chromium and runs the whole panel flow against `test/real_like_site.py` (a local site that reloads after Add and refreshes the form after date/service changes).
- `python3 test/engine_vs_mock_site.py` — runs `src/content.js` against the mock website (needs internet and Python Playwright).

## Form wording
Defaults match the real site: labels "Service date", "Service", "Number of hours", "Total cost", buttons "Add claim" and "Predetermine". To override without editing code, set `options` in the extension's storage from the panel's DevTools console, e.g.
`chrome.storage.local.set({options: {buttonPredetermine: "Predetermine Claim", typingDelayMs: 60}})`.
Keys: `fieldServiceDate, fieldService, fieldHours, fieldCost, buttonAdd, buttonPredetermine, typingDelayMs, fieldTries, maxAttempts, totalsWaitSeconds`.

## Not done / differences from the Python app
- Not tried on the real site yet (only the mock site, including a full group + Predetermine in a real extension).
- No failure screenshots (an extension cannot save to a folder); failed attempts log what the form held instead.
- No icons, no Web Store packaging, no UI for the wording overrides.
- Typing is simulated with DOM events (`isTrusted` is false); the Python version used real key presses.

## Releasing
`.github/workflows/build-extension.yml` (repo root) runs on every app release tag (`app-v*`, `v*`) and on `ext-v*` tags. It runs `npm test`, writes the tag's version number into `manifest.json`, zips `manifest.json`, `background.js`, `panel.*`, `src/`, `vendor/` as `ClaimFiller-extension.zip` and attaches it to the same release (stable link: `releases/latest/download/ClaimFiller-extension.zip`).
- **Switch off for app releases:** set the repository variable `BUILD_EXTENSION` to `false` (Settings > Secrets and variables > Actions > Variables). Remove it or set `true` to switch back on.
- **Extension on its own:** push `ext-vX.Y.Z` or use Run workflow with that tag; it ignores the switch and publishes a pre-release.
- Chrome only installs such a zip via Developer mode; one-click installs need the Chrome Web Store (same zip, $5 one-time fee).
