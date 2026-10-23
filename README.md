# read_claims

Reads a client's monthly hours from the timesheet workbook and enters them as claims in the online form.

## Download

> **Windows/Mac app builds are paused** (only the Chrome extension zip is built on releases for now), so the links below may be missing or out of date. See the extension section at the bottom.

Always the newest version (no account needed):

- **Windows:** [ClaimFiller.exe](https://github.com/mcho98/read_claims/releases/latest/download/ClaimFiller.exe) and [config.json](https://github.com/mcho98/read_claims/releases/latest/download/config.json) — save both in the same folder.
- **Mac (Apple Silicon):** [ClaimFiller-mac.zip](https://github.com/mcho98/read_claims/releases/latest/download/ClaimFiller-mac.zip) — unzip; put `config.json` (above) in the same folder as `ClaimFiller.app`.
- All versions: [Releases page](https://github.com/mcho98/read_claims/releases)

**Google Chrome must be installed** (Windows and Mac). The program opens its own Chrome window; pass any website security check and log in there yourself, then click Fill form. Logins are remembered between runs.

## First run

- **Windows:** if SmartScreen says "Windows protected your PC", click **More info**, then **Run anyway**.
- **Mac:** right-click `ClaimFiller.app`, choose **Open**, then **Open** again (or System Settings > Privacy & Security > **Open Anyway**).

## Publishing a new version (developer)

GitHub > Actions > **Build ClaimFiller** > **Run workflow**, enter a tag such as `app-v1.0.3`. The Windows and Mac builds are attached to that release automatically. (Pushing an `app-v*` tag does the same.) Leave the tag blank for a test build kept only on the run page.

## Chrome extension (preview)

A Chrome-only version that runs inside your own browser (no Python, no exe).

1. Download [ClaimFiller-extension.zip](https://github.com/mcho98/read_claims/releases/latest/download/ClaimFiller-extension.zip) (it is attached to every release).
2. Unzip it. In Chrome open `chrome://extensions`, turn on **Developer mode**, click **Load unpacked** and choose the unzipped folder.
3. Click the Claim Filler toolbar icon to open the side panel. Details: [extension/README.md](extension/README.md).

Publishing (developer): the extension zip is built and attached automatically with every release tag (`app-v*` / `v*` / `ext-v*`; version number taken from the tag). It is the only thing built for now: the extension needs no `config.json` (its settings live in Chrome). To switch that off, set the repository variable `BUILD_EXTENSION` to `false` (Settings > Secrets and variables > Actions > Variables); delete it or set `true` to switch it back on. To release the extension alone, push an `ext-vX.Y.Z` tag or run **Build Chrome extension** from the Actions page with a tag; that release is marked Pre-release.
