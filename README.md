# read_claims

Reads a client's monthly hours from the timesheet workbook and enters them as claims in the online form.

## Download

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

1. Download `ClaimFiller-extension-<version>.zip` from the newest **ext-v…** entry on the [Releases page](https://github.com/mcho98/read_claims/releases) (these are marked "Pre-release").
2. Unzip it. In Chrome open `chrome://extensions`, turn on **Developer mode**, click **Load unpacked** and choose the unzipped folder.
3. Click the Claim Filler toolbar icon to open the side panel. Details: [extension/README.md](extension/README.md).

Publishing (developer): Actions > **Build Chrome extension** > Run workflow with a tag such as `ext-v0.2.0` (or push that tag). It runs the tests, stamps the version into `manifest.json`, zips the runtime files and attaches the zip to a pre-release.
