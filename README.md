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

GitHub > Actions > **Build ClaimFiller** > **Run workflow**, enter a tag such as `v1.0.3`. The Windows and Mac builds are attached to that release automatically. (Pushing a `v*` tag does the same.) Leave the tag blank for a test build kept only on the run page.
