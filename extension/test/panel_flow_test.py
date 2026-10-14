"""Loads the extension into Chromium and runs the whole side-panel flow against
test/real_like_site.py (page reloads after Add claim, refresh that wipes the fields).

    python3 test/panel_flow_test.py            # needs Python Playwright + Chromium
Host permission for the local site is granted in a temporary copy of the manifest, because
the permission prompt cannot be clicked by a script.
"""
import json
import pathlib
import shutil
import sys
import tempfile
import time

from playwright.sync_api import sync_playwright

sys.path.insert(0, str(pathlib.Path(__file__).parent))
import real_like_site  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parents[2]
EXT = ROOT / "extension"
FIRST_ONLY = "--first-only" in sys.argv

REMOTE = next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--url=")), None)  # e.g. --url=https://mock-website-theta.vercel.app
if REMOTE:
    URL = REMOTE.rstrip("/") + "/"
else:
    server = real_like_site.serve(0)
    URL = f"http://127.0.0.1:{server.server_address[1]}/"
work = pathlib.Path(tempfile.mkdtemp())
copy = work / "ext"
shutil.copytree(EXT, copy, ignore=shutil.ignore_patterns("node_modules", "test", "*.zip"))
manifest = json.loads((copy / "manifest.json").read_text())
manifest["host_permissions"] = [f"{URL}*"]  # the origin, e.g. http://127.0.0.1:1234/*
(copy / "manifest.json").write_text(json.dumps(manifest))

# Dates inside the fake site's 60-day window: use the last 4 days.
import datetime as dt
today = dt.date.today()
claims = [{"service_date": (today - dt.timedelta(days=3 - i)).isoformat(), "service": s, "hours": h, "total_cost": c, "cell": "X"}
          for i, (s, h, c) in enumerate([("Personal Care", 4, 136.12), ("Homemaking", 3.5, 119.11),
                                          ("Respite", 2, 68.06), ("Personal Care", 1.53, 52.07)])]
doc = {"format": "claimfiller-claims", "version": 1, "generated_at": "x", "source_file": "t.xlsx", "sheet": "S",
       "client": "Test Client", "rate": 34.03, "date_range": {"start": None, "end": None},
       "max_claims_per_group": 10, "totals": {}, "problems": [], "warnings": [],
       "groups": [{"group": 1, "claims": claims}]}
doc_path = work / "export.json"
doc_path.write_text(json.dumps(doc))

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(str(work / "profile"), headless=False,
        args=[f"--disable-extensions-except={copy}", f"--load-extension={copy}", "--headless=new"])
    sw = ctx.service_workers[0] if ctx.service_workers else ctx.wait_for_event("serviceworker")
    ext_id = sw.url.split("/")[2]
    site = ctx.new_page()
    site.goto(URL)
    panel = ctx.new_page()
    panel.goto(f"chrome-extension://{ext_id}/panel.html")
    panel.on("dialog", lambda d: d.accept())
    panel.set_input_files("#jsonfile", str(doc_path))
    panel.fill("#url", URL)
    if FIRST_ONLY:
        panel.check("#firstOnly")
    site.bring_to_front()
    panel.click("#fill")
    for _ in range(180):
        text = panel.inner_text("#log")
        if "Done for" in text or "STOPPED" in text or "Stopped:" in text:
            break
        time.sleep(1)
    print(panel.inner_text("#log"))
    want = 1 if FIRST_ONLY else len(claims)
    if REMOTE:
        shown = site.inner_text("body")
        print("site total shown:", [l for l in shown.splitlines() if l.strip().startswith("$")][-1:])
        print("rows on the site:", site.locator("table tbody tr").count())
        ok = "Done for" in panel.inner_text("#log") and site.locator("table tbody tr").count() >= want
    else:
        print("server rows:", real_like_site.ROWS)
        ok = "Done for" in panel.inner_text("#log") and len(real_like_site.ROWS) == want
    print("PASS" if ok else "FAIL")
    ctx.close()
    sys.exit(0 if ok else 1)
