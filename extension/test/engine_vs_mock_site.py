"""Runs src/content.js (the fill engine) against the mock claim website in a real browser.

Not part of `npm test` because it needs the internet and Python Playwright:
    python3 test/engine_vs_mock_site.py            # full first group, then Predetermine
The claims come from ../shared/out.json (made by the Python script from mock_data.xlsx).
"""
import json
import pathlib
import sys

from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).resolve().parent
URL = sys.argv[1] if len(sys.argv) > 1 else "https://mock-website-theta.vercel.app"
claims = json.loads((HERE.parent.parent / "shared" / "out.json").read_text())["groups"][0]["claims"]
expected_total = round(sum(c["total_cost"] for c in claims), 2)

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto(URL)
    page.wait_for_timeout(1500)
    page.add_script_tag(path=str(HERE.parent / "src" / "content.js"))
    print("check:", page.evaluate("ClaimFillerEngine.create().check()"))
    logs = []
    page.expose_function("pyLog", lambda text: logs.append(text))
    outcome = page.evaluate(
        "claims => ClaimFillerEngine.create({log: pyLog}).fill(claims, {predetermine: true})", claims)
    print("\n".join(logs))
    print("outcome:", outcome)
    ok = outcome["ok"] and outcome["added"] == len(claims)
    print("expected total", expected_total, "->", "PASS" if ok else "FAIL")
    browser.close()
    sys.exit(0 if ok else 1)
