"""Self-update from GitHub Releases (standard library only).

A release is a tag like v1.0.1 with ClaimFiller.exe attached (built by
.github/workflows/build.yml). The running exe can't overwrite itself on
Windows, so the new file is downloaded next to it and a tiny batch file swaps
them after the app exits, then relaunches it.
"""
import json
import os
import subprocess
import sys
import urllib.request

from version import REPO, VERSION

ASSET_NAME = "ClaimFiller.exe"
MIN_EXE_BYTES = 1_000_000  # a real build is tens of MB; smaller means a bad download


def parse_version(text):
    return tuple(int(p) for p in text.lstrip("vV").split("."))


def latest_release():
    """Return (version_text, download_url) of the newest release, or None if it has no exe."""
    req = urllib.request.Request(f"https://api.github.com/repos/{REPO}/releases/latest",
                                 headers={"Accept": "application/vnd.github+json", "User-Agent": "ClaimFiller"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.load(resp)
    for asset in data.get("assets", []):
        if asset.get("name") == ASSET_NAME:
            return data["tag_name"], asset["browser_download_url"]
    return None


def check():
    """Return (version_text, url) if a newer release exists, else None. Raises on network errors."""
    found = latest_release()
    if found and parse_version(found[0]) > parse_version(VERSION):
        return found
    return None


def download_and_swap(url):
    """Download the new exe and start the swap script. Caller must exit the app right after."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("Updating only works in the packaged ClaimFiller.exe.")
    if not url.startswith("https://github.com/"):
        raise RuntimeError("Unexpected download address.")
    exe = sys.executable
    folder = os.path.dirname(exe)
    new, old = os.path.join(folder, "ClaimFiller.new.exe"), os.path.join(folder, "ClaimFiller.old.exe")
    urllib.request.urlretrieve(url, new)
    if os.path.getsize(new) < MIN_EXE_BYTES:
        os.remove(new)
        raise RuntimeError("The downloaded file looks incomplete. Nothing was changed.")
    script = os.path.join(folder, "update.bat")
    with open(script, "w") as f:
        f.write(f'''@echo off
ping 127.0.0.1 -n 3 > nul
del "{old}" 2> nul
move /Y "{exe}" "{old}" > nul
move /Y "{new}" "{exe}" > nul
start "" "{exe}"
del "%~f0"
''')
    subprocess.Popen(["cmd", "/c", script], creationflags=0x08000000)  # CREATE_NO_WINDOW
