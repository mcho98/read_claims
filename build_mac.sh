#!/bin/sh
# Build ClaimFiller.app on a Mac (needs Python once, on the build machine only).
pip install openpyxl playwright pyinstaller || exit 1
pyinstaller --windowed --name ClaimFiller --collect-all playwright app.py || exit 1
ditto -c -k --keepParent dist/ClaimFiller.app dist/ClaimFiller-mac.zip
echo "Done: dist/ClaimFiller-mac.zip (put config.json next to the unzipped ClaimFiller.app)"
