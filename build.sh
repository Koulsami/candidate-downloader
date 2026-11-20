#!/bin/bash
# ============================================================
#  build.sh — builds "Candidate Downloader.app" (standalone)
#
#  Usage:
#     1. Put this file and candidate_attachment_downloader_mac.py
#        in the same folder
#     2. Open Terminal in that folder and run:   bash build.sh
#
#  Result:  dist/Candidate Downloader.app   (double-clickable,
#           Python bundled inside — nothing else to install)
# ============================================================
set -e

APP_NAME="Candidate Downloader"
SCRIPT="candidate_attachment_downloader_mac.py"

# --- 0. sanity checks -------------------------------------------------
if [ ! -f "$SCRIPT" ]; then
  echo "ERROR: $SCRIPT not found in this folder."
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "ERROR: python3 not found."
  echo "Install it once from https://www.python.org/downloads/ (the"
  echo "python.org installer includes tkinter, which the app needs),"
  echo "then re-run this script."
  exit 1
fi

# tkinter must be present in this Python (python.org build has it)
if ! python3 -c "import tkinter" >/dev/null 2>&1; then
  echo "ERROR: this python3 has no tkinter."
  echo "Install Python from https://www.python.org/downloads/ and re-run."
  exit 1
fi

# --- 1. clean build environment --------------------------------------
echo ">> Creating build environment…"
rm -rf .buildenv build dist "*.spec"
python3 -m venv .buildenv
source .buildenv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet pyinstaller

# --- 2. build the .app ------------------------------------------------
echo ">> Building ${APP_NAME}.app (takes a minute)…"
pyinstaller --noconfirm --windowed --onedir \
  --name "$APP_NAME" \
  "$SCRIPT"

# --- 3. ad-hoc codesign so macOS treats it as a normal local app ------
echo ">> Codesigning (ad-hoc)…"
codesign --force --deep -s - "dist/${APP_NAME}.app" || true

# --- 4. zip it for easy moving/sharing --------------------------------
echo ">> Zipping…"
cd dist
zip -qry "${APP_NAME}.zip" "${APP_NAME}.app"
cd ..

deactivate

echo ""
echo "=============================================================="
echo "  DONE."
echo "  App:  dist/${APP_NAME}.app"
echo "  Zip:  dist/${APP_NAME}.zip"
echo ""
echo "  First launch: RIGHT-CLICK the app -> Open -> Open."
echo "  (Needed once because the app isn't Apple-notarized.)"
echo "  After that it opens with a normal double-click."
echo "=============================================================="
