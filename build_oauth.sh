#!/bin/bash
# ============================================================
#  build_oauth.sh — builds "Candidate Downloader.app"
#                   (Google Sign-In version)
#
#  Put these three files in the same folder:
#     build_oauth.sh
#     candidate_downloader_oauth.py
#     client_secret.json      <-- from Google Cloud Console
#
#  Then run:   bash build_oauth.sh
# ============================================================
set -e

APP_NAME="Candidate Downloader"
SCRIPT="candidate_downloader_oauth.py"
SECRET="client_secret.json"

# --- checks -----------------------------------------------------------
[ -f "$SCRIPT" ] || { echo "ERROR: $SCRIPT not found here."; exit 1; }

if [ ! -f "$SECRET" ]; then
  echo "ERROR: $SECRET not found in this folder."
  echo ""
  echo "Create it once in Google Cloud Console:"
  echo "  console.cloud.google.com -> new project -> enable Gmail API"
  echo "  -> OAuth consent screen -> Credentials -> Create OAuth client ID"
  echo "  -> Application type: Desktop app -> Download JSON"
  echo "Rename the downloaded file to client_secret.json and put it here."
  exit 1
fi

command -v python3 >/dev/null 2>&1 || {
  echo "ERROR: python3 not found. Install from https://www.python.org/downloads/"
  exit 1; }

python3 -c "import tkinter" >/dev/null 2>&1 || {
  echo "ERROR: this python3 has no tkinter."
  echo "Install Python from https://www.python.org/downloads/ and re-run."
  exit 1; }

# --- build environment ------------------------------------------------
echo ">> Creating build environment…"
rm -rf .buildenv build dist *.spec
python3 -m venv .buildenv
source .buildenv/bin/activate
pip install --quiet --upgrade pip
echo ">> Installing Google libraries + PyInstaller…"
pip install --quiet \
  google-api-python-client \
  google-auth-oauthlib \
  google-auth-httplib2 \
  pyinstaller

# --- build ------------------------------------------------------------
echo ">> Building ${APP_NAME}.app (this takes a few minutes)…"
pyinstaller --noconfirm --windowed --onedir \
  --name "$APP_NAME" --target-arch universal2 \
  --add-data "${SECRET}:." \
  --collect-all googleapiclient \
  --collect-all google_auth_oauthlib \
  --collect-all google.auth \
  --hidden-import googleapiclient.discovery \
  "$SCRIPT"

# also drop a copy beside the app, so it can be replaced without rebuilding
cp "$SECRET" "dist/${SECRET}" 2>/dev/null || true

# --- sign + zip -------------------------------------------------------
echo ">> Codesigning (ad-hoc)…"
codesign --force --deep -s - "dist/${APP_NAME}.app" || true

echo ">> Zipping for delivery…"
cd dist
zip -qry "${APP_NAME}.zip" "${APP_NAME}.app"
cd ..

deactivate

echo ""
echo "=============================================================="
echo "  DONE."
echo "  App:  dist/${APP_NAME}.app"
echo "  Send: dist/${APP_NAME}.zip"
echo ""
echo "  Tell the recipient: right-click the app -> Open -> Open"
echo "  the first time, then click 'Sign in with Google'."
echo "=============================================================="
