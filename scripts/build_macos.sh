#!/bin/zsh
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$PROJECT_ROOT"

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "This script supports macOS only."
  exit 1
fi

VERSION="${1:-dev}"
SAFE_VERSION="$(echo "$VERSION" | tr '/ ' '--')"
APP_NAME="ExcelAutoDiff"
APP_PATH="dist/${APP_NAME}.app"
ZIP_PATH="dist/${APP_NAME}-${SAFE_VERSION}-macOS.zip"
DMG_PATH="dist/${APP_NAME}-${SAFE_VERSION}-macOS.dmg"
SIGN_IDENTITY="${APPLE_SIGN_IDENTITY:-}"
NOTARY_KEY_ID="${APPLE_NOTARY_KEY_ID:-}"
NOTARY_ISSUER_ID="${APPLE_NOTARY_ISSUER_ID:-}"
NOTARY_KEY_P8="${APPLE_NOTARY_KEY_P8:-}"
NOTARY_TMP_ZIP=""
NOTARY_KEY_FILE=""

cleanup() {
  if [[ -n "$NOTARY_TMP_ZIP" && -f "$NOTARY_TMP_ZIP" ]]; then
    rm -f "$NOTARY_TMP_ZIP"
  fi
  if [[ -n "$NOTARY_KEY_FILE" && -f "$NOTARY_KEY_FILE" ]]; then
    rm -f "$NOTARY_KEY_FILE"
  fi
}
trap cleanup EXIT

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

PYTHONPYCACHEPREFIX=/tmp/.pycache_excelautodiff python -m py_compile excel_diff.py excel_diff_gui.py

rm -rf build dist

pyinstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  --add-data "assets:assets" \
  excel_diff_gui.py

if [ ! -d "$APP_PATH" ]; then
  echo "App bundle build failed: $APP_PATH"
  exit 1
fi

if [[ -n "$SIGN_IDENTITY" ]]; then
  echo "Signing app with identity: $SIGN_IDENTITY"
  codesign \
    --force \
    --deep \
    --options runtime \
    --timestamp \
    --sign "$SIGN_IDENTITY" \
    "$APP_PATH"
  codesign --verify --deep --strict --verbose=2 "$APP_PATH"
fi

if [[ -n "$NOTARY_KEY_ID" || -n "$NOTARY_ISSUER_ID" || -n "$NOTARY_KEY_P8" ]]; then
  if [[ -z "$SIGN_IDENTITY" ]]; then
    echo "Notarization requested, but APPLE_SIGN_IDENTITY is missing."
    exit 1
  fi
  if [[ -z "$NOTARY_KEY_ID" || -z "$NOTARY_ISSUER_ID" || -z "$NOTARY_KEY_P8" ]]; then
    echo "Notarization requested, but one or more notary credentials are missing."
    echo "Required: APPLE_NOTARY_KEY_ID, APPLE_NOTARY_ISSUER_ID, APPLE_NOTARY_KEY_P8"
    exit 1
  fi

  echo "Submitting app for notarization..."
  NOTARY_KEY_FILE="$(mktemp "${TMPDIR:-/tmp}/AuthKey_${NOTARY_KEY_ID}_XXXXXX.p8")"
  chmod 600 "$NOTARY_KEY_FILE"
  printf '%s' "$NOTARY_KEY_P8" > "$NOTARY_KEY_FILE"

  NOTARY_TMP_ZIP="dist/${APP_NAME}-${SAFE_VERSION}-notary.zip"
  ditto -c -k --sequesterRsrc --keepParent "$APP_PATH" "$NOTARY_TMP_ZIP"

  xcrun notarytool submit "$NOTARY_TMP_ZIP" \
    --key "$NOTARY_KEY_FILE" \
    --key-id "$NOTARY_KEY_ID" \
    --issuer "$NOTARY_ISSUER_ID" \
    --wait

  xcrun stapler staple "$APP_PATH"
  xcrun stapler validate "$APP_PATH"
fi

ditto -c -k --sequesterRsrc --keepParent "$APP_PATH" "$ZIP_PATH"

STAGE_DIR="$(mktemp -d)"
trap 'cleanup; rm -rf "$STAGE_DIR"' EXIT
cp -R "$APP_PATH" "$STAGE_DIR/"
ln -s /Applications "$STAGE_DIR/Applications"
hdiutil create -volname "$APP_NAME" -srcfolder "$STAGE_DIR" -ov -format UDZO "$DMG_PATH"

echo ""
echo "Build complete"
echo "- APP: $APP_PATH"
echo "- ZIP: $ZIP_PATH"
echo "- DMG: $DMG_PATH"
