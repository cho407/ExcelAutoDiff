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

ditto -c -k --sequesterRsrc --keepParent "$APP_PATH" "$ZIP_PATH"

STAGE_DIR="$(mktemp -d)"
trap 'rm -rf "$STAGE_DIR"' EXIT
cp -R "$APP_PATH" "$STAGE_DIR/"
ln -s /Applications "$STAGE_DIR/Applications"
hdiutil create -volname "$APP_NAME" -srcfolder "$STAGE_DIR" -ov -format UDZO "$DMG_PATH"

echo ""
echo "Build complete"
echo "- APP: $APP_PATH"
echo "- ZIP: $ZIP_PATH"
echo "- DMG: $DMG_PATH"

