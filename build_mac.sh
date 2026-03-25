#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_EXE="$PROJECT_ROOT/.venv/bin/python"
APP_NAME="Gyatt-O-Tune"
APP_BUNDLE="$PROJECT_ROOT/dist/${APP_NAME}.app"
DMG_PATH="$PROJECT_ROOT/dist/${APP_NAME}-macOS.dmg"
DMG_STAGING="$PROJECT_ROOT/build/macos_dmg"

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "This script must be run on macOS." >&2
  exit 1
fi

if [[ ! -x "$PYTHON_EXE" ]]; then
  echo "Python virtual environment not found at .venv."
  echo "Create it first and install dependencies, for example:"
  echo "  python3 -m venv .venv"
  echo "  source .venv/bin/activate"
  echo "  pip install -e .[build]"
  exit 1
fi

cd "$PROJECT_ROOT"

echo "==> Building macOS app bundle with PyInstaller..."
"$PYTHON_EXE" -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  --add-data "src/gyatt_o_tune/assets/gyatt-o-tune.svg:gyatt_o_tune/assets" \
  --add-data "src/gyatt_o_tune/assets/gyatt-o-tune.ico:gyatt_o_tune/assets" \
  "src/gyatt_o_tune/main.py"

echo "==> App bundle ready: $APP_BUNDLE"

if [[ ! -d "$APP_BUNDLE" ]]; then
  echo "Expected app bundle was not produced: $APP_BUNDLE" >&2
  exit 1
fi

rm -rf "$DMG_STAGING"
mkdir -p "$DMG_STAGING"
cp -R "$APP_BUNDLE" "$DMG_STAGING/"
ln -s /Applications "$DMG_STAGING/Applications"

if [[ -f "$DMG_PATH" ]]; then
  rm -f "$DMG_PATH"
fi

echo "==> Building DMG installer..."
hdiutil create \
  -volname "$APP_NAME" \
  -srcfolder "$DMG_STAGING" \
  -ov \
  -format UDZO \
  "$DMG_PATH"

echo "==> DMG installer ready: $DMG_PATH"

cat <<'EOF'

Unsigned macOS apps may be blocked by Gatekeeper on first launch.
For local testing, right-click the app and choose Open.
For distribution, sign and notarize the app with Apple Developer tools.
EOF
