#!/usr/bin/env bash
set -euo pipefail

APP_NAME="NPCMode"
ENTRY="main.py"
ICON="resources/icons/app.icns"

echo "==> Installing / upgrading dependencies…"
pip install --upgrade pip wheel
pip install -r requirements.txt pyinstaller

echo "==> Building macOS .app bundle…"
pyinstaller \
  --clean \
  --noconfirm \
  --windowed \
  --name "$APP_NAME" \
  --icon "$ICON" \
  --add-data "resources:resources" \
  "$ENTRY"

echo "==> Done → dist/$APP_NAME.app"
open dist/
