#!/bin/sh
# Builds "Claude Bridge.app", the menu bar app that starts and stops the shared
# Claude Bridge server. Needs the Xcode command line tools (swiftc), not Xcode.
#
# Usage: macos-app/build.sh [install]
#   install: also copy the app to ~/Applications (replacing an older copy).

set -eu

HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$HERE/.." && pwd)"
APP="$HERE/build/Claude Bridge.app"
VERSION="$(sed -n 's/.*"version": *"\([^"]*\)".*/\1/p' "$ROOT/mcp-server/package.json" | head -1)"
SERVER="$ROOT/mcp-server/dist/server.mjs"

[ -f "$SERVER" ] || { echo "Missing $SERVER: run 'npm run build' in mcp-server first." >&2; exit 1; }

rm -rf "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"

swiftc -O -swift-version 5 -target "$(uname -m)-apple-macos13.0" \
  -o "$APP/Contents/MacOS/ClaudeBridge" "$HERE/ClaudeBridgeMenu.swift"

cp "$SERVER" "$APP/Contents/Resources/server.mjs"

cat > "$APP/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleIdentifier</key><string>io.github.maximlefebvre.claude-bridge</string>
  <key>CFBundleName</key><string>Claude Bridge</string>
  <key>CFBundleDisplayName</key><string>Claude Bridge</string>
  <key>CFBundleExecutable</key><string>ClaudeBridge</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>$VERSION</string>
  <key>CFBundleVersion</key><string>$VERSION</string>
  <key>LSMinimumSystemVersion</key><string>13.0</string>
  <key>LSUIElement</key><true/>
  <key>NSHumanReadableCopyright</key><string>Unofficial project, not affiliated with Anthropic or Ascensio System SIA.</string>
</dict>
</plist>
EOF

# Ad-hoc signature: enough to run on this Mac.
codesign --force --sign - "$APP" >/dev/null
echo "Built $APP ($VERSION)"

if [ "${1:-}" = "install" ]; then
  mkdir -p "$HOME/Applications"
  # Quit a running copy so the new one can replace it.
  osascript -e 'tell application id "io.github.maximlefebvre.claude-bridge" to quit' >/dev/null 2>&1 || true
  rm -rf "$HOME/Applications/Claude Bridge.app"
  cp -R "$APP" "$HOME/Applications/"
  echo "Installed to ~/Applications/Claude Bridge.app"
fi
