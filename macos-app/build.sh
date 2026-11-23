#!/bin/sh
# Builds "Claude-Bridge-for-OnlyOffice.app", the menu bar app that starts and
# stops the shared Claude Bridge server. Needs the Xcode command line tools
# (swiftc, iconutil), not Xcode.
#
# Usage: macos-app/build.sh [install]
#   install: also copy the app to ~/Applications (replacing an older copy).

set -eu

NAME="Claude-Bridge-for-OnlyOffice"
BUNDLE_ID="io.github.maximlefebvre.claude-bridge"
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "$HERE/.." && pwd)"
BUILD="$HERE/build"
APP="$BUILD/$NAME.app"
VERSION="$(sed -n 's/.*"version": *"\([^"]*\)".*/\1/p' "$ROOT/mcp-server/package.json" | head -1)"
SERVER="$ROOT/mcp-server/dist/server.mjs"
TARGET="$(uname -m)-apple-macos13.0"

[ -f "$SERVER" ] || { echo "Missing $SERVER: run 'npm run build' in mcp-server first." >&2; exit 1; }

rm -rf "$APP" "$BUILD/AppIcon.iconset"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"

swiftc -O -swift-version 5 -target "$TARGET" \
  -o "$APP/Contents/MacOS/ClaudeBridge" "$HERE/Sources/main.swift" "$HERE/Sources/MenuBarIcon.swift"

# App icon, rendered from the menu bar glyph.
swiftc -O -swift-version 5 -target "$TARGET" -parse-as-library \
  -o "$BUILD/make-icon" "$HERE/Tools/MakeIcon.swift" "$HERE/Sources/MenuBarIcon.swift"
"$BUILD/make-icon" "$BUILD/AppIcon.iconset"
iconutil -c icns -o "$APP/Contents/Resources/AppIcon.icns" "$BUILD/AppIcon.iconset"

cp "$SERVER" "$APP/Contents/Resources/server.mjs"

cat > "$APP/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleIdentifier</key><string>$BUNDLE_ID</string>
  <key>CFBundleName</key><string>$NAME</string>
  <key>CFBundleDisplayName</key><string>Claude Bridge for OnlyOffice</string>
  <key>CFBundleExecutable</key><string>ClaudeBridge</string>
  <key>CFBundleIconFile</key><string>AppIcon</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>$VERSION</string>
  <key>CFBundleVersion</key><string>$VERSION</string>
  <key>LSMinimumSystemVersion</key><string>13.0</string>
  <key>LSUIElement</key><true/>
  <key>NSHumanReadableCopyright</key><string>Unofficial project, not affiliated with Anthropic or Ascensio System SIA (ONLYOFFICE).</string>
</dict>
</plist>
EOF

# Ad-hoc signature: enough to run on this Mac.
codesign --force --sign - "$APP" >/dev/null
echo "Built $APP ($VERSION)"

if [ "${1:-}" = "install" ]; then
  mkdir -p "$HOME/Applications"
  # Quit a running copy (same bundle id, whatever its file name) so it can be replaced.
  osascript -e "tell application id \"$BUNDLE_ID\" to quit" >/dev/null 2>&1 || true
  sleep 1
  rm -rf "$HOME/Applications/$NAME.app" "$HOME/Applications/Claude Bridge.app" # older name
  cp -R "$APP" "$HOME/Applications/"
  echo "Installed to ~/Applications/$NAME.app"
fi
