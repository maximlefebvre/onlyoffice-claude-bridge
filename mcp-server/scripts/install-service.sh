#!/bin/sh
# Installs the shared Claude Bridge server as a macOS login service (launchd agent).
#
# The service owns the WebSocket the OnlyOffice plugin connects to and serves MCP
# over HTTP on 127.0.0.1. Every Claude session then shares it: sessions started
# with the usual stdio command relay to it automatically, and HTTP clients can
# connect directly with the bearer token in ~/.config/onlyoffice-claude-bridge/mcp-token.
#
# Usage: install-service.sh [install|uninstall|status|restart]

set -eu

LABEL="io.github.maximlefebvre.onlyoffice-claude-bridge"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
LOG="$HOME/Library/Logs/onlyoffice-claude-bridge.log"
DOMAIN="gui/$(id -u)"
SERVER="$(cd "$(dirname "$0")/.." && pwd)/dist/server.mjs"
MCP_PORT="${CLAUDE_BRIDGE_MCP_PORT:-47616}"

if [ "$(uname)" != "Darwin" ]; then
  echo "This script supports macOS only. Elsewhere, run 'node $SERVER --http' at login (systemd user unit, Task Scheduler…)." >&2
  exit 1
fi

xml_escape() { printf '%s' "$1" | sed -e 's/&/\&amp;/g' -e 's/</\&lt;/g' -e 's/>/\&gt;/g'; }

install() {
  NODE="$(command -v node || true)"
  if [ -z "$NODE" ]; then echo "Node.js 20+ is required (node not found in PATH)." >&2; exit 1; fi
  if [ ! -f "$SERVER" ]; then echo "Missing $SERVER: run 'npm run build' in mcp-server first." >&2; exit 1; fi

  mkdir -p "$HOME/Library/LaunchAgents" "$HOME/Library/Logs"
  cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>$(xml_escape "$NODE")</string>
    <string>$(xml_escape "$SERVER")</string>
    <string>--http</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>ThrottleInterval</key><integer>5</integer>
  <key>ProcessType</key><string>Interactive</string>
  <key>StandardErrorPath</key><string>$(xml_escape "$LOG")</string>
  <key>StandardOutPath</key><string>$(xml_escape "$LOG")</string>
</dict>
</plist>
EOF

  launchctl bootout "$DOMAIN/$LABEL" 2>/dev/null || true
  launchctl bootstrap "$DOMAIN" "$PLIST"
  echo "Installed $LABEL"
  echo "  server: $SERVER"
  echo "  node:   $NODE"
  echo "  log:    $LOG"
  sleep 2
  status
  echo
  echo "Claude sessions started from now on share this server. Sessions already open"
  echo "still run their own server: restart them (or their MCP server) to switch."
}

uninstall() {
  launchctl bootout "$DOMAIN/$LABEL" 2>/dev/null || true
  rm -f "$PLIST"
  echo "Removed $LABEL. Claude sessions go back to one server each."
}

status() {
  if launchctl print "$DOMAIN/$LABEL" >/dev/null 2>&1; then
    echo "Service: loaded"
  else
    echo "Service: not loaded"
  fi
  if curl -fsS "http://127.0.0.1:$MCP_PORT/health" 2>/dev/null; then
    echo
  else
    echo "Shared server: not answering on port $MCP_PORT (see $LOG)"
  fi
}

case "${1:-install}" in
  install) install ;;
  uninstall) uninstall ;;
  status) status ;;
  restart) launchctl kickstart -k "$DOMAIN/$LABEL" && sleep 2 && status ;;
  *) echo "Usage: $0 [install|uninstall|status|restart]" >&2; exit 1 ;;
esac
