#!/bin/sh
# Runs the Tally server now and at every login (macOS LaunchAgent).
# Remove with ./uninstall.sh
set -e
DIR="$(cd "$(dirname "$0")/.." && pwd)"
NODE="$(command -v node)"
LABEL="local.tally"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
LOG_DIR="$HOME/Library/Logs/Tally"
mkdir -p "$HOME/Library/LaunchAgents" "$LOG_DIR"

# Earlier versions ran as "local.ollama-excel-agent"; stop and remove that so only one server holds the port.
LEGACY="$HOME/Library/LaunchAgents/local.ollama-excel-agent.plist"
if [ -f "$LEGACY" ]; then
  launchctl unload "$LEGACY" 2>/dev/null || true
  rm -f "$LEGACY"
fi

cat > "$PLIST" <<PL
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key><array><string>$NODE</string><string>$DIR/server.js</string></array>
  <key>WorkingDirectory</key><string>$DIR</string>
  <key>EnvironmentVariables</key><dict><key>PATH</key><string>/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin</string></dict>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardOutPath</key><string>$LOG_DIR/server.log</string>
  <key>StandardErrorPath</key><string>$LOG_DIR/server.log</string>
</dict></plist>
PL
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load "$PLIST"
echo "Tally's server now runs at login. Log: $LOG_DIR/server.log"
