#!/bin/bash
# Removes Tally from this Mac: the login service, the Excel registration and (with --remove-certificate)
# the localhost certificate. Ollama, its models and this folder are left alone; the end of the
# output says how to remove them too.
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REMOVE_CERT=0
for arg in "$@"; do
  case "$arg" in
    --remove-certificate) REMOVE_CERT=1 ;;
    *) echo "Unknown option: $arg (only --remove-certificate)" >&2; exit 2 ;;
  esac
done

for label in local.tally local.ollama-excel-agent; do
  plist="$HOME/Library/LaunchAgents/$label.plist"
  if [ -f "$plist" ]; then
    launchctl unload "$plist" 2>/dev/null || true
    rm -f "$plist"
    echo "✓ Stopped and removed the login service ($label)"
  fi
done

WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
for f in tally.xml ollama-agent.xml; do
  if [ -f "$WEF/$f" ]; then rm -f "$WEF/$f"; echo "✓ Removed Tally from Excel ($f)"; fi
done

if [ "$REMOVE_CERT" = 1 ]; then
  (cd "$DIR" && npx --no-install office-addin-dev-certs uninstall) && echo "✓ Removed the localhost certificate"
fi

echo
echo "Tally is removed. Quit and reopen Excel to clear the button."
echo "Left in place, in case you use them for other things:"
echo "  - the model:          ollama rm qwen3.8"
echo "  - this folder:        rm -rf \"$DIR\""
[ "$REMOVE_CERT" = 1 ] || echo "  - the certificate:    $DIR/uninstall.sh --remove-certificate (other Office add-ins in development may use it)"
