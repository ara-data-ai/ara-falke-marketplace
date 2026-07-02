#!/bin/zsh
# Install the Falke Pulse local viewer: launchd-managed localhost server
# (http://127.0.0.1:8787) + optional 7:00 AM weekday refresh.
# Run once per laptop:  ./install.sh [--with-morning-run]
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
AGENTS="$HOME/Library/LaunchAgents"
PYTHON="$(command -v python3)"

if [[ -z "$PYTHON" ]]; then
  echo "[pulse-server] ERROR: python3 not found on PATH." >&2
  exit 1
fi

mkdir -p "$AGENTS" "$HOME/Library/Logs/falke-pulse-server"

install_plist() {
  local name="$1"
  sed -e "s|__PYTHON__|$PYTHON|g" \
      -e "s|__SERVER__|$HERE/server.py|g" \
      -e "s|__HOME__|$HOME|g" \
      "$HERE/launchd/$name.plist" > "$AGENTS/$name.plist"
  launchctl bootout "gui/$(id -u)/$name" 2>/dev/null || true
  launchctl bootstrap "gui/$(id -u)" "$AGENTS/$name.plist"
  echo "[pulse-server] Installed + started: $name"
}

install_plist "com.falke.pulse-server"

if [[ "${1:-}" == "--with-morning-run" ]]; then
  install_plist "com.falke.pulse-morning"
fi

echo "[pulse-server] Done. Bookmark http://127.0.0.1:8787 in Chrome"
echo "[pulse-server] (Chrome > Settings > On startup > Open specific page)."
