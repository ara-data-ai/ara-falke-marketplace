"""Falke Pulse local viewer — serves the latest morning-pulse HTML at
http://127.0.0.1:8787 with a Refresh button that triggers a headless
Claude Code run of the falke-business-pulse skill.

Design constraints (see README):
- SERVING and REFRESHING fail independently: the page always shows the last
  successful pulse; a broken refresh path only disables the button, never the view.
- Bound to 127.0.0.1 ONLY. Never expose on the network.
- Stdlib only — no dependencies beyond Python >=3.10 (already required by the plugin).
"""

from __future__ import annotations

import glob
import json
import os
import re
import subprocess
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOST = "127.0.0.1"
PORT = 8787
CONFIG_PATH = os.path.expanduser("~/.falke-business-pulse/config.json")

# Defaults; override via config.json keys of the same name.
DEFAULTS = {
    # Directory the skill writes pulse-YYYY-MM-DD.html files into.
    "pulse_html_dir": "~/Claude/Projects/Falke-Business-Pulse",
    # Headless run. --permission-mode acceptEdits lets the skill write the
    # HTML + state files without prompting; MCP mail tools may additionally
    # need a permissions allowlist in the project .claude/settings.json.
    "refresh_command": 'claude -p "run my morning pulse" --permission-mode acceptEdits',
}

_state_lock = threading.Lock()
_state = {"refreshing": False, "last_error": None, "last_exit": None}


def _config() -> dict:
    cfg = dict(DEFAULTS)
    try:
        with open(CONFIG_PATH) as f:
            cfg.update(json.load(f))
    except (OSError, ValueError):
        pass
    return cfg


def _latest_pulse() -> str | None:
    d = os.path.expanduser(_config()["pulse_html_dir"])
    files = glob.glob(os.path.join(d, "pulse-*.html"))
    return max(files, key=os.path.getmtime) if files else None


def _stamp(path: str | None) -> str:
    if not path:
        return "no pulse generated yet"
    t = datetime.fromtimestamp(os.path.getmtime(path))
    return "Data last refreshed on " + t.strftime("%H:%M:%S %d/%m/%Y")


TOOLBAR = """
<div id="pulse-toolbar" style="position:sticky;top:0;z-index:9999;display:flex;
  align-items:center;gap:14px;padding:8px 16px;background:#1A2A33;color:#fff;
  font:12px 'Helvetica Neue',Arial,sans-serif;border-bottom:3px solid #F15022;">
  <strong style="letter-spacing:.08em;">FALKE PULSE</strong>
  <span id="pulse-stamp" style="color:#c8d2d8;">__STAMP__</span>
  <button id="pulse-refresh" onclick="pulseRefresh()" style="margin-left:auto;
    background:#F15022;color:#fff;border:none;border-radius:4px;
    padding:6px 14px;font:bold 12px 'Helvetica Neue',Arial,sans-serif;
    cursor:pointer;">Refresh</button>
</div>
<script>
async function pulseRefresh() {
  const b = document.getElementById('pulse-refresh');
  b.disabled = true; b.textContent = 'Refreshing\\u2026';
  try {
    const r = await fetch('/refresh', {method: 'POST'});
    if (!r.ok) throw new Error(await r.text());
    const poll = setInterval(async () => {
      const s = await (await fetch('/status')).json();
      if (!s.refreshing) {
        clearInterval(poll);
        if (s.last_error) {
          b.disabled = false; b.textContent = 'Refresh';
          alert('Refresh failed \\u2014 open Claude Code and run the pulse manually.\\n\\n' + s.last_error);
        } else { location.reload(); }
      }
    }, 5000);
  } catch (e) {
    b.disabled = false; b.textContent = 'Refresh';
    alert('Could not start refresh: ' + e.message);
  }
}
</script>
"""

PLACEHOLDER = """<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Falke Pulse</title></head><body style="font-family:'Helvetica Neue',Arial,sans-serif;">
__TOOLBAR__
<div style="max-width:600px;margin:80px auto;text-align:center;color:#5E6E76;">
  <h2 style="color:#1A2A33;">No pulse yet</h2>
  <p>Click <strong>Refresh</strong> above to run this morning's pulse,
  or run the falke-business-pulse skill from Claude Code.</p>
</div></body></html>"""


def _render() -> bytes:
    path = _latest_pulse()
    bar = TOOLBAR.replace("__STAMP__", _stamp(path))
    if not path:
        return PLACEHOLDER.replace("__TOOLBAR__", bar).encode()
    with open(path, encoding="utf-8") as f:
        html = f.read()
    # Inject the toolbar immediately after the opening <body> tag.
    injected, n = re.subn(r"(<body[^>]*>)", r"\1" + bar.replace("\\", "\\\\"), html, count=1)
    return (injected if n else bar + html).encode()


def _run_refresh() -> None:
    cfg = _config()
    cwd = os.path.expanduser(cfg["pulse_html_dir"])
    try:
        proc = subprocess.run(
            cfg["refresh_command"], shell=True, cwd=cwd,
            capture_output=True, text=True, timeout=900,
        )
        err = None if proc.returncode == 0 else (proc.stderr or proc.stdout)[-500:]
        with _state_lock:
            _state.update(refreshing=False, last_exit=proc.returncode, last_error=err)
    except Exception as e:  # timeout, missing CLI, etc. — fail loud in /status
        with _state_lock:
            _state.update(refreshing=False, last_exit=-1, last_error=str(e)[-500:])


class Handler(BaseHTTPRequestHandler):
    def _send(self, code: int, body: bytes, ctype: str = "text/html; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/status":
            with _state_lock:
                body = json.dumps(_state).encode()
            self._send(200, body, "application/json")
        else:
            self._send(200, _render())

    def do_POST(self):
        if self.path != "/refresh":
            self._send(404, b"not found", "text/plain")
            return
        origin = self.headers.get("Origin", "")
        if origin and not re.match(r"https?://(127\.0\.0\.1|localhost)(:\d+)?$", origin):
            self._send(403, b"forbidden", "text/plain")
            return
        with _state_lock:
            if _state["refreshing"]:
                self._send(409, b"refresh already running", "text/plain")
                return
            _state.update(refreshing=True, last_error=None, last_exit=None)
        threading.Thread(target=_run_refresh, daemon=True).start()
        self._send(202, b"started", "text/plain")

    def log_message(self, *args):  # keep launchd logs quiet
        pass


if __name__ == "__main__":
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
