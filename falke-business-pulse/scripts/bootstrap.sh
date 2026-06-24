#!/usr/bin/env bash
# falke-business-pulse — first-run dependency bootstrap (SessionStart hook).
#
# The bundled apple-mail MCP server has ONE runtime dependency: the MCP Python
# SDK (mcp>=1.2.0, see apple-mail/requirements.txt). This installs it ONCE into a
# persistent venv under ${CLAUDE_PLUGIN_DATA}, reinstalling only when the bundled
# requirements.txt changes (the canonical diff-stamp pattern from
# code.claude.com/docs/en/plugins-reference). .mcp.json points the apple-mail
# server's `command` at ${CLAUDE_PLUGIN_DATA}/venv/bin/python3, so this venv must
# exist before the server first starts — the SessionStart hook guarantees that.
#
# Trimmed from the proven falke-bid-tools bootstrap: NO Chromium/Playwright step
# (this server has no such dependency) and no render-mode marker. Pure-Python dep
# only.
#
# Idempotent and quiet on the happy path. Floyd reviews this for the execution
# threat model before ship, like the bid-tools one.

set -euo pipefail

DATA="${CLAUDE_PLUGIN_DATA:?CLAUDE_PLUGIN_DATA not set}"
ROOT="${CLAUDE_PLUGIN_ROOT:?CLAUDE_PLUGIN_ROOT not set}"

VENV="${DATA}/venv"
REQ="${ROOT}/apple-mail/requirements.txt"
STAMP="${DATA}/requirements.installed.txt"

mkdir -p "${DATA}"

# Pick a Python >=3.10 (the mcp SDK floor). macOS stock python3 is often 3.9,
# which CANNOT install `mcp` — pip would fail with a confusing "no matching
# distribution" error and leave a broken server. Fail LOUD with an actionable
# message instead, consistent with the rest of this system's fail-loud discipline.
PYBIN=""
for cand in python3.13 python3.12 python3.11 python3.10 python3; do
  if command -v "${cand}" >/dev/null 2>&1 && \
     "${cand}" -c 'import sys; raise SystemExit(0 if sys.version_info[:2] >= (3,10) else 1)' 2>/dev/null; then
    PYBIN="${cand}"
    break
  fi
done
if [ -z "${PYBIN}" ]; then
  echo "[falke-business-pulse] ERROR: need Python >=3.10 for the mail tool, none found." >&2
  echo "[falke-business-pulse] Install it (e.g. 'brew install python@3.12') and reopen Cowork." >&2
  exit 1
fi

# Reinstall only if the bundled manifest differs from what we last installed
# (covers both first run and a dependency-changing plugin update).
if [ ! -f "${STAMP}" ] || ! diff -q "${REQ}" "${STAMP}" >/dev/null 2>&1; then
  echo "[falke-business-pulse] Setting up dependencies (first run may take ~30s)..." >&2

  "${PYBIN}" -m venv "${VENV}"
  "${VENV}/bin/pip" install --quiet --upgrade pip
  "${VENV}/bin/pip" install --quiet -r "${REQ}"

  # Stamp success LAST so a mid-install failure retries next session.
  cp "${REQ}" "${STAMP}"
  echo "[falke-business-pulse] Setup complete." >&2
fi

exit 0
