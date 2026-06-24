# Falke Business Pulse (Cowork plugin) — PROTOTYPE, pending Floyd's go-live gate

One plugin that bundles **the morning Chief-of-Staff routine** as a Cowork
plugin: the `/Falke-business-pulse` skill + the host-native **`apple-mail` MCP
server** (read + draft, never send), auto-registered, with a one-Python-dep
bootstrap. Distributed through the same private https-git marketplace
(`ara-falke`) as `falke-bid-tools`.

> **Status: assembled, pending live plugin-install validation on Derick's Mac +
> Floyd's go-live gate.** This packages the **Floyd-gated, live-validated**
> `apple-mail-draft-mcp` server **unchanged** — it is not a new build. The plugin
> adds only: manifest, `.mcp.json`, the SessionStart bootstrap hook, and the
> skill. Requires Cowork (hooks + plugin-bundled local MCP servers run in the
> Cowork desktop app). Still passes Floyd's gate before any ship.

## Layout (mirrors falke-bid-tools; verified against the INSTALL-DESIGN)

```
falke-business-pulse/
├── .claude-plugin/plugin.json    # manifest (name = falke-business-pulse, kebab-case)
├── .mcp.json                     # auto-registers the apple-mail MCP server (command/args/env)
├── hooks/hooks.json              # SessionStart -> scripts/bootstrap.sh
├── scripts/bootstrap.sh          # one-Python-dep persistent-venv bootstrap (trimmed bid-tools pattern)
├── skills/
│   └── falke-business-pulse/     # the /Falke-business-pulse skill (COPY of the phase2-build source)
│       ├── SKILL.md
│       └── reference/            # teams-card.md, categories.md, config.md
└── apple-mail/                   # the EXISTING Floyd-gated MCP server, vendored UNCHANGED
    ├── server.py                 # two tools: read_apple_mail, create_apple_mail_draft (stdio)
    ├── read_core.py / draft_core.py / config.py
    ├── requirements.txt          # mcp>=1.2.0 (the only runtime dep)
    └── applescript/              # the four static .applescript files the core modules call
```

The MCP server is **vendored, not rebuilt** — copied unchanged from
`FALKE/01_Chief_of_Staff/phase2-build/apple-mail-draft-mcp/`. Its core modules
resolve the AppleScript files via `Path(__file__).parent / "applescript"`, so the
`applescript/` directory MUST stay a sibling of `read_core.py`/`draft_core.py` —
which the vendored layout preserves. The server's tests, dev README, and live-test
harness are intentionally NOT vendored (runtime needs only the source + scripts).

## Install (employee — two terminal lines)

> ### ⚠️ One-time prerequisite — Python ≥ 3.10 (do this FIRST)
>
> The bundled mail tool's runtime dependency (the `mcp` SDK) needs **Python
> ≥ 3.10**. **Stock macOS ships Python 3.9**, which is too old — so on a fresh Mac
> you must install a newer Python **before** the install below, or the mail tool
> won't set up.
>
> **Install it once:**
> ```
> brew install python@3.12
> ```
> (Any Python ≥ 3.10 works — 3.10/3.11/3.12/3.13. If you don't have Homebrew, get
> it from https://brew.sh first.)
>
> **This is the one thing that breaks the otherwise dead-simple install if it's
> missed.** You don't have to check your version by hand — the first-run bootstrap
> has a Python-≥3.10 guard that **fails loud** with this exact guidance
> (`ERROR: need Python >=3.10 for the mail tool … Install it (e.g. 'brew install
> python@3.12') and reopen Cowork.`) if no suitable Python is found. Install
> Python, reopen Cowork, and the setup continues.

Then the two terminal lines:

```
/plugin marketplace add https://<private-falke-plugins-repo>.git
/plugin install falke-business-pulse@ara-falke
```

Approve the "Will install" summary (it lists the `apple-mail` MCP server + the
skill — expected), then `/reload-plugins`. First session shows
"Setting up dependencies…" for ~30s while the bootstrap hook prepares the mail
tool. Then: paste the per-person config (below) and grant the one-time macOS
Automation permission. Full step-by-step is the INSTALL-DESIGN + the
employee guide Anna/Maggie polish.

## Per-person config

| What | Where | Default | Secret? |
|---|---|---|---|
| Read account allow-list | `APPLE_MAIL_READ_ALLOWED_ACCOUNTS` (`.mcp.json` env) | `falkecorp.com,falkehoa.com` | no |
| Recipient allow-list (drafts) | `APPLE_MAIL_DRAFT_ALLOWED_DOMAINS` (`.mcp.json` env) | `falkecorp.com` | no |
| From-account allow-list (sender) | `APPLE_MAIL_DRAFT_FROM_ACCOUNTS` (`.mcp.json` env) | `falkecorp.com,falkehoa.com` | no |
| Teams webhook URL | `FALKE_TEAMS_WEBHOOK_URL` (env / keychain) | *(per channel — employee pastes)* | **YES — never commit** |
| Dropbox project folder | local path the skill reads | `~/Library/CloudStorage/Dropbox/<project>` (Available offline) | no |

The three allow-lists ship with safe Falke defaults in `.mcp.json` and are
overridden only if a person's case differs. The **Teams webhook URL is a secret**
and is deliberately NOT in `.mcp.json` — it is read from the environment/keychain
the employee sets, never baked into the plugin or the repo (COND-3).
