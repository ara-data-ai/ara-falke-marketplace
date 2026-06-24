# Per-person config the skill + MCP server read

Five pieces of per-person config drive the whole routine. Four are set as MCP
server **env vars** in the plugin's `.mcp.json` (or the host MCP config); one is a
local **folder path**. None of these are secrets EXCEPT the Teams webhook URL,
which must never be committed or synced.

| What | Where it's read | Default / example | Notes |
|---|---|---|---|
| **Read account allow-list** | `APPLE_MAIL_READ_ALLOWED_ACCOUNTS` (MCP env) | `falkecorp.com,falkehoa.com` | Both Falke domains (R2). Matched on email **domain**, not display name. Empty/garbage ⇒ read nothing (fail closed). A personal account is skipped automatically. |
| **From-account allow-list (sender)** | `APPLE_MAIL_DRAFT_FROM_ACCOUNTS` (MCP env) | `falkecorp.com,falkehoa.com` | Bounds which account a draft can be composed **FROM** — the person's own Falke mailbox(es). `create_apple_mail_draft`'s required `from_account` must be on this list, or the draft is rejected. A nudge is always drafted from the person's own account so it lands in their Drafts and sends from their address. |
| **Recipient allow-list (drafts)** | `APPLE_MAIL_DRAFT_ALLOWED_DOMAINS` (MCP env) | `falkecorp.com` (add `falkehoa.com` if drafting for that domain) | Bounds who an injection could ever draft to. Keep conservative. |
| **Teams webhook URL** | `FALKE_TEAMS_WEBHOOK_URL` (env / keychain — **SECRET**) | *(per channel)* | **Never** commit, never write to the digest, never put in the Dropbox folder. Bound to one channel. Rotate if leaked (COND-3). |
| **Dropbox project folder** | local path the skill reads | `~/Library/CloudStorage/Dropbox/<Falke project folder>` | Must be set **"Available offline"** so Cowork reads files on disk, not cloud placeholders. `[VERIFY]` exact path per Mac. |

Run-state (not config): `state/last-run.txt` in the project folder holds the last
successful run timestamp = the next run's `since_iso` cutoff. Absent ⇒ default to
24h ago.

> The install guide (`INSTALL-falke-business-pulse-USERS.md`) collects these
> values per person in one short form so the employee never has to understand the
> mechanism — they paste a few things (the three allow-lists ship with safe Falke
> defaults) and grant one macOS permission.
