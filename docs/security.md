# Security & threat model

This is an end-user experiment that touches credentials. Read this before
using it on anything you care about.

## What the tool actually does

1. Launches a Chromium instance with a persistent profile under
   `~/.local/share/outlook-spike/browser-profile/`. You sign in interactively
   the first time. Your password and MFA factors are typed into Microsoft's
   real login page — the CLI never sees them.
2. After sign-in, OWA performs its normal silent OAuth handshake and starts
   making API calls. The CLI captures the `Authorization: Bearer …` header
   from those requests and saves it to
   `~/.cache/outlook-spike/auth.json`.
3. Subsequent CLI calls read the cached header and make the same API calls
   from Node `fetch` directly.

## What the tool does **not** do

- No keylogging. Your password is typed into Microsoft's UI.
- No new OAuth consent. The token has only the scopes you previously
  consented to when first signing into OWA.
- No outbound communication to anything except `outlook.office.com` and
  `login.microsoftonline.com`.
- No telemetry. There is no analytics layer.

## What's sensitive

| Path | Sensitivity |
| --- | --- |
| `~/.cache/outlook-spike/auth.json` | Contains a JWT with read+send access to your mailbox. Anyone who reads this file can act as you for ~24h. |
| `~/.local/share/outlook-spike/browser-profile/` | Contains SSO cookies. Anyone with this directory can sign in as you without MFA. |
| `~/.local/share/outlook-spike/learnings.md` | Free-form notes the AI agent has accumulated. May contain names, email addresses, project names, and habits. Plain-text markdown; user-editable. |

These are stored with default user-only permissions. **Don't share them.
Don't sync them to cloud storage. Don't commit them.** The repo `.gitignore`
already protects against `auth.json` and any `auth-cache.json` filename
landing in-tree, but cloud-drive sync (iCloud Desktop, OneDrive, Dropbox
selective-sync) is up to you.

## Threats this is and isn't safe against

| Threat | Posture |
| --- | --- |
| You typing your password into a fake login page | **Safe** — the login UI is Microsoft's. |
| Malware that already runs as your user | **Not safe** — it can read `auth.json`. Same exposure as any other credential file (browser cookies, SSH keys). |
| A shared / multi-user machine | **Not safe** — anyone with shell access as you can use the token. Don't use this on shared machines. |
| Enterprise EDR / DLP monitoring | **Detectable** — mailbox API calls from a non-Outlook process pattern stand out. See "EDR considerations" below. |
| Microsoft revoking the token mid-call | **Safe** — the CLI catches 401, clears the cache, and exits with `E_AUTH_REQUIRED` telling you to re-run. |

## EDR considerations

If your employer runs Crowdstrike, Defender for Endpoint, or similar tooling:

- The pattern of "Node.js process making OAuth API calls to
  `outlook.office.com/api/v2.0/...`" is unusual on most managed devices and
  may trigger alerts.
- Capturing Bearer tokens from page network traffic is exactly the technique
  documented in attack frameworks like MITRE ATT&CK T1528. Your EDR
  probably has detection rules for it.
- The browser profile and token cache files outside the standard browser
  profile locations may be flagged as credential-staging behaviour.

**Recommendation:** if you're using this on a managed work device,
proactively let your security team know what you're doing. "I'm running a
Node script that reads my own Outlook inbox by capturing the Bearer my OWA
session already produces" is a sentence that's far better said before they
see the alert than after.

## Tenant-policy considerations

Even if your machine isn't monitored, your tenant may have policies you'd
unknowingly be skirting:

- **Conditional Access requires app compliance.** If your tenant only
  allows access from compliant devices, the Bearer the CLI captures is
  scoped to that compliance — which is fine if your Mac is enrolled. It's
  not a way to use Outlook from an unmanaged device.
- **Acceptable Use Policy.** Some organisations explicitly prohibit
  unsanctioned tooling that accesses corporate data programmatically, even
  by the data owner. Check yours.
- **Data Loss Prevention (DLP).** Sending mail via this CLI bypasses any
  client-side DLP plugins that hook into Outlook. Server-side DLP still
  applies, but client-side scanning won't.

## What to do if something goes wrong

| Situation | What to do |
| --- | --- |
| Token cache file leaked (sent in an email, committed by mistake, …) | Sign out of OWA in your browser. The cached Bearer is still valid for up to ~24h regardless — you can't directly revoke a delegated JWT, but signing out invalidates the refresh chain so no new token can be issued. Then `outlook logout` locally. |
| Browser profile leaked | Change your Microsoft account password. Sign out of all sessions from `https://account.microsoft.com/security`. |
| You suspect EDR flagged you | Reach out to your security team first, don't try to delete logs. Be honest about what you did. |

## Self-learning behaviour

The CLI maintains a `learnings.md` file the AI agent reads at session
start (`outlook context`) and appends to (`outlook learn add`). The
intent is to let your agent improve over time — knowing aliases like
"the team", your sign-off, your preferred tone, etc.

**Privacy posture:**
- The file is **local-only**. It's never uploaded, synced, or
  transmitted. The CLI never reads it during API calls.
- The user (you) can view, edit, or wipe the file at any time:
  - `outlook learn` → list all entries
  - `outlook learn forget "<text>"` → remove matching entries
  - `outlook learn clear` → wipe entirely
  - Or just edit the markdown file directly with your favourite editor.
- The SKILL.md tells the agent: never record sensitive content (financial
  details, HR context, private contact info) without asking first.
- The file is plain-text markdown — if a future you (or a colleague
  reviewing the file) finds an entry uncomfortable, deleting it is
  one command.

**Recommendations:**
- Don't sync `~/.local/share/outlook-spike/` to cloud storage if you
  share your account with others or use the machine for both work and
  personal contexts.
- Periodically run `outlook learn` to audit what's been recorded.
- After job/project changes that touch your inbox patterns, run
  `outlook learn clear` to reset; the agent will rebuild context based
  on new behaviour.

## Open security questions / unfinished work

- The token cache is plaintext JSON. A future version should at least file-
  permission-restrict it to `0600` explicitly and ideally seal it with the
  macOS keychain via `security add-generic-password` so reading requires a
  Touch ID prompt.
- The browser profile inherits Chromium's default permission model.
  Sandboxing it inside a separate user account would be a defence-in-depth
  improvement but is outside the scope of a CLI tool.
