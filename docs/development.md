# Development

How to work on this codebase, diagnose new endpoints, and add features.

## Local setup

```bash
npm install
npm run setup           # playwright install chromium
./src/cli.mjs auth      # sign in once
./src/cli.mjs list -n 3 # smoke test
```

The shebang on `src/cli.mjs` makes it executable directly. `npm link` (once)
adds `outlook` to your `$PATH`.

## File-by-file map

| File | Touch when |
| --- | --- |
| `src/cli.mjs` | Adding/removing subcommands, changing the CLI surface |
| `src/auth.mjs` | Token cache load/save/clear, JWT expiry decoding |
| `src/capture.mjs` | Playwright launcher + Bearer-header capture (lazy-imported) |
| `src/client.mjs` | Pure-fetch `call()` wrapper, `OUTLOOK_API_BASE` override |
| `src/resources.mjs` | Microsoft resource registry (outlook/graph/substrate): base URLs, audiences, token classification |
| `src/audit.mjs` | `token-audit` — decode cached tokens + group scopes into capability areas (offline) |
| `src/odata.mjs` | Mail filter / query builders (`--unread`, `--from`, `--folder`, …) |
| `src/calendar.mjs` | Calendar time parsing, calendarView URL builder, event filters |
| `src/learn.mjs` | Persistent learnings file (read/append/forget/clear) |
| `src/jwt.mjs` | Tiny JWT payload decoder shared by auth + cli |
| `src/output.mjs` | Adjusting stdout/stderr discipline or JSON formatting |
| `src/errors.mjs` | Adding a new error code or exit code |
| `src/paths.mjs` | Moving cache or profile directories |
| `src/diagnose.mjs` | Dev tool — sniff OWA endpoints. Edit when investigating new APIs |
| `src/sniff.mjs` | Dev tool — PII-safe live capture: logs method + redacted path + token audience + JSON shape while you click, and saves per-resource tokens. Used to reverse-engineer Teams/Copilot (LEARNINGS §14) |

## Code style

- ESM only (`.mjs`, `"type": "module"`).
- Plain JS with JSDoc where useful. No TypeScript.
- 2-space indent, single quotes, trailing commas in multi-line literals.
- **stdout = command output only.** Everything else goes through `info()`,
  `debug()`, or `errorBlock()` in `output.mjs`. This is what makes `outlook
  list | jq …` work.
- New error paths construct an `AppError` with a `code` (from `E.*`) and a
  `hint` written for the user, not the developer.
- Keep dependencies minimal. Right now: `commander`, `playwright`. Justify
  any addition.

## Adding a new subcommand

1. Add a `program.command(...)` block in `cli.mjs` with description,
   arguments, and options.
2. The action handler calls `runApi(path, init)` for anything that hits the
   Outlook REST API. Return the parsed body to `printJson()`.
3. If the command involves a new failure mode, add an entry to `E.*` in
   `errors.mjs` and an exit-code mapping in `exitCodeFor()`.
4. Document the new command in `docs/usage.md`.
5. Smoke-test with `--debug` to see the request URL and confirm the
   captured headers are still valid.

Example — fetching a single calendar event:

```js
program
  .command('event')
  .argument('<id>', 'event Id')
  .description('Read a single calendar event.')
  .action(async (id) => {
    const body = await runApi(`/events/${encodeURIComponent(id)}`);
    printJson(body);
  });
```

## Adding a new endpoint that needs a different audience

Right now the captured Bearer is audience `https://outlook.office.com/`. If
you want to call Graph (audience `https://graph.microsoft.com/`), you need
a second silent acquisition.

The cleanest way: in `capture.mjs`'s `captureAuth()`, after the OWA Bearer is
captured, evaluate a small script in the page that triggers OWA's internal
`acquireTokenSilent({scopes: ['https://graph.microsoft.com/.default']})`
call — Outlook's MSAL.js instance is exposed as
`window.MSALWrapper`-style globals in some builds. Watch the next set of
requests for a Bearer to graph.microsoft.com and capture it too.

This isn't implemented yet — file an issue if you need it.

## Diagnosing changed endpoints

When Microsoft changes things, the diagnostic sniffer is your best friend:

```bash
node src/diagnose.mjs
```

It opens OWA in **headed** mode, lets it boot completely, then prints every
API endpoint OWA called in the first 60 seconds along with:

- HTTP method
- Full URL (truncated to 140 chars in the summary)
- Whether the request used Bearer or cookie auth
- Any `x-*`, `prefer`, `action`, `content-type` headers

Use the output to spot:

- **New URL patterns.** If `/api/v2.0/me/messages` is gone, OWA will be
  hitting whatever replaced it.
- **Required headers we're missing.** If a request OWA makes successfully
  has `x-foo-bar: baz` and we don't forward it, copy it into
  `FORWARDED_HEADERS` in `capture.mjs`.
- **Different bearer audiences.** Watch for requests to `graph.microsoft.com`
  or `substrate.office.com` — those use different JWTs OWA acquires
  separately.

## Testing

```bash
npm test          # all unit + integration tests, no network, no Playwright
```

The suite uses Node's built-in `node:test` runner — no Jest, no Vitest. It
lives in `test/` and has two kinds of tests:

**Unit tests** import modules directly and assert against pure functions:
- `test/errors.test.mjs` — `AppError`, `exitCodeFor` mapping, frozen enums
- `test/auth.test.mjs` — token-cache lifecycle: load / save / clear / expiry windows / malformed JSON

**Integration tests** spawn the CLI as a subprocess against a local mock
HTTP server (`test/cli.test.mjs`). This catches real wire-format bugs and
exit-code regressions in one shot. The helpers in `test/helpers.mjs`:

- `makeFakeJwt({ expSeconds })` — synthesise a JWT with a chosen expiry.
- `seedTokenCache({ secondsFromNow })` — write a fake auth cache file and
  return its path; pass via `OUTLOOK_TOKEN_CACHE` env var.
- `seedCaptureFixture({ secondsFromNow })` — write a captured-headers file
  for the `OUTLOOK_CAPTURE_FIXTURE` seam, letting tests exercise the
  automatic 401-recapture path without Playwright.
- `startMockServer(handler)` — `http.createServer` on a random port, returns
  `{ url, close }`. Pass via `OUTLOOK_API_BASE` env var.
- `runCli(args, { env, stdin, timeoutMs })` — spawn `src/cli.mjs`, return
  `{ code, stdout, stderr }`. Sets `OUTLOOK_NO_CAPTURE=1` by default so no
  test can accidentally launch Chromium.

Two env seams in `auth.mjs` keep the browser out of tests:

| Variable | Effect |
| --- | --- |
| `OUTLOOK_NO_CAPTURE=1` | Any code path that would launch Chromium throws `E_AUTH_REQUIRED` instead |
| `OUTLOOK_CAPTURE_FIXTURE=/path` | "Capture" returns the headers JSON in that file — simulates a successful browser run |

### Adding a test for a new command

In `test/cli.test.mjs`:

```js
test('events lists calendar items from /events', async () => {
  let observedPath = null;
  const mock = await startMockServer((req, res) => {
    observedPath = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [{ Subject: 'standup' }] }));
  });

  try {
    const { code, stdout } = await runCli(['events'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedPath, '/api/v2.0/me/events');
    assert.equal(JSON.parse(stdout).value[0].Subject, 'standup');
  } finally {
    await mock.close();
  }
});
```

### A real send test (manual, uses your live mailbox)

```bash
cat <<'JSON' | ./src/cli.mjs send
{ "Subject": "test", "Body": {"ContentType":"Text","Content":"hi"},
  "ToRecipients": [{"EmailAddress":{"Address":"your-other-account@example.com"}}] }
JSON
```

### Don't add to the suite

- Tests that hit `outlook.office.com` for real. Mock the network.
- Tests that exercise Playwright. The auth-capture path is e2e-only; verify
  it manually with `./src/cli.mjs auth`.
- Tests that depend on a specific JWT lifetime. We use `secondsFromNow` so
  cache freshness is computed at test time.

## Release / publish

Not yet published to npm. If/when we do:

1. Bump `version` in `package.json`.
2. Update `LEARNINGS.md` if any new dead ends were discovered.
3. Tag the commit: `git tag v0.x.0 && git push --tags`.
4. `npm publish --access public` (assuming the name is free).
5. Verify `npx outlook-experiment --help` works in a clean directory.

The `files` array in `package.json` already limits what gets published to
`src/`, `README.md`, and `LICENSE`. No docs, no LEARNINGS, no `.gitignore`.
