#!/usr/bin/env node
// Outlook CLI — single binary, subcommand-based, designed per the patterns in
// https://github.com/lirantal/nodejs-cli-apps-best-practices.
//
// Conventions:
//   - stdout: command result (JSON), pretty-printed when on a TTY, compact
//     when piped — so `outlook list | jq …` Just Works.
//   - stderr: diagnostics, progress, errors. Never pollutes stdout.
//   - Exit codes: 0 ok, 1 generic, 2 auth, 3 HTTP, 64 usage, 130 SIGINT.
//   - Errors carry trackable codes (E_AUTH_*, E_HTTP, E_ARGS) + actionable hints.
//   - STDIN: `send` accepts message JSON via STDIN when no arg is given.

import { Command, CommanderError } from 'commander';
import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';

// NOTE: capture.mjs (and with it Playwright) must never be imported
// statically here — auth.mjs lazy-loads it only when a browser is really
// needed, so cache-hit commands start fast.
import { getAuth, clearAuth, loadCachedAuth, refreshAuth } from './auth.mjs';
import { call } from './client.mjs';
import { printJson, info, errorBlock, debug } from './output.mjs';
import { AppError, E, EXIT, exitCodeFor } from './errors.mjs';
import { buildFilter, buildQuery, resolveFolder, parseCount } from './odata.mjs';
import { RESOURCES } from './resources.mjs';
import { auditAll } from './audit.mjs';
import {
  resolveEventRange,
  calendarViewPath,
  buildEventFilter,
} from './calendar.mjs';
import {
  loadLearnings,
  addLearning,
  removeLearning,
  clearLearnings,
  learningsFile,
} from './learn.mjs';
import { decodePayload } from './jwt.mjs';
import {
  buildFileAttachment,
  attachFilesToDraft,
  validateAttachPath,
} from './attachments.mjs';

/** Commander collector for repeated --attach options. */
function collectAttach(value, previous) {
  return previous ? previous.concat([value]) : [value];
}

const HERE = dirname(fileURLToPath(import.meta.url));
const pkg = JSON.parse(readFileSync(resolve(HERE, '..', 'package.json'), 'utf8'));

const DEFAULT_SELECT = 'Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments,Importance';
const DEFAULT_ORDER_BY = 'ReceivedDateTime desc';

/**
 * Build the `/messages` (or `/mailfolders/{folder}/messages`) path for a
 * list/search call. Shared between `list` and `search` so flag semantics
 * stay identical.
 */
function messagesPath(opts) {
  const base = opts.folder
    ? `/mailfolders/${encodeURIComponent(resolveFolder(opts.folder))}/messages`
    : '/messages';
  const filter = buildFilter({
    unread: opts.unread,
    from: opts.from,
    to: opts.to,
    since: opts.since,
    until: opts.until,
    hasAttachments: opts.hasAttachments,
    importance: opts.importance,
    subject: opts.subject,
    raw: opts.filter,
  });
  const query = buildQuery({
    top: opts.top,
    skip: opts.skip,
    filter,
    orderBy: opts.orderBy ?? DEFAULT_ORDER_BY,
    select: opts.select ?? DEFAULT_SELECT,
    search: opts.search,
  });
  return base + query;
}

// ---------------------------------------------------------------------------
// Signal handling — make sure Ctrl-C doesn't leave stray Chromium processes.

let shuttingDown = false;
function onSignal(sig) {
  if (shuttingDown) return;
  shuttingDown = true;
  info(`Received ${sig}, exiting.`);
  process.exit(EXIT.SIGINT);
}
process.on('SIGINT', () => onSignal('SIGINT'));
process.on('SIGTERM', () => onSignal('SIGTERM'));

// ---------------------------------------------------------------------------
// Shared helpers

/** Human-readable summary of an HTTP error body for the error hint. */
function httpErrorHint(body) {
  // Outlook errors follow the OData shape { error: { code, message } } —
  // surface just that instead of a wall of JSON when we can.
  const err = body?.error;
  if (err?.message) return err.code ? `${err.code}: ${err.message}` : err.message;
  return typeof body === 'object'
    ? JSON.stringify(body).slice(0, 500)
    : String(body).slice(0, 500);
}

function httpError(res, method, path) {
  return new AppError({
    code: E.HTTP,
    message: `${method} ${path.split('?')[0]} returned HTTP ${res.status}.`,
    hint: httpErrorHint(res.body),
  });
}

/**
 * Authenticated API call with automatic 401 recovery: if the cached token
 * is rejected server-side (revoked, CA policy change) we clear it, capture
 * a fresh one, and retry the request once before giving up. Transient
 * network/throttling retries live one level down, in client.call().
 *
 * `init.resource` selects which Microsoft API + token to use (default
 * "outlook"). For non-default resources, getAuth() will not silently open
 * the browser — it errors with a `outlook auth --all` hint instead — so a
 * `graph` call never pops Chromium unexpectedly.
 */
async function runApi(path, init = {}) {
  const method = init.method ?? 'GET';
  const res_ = init.resource ?? 'outlook';
  let auth = await getAuth({ resource: res_ });
  let res = await call(auth, path, init);

  if (res.status === 401) {
    debug(`cached ${res_} token rejected (401); recapturing and retrying`);
    clearAuth(res_);
    auth = await getAuth({ resource: res_ });
    res = await call(auth, path, init);
    if (res.status === 401) {
      clearAuth(res_);
      throw new AppError({
        code: E.AUTH_REQUIRED,
        message: `API rejected a freshly captured ${res_} token (HTTP 401).`,
        hint:
          res_ === 'outlook'
            ? 'Run `outlook auth` to sign in interactively — your session may need MFA or a policy step-up.'
            : `Run \`outlook auth --all\` to re-capture a ${res_} token — your session may need MFA or a policy step-up.`,
      });
    }
  }
  if (res.status >= 400) throw httpError(res, method, path);
  return res.body;
}

async function readStdin() {
  return new Promise((res, rej) => {
    let buf = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', (d) => (buf += d));
    process.stdin.on('end', () => res(buf));
    process.stdin.on('error', rej);
  });
}

/**
 * Resolve a command's JSON payload from its optional argument or STDIN,
 * then parse it. Every mutating command (send/draft/event-*) shares these
 * exact semantics and error messages.
 */
async function readJsonPayload(jsonArg, what) {
  const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
  if (!raw.trim()) {
    throw new AppError({
      code: E.ARGS,
      message: `No ${what} JSON provided.`,
      hint: 'Pass JSON as an argument or pipe it via STDIN.',
    });
  }
  try {
    return JSON.parse(raw);
  } catch (cause) {
    throw new AppError({
      code: E.ARGS,
      message: `${what[0].toUpperCase()}${what.slice(1)} payload was not valid JSON.`,
      hint: 'Validate with `jq .` first, then retry.',
      cause,
    });
  }
}

// ---------------------------------------------------------------------------
// CLI surface

const program = new Command();

// Route commander's own exits through our exit-code scheme: usage mistakes
// (unknown option, missing argument) land on EXIT.USAGE like every other
// bad-input error, while --help / --version stay 0. Must be set before the
// subcommands are created — they copy this setting at creation time.
program.exitOverride();

program
  .name('outlook')
  .description(
    'Outlook mailbox CLI that piggybacks on the Outlook Web (OWA) session.\n' +
      'No app registration, no admin consent — just your existing OWA login.',
  )
  .version(pkg.version, '-V, --version', 'print the CLI version')
  .option('--debug', 'verbose diagnostic output to stderr')
  .hook('preAction', (cmd) => {
    if (cmd.optsWithGlobals().debug) process.env.OUTLOOK_DEBUG = '1';
    debug(`outlook ${pkg.version}, node ${process.version}`);
  });

program
  .command('setup')
  .description(
    'Download Chromium (needed for the first sign-in) and optionally\n' +
      'install the agent skill for Claude Code / Codex / Cursor.',
  )
  .option('--with-skill', 'also install the agent skill into every detected AI agent (Claude/Codex/Cursor)')
  .option('--skill-target <list>', 'comma-separated targets, or "all" / "auto"', 'auto')
  .option('--skip-browser', 'do not download Chromium (you must run `npx playwright install chromium` later)')
  .action(async (opts) => {
    if (!opts.skipBrowser) {
      info('Downloading Chromium via Playwright (~150 MB)…');
      await runChildToCompletion('npx', ['playwright', 'install', 'chromium']);
    }
    if (opts.withSkill) {
      info('Installing skill + slash commands into detected AI agents…');
      const installerPath = resolve(HERE, '..', 'skill', 'install.mjs');
      await runChildToCompletion(process.execPath, [
        installerPath,
        '--target',
        opts.skillTarget,
      ]);
    }
    info('');
    info('Setup complete. Next: run `outlook auth` to sign in (one-time, opens a browser).');
  });

/**
 * Spawn a child process and wait for it to exit. Throws an AppError when
 * the child exits non-zero so the top-level error handler emits a clean
 * message + correct exit code.
 */
function runChildToCompletion(command, args) {
  return new Promise((resolveRun, rejectRun) => {
    const child = spawn(command, args, { stdio: 'inherit' });
    child.on('error', rejectRun);
    child.on('close', (code) => {
      if (code === 0) return resolveRun();
      rejectRun(
        new AppError({
          code: E.UNEXPECTED,
          message: `${command} ${args.join(' ')} exited with code ${code}`,
          hint: 'Re-run with the same command manually to see the underlying error.',
        }),
      );
    });
  });
}

program
  .command('auth')
  .description(
    'Interactive sign-in. Opens a Chromium window pointed at OWA; waits up\n' +
      'to 10 minutes for you to complete SSO + MFA; caches the Bearer token.\n' +
      'With --all it also opens Teams + Copilot and captures a token per\n' +
      'Microsoft resource (Graph, Substrate) in the same session.',
  )
  .option('--all', 'capture tokens for every reachable resource (Graph, Substrate), not just Outlook')
  .action(async (opts) => {
    await refreshAuth({
      timeoutMs: 10 * 60 * 1000,
      interactive: true,
      waitForAll: !!opts.all,
      openExtraSurfaces: !!opts.all,
    });
    info(opts.all ? 'Signed in. Tokens cached. Run `outlook token-audit` to see what you can reach.' : 'Signed in. Token cached.');
  });

program
  .command('refresh')
  .description('Force-refresh the cached Bearer token(s) (opens the browser briefly).')
  .option('--all', 'refresh every resource token, not just Outlook')
  .action(async (opts) => {
    await refreshAuth({ waitForAll: !!opts.all, openExtraSurfaces: !!opts.all });
    info('Token refreshed.');
  });

program
  .command('token-audit')
  .description(
    'Report which Microsoft resources you can reach (Outlook, Graph,\n' +
      'Substrate) and what each cached token is scoped for. Decodes tokens\n' +
      'locally — no network, no browser. Use this to see whether Teams /\n' +
      'Copilot / Graph access is available before calling those commands.',
  )
  .action(() => {
    printJson(auditAll());
  });

// ---------------------------------------------------------------------------
// Cross-resource passthrough. Graph (and Substrate) speak a different base
// URL + token audience than Outlook REST; `graph` is a thin authenticated
// proxy so new endpoints can be exercised without a bespoke subcommand each.

program
  .command('graph')
  .argument('<path>', 'Graph path, e.g. "/me", "/me/chats?$top=10", "/me/joinedTeams"')
  .argument('[json]', 'request body for POST/PATCH/PUT; reads STDIN if omitted')
  .option('-X, --method <verb>', 'HTTP method (GET, POST, PATCH, PUT, DELETE)', 'GET')
  .option('--resource <name>', `target resource: ${Object.keys(RESOURCES).join(' | ')}`, 'graph')
  .description(
    'Authenticated passthrough to Microsoft Graph (or another resource via\n' +
      '--resource). Requires a token for that resource: run `outlook auth --all`\n' +
      'first, then `outlook token-audit` to confirm it was captured.',
  )
  .action(async (path, jsonArg, opts) => {
    if (!RESOURCES[opts.resource]) {
      throw new AppError({
        code: E.ARGS,
        message: `Unknown --resource: ${opts.resource}`,
        hint: `Choose one of: ${Object.keys(RESOURCES).join(', ')}.`,
      });
    }
    const method = opts.method.toUpperCase();
    const init = { method, resource: opts.resource };
    if (method !== 'GET' && method !== 'DELETE') {
      // Body is optional (some Graph actions take none); only attach if given.
      const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
      if (raw.trim()) {
        let body;
        try {
          body = JSON.parse(raw);
        } catch (cause) {
          throw new AppError({
            code: E.ARGS,
            message: 'Request payload was not valid JSON.',
            hint: 'Validate with `jq .` first, then retry.',
            cause,
          });
        }
        init.headers = { 'Content-Type': 'application/json' };
        init.body = JSON.stringify(body);
      }
    }
    const p = path.startsWith('/') ? path : `/${path}`;
    const body = await runApi(p, init);
    printJson(body ?? { ok: true });
  });

// ---------------------------------------------------------------------------
// Microsoft Teams (via Graph). Friendly verbs over the documented Graph
// endpoints, all using the captured graph-audience token (run
// `outlook auth --all` first). Reachability of each depends on the scopes
// your token carries — reads that need Chat.Read may 403 in tenants that
// only grant Chat.ReadBasic; the error hint says which scope is missing.

const gget = (path) => runApi(path, { resource: 'graph' });

program
  .command('teams')
  .description('List the Teams you belong to (Graph /me/joinedTeams).')
  .action(async () => {
    printJson(await gget('/me/joinedTeams?$select=id,displayName,description'));
  });

program
  .command('teams-channels')
  .argument('<teamId>', 'team Id (from `outlook teams`)')
  .description('List channels in a team (Graph /teams/{id}/channels).')
  .action(async (teamId) => {
    printJson(
      await gget(
        `/teams/${encodeURIComponent(teamId)}/channels?$select=id,displayName,description,membershipType`,
      ),
    );
  });

program
  .command('teams-chats')
  .description('List your Teams chats (1:1 and group) — Graph /me/chats.')
  .option('-n, --top <count>', 'how many chats to fetch', '20')
  .action(async (opts) => {
    const top = parseCount(opts.top, '--top');
    printJson(
      await gget(
        `/me/chats?$top=${top}&$select=id,topic,chatType,lastUpdatedDateTime`,
      ),
    );
  });

program
  .command('teams-members')
  .argument('<chatId>', 'chat Id (from `outlook teams-chats`)')
  .description('List members of a Teams chat (Graph /chats/{id}/members).')
  .action(async (chatId) => {
    printJson(await gget(`/chats/${encodeURIComponent(chatId)}/members`));
  });

program
  .command('teams-messages')
  .argument('<chatId>', 'chat Id (from `outlook teams-chats`)')
  .option('-n, --top <count>', 'how many messages to fetch', '20')
  .description(
    'Read messages in a Teams chat (Graph /chats/{id}/messages).\n' +
      'Requires a token with Chat.Read; tenants granting only Chat.ReadBasic\n' +
      'will get a 403 with a scope hint (the Teams web client reads messages\n' +
      'via a separate, undocumented service).',
  )
  .action(async (chatId, opts) => {
    const top = parseCount(opts.top, '--top');
    printJson(await gget(`/chats/${encodeURIComponent(chatId)}/messages?$top=${top}`));
  });

program
  .command('teams-send')
  .argument('<chatId>', 'chat Id to post into (from `outlook teams-chats`)')
  .argument('[text]', 'message text; reads STDIN if omitted')
  .option('--html', 'send the text as HTML instead of plain text')
  .description(
    'Post a message to a Teams chat (Graph /chats/{id}/messages).\n' +
      '**Sends immediately** — confirm the chat + text with the user first,\n' +
      'the same as `send` for email. Requires ChatMessage.Send.',
  )
  .action(async (chatId, text, opts) => {
    const content = (text ?? (process.stdin.isTTY ? '' : await readStdin())).trim();
    if (!content) {
      throw new AppError({
        code: E.ARGS,
        message: 'No message text provided.',
        hint: 'Pass the text as an argument or pipe it via STDIN.',
      });
    }
    const body = await runApi(`/chats/${encodeURIComponent(chatId)}/messages`, {
      resource: 'graph',
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        body: { contentType: opts.html ? 'html' : 'text', content },
      }),
    });
    printJson({ sent: true, MessageId: body?.id ?? null, ChatId: chatId });
  });

// ---------------------------------------------------------------------------
// Self-learning context. The agent reads `outlook context` at the start of
// every mail-related interaction and appends durable observations via
// `outlook learn add "<observation>"`. The file is plain-text markdown the
// user can read/edit/clear at any time.

program
  .command('context')
  .description(
    'Print the context an AI agent should load at session start:\n' +
      'user identity (from cached JWT), accumulated learnings, paths to the\n' +
      'learnings file. Pure read; never makes API calls.',
  )
  .action(() => {
    // Read from cache directly — context must never trigger a Chromium
    // relaunch (it's called at every agent session start, often before the
    // user even asks for mail).
    let user = null;
    let tokenMinutesUntilExpiry = null;
    const cached = loadCachedAuth();
    if (cached) {
      const claims = decodePayload(cached.headers.authorization);
      if (claims) {
        user = {
          email: claims.upn ?? claims.unique_name ?? null,
          name: claims.name ?? null,
          tenant_id: claims.tid ?? null,
        };
      }
      tokenMinutesUntilExpiry = Math.floor(
        (cached.expiresAt - Date.now()) / 60_000,
      );
    }

    printJson({
      user,
      tokenMinutesUntilExpiry,
      learnings: loadLearnings(),
      learningsFile: learningsFile(),
    });
  });

const learn = program
  .command('learn')
  .description('Show, add, forget, or clear the agent\'s learnings about the user.\n' +
    'Examples:\n' +
    '  outlook learn                                # list all learnings\n' +
    '  outlook learn add "Signs off as Sam"         # record an observation\n' +
    '  outlook learn forget "Signs off"             # remove matching entries\n' +
    '  outlook learn clear                          # wipe the file');

learn
  .command('add <text>')
  .description('Record a new observation.')
  .action((text) => {
    const added = addLearning(text);
    printJson({ added, text: text.trim(), file: learningsFile() });
  });

learn
  .command('forget <substring>')
  .description('Remove every learning containing the substring (case-insensitive).')
  .action((substring) => {
    const removed = removeLearning(substring);
    printJson({ removed, substring });
  });

learn
  .command('clear')
  .description('Delete every learning. No confirmation prompt.')
  .action(() => {
    clearLearnings();
    printJson({ cleared: true, file: learningsFile() });
  });

learn
  .command('list', { isDefault: true })
  .description('List all current learnings.')
  .action(() => {
    const items = loadLearnings();
    printJson({ count: items.length, learnings: items, file: learningsFile() });
  });

// ---------------------------------------------------------------------------
// Calendar

const EVENT_SELECT =
  'Id,Subject,Start,End,Location,Organizer,Attendees,IsAllDay,ShowAs,IsCancelled,IsOnlineMeeting,OnlineMeetingUrl,ResponseStatus,BodyPreview';

program
  .command('agenda')
  .description(
    'Show events from your calendar in a time window.\n' +
      'Defaults: from now, for the next 7 days, primary calendar.',
  )
  .option('--from <when>', 'window start (default: now)')
  .option('--to <when>', 'window end (default: 7 days from start)')
  .option('--days <n>', 'shortcut for --to: window length in days')
  .option('-n, --top <count>', 'max events to return', '50')
  .option('-s, --skip <count>', 'pagination offset')
  .option('--calendar <id>', 'calendar Id (default: primary)')
  .option('--organizer <addr>', 'only events organised by this person')
  .option('--show-as <state>', 'Free | Tentative | Busy | Oof | WorkingElsewhere | Unknown')
  .option('--subject <text>', 'subject contains substring (case-insensitive)')
  .option('--filter <odata>', 'raw $filter expression — ANDed with the friendly flags')
  .option('--order-by <expr>', 'OData $orderby (default: "Start/DateTime asc")')
  .option('--select <fields>', 'override the default field set')
  .action(async (opts) => {
    const { start, end } = resolveEventRange({
      from: opts.from,
      to: opts.to,
      days: opts.days,
    });
    const filter = buildEventFilter({
      organizer: opts.organizer,
      showAs: opts.showAs,
      subject: opts.subject,
      raw: opts.filter,
    });
    const base = opts.calendar
      ? `/calendars/${encodeURIComponent(opts.calendar)}`
      : '';
    const path =
      base +
      calendarViewPath(
        { start, end },
        {
          top: opts.top,
          skip: opts.skip,
          filter,
          orderBy: opts.orderBy ?? 'Start/DateTime asc',
          select: opts.select ?? EVENT_SELECT,
        },
      );
    const body = await runApi(path);
    printJson(body);
  });

program
  .command('events')
  .description(
    'Generic event list (no time-range expansion). Use this for filter-\n' +
      'heavy queries or to see recurring master series; for "what\'s on my\n' +
      'calendar this week", use `agenda` instead.',
  )
  .option('-n, --top <count>', 'how many events to fetch', '25')
  .option('-s, --skip <count>', 'pagination offset')
  .option('--organizer <addr>', 'only events organised by this person')
  .option('--show-as <state>', 'Free | Tentative | Busy | Oof | WorkingElsewhere | Unknown')
  .option('--subject <text>', 'subject contains substring')
  .option('--all-day', 'only all-day events')
  .option('--cancelled', 'only cancelled events')
  .option('--filter <odata>', 'raw $filter expression')
  .option('--order-by <expr>', `OData $orderby (default: "Start/DateTime desc")`)
  .option('--select <fields>', 'override the default field set')
  .action(async (opts) => {
    const filter = buildEventFilter({
      organizer: opts.organizer,
      showAs: opts.showAs,
      subject: opts.subject,
      isAllDay: opts.allDay,
      isCancelled: opts.cancelled,
      raw: opts.filter,
    });
    const query = buildQuery({
      top: opts.top,
      skip: opts.skip,
      filter,
      orderBy: opts.orderBy ?? 'Start/DateTime desc',
      select: opts.select ?? EVENT_SELECT,
    });
    const body = await runApi(`/events${query}`);
    printJson(body);
  });

program
  .command('event-read')
  .argument('<id>', 'event Id')
  .description('Read a single event in full.')
  .action(async (id) => {
    const body = await runApi(`/events/${encodeURIComponent(id)}`);
    printJson(body);
  });

program
  .command('calendars')
  .description('List the user\'s calendars (primary + secondary + shared).')
  .action(async () => {
    const body = await runApi('/calendars');
    printJson(body);
  });

program
  .command('event-create')
  .argument('[json]', 'Outlook Event JSON; reads STDIN if omitted')
  .description(
    'Create a calendar event. **If the event has Attendees, invitations\n' +
      'are sent IMMEDIATELY** — confirm with the user first. For meetings,\n' +
      'consider running with no Attendees first, then adding them in\n' +
      'Outlook after review.',
  )
  .action(async (jsonArg) => {
    const message = await readJsonPayload(jsonArg, 'event');
    const body = await runApi('/events', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(message),
    });
    printJson(body);
  });

program
  .command('event-update')
  .argument('<id>', 'event Id to modify')
  .argument('[json]', 'partial Event override; reads STDIN if omitted')
  .description(
    'Update an event. PATCH semantics — only fields in the JSON are\n' +
      'changed. If the event has attendees, an update notification is\n' +
      'usually sent.',
  )
  .action(async (id, jsonArg) => {
    const patch = await readJsonPayload(jsonArg, 'override');
    const body = await runApi(`/events/${encodeURIComponent(id)}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(patch),
    });
    printJson(body);
  });

program
  .command('event-cancel')
  .argument('<id>', 'event Id to cancel')
  .description(
    'Cancel/delete an event. If the event has attendees, **cancellation\n' +
      'notifications are sent immediately**. Confirm with the user first.',
  )
  .action(async (id) => {
    await runApi(`/events/${encodeURIComponent(id)}`, { method: 'DELETE' });
    printJson({ cancelled: true, EventId: id });
  });

for (const verb of ['accept', 'decline', 'tentative']) {
  const action =
    verb === 'tentative' ? 'tentativelyAccept' : verb; // Outlook uses tentativelyAccept
  program
    .command(verb)
    .argument('<id>', 'event Id')
    .option('-c, --comment <text>', 'response comment included in the RSVP')
    .option('--no-respond', 'do not notify the organiser of your response')
    .description(`RSVP "${verb}" to a meeting.`)
    .action(async (id, opts) => {
      const payload = { SendResponse: opts.respond !== false };
      if (opts.comment) payload.Comment = opts.comment;
      await runApi(`/events/${encodeURIComponent(id)}/${action}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      printJson({ rsvp: verb, EventId: id, SendResponse: payload.SendResponse });
    });
}

program
  .command('free-busy')
  .description(
    'Get schedule (free/busy) info for one or more people in a window.\n' +
      'Defaults: now → +1d, 30-minute slot granularity.',
  )
  .argument('<emails...>', 'one or more email addresses')
  .option('--from <when>', 'window start (default: now)')
  .option('--to <when>', 'window end (default: 24h from start)')
  .option('--interval <minutes>', 'slot granularity', '30')
  .action(async (emails, opts) => {
    const { start, end } = resolveEventRange({
      from: opts.from,
      to: opts.to,
      days: opts.to ? null : 1, // default window is 24h not 7d
    });
    const body = await runApi('/getSchedule', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        Schedules: emails,
        StartTime: { DateTime: start.toISOString(), TimeZone: 'UTC' },
        EndTime: { DateTime: end.toISOString(), TimeZone: 'UTC' },
        AvailabilityViewInterval: parseCount(opts.interval, '--interval'),
      }),
    });
    printJson(body);
  });

program
  .command('logout')
  .description('Clear the local token cache (does not sign you out of OWA).')
  .action(() => {
    clearAuth();
    info('Local token cache cleared.');
  });

/**
 * Attach the shared filter/pagination/selection options to a command.
 * Used by both `list` and `search` so they share semantics.
 */
function withListOptions(cmd) {
  return cmd
    .option('-n, --top <count>', 'how many messages to fetch', '10')
    .option('-s, --skip <count>', 'pagination offset', '0')
    .option('--folder <name>', 'folder to read from (Inbox, Sent, Drafts, Deleted, Junk, Outbox, Archive, or a folder Id)')
    .option('--unread', 'only unread messages')
    .option('--has-attachments', 'only messages with attachments')
    .option('--from <addr>', 'messages from this address')
    .option('--to <addr>', 'messages addressed to this address')
    .option('--since <when>', 'received after this time (e.g. 7d, 24h, 2026-01-15)')
    .option('--until <when>', 'received before this time (same syntax as --since)')
    .option('--importance <level>', 'Low | Normal | High')
    .option('--subject <text>', 'subject contains this substring (case-insensitive)')
    .option('--filter <odata>', 'raw $filter expression — ANDed with the friendly flags')
    .option('--order-by <expr>', `OData $orderby (default: "${DEFAULT_ORDER_BY}")`)
    .option('--select <fields>', 'CSV of fields to return (default: a sensible set)');
}

withListOptions(
  program
    .command('list')
    .description('List messages from a folder with optional filters.\nDefaults: 10 most recent from Inbox.'),
).action(async (opts) => {
  const body = await runApi(messagesPath(opts));
  printJson(body);
});

withListOptions(
  program
    .command('search')
    .description('Search messages. Same filters as `list` plus a free-text query.')
    .argument('<query>', 'KQL-style search expression (quoted by the CLI)'),
).action(async (query, opts) => {
  // $search and $orderby are mutually exclusive server-side; buildQuery
  // drops orderBy automatically when search is set.
  const body = await runApi(messagesPath({ ...opts, search: query }));
  printJson(body);
});

withListOptions(
  program
    .command('unread')
    .description('Shortcut for `list --unread`. Accepts the same filters as `list`.'),
).action(async (opts, cmd) => {
  // Same surface as `list`, but unread-only and a larger default page.
  const top = cmd.getOptionValueSource('top') === 'default' ? '25' : opts.top;
  const body = await runApi(messagesPath({ ...opts, top, unread: true }));
  printJson(body);
});

program
  .command('read')
  .argument('<id>', 'message Id (from `outlook list`)')
  .description('Read a single message.')
  .action(async (id) => {
    const body = await runApi(`/messages/${encodeURIComponent(id)}`);
    printJson(body);
  });

program
  .command('folders')
  .description('List mail folders.')
  .action(async () => {
    const body = await runApi('/mailfolders');
    printJson(body);
  });

/**
 * Common implementation for draft-* commands.
 *
 * The bug we're avoiding: if you call createReply / createReplyAll /
 * createForward and then PATCH the resulting draft with a full
 * `Body` resource, the server REPLACES the body entirely — wiping out
 * the quoted original thread that createReply had just populated.
 *
 * Fix: the user's new reply text goes into the `Comment` field of the
 * createReply request body. The server then writes a draft whose body
 * looks like `<your comment>\n\n<quoted thread>` (or the HTML
 * equivalent), preserving the thread. Other override fields
 * (CcRecipients, Subject changes, etc.) are PATCHed after, since they
 * don't conflict with body composition.
 *
 * Returns the draft Id and the OWA WebLink (deep link the user can click).
 */
async function makeDraftFrom(messageId, action, overridesJson, attachPaths = []) {
  // Validate attachment paths BEFORE creating the draft so a bad path
  // doesn't leave an orphan draft in the user's mailbox.
  for (const p of attachPaths) validateAttachPath(p);

  // Parse overrides up-front so we can fail fast and split body vs non-body.
  let overrides = null;
  if (overridesJson?.trim()) {
    try {
      overrides = JSON.parse(overridesJson);
    } catch (cause) {
      throw new AppError({
        code: E.ARGS,
        message: 'Override payload was not valid JSON.',
        hint: 'Pass a partial Outlook Message resource (Body, ToRecipients, CcRecipients, …).',
        cause,
      });
    }
  }

  let bodyComment = null;
  let nonBodyOverrides = null;
  if (overrides) {
    const { Body, ...rest } = overrides;
    // Comment is plain text per the Outlook API contract. Server formats it
    // into whatever ContentType matches the thread.
    if (Body && typeof Body.Content === 'string') {
      bodyComment = Body.Content;
    }
    if (Object.keys(rest).length > 0) {
      nonBodyOverrides = rest;
    }
  }

  // ---- 1. createReply / createReplyAll / createForward (with Comment) ----
  const createInit = { method: 'POST' };
  if (bodyComment !== null) {
    createInit.headers = { 'Content-Type': 'application/json' };
    createInit.body = JSON.stringify({ Comment: bodyComment });
  }
  const draft = await runApi(
    `/messages/${encodeURIComponent(messageId)}/${action}`,
    createInit,
  );

  // ---- 2. PATCH non-body overrides (CcRecipients, Subject, etc.) --------
  if (nonBodyOverrides) {
    await runApi(`/messages/${encodeURIComponent(draft.Id)}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(nonBodyOverrides),
    });
  }

  // ---- 3. Attach files (if any) ----------------------------------------
  let attached = null;
  if (attachPaths.length > 0) {
    attached = await attachFilesToDraft(runApi, draft.Id, attachPaths);
  }

  return {
    DraftId: draft.Id,
    WebLink: draft.WebLink ?? null,
    ConversationId: draft.ConversationId ?? null,
    Attachments: attached,
    message: 'Draft saved to your Drafts folder. Open Outlook to review and send.',
  };
}

program
  .command('draft')
  .argument('[json]', 'Outlook Message JSON; reads STDIN if omitted')
  .option('-a, --attach <path>', 'attach a local file (repeatable, ≤3MB each)', collectAttach)
  .description(
    'Create a new draft (does NOT send). Same JSON shape as `send`. The\n' +
      'draft appears in your Drafts folder for review.',
  )
  .action(async (jsonArg, opts) => {
    const message = await readJsonPayload(jsonArg, 'message');
    // Validate attachments upfront so a bad path fails before we create
    // an empty orphan draft.
    if (opts.attach?.length) {
      for (const p of opts.attach) validateAttachPath(p);
    }
    const body = await runApi('/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(message),
    });
    let attached = null;
    if (opts.attach?.length) {
      attached = await attachFilesToDraft(runApi, body.Id, opts.attach);
    }
    printJson({
      DraftId: body.Id,
      WebLink: body.WebLink ?? null,
      Attachments: attached,
      message: 'Draft saved to your Drafts folder. Open Outlook to review and send.',
    });
  });

program
  .command('draft-reply')
  .argument('<id>', 'message Id to reply to')
  .argument(
    '[json]',
    'partial Message override (e.g. {"Body":{"ContentType":"Text","Content":"…"}}); reads STDIN if omitted',
  )
  .option('-a, --attach <path>', 'attach a local file (repeatable, ≤3MB each)', collectAttach)
  .description(
    'Create a draft reply to a message. Outlook automatically fills in the\n' +
      'quoted thread; your override JSON sets the new body or extra recipients.',
  )
  .action(async (id, jsonArg, opts) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    printJson(await makeDraftFrom(id, 'createReply', raw, opts.attach ?? []));
  });

program
  .command('draft-reply-all')
  .argument('<id>', 'message Id to reply-all to')
  .argument('[json]', 'partial Message override; STDIN if omitted')
  .option('-a, --attach <path>', 'attach a local file (repeatable, ≤3MB each)', collectAttach)
  .description('Like `draft-reply`, but addresses everyone on the original thread.')
  .action(async (id, jsonArg, opts) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    printJson(await makeDraftFrom(id, 'createReplyAll', raw, opts.attach ?? []));
  });

program
  .command('draft-forward')
  .argument('<id>', 'message Id to forward')
  .argument('[json]', 'partial Message override (typically ToRecipients + Body); STDIN if omitted')
  .option('-a, --attach <path>', 'attach a local file (repeatable, ≤3MB each)', collectAttach)
  .description('Create a draft forward of a message.')
  .action(async (id, jsonArg, opts) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    printJson(await makeDraftFrom(id, 'createForward', raw, opts.attach ?? []));
  });

program
  .command('discard-draft')
  .argument('<id>', 'draft Id to delete')
  .description('Permanently delete a draft.')
  .action(async (id) => {
    // 204 No Content on success; runApi treats that as a non-error and returns
    // an empty body. We synthesise the success payload ourselves.
    await runApi(`/messages/${encodeURIComponent(id)}`, { method: 'DELETE' });
    printJson({ discarded: true, DraftId: id });
  });

program
  .command('send')
  .argument(
    '[json]',
    'message JSON (Outlook REST "Message" resource); reads STDIN if omitted',
  )
  .option('-a, --attach <path>', 'attach a local file (repeatable, ≤3MB each)', collectAttach)
  .description(
    'Send a message.\n' +
      'Example STDIN payload:\n' +
      '  { "Subject": "hi", "Body": {"ContentType": "Text", "Content": "…"},\n' +
      '    "ToRecipients": [{"EmailAddress": {"Address": "x@y.com"}}] }',
  )
  .action(async (jsonArg, opts) => {
    const message = await readJsonPayload(jsonArg, 'message');
    if (opts.attach?.length) {
      const built = opts.attach.map(buildFileAttachment);
      message.Attachments = [...(message.Attachments ?? []), ...built];
    }
    const body = await runApi('/sendmail', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ Message: message, SaveToSentItems: true }),
    });
    // sendmail returns 202 with no body on success.
    printJson(body || { sent: true });
  });

// ---------------------------------------------------------------------------
// Run

try {
  await program.parseAsync(process.argv);
} catch (e) {
  if (e instanceof CommanderError) {
    // Commander has already printed its message to the right stream.
    process.exit(e.exitCode === 0 ? EXIT.OK : EXIT.USAGE);
  }
  if (e instanceof AppError) {
    errorBlock(e.code, e.message, e.hint);
  } else {
    errorBlock(E.UNEXPECTED, e.message ?? String(e));
    if (process.env.OUTLOOK_DEBUG) process.stderr.write((e.stack ?? '') + '\n');
  }
  process.exit(exitCodeFor(e));
}
