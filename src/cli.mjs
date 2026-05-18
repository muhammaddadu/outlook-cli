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

import { Command } from 'commander';
import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';

import { getAuth, saveAuth, clearAuth } from './auth.mjs';
import { captureAuth } from './capture.mjs';
import { call } from './client.mjs';
import { printJson, info, errorBlock, debug } from './output.mjs';
import { AppError, E, EXIT, exitCodeFor } from './errors.mjs';
import { buildFilter, buildQuery, resolveFolder } from './odata.mjs';

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
// Shared helper

async function runApi(path, init = {}) {
  const auth = await getAuth();
  const res = await call(auth, path, init);

  if (res.status === 401) {
    clearAuth();
    throw new AppError({
      code: E.AUTH_REQUIRED,
      message: 'API returned 401; cached token was rejected.',
      hint: 'Re-run the same command — the cache has been cleared and a fresh token will be captured.',
    });
  }
  if (res.status >= 400) {
    throw new AppError({
      code: E.HTTP,
      message: `API returned HTTP ${res.status}.`,
      hint:
        typeof res.body === 'object'
          ? JSON.stringify(res.body)
          : String(res.body).slice(0, 500),
    });
  }
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

// ---------------------------------------------------------------------------
// CLI surface

const program = new Command();

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
      'to 10 minutes for you to complete SSO + MFA; caches the Bearer token.',
  )
  .action(async () => {
    const auth = await captureAuth({ timeoutMs: 10 * 60 * 1000 });
    saveAuth(auth);
    info('Signed in. Token cached.');
  });

program
  .command('refresh')
  .description('Force-refresh the cached Bearer token (opens OWA briefly).')
  .action(async () => {
    const auth = await captureAuth();
    saveAuth(auth);
    info('Token refreshed.');
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

program
  .command('unread')
  .description('Shortcut for `list --unread`.')
  .option('-n, --top <count>', 'how many messages to fetch', '25')
  .action(async (opts) => {
    const body = await runApi(messagesPath({ ...opts, unread: true }));
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
 * Common implementation for draft-* commands: createReply / createReplyAll /
 * createForward all return a freshly-minted draft message in the Drafts
 * folder. If the user supplied JSON, PATCH the draft so the body / extra
 * recipients land on it before the user reviews in Outlook.
 *
 * Returns the draft Id and the OWA WebLink (deep link the user can click).
 */
async function makeDraftFrom(messageId, action, overridesJson) {
  const auth = await getAuth();
  const create = await call(
    auth,
    `/messages/${encodeURIComponent(messageId)}/${action}`,
    { method: 'POST' },
  );
  if (create.status === 401) {
    clearAuth();
    throw new AppError({
      code: E.AUTH_REQUIRED,
      message: 'API returned 401; cached token was rejected.',
      hint: 'Re-run the same command.',
    });
  }
  if (create.status >= 400) {
    throw new AppError({
      code: E.HTTP,
      message: `createReply/Forward returned HTTP ${create.status}.`,
      hint:
        typeof create.body === 'object'
          ? JSON.stringify(create.body)
          : String(create.body).slice(0, 500),
    });
  }
  const draft = create.body;

  if (overridesJson?.trim()) {
    let overrides;
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
    const patch = await call(
      auth,
      `/messages/${encodeURIComponent(draft.Id)}`,
      {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(overrides),
      },
    );
    if (patch.status >= 400) {
      throw new AppError({
        code: E.HTTP,
        message: `PATCH on draft returned HTTP ${patch.status}.`,
        hint:
          typeof patch.body === 'object'
            ? JSON.stringify(patch.body)
            : String(patch.body).slice(0, 500),
      });
    }
  }

  return {
    DraftId: draft.Id,
    WebLink: draft.WebLink ?? null,
    ConversationId: draft.ConversationId ?? null,
    message: 'Draft saved to your Drafts folder. Open Outlook to review and send.',
  };
}

program
  .command('draft')
  .argument('[json]', 'Outlook Message JSON; reads STDIN if omitted')
  .description(
    'Create a new draft (does NOT send). Same JSON shape as `send`. The\n' +
      'draft appears in your Drafts folder for review.',
  )
  .action(async (jsonArg) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    if (!raw.trim()) {
      throw new AppError({
        code: E.ARGS,
        message: 'No message JSON provided.',
        hint: 'Pass JSON as an argument or pipe it via STDIN.',
      });
    }
    let message;
    try {
      message = JSON.parse(raw);
    } catch (cause) {
      throw new AppError({
        code: E.ARGS,
        message: 'Message payload was not valid JSON.',
        hint: 'Validate with `jq .` first, then retry.',
        cause,
      });
    }
    const body = await runApi('/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(message),
    });
    printJson({
      DraftId: body.Id,
      WebLink: body.WebLink ?? null,
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
  .description(
    'Create a draft reply to a message. Outlook automatically fills in the\n' +
      'quoted thread; your override JSON sets the new body or extra recipients.',
  )
  .action(async (id, jsonArg) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    printJson(await makeDraftFrom(id, 'createReply', raw));
  });

program
  .command('draft-reply-all')
  .argument('<id>', 'message Id to reply-all to')
  .argument('[json]', 'partial Message override; STDIN if omitted')
  .description('Like `draft-reply`, but addresses everyone on the original thread.')
  .action(async (id, jsonArg) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    printJson(await makeDraftFrom(id, 'createReplyAll', raw));
  });

program
  .command('draft-forward')
  .argument('<id>', 'message Id to forward')
  .argument('[json]', 'partial Message override (typically ToRecipients + Body); STDIN if omitted')
  .description('Create a draft forward of a message.')
  .action(async (id, jsonArg) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    printJson(await makeDraftFrom(id, 'createForward', raw));
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
  .description(
    'Send a message.\n' +
      'Example STDIN payload:\n' +
      '  { "Subject": "hi", "Body": {"ContentType": "Text", "Content": "…"},\n' +
      '    "ToRecipients": [{"EmailAddress": {"Address": "x@y.com"}}] }',
  )
  .action(async (jsonArg) => {
    const raw = jsonArg ?? (process.stdin.isTTY ? '' : await readStdin());
    if (!raw.trim()) {
      throw new AppError({
        code: E.ARGS,
        message: 'No message JSON provided.',
        hint: 'Pass JSON as an argument or pipe it via STDIN (see `outlook send --help`).',
      });
    }
    let message;
    try {
      message = JSON.parse(raw);
    } catch (cause) {
      throw new AppError({
        code: E.ARGS,
        message: 'Message payload was not valid JSON.',
        hint: 'Validate with `jq .` first, then retry.',
        cause,
      });
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
  if (e instanceof AppError) {
    errorBlock(e.code, e.message, e.hint);
  } else {
    errorBlock(E.UNEXPECTED, e.message ?? String(e));
    if (process.env.OUTLOOK_DEBUG) process.stderr.write((e.stack ?? '') + '\n');
  }
  process.exit(exitCodeFor(e));
}
