// Integration tests for the CLI binary.
//
// Strategy: spawn `src/cli.mjs` as a subprocess with a pre-seeded fake
// token cache (so the Playwright capture path is never triggered) and an
// OUTLOOK_API_BASE pointing at a local mock HTTP server. This exercises
// the full real CLI surface end-to-end without touching the internet.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startMockServer, freshLearningsPath } from './helpers.mjs';

// ---------------------------------------------------------------------------
// Smoke tests — no auth or network involved

test('--version prints semver to stdout, exits 0', async () => {
  const { code, stdout, stderr } = await runCli(['--version']);
  assert.equal(code, 0);
  assert.match(stdout.trim(), /^\d+\.\d+\.\d+$/);
  assert.equal(stderr, '');
});

test('--help lists all commands', async () => {
  const { code, stdout } = await runCli(['--help']);
  assert.equal(code, 0);
  for (const cmd of ['auth', 'refresh', 'logout', 'list', 'search', 'read', 'folders', 'send']) {
    assert.match(stdout, new RegExp(`^\\s*${cmd}\\b`, 'm'), `expected to see "${cmd}" in help`);
  }
});

test('unknown subcommand exits non-zero and prints help to stderr', async () => {
  const { code, stderr } = await runCli(['no-such-command']);
  assert.notEqual(code, 0);
  assert.match(stderr, /unknown command|error/i);
});

// ---------------------------------------------------------------------------
// Argument validation — exercised without a network

test('send with no argument and no stdin exits 64 (E_ARGS)', async () => {
  const { code, stderr } = await runCli(['send'], {
    env: { OUTLOOK_TOKEN_CACHE: seedTokenCache() },
  });
  assert.equal(code, 64);
  assert.match(stderr, /E_ARGS/);
  assert.match(stderr, /No message JSON provided/);
});

test('send with malformed JSON exits 64 (E_ARGS)', async () => {
  const { code, stderr } = await runCli(['send', '{not-json'], {
    env: { OUTLOOK_TOKEN_CACHE: seedTokenCache() },
  });
  assert.equal(code, 64);
  assert.match(stderr, /E_ARGS/);
  assert.match(stderr, /not valid JSON/);
});

// ---------------------------------------------------------------------------
// Happy-path API calls against a mock server

test('list returns the mock server JSON on stdout, nothing on stderr', async () => {
  const cache = seedTokenCache();
  const messages = {
    '@odata.context': 'mock',
    value: [{ Id: 'abc', Subject: 'hello', From: { EmailAddress: { Name: 'A' } } }],
  };
  const mock = await startMockServer((req, res) => {
    assert.match(req.url, /messages\?\$top=10/);
    assert.equal(req.headers.authorization?.startsWith('Bearer '), true);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(messages));
  });

  try {
    const { code, stdout, stderr } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: cache,
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(stderr, '');
    assert.deepEqual(JSON.parse(stdout), messages);
  } finally {
    await mock.close();
  }
});

test('search url-encodes the query and respects --top', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });

  try {
    const { code } = await runCli(['search', 'quarterly review', '-n', '3'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.match(observedUrl, /\$search=%22quarterly%20review%22/);
    assert.match(observedUrl, /\$top=3/);
  } finally {
    await mock.close();
  }
});

test('folders calls /mailfolders', async () => {
  let observedPath = null;
  const mock = await startMockServer((req, res) => {
    observedPath = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [{ DisplayName: 'Inbox' }] }));
  });

  try {
    const { code, stdout } = await runCli(['folders'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedPath, '/api/v2.0/me/mailfolders');
    assert.equal(JSON.parse(stdout).value[0].DisplayName, 'Inbox');
  } finally {
    await mock.close();
  }
});

test('send POSTs to /sendmail and reads JSON from stdin', async () => {
  let captured = null;
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      captured = { method: req.method, url: req.url, body };
      res.writeHead(202);
      res.end();
    });
  });

  try {
    const payload = JSON.stringify({
      Subject: 'hi',
      Body: { ContentType: 'Text', Content: 'from a test' },
      ToRecipients: [{ EmailAddress: { Address: 'me@example.com' } }],
    });
    const { code, stdout } = await runCli(['send'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
      stdin: payload,
    });
    assert.equal(code, 0);
    assert.equal(captured.method, 'POST');
    assert.equal(captured.url, '/api/v2.0/me/sendmail');
    const sent = JSON.parse(captured.body);
    assert.equal(sent.Message.Subject, 'hi');
    assert.equal(sent.SaveToSentItems, true);
    // 202 with empty body should print { sent: true }
    assert.deepEqual(JSON.parse(stdout), { sent: true });
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// Error paths from the API

test('401 from API clears cache and exits 2 with E_AUTH_REQUIRED', async () => {
  const { existsSync } = await import('node:fs');
  const cache = seedTokenCache();
  assert.ok(existsSync(cache));

  const mock = await startMockServer((req, res) => {
    res.writeHead(401);
    res.end('{"error":"unauthorized"}');
  });

  try {
    const { code, stderr } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: cache,
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 2);
    assert.match(stderr, /E_AUTH_REQUIRED/);
    // Cache should have been deleted as part of the 401 handler.
    assert.equal(existsSync(cache), false);
  } finally {
    await mock.close();
  }
});

test('500 from API exits 3 with E_HTTP', async () => {
  const mock = await startMockServer((req, res) => {
    res.writeHead(500);
    res.end('{"error":"boom"}');
  });

  try {
    const { code, stderr } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 3);
    assert.match(stderr, /E_HTTP/);
    assert.match(stderr, /HTTP 500/);
  } finally {
    await mock.close();
  }
});

test('read encodes message id (so slashes and "+" survive)', async () => {
  let observedPath = null;
  const mock = await startMockServer((req, res) => {
    observedPath = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ Id: 'x' }));
  });

  try {
    const { code } = await runCli(['read', 'AAMk/abc+def='], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedPath, '/api/v2.0/me/messages/AAMk%2Fabc%2Bdef%3D');
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// Output discipline

test('stdout is parseable JSON when piped (compact mode in non-TTY)', async () => {
  const mock = await startMockServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [{ Id: '1' }] }));
  });

  try {
    const { stdout } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    // Compact: should fit on a single line (no pretty-print newlines inside).
    const lines = stdout.trimEnd().split('\n');
    assert.equal(lines.length, 1, 'expected compact (single-line) JSON when stdout is a pipe');
    assert.deepEqual(JSON.parse(lines[0]), { value: [{ Id: '1' }] });
  } finally {
    await mock.close();
  }
});

test('--debug puts diagnostics on stderr, leaves stdout clean', async () => {
  const mock = await startMockServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true }));
  });

  try {
    const { stdout, stderr } = await runCli(['--debug', 'list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.deepEqual(JSON.parse(stdout), { ok: true });
    assert.match(stderr, /\[debug\] fetch GET/);
    assert.match(stderr, /\[debug\] using cached token/);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// Filter flags

test('list --unread emits IsRead eq false filter', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['list', '--unread'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.match(decodeURIComponent(observedUrl), /\$filter=IsRead eq false/);
  } finally {
    await mock.close();
  }
});

test('list combines --unread --from --since into one filter', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(
      ['list', '--unread', '--from', 'alice@example.com', '--since', '7d'],
      {
        env: {
          OUTLOOK_TOKEN_CACHE: seedTokenCache(),
          OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
        },
      },
    );
    assert.equal(code, 0);
    const filter = decodeURIComponent(observedUrl);
    assert.match(filter, /IsRead eq false/);
    assert.match(filter, /From\/EmailAddress\/Address eq 'alice@example\.com'/);
    assert.match(filter, /ReceivedDateTime gt \d{4}-\d{2}-\d{2}T/);
    assert.match(filter, / and /); // multiple clauses joined
  } finally {
    await mock.close();
  }
});

test('list --folder Sent routes to /mailfolders/SentItems/messages', async () => {
  let observedPath = null;
  const mock = await startMockServer((req, res) => {
    observedPath = req.url.split('?')[0];
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['list', '--folder', 'Sent'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedPath, '/api/v2.0/me/mailfolders/SentItems/messages');
  } finally {
    await mock.close();
  }
});

test('list --folder with unknown short name exits E_ARGS', async () => {
  const { code, stderr } = await runCli(['list', '--folder', 'Bogus'], {
    env: { OUTLOOK_TOKEN_CACHE: seedTokenCache() },
  });
  assert.equal(code, 64);
  assert.match(stderr, /E_ARGS/);
  assert.match(stderr, /Unknown folder/);
});

test('list --filter raw expression is preserved and parenthesised', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(
      ['list', '--filter', "Categories/any(c: c eq 'red')"],
      {
        env: {
          OUTLOOK_TOKEN_CACHE: seedTokenCache(),
          OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
        },
      },
    );
    assert.equal(code, 0);
    assert.match(decodeURIComponent(observedUrl), /\(Categories\/any\(c: c eq 'red'\)\)/);
  } finally {
    await mock.close();
  }
});

test('list --skip and --top set pagination params', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['list', '--top', '50', '--skip', '100'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.match(observedUrl, /\$top=50/);
    assert.match(observedUrl, /\$skip=100/);
  } finally {
    await mock.close();
  }
});

test('unread shortcut is equivalent to list --unread', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['unread'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.match(decodeURIComponent(observedUrl), /\$filter=IsRead eq false/);
    assert.match(observedUrl, /\$top=25/); // default for unread
  } finally {
    await mock.close();
  }
});

test('search drops $orderby when combined with --order-by (server rule)', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(
      ['search', 'deploy', '--order-by', 'ReceivedDateTime desc'],
      {
        env: {
          OUTLOOK_TOKEN_CACHE: seedTokenCache(),
          OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
        },
      },
    );
    assert.equal(code, 0);
    assert.match(observedUrl, /\$search=/);
    assert.doesNotMatch(observedUrl, /\$orderby/);
  } finally {
    await mock.close();
  }
});

test('list --select overrides the default field list', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['list', '--select', 'Id,Subject'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.match(observedUrl, /\$select=Id%2CSubject/);
    assert.doesNotMatch(observedUrl, /BodyPreview/);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// Drafts

test('draft posts to /messages with the user-supplied JSON', async () => {
  const captured = { method: null, url: null, body: null };
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      captured.method = req.method;
      captured.url = req.url;
      captured.body = body;
      res.writeHead(201, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ Id: 'DRAFT123', WebLink: 'https://outlook.office.com/owa/?ItemID=DRAFT123' }));
    });
  });
  try {
    const payload = JSON.stringify({
      Subject: 'draft me',
      Body: { ContentType: 'Text', Content: 'hello' },
      ToRecipients: [{ EmailAddress: { Address: 'a@b.com' } }],
    });
    const { code, stdout } = await runCli(['draft'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
      stdin: payload,
    });
    assert.equal(code, 0);
    assert.equal(captured.method, 'POST');
    assert.equal(captured.url, '/api/v2.0/me/messages');
    assert.equal(JSON.parse(captured.body).Subject, 'draft me');

    const out = JSON.parse(stdout);
    assert.equal(out.DraftId, 'DRAFT123');
    assert.match(out.WebLink, /outlook\.office\.com/);
  } finally {
    await mock.close();
  }
});

test('draft-reply calls createReply and prints the new draft id + WebLink', async () => {
  const reqs = [];
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      reqs.push({ method: req.method, url: req.url, body });
      if (req.url.endsWith('/createReply')) {
        res.writeHead(201, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ Id: 'REPLY99', WebLink: 'https://outlook.office.com/owa/?ItemID=REPLY99' }));
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end('{}');
      }
    });
  });
  try {
    const { code, stdout } = await runCli(['draft-reply', 'ORIG-123'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(reqs.length, 1, 'no PATCH when no override JSON is provided');
    assert.equal(reqs[0].method, 'POST');
    assert.equal(reqs[0].url, '/api/v2.0/me/messages/ORIG-123/createReply');

    const out = JSON.parse(stdout);
    assert.equal(out.DraftId, 'REPLY99');
    assert.match(out.WebLink, /outlook\.office\.com/);
    assert.match(out.message, /Drafts folder/);
  } finally {
    await mock.close();
  }
});

test('draft-reply with body override sends Comment in createReply (preserves thread)', async () => {
  const reqs = [];
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      reqs.push({ method: req.method, url: req.url, body });
      if (req.url.endsWith('/createReply')) {
        res.writeHead(201, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ Id: 'REPLY-XYZ' }));
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end('{}');
      }
    });
  });
  try {
    const override = JSON.stringify({
      Body: { ContentType: 'Text', Content: 'thanks — looks great' },
    });
    const { code } = await runCli(['draft-reply', 'ORIG-1', override], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    // Only one call should fire — createReply with Comment. No PATCH, so
    // the server-populated quoted thread is preserved.
    assert.equal(reqs.length, 1, 'expected exactly one request (createReply)');
    assert.equal(reqs[0].method, 'POST');
    assert.equal(reqs[0].url, '/api/v2.0/me/messages/ORIG-1/createReply');
    assert.deepEqual(JSON.parse(reqs[0].body), { Comment: 'thanks — looks great' });
  } finally {
    await mock.close();
  }
});

test('draft-reply with non-body overrides still PATCHes those fields', async () => {
  const reqs = [];
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      reqs.push({ method: req.method, url: req.url, body });
      if (req.url.endsWith('/createReply')) {
        res.writeHead(201, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ Id: 'REPLY-XYZ' }));
      } else {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end('{}');
      }
    });
  });
  try {
    const override = JSON.stringify({
      Body: { ContentType: 'Text', Content: 'thanks' },
      CcRecipients: [{ EmailAddress: { Address: 'cc@example.com' } }],
    });
    const { code } = await runCli(['draft-reply', 'ORIG-1', override], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(reqs.length, 2);
    // First: createReply with Comment (body)
    assert.equal(reqs[0].method, 'POST');
    assert.equal(reqs[0].url, '/api/v2.0/me/messages/ORIG-1/createReply');
    assert.deepEqual(JSON.parse(reqs[0].body), { Comment: 'thanks' });
    // Second: PATCH only the non-body fields
    assert.equal(reqs[1].method, 'PATCH');
    assert.equal(reqs[1].url, '/api/v2.0/me/messages/REPLY-XYZ');
    const patchBody = JSON.parse(reqs[1].body);
    assert.equal(patchBody.Body, undefined, 'PATCH must not include Body');
    assert.deepEqual(patchBody.CcRecipients, [{ EmailAddress: { Address: 'cc@example.com' } }]);
  } finally {
    await mock.close();
  }
});

test('draft-reply-all hits createReplyAll', async () => {
  let observedPath = null;
  const mock = await startMockServer((req, res) => {
    observedPath = req.url;
    res.writeHead(201, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ Id: 'X' }));
  });
  try {
    const { code } = await runCli(['draft-reply-all', 'ORIG-1'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedPath, '/api/v2.0/me/messages/ORIG-1/createReplyAll');
  } finally {
    await mock.close();
  }
});

test('draft-forward hits createForward', async () => {
  let observedPath = null;
  const mock = await startMockServer((req, res) => {
    observedPath = req.url;
    res.writeHead(201, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ Id: 'X' }));
  });
  try {
    const { code } = await runCli(['draft-forward', 'ORIG-1'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedPath, '/api/v2.0/me/messages/ORIG-1/createForward');
  } finally {
    await mock.close();
  }
});

test('discard-draft DELETEs the message and prints a success payload', async () => {
  let observed = null;
  const mock = await startMockServer((req, res) => {
    observed = { method: req.method, url: req.url };
    res.writeHead(204);
    res.end();
  });
  try {
    const { code, stdout } = await runCli(['discard-draft', 'DRAFT-XYZ'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observed.method, 'DELETE');
    assert.equal(observed.url, '/api/v2.0/me/messages/DRAFT-XYZ');
    assert.deepEqual(JSON.parse(stdout), { discarded: true, DraftId: 'DRAFT-XYZ' });
  } finally {
    await mock.close();
  }
});

test('draft-reply with malformed override JSON exits 64', async () => {
  // Need a working mock so createReply succeeds before PATCH parse fails.
  const mock = await startMockServer((req, res) => {
    res.writeHead(201, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ Id: 'X' }));
  });
  try {
    const { code, stderr } = await runCli(['draft-reply', 'ORIG-1', '{not-json'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 64);
    assert.match(stderr, /E_ARGS/);
    assert.match(stderr, /not valid JSON/);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// Self-learning context

test('context with no cached auth returns null user and empty learnings', async () => {
  const learnings = freshLearningsPath();
  const { code, stdout } = await runCli(['context'], {
    env: {
      OUTLOOK_TOKEN_CACHE: '/nonexistent/auth.json',
      OUTLOOK_LEARNINGS: learnings,
    },
  });
  assert.equal(code, 0);
  const out = JSON.parse(stdout);
  assert.equal(out.user, null);
  assert.deepEqual(out.learnings, []);
  assert.equal(out.learningsFile, learnings);
});

test('context decodes user info from the cached JWT claims', async () => {
  const cache = seedTokenCache({
    claims: {
      upn: 'alice@example.com',
      name: 'Alice Example',
      tid: 'tenant-uuid',
    },
  });
  const { code, stdout } = await runCli(['context'], {
    env: { OUTLOOK_TOKEN_CACHE: cache, OUTLOOK_LEARNINGS: freshLearningsPath() },
  });
  assert.equal(code, 0);
  const out = JSON.parse(stdout);
  assert.deepEqual(out.user, {
    email: 'alice@example.com',
    name: 'Alice Example',
    tenant_id: 'tenant-uuid',
  });
  assert.ok(out.tokenMinutesUntilExpiry > 0);
});

test('learn add records an observation that learn (list) reads back', async () => {
  const learnings = freshLearningsPath();
  const env = { OUTLOOK_LEARNINGS: learnings };

  const add = await runCli(['learn', 'add', 'Signs off as Sam'], { env });
  assert.equal(add.code, 0);
  assert.equal(JSON.parse(add.stdout).added, true);

  const list = await runCli(['learn'], { env });
  assert.equal(list.code, 0);
  const out = JSON.parse(list.stdout);
  assert.equal(out.count, 1);
  assert.match(out.learnings[0], /Signs off as Sam$/);
});

test('learn forget removes matching learnings', async () => {
  const learnings = freshLearningsPath();
  const env = { OUTLOOK_LEARNINGS: learnings };
  await runCli(['learn', 'add', 'Signs off as Sam'], { env });
  await runCli(['learn', 'add', 'Prefers terse replies'], { env });
  await runCli(['learn', 'add', 'Boss is alice@example.com'], { env });

  const forget = await runCli(['learn', 'forget', 'alice'], { env });
  assert.equal(forget.code, 0);
  assert.equal(JSON.parse(forget.stdout).removed, 1);

  const list = await runCli(['learn'], { env });
  const out = JSON.parse(list.stdout);
  assert.equal(out.count, 2);
});

test('learn clear wipes everything', async () => {
  const learnings = freshLearningsPath();
  const env = { OUTLOOK_LEARNINGS: learnings };
  await runCli(['learn', 'add', 'x'], { env });
  await runCli(['learn', 'add', 'y'], { env });

  const clear = await runCli(['learn', 'clear'], { env });
  assert.equal(clear.code, 0);
  assert.equal(JSON.parse(clear.stdout).cleared, true);

  const list = await runCli(['learn'], { env });
  assert.equal(JSON.parse(list.stdout).count, 0);
});

test('learn add deduplicates near-identical observations', async () => {
  const learnings = freshLearningsPath();
  const env = { OUTLOOK_LEARNINGS: learnings };
  const first = await runCli(['learn', 'add', 'Likes coffee'], { env });
  assert.equal(JSON.parse(first.stdout).added, true);
  const second = await runCli(['learn', 'add', 'Likes coffee'], { env });
  assert.equal(JSON.parse(second.stdout).added, false);
  const list = await runCli(['learn'], { env });
  assert.equal(JSON.parse(list.stdout).count, 1);
});

// ---------------------------------------------------------------------------
// Calendar

test('agenda calls /calendarView with startDateTime/endDateTime', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [{ Subject: 'Standup' }] }));
  });
  try {
    const { code, stdout } = await runCli(['agenda', '--days', '3'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.match(observedUrl, /\/calendarView/);
    assert.match(observedUrl, /startDateTime=/);
    assert.match(observedUrl, /endDateTime=/);
    assert.match(observedUrl, /\$top=50/);
    assert.match(observedUrl, /\$orderby=Start%2FDateTime%20asc/);
    assert.equal(JSON.parse(stdout).value[0].Subject, 'Standup');
  } finally {
    await mock.close();
  }
});

test('agenda --calendar routes to /calendars/<id>/calendarView', async () => {
  let observedPath = null;
  const mock = await startMockServer((req, res) => {
    observedPath = req.url.split('?')[0];
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['agenda', '--calendar', 'CAL-123'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedPath, '/api/v2.0/me/calendars/CAL-123/calendarView');
  } finally {
    await mock.close();
  }
});

test('events hits /events without time-range expansion', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(
      ['events', '--organizer', 'boss@example.com', '--subject', 'review', '-n', '5'],
      {
        env: {
          OUTLOOK_TOKEN_CACHE: seedTokenCache(),
          OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
        },
      },
    );
    assert.equal(code, 0);
    const decoded = decodeURIComponent(observedUrl);
    assert.match(observedUrl, /\/events\?/);
    assert.match(observedUrl, /\$top=5/);
    assert.match(decoded, /Organizer\/EmailAddress\/Address eq 'boss@example\.com'/);
    assert.match(decoded, /contains\(Subject, 'review'\)/);
  } finally {
    await mock.close();
  }
});

test('event-create POSTs the JSON to /events', async () => {
  const captured = { method: null, url: null, body: null };
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      captured.method = req.method;
      captured.url = req.url;
      captured.body = body;
      res.writeHead(201, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ Id: 'NEW-EVT', Subject: JSON.parse(body).Subject }));
    });
  });
  try {
    const payload = JSON.stringify({
      Subject: '1:1',
      Start: { DateTime: '2026-06-01T10:00:00', TimeZone: 'UTC' },
      End: { DateTime: '2026-06-01T10:30:00', TimeZone: 'UTC' },
    });
    const { code, stdout } = await runCli(['event-create'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
      stdin: payload,
    });
    assert.equal(code, 0);
    assert.equal(captured.method, 'POST');
    assert.equal(captured.url, '/api/v2.0/me/events');
    assert.equal(JSON.parse(captured.body).Subject, '1:1');
    assert.equal(JSON.parse(stdout).Id, 'NEW-EVT');
  } finally {
    await mock.close();
  }
});

test('event-update PATCHes /events/<id>', async () => {
  const reqs = [];
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      reqs.push({ method: req.method, url: req.url, body });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ Id: 'EVT-1' }));
    });
  });
  try {
    const { code } = await runCli(
      ['event-update', 'EVT-1', '{"Subject":"Renamed"}'],
      {
        env: {
          OUTLOOK_TOKEN_CACHE: seedTokenCache(),
          OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
        },
      },
    );
    assert.equal(code, 0);
    assert.equal(reqs[0].method, 'PATCH');
    assert.equal(reqs[0].url, '/api/v2.0/me/events/EVT-1');
    assert.equal(JSON.parse(reqs[0].body).Subject, 'Renamed');
  } finally {
    await mock.close();
  }
});

test('event-cancel DELETEs and reports success', async () => {
  let observed = null;
  const mock = await startMockServer((req, res) => {
    observed = { method: req.method, url: req.url };
    res.writeHead(204);
    res.end();
  });
  try {
    const { code, stdout } = await runCli(['event-cancel', 'EVT-9'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observed.method, 'DELETE');
    assert.equal(observed.url, '/api/v2.0/me/events/EVT-9');
    assert.deepEqual(JSON.parse(stdout), { cancelled: true, EventId: 'EVT-9' });
  } finally {
    await mock.close();
  }
});

for (const [verb, expectedAction] of [
  ['accept', 'accept'],
  ['decline', 'decline'],
  ['tentative', 'tentativelyAccept'],
]) {
  test(`${verb} POSTs to /events/<id>/${expectedAction}`, async () => {
    let observed = null;
    const mock = await startMockServer((req, res) => {
      let body = '';
      req.on('data', (d) => (body += d));
      req.on('end', () => {
        observed = { method: req.method, url: req.url, body };
        res.writeHead(202);
        res.end();
      });
    });
    try {
      const { code, stdout } = await runCli(
        [verb, 'EVT-1', '-c', 'sounds good'],
        {
          env: {
            OUTLOOK_TOKEN_CACHE: seedTokenCache(),
            OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
          },
        },
      );
      assert.equal(code, 0);
      assert.equal(observed.method, 'POST');
      assert.equal(observed.url, `/api/v2.0/me/events/EVT-1/${expectedAction}`);
      const body = JSON.parse(observed.body);
      assert.equal(body.Comment, 'sounds good');
      assert.equal(body.SendResponse, true);
      const out = JSON.parse(stdout);
      assert.equal(out.rsvp, verb);
      assert.equal(out.EventId, 'EVT-1');
    } finally {
      await mock.close();
    }
  });
}

test('free-busy POSTs the right schedule shape to /getSchedule', async () => {
  const captured = { method: null, url: null, body: null };
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      captured.method = req.method;
      captured.url = req.url;
      captured.body = body;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ value: [] }));
    });
  });
  try {
    const { code } = await runCli(
      ['free-busy', 'a@b.com', 'c@d.com', '--interval', '60'],
      {
        env: {
          OUTLOOK_TOKEN_CACHE: seedTokenCache(),
          OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
        },
      },
    );
    assert.equal(code, 0);
    assert.equal(captured.method, 'POST');
    assert.equal(captured.url, '/api/v2.0/me/getSchedule');
    const body = JSON.parse(captured.body);
    assert.deepEqual(body.Schedules, ['a@b.com', 'c@d.com']);
    assert.equal(body.AvailabilityViewInterval, 60);
    assert.ok(body.StartTime.DateTime);
    assert.ok(body.EndTime.DateTime);
  } finally {
    await mock.close();
  }
});

test('logout removes the cache file and exits 0', async () => {
  const { existsSync } = await import('node:fs');
  const cache = seedTokenCache();
  assert.ok(existsSync(cache));

  const { code, stderr } = await runCli(['logout'], {
    env: { OUTLOOK_TOKEN_CACHE: cache },
  });
  assert.equal(code, 0);
  assert.equal(existsSync(cache), false);
  assert.match(stderr, /cache cleared/i);
});
