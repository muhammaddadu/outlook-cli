// Stability & recovery behaviour: automatic 401 re-capture, transient
// retry (429/503 + Retry-After), network-error mapping, and input
// validation that used to leak to the server as cryptic 400s.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { existsSync, readFileSync } from 'node:fs';

import {
  runCli,
  seedTokenCache,
  seedCaptureFixture,
  startMockServer,
} from './helpers.mjs';

// ---------------------------------------------------------------------------
// 401 auto-recovery

test('401 triggers a fresh capture and the request is retried once', async () => {
  const cache = seedTokenCache({ claims: { upn: 'stale@test' } });
  const fixture = seedCaptureFixture({ claims: { upn: 'fresh@test' } });

  const seenAuth = [];
  const mock = await startMockServer((req, res) => {
    seenAuth.push(req.headers.authorization);
    if (seenAuth.length === 1) {
      res.writeHead(401);
      res.end('{"error":"token expired server-side"}');
      return;
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [{ Id: 'recovered' }] }));
  });

  try {
    const { code, stdout } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: cache,
        OUTLOOK_CAPTURE_FIXTURE: fixture,
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0, 'expected the retried request to succeed');
    assert.equal(seenAuth.length, 2, 'expected exactly one retry');
    assert.notEqual(seenAuth[0], seenAuth[1], 'retry must use the freshly captured token');
    assert.deepEqual(JSON.parse(stdout), { value: [{ Id: 'recovered' }] });
    // The fresh token must have been persisted for the next invocation.
    const persisted = JSON.parse(readFileSync(cache, 'utf8'));
    assert.equal(`${persisted.headers.authorization}`, seenAuth[1]);
  } finally {
    await mock.close();
  }
});

test('persistent 401 after recapture exits 2 and clears the cache', async () => {
  const cache = seedTokenCache();
  const fixture = seedCaptureFixture();

  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    res.writeHead(401);
    res.end('{"error":"still unauthorized"}');
  });

  try {
    const { code, stderr } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: cache,
        OUTLOOK_CAPTURE_FIXTURE: fixture,
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 2);
    assert.equal(hits, 2, 'should give up after one recapture attempt');
    assert.match(stderr, /E_AUTH_REQUIRED/);
    assert.match(stderr, /outlook auth/);
    assert.equal(existsSync(cache), false);
  } finally {
    await mock.close();
  }
});

test('401 with capture disabled surfaces E_AUTH_REQUIRED without looping', async () => {
  const cache = seedTokenCache();
  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    res.writeHead(401);
    res.end('{}');
  });

  try {
    const { code, stderr } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: cache,
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
        // OUTLOOK_NO_CAPTURE=1 is runCli's default — recapture must fail fast.
      },
    });
    assert.equal(code, 2);
    assert.equal(hits, 1);
    assert.match(stderr, /E_AUTH_REQUIRED/);
    assert.equal(existsSync(cache), false);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// Transient retry

test('429 with Retry-After is retried and succeeds', async () => {
  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    if (hits === 1) {
      res.writeHead(429, { 'Retry-After': '0' });
      res.end('{"error":{"code":"TooManyRequests","message":"throttled"}}');
      return;
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });

  try {
    const { code, stderr } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.equal(hits, 2);
    assert.equal(stderr, '');
  } finally {
    await mock.close();
  }
});

test('429 on a POST is also retried (server did not process it)', async () => {
  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    if (hits === 1) {
      res.writeHead(429, { 'Retry-After': '0' });
      res.end('{}');
      return;
    }
    res.writeHead(202);
    res.end();
  });

  try {
    const { code } = await runCli(['send'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
      stdin: JSON.stringify({ Subject: 'x', Body: { ContentType: 'Text', Content: 'y' } }),
    });
    assert.equal(code, 0);
    assert.equal(hits, 2);
  } finally {
    await mock.close();
  }
});

test('persistent 429 gives up with E_HTTP and a parsed OData message', async () => {
  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    res.writeHead(429, { 'Retry-After': '0' });
    res.end('{"error":{"code":"TooManyRequests","message":"slow down"}}');
  });

  try {
    const { code, stderr } = await runCli(['list'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 3);
    assert.equal(hits, 3, 'three attempts total');
    assert.match(stderr, /E_HTTP/);
    assert.match(stderr, /HTTP 429/);
    assert.match(stderr, /TooManyRequests: slow down/);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// Network errors

test('unreachable API maps to E_NETWORK with exit 3, not a stack trace', async () => {
  // Grab a port that is guaranteed closed: bind, note it, release it.
  const probe = await startMockServer(() => {});
  const deadUrl = probe.url;
  await probe.close();

  const { code, stderr } = await runCli(['list'], {
    env: {
      OUTLOOK_TOKEN_CACHE: seedTokenCache(),
      OUTLOOK_API_BASE: `${deadUrl}/api/v2.0/me`,
    },
  });
  assert.equal(code, 3);
  assert.match(stderr, /E_NETWORK/);
  assert.match(stderr, /Could not reach 127\.0\.0\.1/);
  assert.match(stderr, /Hint:/);
  assert.doesNotMatch(stderr, /E_UNEXPECTED/);
});

// ---------------------------------------------------------------------------
// Input validation that used to reach the server

test('list with non-numeric --top exits 64 before any request', async () => {
  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end('{}');
  });
  try {
    const { code, stderr } = await runCli(['list', '-n', 'lots'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 64);
    assert.equal(hits, 0);
    assert.match(stderr, /E_ARGS/);
    assert.match(stderr, /--top/);
  } finally {
    await mock.close();
  }
});

test('free-busy with non-numeric --interval exits 64', async () => {
  const { code, stderr } = await runCli(
    ['free-busy', 'a@b.com', '--interval', 'often'],
    { env: { OUTLOOK_TOKEN_CACHE: seedTokenCache() } },
  );
  assert.equal(code, 64);
  assert.match(stderr, /E_ARGS/);
  assert.match(stderr, /--interval/);
});

test('unknown option exits 64 (usage), matching E_ARGS semantics', async () => {
  const { code, stderr } = await runCli(['list', '--no-such-flag'], {
    env: { OUTLOOK_TOKEN_CACHE: seedTokenCache() },
  });
  assert.equal(code, 64);
  assert.match(stderr, /unknown option/i);
});

// ---------------------------------------------------------------------------
// Documented-but-broken surface (regression)

test('unread accepts the full list filter set (README example)', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['unread', '--from', 'boss@example.com'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    const decoded = decodeURIComponent(observedUrl);
    assert.match(decoded, /IsRead eq false/);
    assert.match(decoded, /From\/EmailAddress\/Address eq 'boss@example\.com'/);
    assert.match(observedUrl, /\$top=25/); // unread keeps its larger default page
  } finally {
    await mock.close();
  }
});

test('unread respects an explicit --top over its default', async () => {
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['unread', '-n', '10'], {
      env: {
        OUTLOOK_TOKEN_CACHE: seedTokenCache(),
        OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
      },
    });
    assert.equal(code, 0);
    assert.match(observedUrl, /\$top=10/);
  } finally {
    await mock.close();
  }
});
