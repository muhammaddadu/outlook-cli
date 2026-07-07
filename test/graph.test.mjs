// Cross-resource behaviour: the `graph` passthrough routes to the Graph
// base with the Graph-audience token, `token-audit` classifies cached
// tokens offline, and multi-resource capture persists one file per resource.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { existsSync, readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import {
  runCli,
  seedTokenCache,
  startMockServer,
  makeFakeJwt,
} from './helpers.mjs';

// A cache path guaranteed not to resolve to the developer's real token, so
// audits/graph calls in tests never touch ~/.cache/outlook-spike.
const NO_OUTLOOK = '/nonexistent/outlook-auth.json';

// ---------------------------------------------------------------------------
// graph passthrough

test('graph GET routes to the Graph base with the Graph token', async () => {
  const graphCache = seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } });
  let observed = null;
  const mock = await startMockServer((req, res) => {
    observed = { url: req.url, method: req.method, auth: req.headers.authorization };
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ displayName: 'Test User' }));
  });
  try {
    const { code, stdout } = await runCli(['graph', '/me'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: graphCache,
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observed.method, 'GET');
    assert.equal(observed.url, '/v1.0/me');
    assert.ok(observed.auth?.startsWith('Bearer '));
    assert.equal(JSON.parse(stdout).displayName, 'Test User');
  } finally {
    await mock.close();
  }
});

test('graph normalises a path that omits the leading slash', async () => {
  const graphCache = seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } });
  let observedUrl = null;
  const mock = await startMockServer((req, res) => {
    observedUrl = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ value: [] }));
  });
  try {
    const { code } = await runCli(['graph', 'me/chats?$top=5'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: graphCache,
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
      },
    });
    assert.equal(code, 0);
    assert.equal(observedUrl, '/v1.0/me/chats?$top=5');
  } finally {
    await mock.close();
  }
});

test('graph POST sends the JSON body with a Content-Type', async () => {
  const graphCache = seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } });
  let captured = null;
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      captured = { method: req.method, ct: req.headers['content-type'], body };
      res.writeHead(201, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ id: 'msg-1' }));
    });
  });
  try {
    const payload = JSON.stringify({ body: { content: 'hi team' } });
    const { code } = await runCli(
      ['graph', '/chats/abc/messages', payload, '-X', 'POST'],
      {
        env: {
          OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
          OUTLOOK_TOKEN_CACHE_GRAPH: graphCache,
          OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
        },
      },
    );
    assert.equal(code, 0);
    assert.equal(captured.method, 'POST');
    assert.match(captured.ct, /application\/json/);
    assert.equal(JSON.parse(captured.body).body.content, 'hi team');
  } finally {
    await mock.close();
  }
});

test('graph without a cached Graph token exits 2 and points at `auth --all`', async () => {
  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    res.writeHead(200);
    res.end('{}');
  });
  try {
    const { code, stderr } = await runCli(['graph', '/me'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: '/nonexistent/graph-auth.json',
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
        // OUTLOOK_NO_CAPTURE=1 is runCli's default — no browser fallback.
      },
    });
    assert.equal(code, 2);
    assert.equal(hits, 0, 'must not call Graph without a token');
    assert.match(stderr, /E_AUTH_REQUIRED/);
    assert.match(stderr, /auth --all/);
  } finally {
    await mock.close();
  }
});

test('graph with an unknown --resource exits 64', async () => {
  const { code, stderr } = await runCli(['graph', '/me', '--resource', 'teams'], {
    env: { OUTLOOK_TOKEN_CACHE: NO_OUTLOOK },
  });
  assert.equal(code, 64);
  assert.match(stderr, /E_ARGS/);
  assert.match(stderr, /Unknown --resource/);
});

test('graph does not auto-launch the browser on a rejected token (401)', async () => {
  const graphCache = seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } });
  let hits = 0;
  const mock = await startMockServer((req, res) => {
    hits++;
    res.writeHead(401);
    res.end('{"error":{"code":"InvalidAuthenticationToken","message":"expired"}}');
  });
  try {
    const { code, stderr } = await runCli(['graph', '/me'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: graphCache,
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
      },
    });
    // First 401 → clear graph cache → getAuth(graph) can't recapture
    // (capture disabled) → E_AUTH_REQUIRED, only one server hit.
    assert.equal(code, 2);
    assert.equal(hits, 1);
    assert.match(stderr, /E_AUTH_REQUIRED/);
  } finally {
    await mock.close();
  }
});

// ---------------------------------------------------------------------------
// token-audit

test('token-audit classifies a live token and lists it as reachable', async () => {
  const graphCache = seedTokenCache({
    claims: {
      aud: 'https://graph.microsoft.com',
      appid: 'app-123',
      app_displayname: 'Test App',
      upn: 'user@test',
      scp: 'Chat.ReadWrite Team.ReadBasic.All Files.ReadWrite.All User.Read',
    },
  });
  const { code, stdout } = await runCli(['token-audit'], {
    env: {
      OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
      OUTLOOK_TOKEN_CACHE_GRAPH: graphCache,
      OUTLOOK_TOKEN_CACHE_SUBSTRATE: '/nonexistent/substrate.json',
    },
  });
  assert.equal(code, 0);
  const out = JSON.parse(stdout);
  assert.deepEqual(out.reachable, ['graph']);

  const outlook = out.resources.find((r) => r.resource === 'outlook');
  assert.equal(outlook.status, 'absent');
  const graph = out.resources.find((r) => r.resource === 'graph');
  assert.equal(graph.status, 'live');
  assert.equal(graph.audience, 'https://graph.microsoft.com');
  assert.equal(graph.user, 'user@test');
  assert.ok(graph.capabilities['teams-chat'].includes('Chat.ReadWrite'));
  assert.ok(graph.capabilities.files.includes('Files.ReadWrite.All'));
});

test('token-audit reports an expired token as expired, not reachable', async () => {
  const expired = seedTokenCache({
    secondsFromNow: -3600,
    claims: { aud: 'https://graph.microsoft.com', scp: 'User.Read' },
  });
  const { code, stdout } = await runCli(['token-audit'], {
    env: {
      OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
      OUTLOOK_TOKEN_CACHE_GRAPH: expired,
    },
  });
  assert.equal(code, 0);
  const out = JSON.parse(stdout);
  assert.deepEqual(out.reachable, []);
  const graph = out.resources.find((r) => r.resource === 'graph');
  assert.equal(graph.status, 'expired');
  assert.ok(graph.minutesUntilExpiry < 0);
});

// ---------------------------------------------------------------------------
// multi-resource capture persistence (via the fixture seam)

test('auth --all persists one cache file per captured resource', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-multi-'));
  const outlookCache = join(dir, 'auth.json');
  const graphCache = join(dir, 'auth-graph.json');

  // Fixture the capture would have produced: a resource->headers map.
  const fixture = join(dir, 'capture.json');
  writeFileSync(
    fixture,
    JSON.stringify({
      outlook: {
        authorization: `Bearer ${makeFakeJwt({ expSeconds: Math.floor(Date.now() / 1000) + 3600, claims: { aud: 'https://outlook.office.com' } })}`,
        'x-anchormailbox': 'PUID:me@test',
      },
      graph: {
        authorization: `Bearer ${makeFakeJwt({ expSeconds: Math.floor(Date.now() / 1000) + 3600, claims: { aud: 'https://graph.microsoft.com' } })}`,
      },
    }),
  );

  const { code } = await runCli(['auth', '--all'], {
    env: {
      OUTLOOK_TOKEN_CACHE: outlookCache,
      OUTLOOK_TOKEN_CACHE_GRAPH: graphCache,
      OUTLOOK_CAPTURE_FIXTURE: fixture,
      // Clear the guard so the fixture seam (not a real browser) is used.
      OUTLOOK_NO_CAPTURE: '',
    },
  });
  assert.equal(code, 0);
  assert.ok(existsSync(outlookCache), 'outlook token persisted');
  assert.ok(existsSync(graphCache), 'graph token persisted');
  assert.equal(
    JSON.parse(readFileSync(graphCache, 'utf8')).headers.authorization.startsWith('Bearer '),
    true,
  );
});
