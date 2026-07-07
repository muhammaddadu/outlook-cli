// Teams command tests. All Teams verbs route to the Graph base with the
// graph-audience token; these drive the CLI against a mock Graph server.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startMockServer } from './helpers.mjs';

const NO_OUTLOOK = '/nonexistent/outlook-auth.json';

/** Common env: no outlook token, a live graph token, graph base → mock. */
function graphEnv(mockUrl) {
  return {
    OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
    OUTLOOK_TOKEN_CACHE_GRAPH: seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } }),
    OUTLOOK_GRAPH_BASE: `${mockUrl}/v1.0`,
  };
}

async function withMock(handler, run) {
  const mock = await startMockServer(handler);
  try {
    return await run(mock);
  } finally {
    await mock.close();
  }
}

test('teams lists joined teams from Graph', async () => {
  await withMock(
    (req, res) => {
      assert.match(req.url, /^\/v1\.0\/me\/joinedTeams/);
      assert.ok(req.headers.authorization?.startsWith('Bearer '));
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ value: [{ id: 't1', displayName: 'Eng' }] }));
    },
    async (mock) => {
      const { code, stdout } = await runCli(['teams'], { env: graphEnv(mock.url) });
      assert.equal(code, 0);
      assert.equal(JSON.parse(stdout).value[0].displayName, 'Eng');
    },
  );
});

test('teams-channels encodes the team id and calls /teams/{id}/channels', async () => {
  let url = null;
  await withMock(
    (req, res) => {
      url = req.url;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ value: [] }));
    },
    async (mock) => {
      const { code } = await runCli(['teams-channels', 'team/1+2'], { env: graphEnv(mock.url) });
      assert.equal(code, 0);
      assert.match(url, /^\/v1\.0\/teams\/team%2F1%2B2\/channels/);
    },
  );
});

test('teams-chats honours --top', async () => {
  let url = null;
  await withMock(
    (req, res) => {
      url = req.url;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ value: [] }));
    },
    async (mock) => {
      const { code } = await runCli(['teams-chats', '-n', '5'], { env: graphEnv(mock.url) });
      assert.equal(code, 0);
      assert.match(url, /\/me\/chats\?\$top=5/);
    },
  );
});

test('teams-members calls /chats/{id}/members', async () => {
  let url = null;
  await withMock(
    (req, res) => {
      url = req.url;
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ value: [] }));
    },
    async (mock) => {
      const { code } = await runCli(['teams-members', '19:abc@thread.v2'], { env: graphEnv(mock.url) });
      assert.equal(code, 0);
      assert.equal(url, '/v1.0/chats/19%3Aabc%40thread.v2/members');
    },
  );
});

test('teams-send POSTs the message body and reports sent', async () => {
  let captured = null;
  await withMock(
    (req, res) => {
      let body = '';
      req.on('data', (d) => (body += d));
      req.on('end', () => {
        captured = { method: req.method, url: req.url, body };
        res.writeHead(201, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ id: 'msg-9' }));
      });
    },
    async (mock) => {
      const { code, stdout } = await runCli(
        ['teams-send', '19:abc@thread.v2', 'ship it'],
        { env: graphEnv(mock.url) },
      );
      assert.equal(code, 0);
      assert.equal(captured.method, 'POST');
      assert.equal(captured.url, '/v1.0/chats/19%3Aabc%40thread.v2/messages');
      const sent = JSON.parse(captured.body);
      assert.equal(sent.body.content, 'ship it');
      assert.equal(sent.body.contentType, 'text');
      const out = JSON.parse(stdout);
      assert.equal(out.sent, true);
      assert.equal(out.MessageId, 'msg-9');
    },
  );
});

test('teams-send --html sets contentType html', async () => {
  let body = null;
  await withMock(
    (req, res) => {
      let b = '';
      req.on('data', (d) => (b += d));
      req.on('end', () => {
        body = b;
        res.writeHead(201, { 'Content-Type': 'application/json' });
        res.end('{"id":"m"}');
      });
    },
    async (mock) => {
      const { code } = await runCli(
        ['teams-send', 'C1', '<b>hi</b>', '--html'],
        { env: graphEnv(mock.url) },
      );
      assert.equal(code, 0);
      assert.equal(JSON.parse(body).body.contentType, 'html');
    },
  );
});

test('teams-send with no text exits 64', async () => {
  const { code, stderr } = await runCli(['teams-send', 'C1'], {
    env: { OUTLOOK_TOKEN_CACHE: NO_OUTLOOK, OUTLOOK_TOKEN_CACHE_GRAPH: seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } }) },
  });
  assert.equal(code, 64);
  assert.match(stderr, /E_ARGS/);
  assert.match(stderr, /No message text/);
});

test('teams-messages surfaces the Graph 403 scope hint verbatim', async () => {
  await withMock(
    (req, res) => {
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: { code: 'Forbidden', message: "requires one of 'Chat.Read, Chat.ReadWrite'" } }));
    },
    async (mock) => {
      const { code, stderr } = await runCli(['teams-messages', 'C1'], { env: graphEnv(mock.url) });
      assert.equal(code, 3);
      assert.match(stderr, /E_HTTP/);
      assert.match(stderr, /HTTP 403/);
      assert.match(stderr, /Chat\.Read/);
    },
  );
});

test('teams without a graph token exits 2 toward `auth --all`', async () => {
  const { code, stderr } = await runCli(['teams'], {
    env: { OUTLOOK_TOKEN_CACHE: NO_OUTLOOK, OUTLOOK_TOKEN_CACHE_GRAPH: '/nonexistent/graph.json' },
  });
  assert.equal(code, 2);
  assert.match(stderr, /E_AUTH_REQUIRED/);
  assert.match(stderr, /auth --all/);
});
