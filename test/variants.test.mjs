// Per-app token variants: the same resource (Graph) is issued to different
// web apps with different scopes. A 403 "missing scope" retries with the
// cached sibling that has the scope.
import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { mkdtempSync, writeFileSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { runCli, startMockServer, makeFakeJwt } from './helpers.mjs';
import { saveVariant, loadVariants, requiredScopesFrom, variantWithScopes, saveAuth } from '../src/auth.mjs';

const tok = (appid, scp, secs = 3600) =>
  `Bearer ${makeFakeJwt({ expSeconds: Math.floor(Date.now() / 1000) + secs, claims: { aud: 'https://graph.microsoft.com', appid, scp } })}`;

function isolatedGraphCache() {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-var-'));
  const file = join(dir, 'auth-graph.json');
  process.env.OUTLOOK_TOKEN_CACHE_GRAPH = file;
  return file;
}

test('requiredScopesFrom parses the Graph 403 message', () => {
  const body = { error: { code: 'Forbidden', message: "Missing scope permissions on the request. API requires one of 'ChannelMessage.Read.All, Group.Read.All'. Scopes on the request 'x'" } };
  assert.deepEqual(requiredScopesFrom(body), ['ChannelMessage.Read.All', 'Group.Read.All']);
  assert.deepEqual(requiredScopesFrom({ error: { message: 'nope' } }), []);
});

test('saveAuth keeps one variant per app id; later-expiring wins per app', () => {
  const file = isolatedGraphCache();
  try {
    saveAuth({ authorization: tok('teams', 'ChannelMessage.Read.All') }, 'graph');
    saveAuth({ authorization: tok('owa', 'Chat.Read', 86400) }, 'graph');
    assert.equal(saveVariant({ authorization: tok('teams', 'ChannelMessage.Read.All', 60 * 10) }, 'graph'), false);
    const files = readdirSync(dirname(file)).sort();
    assert.deepEqual(files, ['auth-graph.json', 'auth-graph@owa.json', 'auth-graph@teams.json']);
    const vs = loadVariants('graph');
    assert.equal(vs.length, 2);
    assert.ok(vs[0].scopes.has('Chat.Read')); // longest-lived first
    const ch = variantWithScopes('graph', ['ChannelMessage.Read.All'], vs[0].headers.authorization);
    assert.ok(ch.scopes.has('ChannelMessage.Read.All'));
  } finally {
    delete process.env.OUTLOOK_TOKEN_CACHE_GRAPH;
  }
});

test('graph call retries a missing-scope 403 with a sibling token that has the scope', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-var-'));
  const primary = join(dir, 'auth-graph.json');
  const owa = tok('owa', 'Chat.Read', 86400);
  const teams = tok('teams', 'ChannelMessage.Read.All');
  const write = (f, a) => writeFileSync(f, JSON.stringify({ headers: { authorization: a } }));
  write(primary, owa);
  write(join(dir, 'auth-graph@owa.json'), owa);
  write(join(dir, 'auth-graph@teams.json'), teams);
  const seen = [];
  const mock = await startMockServer((req, res) => {
    seen.push(req.headers.authorization);
    if (req.headers.authorization === teams) {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ value: [{ id: 'm1' }] }));
    }
    res.writeHead(403, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: { code: 'Forbidden', message: "Missing scope permissions on the request. API requires one of 'ChannelMessage.Read.All'." } }));
  });
  try {
    const { code, stdout } = await runCli(['graph', '/teams/t/channels/c/messages'], {
      env: { OUTLOOK_TOKEN_CACHE: '/nonexistent/o.json', OUTLOOK_TOKEN_CACHE_GRAPH: primary, OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0` },
    });
    assert.equal(code, 0);
    assert.deepEqual(seen, [owa, teams]);
    assert.equal(JSON.parse(stdout).value[0].id, 'm1');
  } finally {
    await mock.close();
  }
});
