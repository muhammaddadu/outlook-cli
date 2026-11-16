// Keepalive renewal of on-demand tokens (Teams Graph, recap Substrate,
// SharePoint), and `refresh --all` waiting for them.
import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { makeFakeJwt } from './helpers.mjs';
import {
  TEAMS_APPID,
  RENEW_MARGIN_MS,
  RENEW_RETRY_MS,
  tokenExpiries,
  staleTokens,
  tabsToRenew,
} from '../src/renewal.mjs';
import { tabName } from '../src/keepalive.mjs';

const NOW = Date.parse('2026-11-16T09:00:00Z');
const exp = (minutes) => Math.floor((NOW + minutes * 60_000) / 1000);

test('tokenExpiries picks the right token per renewed kind', () => {
  const e = tokenExpiries({
    graph: [
      { appid: 'owa', exp: exp(600) }, // OWA's Graph token does not count for Teams
      { appid: TEAMS_APPID, exp: exp(20) },
    ],
    substrate: [
      { aud: 'https://substrate.office.com/sydney', exp: exp(60) }, // Copilot, wrong audience
      { aud: 'https://substrate.office.com', exp: exp(-5) },
      { aud: 'https://substrate.office.com/', exp: exp(30) },
    ],
  });
  assert.equal(e['teams-graph'], exp(20) * 1000);
  assert.equal(e['recap-substrate'], exp(30) * 1000);
  assert.equal(e.sharepoint, null);
});

test('staleTokens: missing, expired, or inside the margin', () => {
  const e = { 'teams-graph': NOW + RENEW_MARGIN_MS + 60_000, 'recap-substrate': NOW + 60_000, sharepoint: null };
  assert.deepEqual(staleTokens(e, NOW), ['recap-substrate', 'sharepoint']);
  assert.deepEqual(staleTokens(e, NOW, 30_000), ['sharepoint']);
});

test('tabsToRenew: only tabs that exist, own a stale token, and were not just opened', () => {
  const expiries = { 'teams-graph': NOW + 3_600_000, 'recap-substrate': NOW - 1, sharepoint: NOW + 120_000 };
  const tabs = new Set(['outlook', 'teams', 'onedrive']);
  assert.deepEqual(tabsToRenew({ expiries, now: NOW, tabs }).sort(), ['onedrive', 'teams']);
  // Teams was re-opened a minute ago: give it time before trying again.
  assert.deepEqual(tabsToRenew({ expiries, now: NOW, tabs, lastOpened: { teams: NOW - 60_000 } }), ['onedrive']);
  assert.deepEqual(
    tabsToRenew({ expiries, now: NOW, tabs, lastOpened: { teams: NOW - RENEW_RETRY_MS } }).sort(),
    ['onedrive', 'teams'],
  );
  // Only SharePoint stale: Teams web mints it too, so both tabs re-open.
  const spOnly = { 'teams-graph': NOW + 3_600_000, 'recap-substrate': NOW + 3_600_000, sharepoint: NOW + 60_000 };
  assert.deepEqual(tabsToRenew({ expiries: spOnly, now: NOW, tabs }).sort(), ['onedrive', 'teams']);
  // No OneDrive tab (no Graph token at start): Teams alone renews SharePoint.
  assert.deepEqual(tabsToRenew({ expiries: spOnly, now: NOW, tabs: new Set(['outlook', 'teams']) }), ['teams']);
  // Everything fresh: nothing to do.
  const fresh = { 'teams-graph': NOW + 3_600_000, 'recap-substrate': NOW + 3_600_000, sharepoint: NOW + 3_600_000 };
  assert.deepEqual(tabsToRenew({ expiries: fresh, now: NOW, tabs }), []);
});

test('tabName maps keepalive URLs to the tab names renewal uses', () => {
  assert.equal(tabName('https://teams.microsoft.com/_#/l/meetingrecap?x=1'), 'teams');
  assert.equal(tabName('https://teams.microsoft.com/v2/'), 'teams');
  assert.equal(tabName('https://contoso-my.sharepoint.com/personal/me/_layouts/15/onedrive.aspx'), 'onedrive');
  assert.equal(tabName('https://outlook.office.com/mail/'), 'outlook');
  assert.equal(tabName('https://m365.cloud.microsoft/chat/'), 'copilot');
});

test('refresh --all with a keepalive waits until the renewed tokens are fresh', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-renew-'));
  const env = {
    OUTLOOK_KEEPALIVE_PID: join(dir, 'keepalive.pid'),
    OUTLOOK_KEEPALIVE_STATUS: join(dir, 'keepalive.json'),
    OUTLOOK_TOKEN_CACHE: join(dir, 'auth.json'),
    OUTLOOK_TOKEN_CACHE_GRAPH: join(dir, 'auth-graph.json'),
    OUTLOOK_TOKEN_CACHE_SUBSTRATE: join(dir, 'auth-substrate.json'),
    OUTLOOK_TOKEN_CACHE_SHAREPOINT: join(dir, 'auth-sharepoint.json'),
  };
  Object.assign(process.env, env);
  const jwt = (claims) => `Bearer ${makeFakeJwt({ expSeconds: Math.floor(Date.now() / 1000) + 3600, claims })}`;
  const put = (file, claims) => writeFileSync(file, JSON.stringify({ headers: { authorization: jwt(claims) } }));
  writeFileSync(env.OUTLOOK_KEEPALIVE_PID, String(process.pid));
  // Teams tab, no OneDrive tab. SharePoint is fresh from the start, so the
  // wait is for the Teams Graph and recap Substrate tokens.
  writeFileSync(env.OUTLOOK_KEEPALIVE_STATUS, JSON.stringify({ state: 'signed-in', tabs: ['outlook', 'teams'] }));
  put(env.OUTLOOK_TOKEN_CACHE, {});
  put(env.OUTLOOK_TOKEN_CACHE_SHAREPOINT, { aud: '00000003-0000-0ff1-ce00-000000000000' });
  // Only OWA's tokens are fresh at first, as on 2026-11-16.
  put(join(dir, 'auth-graph@owa.json'), { aud: 'https://graph.microsoft.com', appid: 'owa' });
  const onUsr1 = () => {};
  process.on('SIGUSR1', onUsr1);
  try {
    const { refreshAuth } = await import('../src/auth.mjs');
    const started = Date.now();
    setTimeout(() => put(join(dir, `auth-graph@${TEAMS_APPID}.json`), { aud: 'https://graph.microsoft.com', appid: TEAMS_APPID }), 300);
    setTimeout(() => put(env.OUTLOOK_TOKEN_CACHE_SUBSTRATE, { aud: 'https://substrate.office.com', appid: TEAMS_APPID }), 2_500);
    const map = await refreshAuth({ wantResource: null, silent: true, waitForAll: true });
    assert.ok(Date.now() - started >= 2_400, 'returned before the recap token was renewed');
    assert.ok(map.outlook && map.substrate);
  } finally {
    process.off('SIGUSR1', onUsr1);
    for (const k of Object.keys(env)) delete process.env[k];
  }
});
