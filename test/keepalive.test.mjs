// Pure helpers behind `outlook keepalive`, plus the refresh → keepalive hand-off.
import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { mkdtempSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { makeFakeJwt } from './helpers.mjs';
import {
  isWithinHours,
  isNewerToken,
  keepalivePid,
  launchdPlist,
  LAUNCHD_LABEL,
} from '../src/keepalive.mjs';

const at = (iso) => new Date(iso);
const bearer = (secondsFromNow) =>
  `Bearer ${makeFakeJwt({ expSeconds: Math.floor(Date.now() / 1000) + secondsFromNow })}`;

test('isWithinHours: weekday inside the window', () => {
  assert.equal(isWithinHours(at('2026-11-11T10:00:00')), true); // Wednesday
});

test('isWithinHours: before start, at end, and weekend are outside', () => {
  assert.equal(isWithinHours(at('2026-11-11T07:30:00')), false);
  assert.equal(isWithinHours(at('2026-11-11T19:00:00')), false);
  assert.equal(isWithinHours(at('2026-11-14T10:00:00')), false); // Saturday
});

test('isNewerToken: saves when nothing is cached', () => {
  assert.equal(isNewerToken(null, { authorization: bearer(3600) }), true);
});

test('isNewerToken: ignores the same token and older tokens', () => {
  const cur = { authorization: bearer(3600) };
  assert.equal(isNewerToken(cur, { ...cur }), false);
  assert.equal(isNewerToken(cur, { authorization: bearer(600) }), false);
});

test('isNewerToken: accepts a later-expiring token', () => {
  assert.equal(isNewerToken({ authorization: bearer(600) }, { authorization: bearer(4000) }), true);
});

test('keepalivePid: live pid is reported, stale pid file is removed', () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-ka-'));
  const file = join(dir, 'keepalive.pid');
  process.env.OUTLOOK_KEEPALIVE_PID = file;
  try {
    writeFileSync(file, String(process.pid));
    assert.equal(keepalivePid(), process.pid);
    writeFileSync(file, '999999');
    assert.equal(keepalivePid(), null);
    assert.equal(existsSync(file), false);
  } finally {
    delete process.env.OUTLOOK_KEEPALIVE_PID;
  }
});

test('launchdPlist: runs `keepalive`, at login and weekdays 07:55', () => {
  const xml = launchdPlist({ nodePath: '/n/node', cliPath: '/c/cli.mjs', logFile: '/l/k.log' });
  assert.match(xml, new RegExp(`<string>${LAUNCHD_LABEL}</string>`));
  assert.match(xml, /<string>\/n\/node<\/string>\s*<string>\/c\/cli\.mjs<\/string>\s*<string>keepalive<\/string>/);
  assert.match(xml, /<key>RunAtLoad<\/key><true\/>/);
  for (const d of [1, 2, 3, 4, 5]) assert.match(xml, new RegExp(`<key>Weekday</key><integer>${d}</integer><key>Hour</key><integer>7</integer><key>Minute</key><integer>55</integer>`));
  assert.doesNotMatch(xml, /<key>Weekday<\/key><integer>[06]<\/integer>/);
});

test('refreshAuth waits for a running keepalive instead of opening a browser', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-ka-'));
  process.env.OUTLOOK_KEEPALIVE_PID = join(dir, 'keepalive.pid');
  process.env.OUTLOOK_TOKEN_CACHE = join(dir, 'auth.json');
  writeFileSync(process.env.OUTLOOK_KEEPALIVE_PID, String(process.pid));
  // Our own pid stands in for keepalive; swallow the SIGUSR1 poke.
  const onUsr1 = () => {};
  process.on('SIGUSR1', onUsr1);
  try {
    const { refreshAuth } = await import('../src/auth.mjs');
    const auth = bearer(3600);
    setTimeout(() => writeFileSync(process.env.OUTLOOK_TOKEN_CACHE, JSON.stringify({ headers: { authorization: auth } })), 300);
    const headers = await refreshAuth({ silent: true });
    assert.equal(headers.authorization, auth);
    await assert.rejects(refreshAuth({ interactive: true }), /keepalive/);
  } finally {
    process.off('SIGUSR1', onUsr1);
    delete process.env.OUTLOOK_KEEPALIVE_PID;
    delete process.env.OUTLOOK_TOKEN_CACHE;
  }
});
