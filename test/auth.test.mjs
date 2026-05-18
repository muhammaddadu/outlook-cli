// Unit tests for the token-cache lifecycle.
//
// `src/paths.mjs` reads env vars on every call, so we can isolate each test
// by setting OUTLOOK_TOKEN_CACHE to a unique tmp file via `t.before()`.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { mkdtempSync, writeFileSync, existsSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { makeFakeJwt } from './helpers.mjs';
import {
  loadCachedAuth,
  saveAuth,
  clearAuth,
} from '../src/auth.mjs';

function isolatedCacheFile() {
  const file = join(mkdtempSync(join(tmpdir(), 'outlook-auth-')), 'auth.json');
  process.env.OUTLOOK_TOKEN_CACHE = file;
  return file;
}

test('loadCachedAuth returns null when cache file is missing', () => {
  isolatedCacheFile();
  assert.equal(loadCachedAuth(), null);
});

test('loadCachedAuth returns null when JWT has already expired', () => {
  const file = isolatedCacheFile();
  const exp = Math.floor(Date.now() / 1000) - 60;
  writeFileSync(
    file,
    JSON.stringify({
      headers: { authorization: `Bearer ${makeFakeJwt({ expSeconds: exp })}` },
      expiresAt: exp * 1000,
    }),
  );
  assert.equal(loadCachedAuth(), null);
});

test('loadCachedAuth returns null when JWT is within the 5-minute refresh window', () => {
  const file = isolatedCacheFile();
  const exp = Math.floor(Date.now() / 1000) + 120; // under MIN_LIFETIME_MS
  writeFileSync(
    file,
    JSON.stringify({
      headers: { authorization: `Bearer ${makeFakeJwt({ expSeconds: exp })}` },
      expiresAt: exp * 1000,
    }),
  );
  assert.equal(loadCachedAuth(), null);
});

test('loadCachedAuth returns headers when JWT is comfortably valid', () => {
  const file = isolatedCacheFile();
  const exp = Math.floor(Date.now() / 1000) + 3600;
  const headers = {
    authorization: `Bearer ${makeFakeJwt({ expSeconds: exp })}`,
    'x-anchormailbox': 'PUID:test',
  };
  writeFileSync(file, JSON.stringify({ headers, expiresAt: exp * 1000 }));

  const cached = loadCachedAuth();
  assert.ok(cached, 'expected fresh cache to load');
  assert.equal(cached.headers.authorization, headers.authorization);
  assert.equal(cached.headers['x-anchormailbox'], 'PUID:test');
});

test('loadCachedAuth returns null when JSON is malformed', () => {
  const file = isolatedCacheFile();
  writeFileSync(file, 'this is not json');
  assert.equal(loadCachedAuth(), null);
});

test('loadCachedAuth returns null when authorization header is missing', () => {
  const file = isolatedCacheFile();
  writeFileSync(file, JSON.stringify({ headers: { foo: 'bar' } }));
  assert.equal(loadCachedAuth(), null);
});

test('saveAuth + loadCachedAuth roundtrip preserves headers', () => {
  const file = isolatedCacheFile();
  const exp = Math.floor(Date.now() / 1000) + 3600;
  const headers = {
    authorization: `Bearer ${makeFakeJwt({ expSeconds: exp })}`,
    'x-anchormailbox': 'PUID:abc',
    'x-tenantid': 'tenant-uuid',
  };
  saveAuth(headers);

  const cached = loadCachedAuth();
  assert.deepEqual(cached.headers, headers);

  // File on disk has the expected shape.
  const onDisk = JSON.parse(readFileSync(file, 'utf8'));
  assert.ok(onDisk.savedAt);
  assert.ok(onDisk.expiresAt);
});

test('clearAuth removes the cache file', () => {
  const file = isolatedCacheFile();
  const exp = Math.floor(Date.now() / 1000) + 3600;
  saveAuth({ authorization: `Bearer ${makeFakeJwt({ expSeconds: exp })}` });
  assert.ok(existsSync(file));
  clearAuth();
  assert.equal(existsSync(file), false);
});

test('clearAuth is a no-op when no cache file exists', () => {
  isolatedCacheFile();
  assert.doesNotThrow(() => clearAuth());
});
