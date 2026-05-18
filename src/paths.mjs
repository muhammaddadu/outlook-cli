// XDG-compliant paths for the CLI's stateful data and caches.
//
// All path resolvers are exported as functions (not constants) so they
// re-read env vars on each call. That's what makes the test suite able to
// stub `OUTLOOK_TOKEN_CACHE` per-test by setting `process.env` before
// invoking the code under test, without having to evict modules from
// Node's loader cache.

import { homedir } from 'node:os';
import { resolve } from 'node:path';

const APP = 'outlook-spike';

function xdgData() {
  return process.env.XDG_DATA_HOME ?? resolve(homedir(), '.local', 'share');
}
function xdgCache() {
  return process.env.XDG_CACHE_HOME ?? resolve(homedir(), '.cache');
}

export function dataDir() {
  return resolve(xdgData(), APP);
}
export function cacheDir() {
  return resolve(xdgCache(), APP);
}
export function profileDir() {
  return process.env.OUTLOOK_PROFILE ?? resolve(dataDir(), 'browser-profile');
}
export function tokenCacheFile() {
  return process.env.OUTLOOK_TOKEN_CACHE ?? resolve(cacheDir(), 'auth.json');
}
