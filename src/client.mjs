// Pure-fetch wrapper around the Outlook REST v2 API.
//
// All Playwright-dependent code lives in capture.mjs. Keeping this module
// dependency-free makes it cheap to unit-test and means cache-hit calls
// don't pay the Playwright cold-load cost.
//
// `OUTLOOK_API_BASE` overrides the base URL — useful for tests pointing at
// a local mock server.

import { debug } from './output.mjs';

const REST_BASE =
  process.env.OUTLOOK_API_BASE ?? 'https://outlook.office.com/api/v2.0/me';

/** Call the Outlook REST v2 API with the captured headers. */
export async function call(auth, path, init = {}) {
  const url = `${REST_BASE}${path}`;
  debug('fetch', init.method ?? 'GET', url);
  const res = await fetch(url, {
    ...init,
    headers: {
      ...auth,
      Accept: 'application/json',
      ...(init.headers ?? {}),
    },
  });
  const text = await res.text();
  let body;
  try {
    body = JSON.parse(text);
  } catch {
    body = text;
  }
  return { status: res.status, body };
}
