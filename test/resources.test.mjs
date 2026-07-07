// Unit tests for the resource registry: audience/host classification and
// base-URL resolution (including env overrides used by the CLI tests).

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  RESOURCES,
  DEFAULT_RESOURCE,
  resource,
  resourceBase,
  classifyToken,
} from '../src/resources.mjs';

test('default resource is outlook and exists in the registry', () => {
  assert.equal(DEFAULT_RESOURCE, 'outlook');
  assert.ok(RESOURCES.outlook);
});

test('classifyToken maps known audiences to resources', () => {
  assert.equal(classifyToken({ aud: 'https://outlook.office.com' }), 'outlook');
  assert.equal(classifyToken({ aud: 'https://outlook.office.com/' }), 'outlook');
  assert.equal(classifyToken({ aud: 'https://graph.microsoft.com' }), 'graph');
  // Graph is sometimes issued under its app GUID rather than the URL.
  assert.equal(
    classifyToken({ aud: '00000003-0000-0000-c000-000000000000' }),
    'graph',
  );
  assert.equal(classifyToken({ aud: 'https://substrate.office.com' }), 'substrate');
});

test('classifyToken falls back to request host when audience is unknown', () => {
  assert.equal(classifyToken({ aud: 'urn:weird', host: 'graph.microsoft.com' }), 'graph');
  assert.equal(classifyToken({ host: 'outlook.office.com' }), 'outlook');
});

test('classifyToken returns null for resources we do not model', () => {
  assert.equal(classifyToken({ aud: 'https://api.spaces.skype.com' }), null);
  assert.equal(classifyToken({}), null);
});

test('resourceBase honours the per-resource env override', () => {
  const prev = process.env.OUTLOOK_GRAPH_BASE;
  try {
    process.env.OUTLOOK_GRAPH_BASE = 'http://127.0.0.1:9/graph';
    assert.equal(resourceBase('graph'), 'http://127.0.0.1:9/graph');
  } finally {
    if (prev === undefined) delete process.env.OUTLOOK_GRAPH_BASE;
    else process.env.OUTLOOK_GRAPH_BASE = prev;
  }
});

test('resource() throws a helpful error for an unknown key', () => {
  assert.throws(() => resource('teams'), /Unknown resource "teams"/);
});
