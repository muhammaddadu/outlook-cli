// Meeting transcripts: pure helpers (listing rows, VTT/JSON parsing, turn
// merging), SharePoint token classification, and the `transcripts` /
// `transcript` / `graph --beta` CLI surface against a mock Graph + SharePoint.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startMockServer } from './helpers.mjs';
import { classifyToken, audienceOrigin } from '../src/resources.mjs';
import {
  subjectFromName,
  encodeRecordingId,
  decodeRecordingId,
  driveItemPath,
  toListing,
  mergeListings,
  entriesFromJson,
  entriesFromVtt,
  mergeTurns,
  toText,
  pickTranscript,
} from '../src/transcripts.mjs';

const NO_OUTLOOK = '/nonexistent/outlook-auth.json';
const SPO = 'https://contoso-my.sharepoint.com';

// ---------------------------------------------------------------------------
// resources: the per-tenant SharePoint audience

test('classifyToken files OneDrive (-my) tokens under sharepoint, in both audience forms', () => {
  assert.equal(classifyToken({ aud: SPO }), 'sharepoint');
  assert.equal(classifyToken({ aud: `${SPO}/` }), 'sharepoint');
  assert.equal(
    classifyToken({ aud: '00000003-0000-0ff1-ce00-000000000000/contoso-my.sharepoint.com@tenant-id' }),
    'sharepoint',
  );
  // Tenant-wide forms (OneDrive web and Teams web use these).
  assert.equal(classifyToken({ aud: '00000003-0000-0ff1-ce00-000000000000' }), 'sharepoint');
  assert.equal(classifyToken({ aud: '00000003-0000-0ff1-ce00-000000000000@tenant-id' }), 'sharepoint');
  assert.equal(classifyToken({ aud: '00000003-0000-0ff1-ce00-000000000000/contoso.sharepoint.com@t' }), null);
});

test('classifyToken ignores team-site SharePoint tokens', () => {
  assert.equal(classifyToken({ aud: 'https://contoso.sharepoint.com', host: 'contoso.sharepoint.com' }), null);
  // A Graph-audience token sent to the OneDrive host stays a Graph token.
  assert.equal(classifyToken({ aud: 'https://graph.microsoft.com', host: 'contoso-my.sharepoint.com' }), 'graph');
});

test('classifyToken still uses the static host fallback for unknown audiences', () => {
  assert.equal(classifyToken({ aud: '66a88757-0000-0000-0000-000000000000', host: 'substrate.office.com' }), 'substrate');
});

test('audienceOrigin reads the host from both audience forms', () => {
  assert.equal(audienceOrigin('https://Contoso-my.sharepoint.com/'), SPO);
  assert.equal(audienceOrigin('00000003-0000-0ff1-ce00-000000000000/contoso-my.sharepoint.com@t'), SPO);
  assert.equal(audienceOrigin(undefined), null);
});

// ---------------------------------------------------------------------------
// pure helpers

test('subjectFromName strips the Teams timestamp and suffix', () => {
  assert.equal(subjectFromName('Weekly Sync-20261111_102639UTC-Meeting Recording.mp4'), 'Weekly Sync');
  assert.equal(subjectFromName('Mo-Jen 1:1-20261104_164056-Meeting Transcript.mp4'), 'Mo-Jen 1:1');
});

test('recording ids round-trip, and a bare id means my own drive', () => {
  const id = encodeRecordingId('b!abc-_x', '01ITEM');
  assert.deepEqual(decodeRecordingId(id), { driveId: 'b!abc-_x', itemId: '01ITEM' });
  assert.deepEqual(decodeRecordingId('01ITEM'), { driveId: null, itemId: '01ITEM' });
  assert.equal(driveItemPath({ driveId: null, itemId: '01ITEM' }), '/me/drive/items/01ITEM');
  assert.equal(driveItemPath({ driveId: 'b!d', itemId: 'I' }), '/drives/b!d/items/I');
});

const ITEM = {
  id: 'I1',
  name: 'Weekly Sync-20261111_102639UTC-Meeting Recording.mp4',
  createdDateTime: '2026-11-11T10:26:39Z',
  webUrl: `${SPO}/personal/me/Documents/Recordings/x.mp4`,
  parentReference: { driveId: 'b!D1' },
  source: { iCalUid: 'UID1', threadId: '19:meeting_x@thread.v2', meetingOrganizerId: 'org-1' },
  media: {
    mediaSource: { contentType: 'Transcript' },
    recordingStartDateTime: '2026-11-11T10:26:39Z',
    recordingEndDateTime: '2026-11-11T11:02:39Z',
  },
};

test('toListing builds a row with meeting links and transcript flag', () => {
  const row = toListing(ITEM, 'me');
  assert.equal(row.id, 'b!D1:I1');
  assert.equal(row.subject, 'Weekly Sync');
  assert.equal(row.durationMinutes, 36);
  assert.equal(row.hasTranscript, true);
  assert.equal(row.kind, 'recording');
  assert.equal(row.iCalUid, 'UID1');
  assert.equal(row.threadId, '19:meeting_x@thread.v2');
  const plain = toListing({ ...ITEM, media: { mediaSource: { contentType: 'Recording' } } }, 'shared');
  assert.equal(plain.hasTranscript, false);
  assert.equal(plain.startedAt, ITEM.createdDateTime);
});

test('mergeListings de-duplicates and sorts newest first', () => {
  const rows = mergeListings([
    { id: 'a', startedAt: '2026-11-01T00:00:00Z' },
    { id: 'b', startedAt: '2026-11-05T00:00:00Z' },
    { id: 'a', startedAt: '2026-11-01T00:00:00Z' },
  ]);
  assert.deepEqual(rows.map((r) => r.id), ['b', 'a']);
});

test('entriesFromJson maps SharePoint entries and drops empty text', () => {
  const entries = entriesFromJson({
    entries: [
      { speakerDisplayName: 'Jane Doe', startOffset: '00:00:03.6100000', endOffset: '00:00:05.0500000', text: ' Hello. ' },
      { speakerDisplayName: 'Jane Doe', startOffset: '00:00:06', endOffset: '00:00:07', text: '' },
    ],
  });
  assert.deepEqual(entries, [{ speaker: 'Jane Doe', start: '00:00:03', end: '00:00:05', text: 'Hello.' }]);
  assert.deepEqual(entriesFromJson({}), []);
});

test('entriesFromVtt parses Teams voice tags', () => {
  const vtt = [
    'WEBVTT',
    '',
    'abc/12-0',
    '00:00:03.610 --> 00:00:05.050',
    '<v Jane Doe>Hello everyone.</v>',
    '',
    '00:00:05.100 --> 00:00:06.000',
    '<v Mo Dadu>Hi Jane.</v>',
    '',
  ].join('\r\n');
  assert.deepEqual(entriesFromVtt(vtt), [
    { speaker: 'Jane Doe', start: '00:00:03', end: '00:00:05', text: 'Hello everyone.' },
    { speaker: 'Mo Dadu', start: '00:00:05', end: '00:00:06', text: 'Hi Jane.' },
  ]);
});

test('mergeTurns joins consecutive lines by one speaker; toText formats them', () => {
  const merged = mergeTurns([
    { speaker: 'A', start: '00:00:01', end: '00:00:02', text: 'One.' },
    { speaker: 'A', start: '00:00:02', end: '00:00:03', text: 'Two.' },
    { speaker: 'B', start: '00:00:04', end: '00:00:05', text: 'Three.' },
  ]);
  assert.deepEqual(merged, [
    { speaker: 'A', start: '00:00:01', end: '00:00:03', text: 'One. Two.' },
    { speaker: 'B', start: '00:00:04', end: '00:00:05', text: 'Three.' },
  ]);
  assert.equal(toText(merged), '[00:00:01] A: One. Two.\n[00:00:04] B: Three.');
});

test('pickTranscript prefers the requested id, then default, then newest', () => {
  const list = [
    { id: 'old', createdDateTime: '2026-01-01T00:00:00Z' },
    { id: 'new', createdDateTime: '2026-02-01T00:00:00Z' },
  ];
  assert.equal(pickTranscript(list).id, 'new');
  assert.equal(pickTranscript([...list, { id: 'def', isDefault: true }]).id, 'def');
  assert.equal(pickTranscript(list, 'old').id, 'old');
  assert.equal(pickTranscript(list, 'nope'), null);
  assert.equal(pickTranscript([]), null);
});

// ---------------------------------------------------------------------------
// CLI surface

function graphCache() {
  return seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } });
}

test('transcripts lists own and shared recordings with transcripts, newest first', async () => {
  const now = Date.now();
  const iso = (daysAgo) => new Date(now - daysAgo * 86_400_000).toISOString();
  const mine = { ...ITEM, id: 'MINE', createdDateTime: iso(1), media: { mediaSource: { contentType: 'Transcript' }, recordingStartDateTime: iso(1) } };
  const noTx = { ...ITEM, id: 'NOTX', createdDateTime: iso(2), media: { mediaSource: { contentType: 'Recording' }, recordingStartDateTime: iso(2) } };
  const old = { ...ITEM, id: 'OLD', createdDateTime: iso(40), media: { mediaSource: { contentType: 'Transcript' }, recordingStartDateTime: iso(40) } };
  const sharedFull = { ...ITEM, id: 'SH', parentReference: { driveId: 'b!OTHER' }, createdDateTime: iso(0.5), media: { mediaSource: { contentType: 'Transcript' }, recordingStartDateTime: iso(0.5) } };
  const seen = [];
  const mock = await startMockServer((req, res) => {
    seen.push(`${req.method} ${req.url}`);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    if (req.url.startsWith('/v1.0/me/drive/root:/Recordings:/children')) {
      res.end(JSON.stringify({ value: [mine, noTx, old, { id: 'DOC', name: 'notes.docx' }] }));
    } else if (req.url.startsWith('/v1.0/me/drive/sharedWithMe')) {
      res.end(JSON.stringify({
        value: [
          { name: sharedFull.name, remoteItem: { id: 'SH', createdDateTime: iso(0.5), parentReference: { driveId: 'b!OTHER' } } },
          { name: 'Budget.xlsx', remoteItem: { id: 'X', createdDateTime: iso(0.5), parentReference: { driveId: 'b!OTHER' } } },
        ],
      }));
    } else if (req.url.startsWith('/v1.0/drives/b!OTHER/items/SH')) {
      res.end(JSON.stringify(sharedFull));
    } else {
      res.writeHead(404);
      res.end('{}');
    }
  });
  try {
    const env = { OUTLOOK_TOKEN_CACHE: NO_OUTLOOK, OUTLOOK_TOKEN_CACHE_GRAPH: graphCache(), OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0` };
    const { code, stdout, stderr } = await runCli(['transcripts', '--since', '14d'], { env });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.deepEqual(out.value.map((r) => r.id), ['b!OTHER:SH', 'b!D1:MINE']);
    assert.deepEqual(out.value.map((r) => r.owner), ['shared', 'me']);
    assert.ok(seen.every((s) => s.startsWith('GET ')), 'only GET requests');

    const all = JSON.parse((await runCli(['transcripts', '--all', '--mine'], { env })).stdout);
    assert.deepEqual(all.value.map((r) => r.id), ['b!D1:MINE', 'b!D1:NOTX']);
  } finally {
    await mock.close();
  }
});

test('transcripts rejects a bad --since with exit 64', async () => {
  const { code, stderr } = await runCli(['transcripts', '--since', 'last tuesday'], {
    env: { OUTLOOK_TOKEN_CACHE: NO_OUTLOOK },
  });
  assert.equal(code, 64);
  assert.match(stderr, /Invalid --since/);
});

function spoCache(origin) {
  return seedTokenCache({ claims: { aud: origin } });
}

async function transcriptMock({ contentJson = true } = {}) {
  const seen = [];
  const mock = await startMockServer((req, res) => {
    seen.push({ method: req.method, url: req.url, auth: req.headers.authorization });
    const json = (status, body) => {
      res.writeHead(status, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(body));
    };
    if (req.url.startsWith('/v1.0/drives/b!D1/items/I1')) return json(200, { ...ITEM, webUrl: `${mock.url}/personal/me/x.mp4` });
    if (req.url === '/_api/v2.1/drives/b!D1/items/I1/media/transcripts') {
      return json(200, { value: [{ id: 'T1', languageTag: 'en-gb', createdDateTime: '2026-11-11T11:03:00Z' }] });
    }
    if (req.url === '/_api/v2.1/drives/b!D1/items/I1/media/transcripts/T1/streamContent?format=json') {
      if (!contentJson) return json(400, { error: { code: 'invalidRequest', message: 'format' } });
      return json(200, {
        entries: [
          { speakerDisplayName: 'Jane Doe', startOffset: '00:00:01.0000000', endOffset: '00:00:02.0000000', text: 'Hi.' },
          { speakerDisplayName: 'Jane Doe', startOffset: '00:00:02.0000000', endOffset: '00:00:03.0000000', text: 'Agenda first.' },
          { speakerDisplayName: 'Mo Dadu', startOffset: '00:00:04.0000000', endOffset: '00:00:05.0000000', text: 'Sounds good.' },
        ],
      });
    }
    if (req.url === '/_api/v2.1/drives/b!D1/items/I1/media/transcripts/T1/streamContent') {
      res.writeHead(200, { 'Content-Type': 'text/vtt' });
      return res.end('WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Jane Doe>Hi.</v>\n');
    }
    json(404, { error: { code: 'itemNotFound', message: req.url } });
  });
  return { mock, seen };
}

test('transcript prints merged, speaker-attributed entries using the SharePoint token', async () => {
  const { mock, seen } = await transcriptMock();
  try {
    const { code, stdout, stderr } = await runCli(['transcript', 'b!D1:I1'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: graphCache(),
        OUTLOOK_TOKEN_CACHE_SHAREPOINT: spoCache(mock.url),
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
      },
    });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.equal(out.subject, 'Weekly Sync');
    assert.equal(out.transcriptId, 'T1');
    assert.equal(out.language, 'en-gb');
    assert.equal(out.iCalUid, 'UID1');
    assert.deepEqual(out.speakers, ['Jane Doe', 'Mo Dadu']);
    assert.deepEqual(out.entries[0], { speaker: 'Jane Doe', start: '00:00:01', end: '00:00:03', text: 'Hi. Agenda first.' });
    assert.ok(seen.every((s) => s.method === 'GET'), 'only GET requests');
    // The SharePoint calls carry a different token from the Graph call.
    const graphAuth = seen.find((s) => s.url.startsWith('/v1.0/')).auth;
    const spoAuth = seen.find((s) => s.url.startsWith('/_api/')).auth;
    assert.notEqual(graphAuth, spoAuth);
  } finally {
    await mock.close();
  }
});

test('transcript --format text falls back to VTT when JSON is refused', async () => {
  const { mock } = await transcriptMock({ contentJson: false });
  try {
    const { code, stdout, stderr } = await runCli(['transcript', 'b!D1:I1', '--format', 'text'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: graphCache(),
        OUTLOOK_TOKEN_CACHE_SHAREPOINT: spoCache(mock.url),
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
      },
    });
    assert.equal(code, 0, stderr);
    assert.equal(stdout, '[00:00:01] Jane Doe: Hi.\n');
  } finally {
    await mock.close();
  }
});

test('transcript without a SharePoint token exits 2 and points at the keepalive', async () => {
  const { mock } = await transcriptMock();
  try {
    const { code, stderr } = await runCli(['transcript', 'b!D1:I1'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: graphCache(),
        OUTLOOK_TOKEN_CACHE_SHAREPOINT: '/nonexistent/auth-sharepoint.json',
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
      },
    });
    assert.equal(code, 2);
    assert.match(stderr, /SharePoint/);
    assert.match(stderr, /keepalive/);
  } finally {
    await mock.close();
  }
});

test('graph --beta swaps the version segment; rejects non-graph resources', async () => {
  let observed = null;
  const mock = await startMockServer((req, res) => {
    observed = req.url;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end('{}');
  });
  try {
    const env = { OUTLOOK_TOKEN_CACHE: NO_OUTLOOK, OUTLOOK_TOKEN_CACHE_GRAPH: graphCache(), OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0` };
    const { code } = await runCli(['graph', '--beta', '/me'], { env });
    assert.equal(code, 0);
    assert.equal(observed, '/beta/me');
    const bad = await runCli(['graph', '--beta', '--resource', 'substrate', '/x'], { env });
    assert.equal(bad.code, 64);
  } finally {
    await mock.close();
  }
});
