// AI recap: entity id + request shape, recap page URL for the keepalive,
// and the `recap` CLI against a mock Graph + Substrate.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startMockServer } from './helpers.mjs';
import { recapEntityId, recapRequestBody, RECAP_PATH, parseRecap, recapToText } from '../src/transcripts.mjs';
import { recapUrl } from '../src/surfaces.mjs';

const NO_OUTLOOK = '/nonexistent/outlook-auth.json';
const IDS = {
  siteId: '541d5497-4d22-4e51-8e54-ef4c3bcf50aa',
  webId: 'db895ba1-1ed2-4d9a-8945-277f48705648',
  listId: '49ac58e2-2baa-4115-b679-8a7b58dcc441',
};

test('recapEntityId is SPO_ + unpadded base64(site,web,list) + _ + item id', () => {
  // Value Teams web sent for this recording (2026-11-12).
  assert.equal(
    recapEntityId(IDS, '01W7VTZ2UOC5KU3SB7HRCJ7OA5725RRH6L'),
    'SPO_NTQxZDU0OTctNGQyMi00ZTUxLThlNTQtZWY0YzNiY2Y1MGFhLGRiODk1YmExLTFlZDItNGQ5YS04OTQ1LTI3N2Y0ODcwNTY0OCw0OWFjNThlMi0yYmFhLTQxMTUtYjY3OS04YTdiNThkY2M0NDE_01W7VTZ2UOC5KU3SB7HRCJ7OA5725RRH6L',
  );
  assert.equal(recapEntityId({ siteId: 'x' }, 'I'), null);
});

test('recapRequestBody asks for MeetingCatchUp on that entity', () => {
  const b = recapRequestBody('SPO_x_I');
  assert.equal(b.EntityRequests[0].Context.EntityId, 'SPO_x_I');
  assert.equal(b.EntityRequests[0].QueryParameters[0].EntityType, 'MeetingCatchUp');
  assert.match(b.Scenario.Name, /^MeetingCatchUp\./);
  assert.match(RECAP_PATH, /AiNotesV3/);
});

test('recapUrl builds a Teams web recap route, or null without meeting links', () => {
  const url = recapUrl({
    id: 'I1',
    webUrl: 'https://contoso-my.sharepoint.com/personal/me/Documents/Recordings/x.mp4',
    parentReference: { driveId: 'b!D1' },
    source: { threadId: '19:meeting_x@thread.v2', iCalUid: 'UID', meetingOrganizerId: 'org', meetingOrganizerTenantId: 'tid', externalId: 'call' },
  });
  assert.ok(url.startsWith('https://teams.microsoft.com/_#/l/meetingrecap?'));
  const q = new URLSearchParams(url.split('?')[1]);
  assert.equal(q.get('driveId'), 'b!D1');
  assert.equal(q.get('driveItemId'), 'I1');
  assert.equal(q.get('threadId'), '19:meeting_x@thread.v2');
  assert.equal(q.get('callId'), 'call');
  assert.equal(recapUrl({ id: 'I1', webUrl: 'u', source: {} }), null);
});

test('recap without a substrate.office.com token exits 2 toward the keepalive', async () => {
  // A Copilot (sydney) Substrate token is cached, but it is the wrong audience.
  const sydney = seedTokenCache({ claims: { aud: 'https://substrate.office.com/sydney' } });
  const mock = await startMockServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ id: 'I1', name: 'Sync-20261111_102639UTC-Meeting Recording.mp4', sharepointIds: IDS, parentReference: { driveId: 'b!D1' } }));
  });
  try {
    const { code, stderr } = await runCli(['recap', 'b!D1:I1'], {
      env: {
        OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
        OUTLOOK_TOKEN_CACHE_GRAPH: seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } }),
        OUTLOOK_TOKEN_CACHE_SUBSTRATE: sydney,
        OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
      },
    });
    assert.equal(code, 2);
    assert.match(stderr, /recap/i);
  } finally {
    await mock.close();
  }
});

// Synthetic MeetingCatchUp response with the live shape (2026-11-13):
// MeetingSummary / MeetingHighlights are JSON strings inside Source.
function recapBody({ summary = true } = {}) {
  return {
    ApiVersion: '1',
    EntitySets: [
      {
        MeetingInsights: 'Premium',
        ResultSets: [
          {
            Total: 1,
            Results: [
              {
                Source: {
                  MeetingSubject: 'Weekly sync',
                  MeetingEndTime: '2026-11-11T11:00:00Z',
                  AttendeesCount: 3,
                  Organizer: { DisplayName: 'Doe, J' },
                  Speakers: [{ DisplayName: 'Jane Doe' }, { DisplayName: 'Sam Lee' }],
                  TeamsRecapUrl: 'https://teams.example/recap',
                  MeetingSummary: summary
                    ? JSON.stringify({
                        ProcessingResult: 'ProcessedSuccessfully',
                        KeyTopics: [
                          {
                            Headline: 'Launch plan',
                            Summary: 'Agreed the pilot date.',
                            Citations: [{ UtteranceStartOffsetTimespan: '00:06:44.1200000' }],
                            DetailedSummaries: [{ topic: 'Pilot', text: 'Pilot starts in two stores.' }, { topic: 'x', text: '' }],
                          },
                        ],
                      })
                    : null,
                  MeetingHighlights: JSON.stringify({ highlights: [{ title: 'Kick-off', startTimeRelativeInMillisecond: 65_000, durationInSec: 30 }] }),
                  PointsOfInterest: {
                    ActionItems: [
                      { OwnerDisplayName: 'Sam Lee', DisplayCleanContent: 'Send the store list. ', Headline: 'Store list' },
                      { OwnerDisplayName: null, DisplayCleanContent: '' },
                    ],
                    MentionEvents: null,
                  },
                  InsightProcessingStates: {
                    MeetingSummary: summary ? 'ProcessedSuccessfully' : 'Processing',
                    Headline: 'NotRequested',
                    UserMentionCount: 'NotProcessedUnexpectedState',
                  },
                },
              },
            ],
          },
        ],
      },
    ],
  };
}

test('parseRecap pulls notes, follow-up tasks, highlights and speakers', () => {
  const r = parseRecap(recapBody());
  assert.equal(r.status, 'ready');
  assert.equal(r.organizer, 'Doe, J');
  assert.equal(r.attendeesCount, 3);
  assert.deepEqual(r.speakers, ['Jane Doe', 'Sam Lee']);
  assert.deepEqual(r.topics, [
    { headline: 'Launch plan', summary: 'Agreed the pilot date.', at: '00:06:44', details: [{ topic: 'Pilot', text: 'Pilot starts in two stores.' }] },
  ]);
  assert.deepEqual(r.actionItems, [{ owner: 'Sam Lee', text: 'Send the store list.' }]);
  assert.deepEqual(r.highlights, [{ title: 'Kick-off', at: '00:01:05', durationSec: 30 }]);
  assert.deepEqual(r.mentions, []);
  assert.deepEqual(r.pending, ['UserMentionCount: NotProcessedUnexpectedState']);
  assert.equal(r.recapUrl, 'https://teams.example/recap');
});

test('parseRecap reports processing and none', () => {
  const p = parseRecap(recapBody({ summary: false }));
  assert.equal(p.topics.length, 0);
  assert.equal(p.status, 'ready'); // follow-up tasks alone still count
  const body = recapBody({ summary: false });
  body.EntitySets[0].ResultSets[0].Results[0].Source.PointsOfInterest.ActionItems = [];
  assert.equal(parseRecap(body).status, 'processing');
  assert.equal(parseRecap({ EntitySets: [{ ResultSets: [{ Results: [] }] }] }).status, 'none');
  assert.equal(parseRecap(null).status, 'none');
});

test('recapToText lists notes then follow-up tasks', () => {
  const text = recapToText({ subject: 'Weekly sync', startedAt: '2026-11-11T10:00:00Z', ...parseRecap(recapBody()) });
  assert.equal(
    text,
    [
      '# Weekly sync (2026-11-11T10:00:00Z)',
      'Speakers: Jane Doe, Sam Lee',
      '',
      '## Notes',
      '- Launch plan [00:06:44]: Agreed the pilot date.',
      '  - Pilot: Pilot starts in two stores.',
      '',
      '## Follow-up tasks',
      '- Sam Lee: Send the store list.',
    ].join('\n'),
  );
  assert.match(recapToText({ status: 'none', topics: [], actionItems: [] }), /No AI notes yet: none/);
});

test('recap POSTs the MeetingCatchUp query with the mailbox anchor and prints parsed JSON / text', async () => {
  const seen = [];
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      seen.push({ method: req.method, url: req.url, anchor: req.headers['x-anchormailbox'], body });
      res.writeHead(200, { 'Content-Type': 'application/json' });
      if (req.url.startsWith('/v1.0/')) {
        res.end(JSON.stringify({ id: 'I1', name: 'Sync-20261111_102639UTC-Meeting Recording.mp4', sharepointIds: IDS, parentReference: { driveId: 'b!D1' }, source: { iCalUid: 'ical' } }));
      } else {
        res.end(JSON.stringify(recapBody()));
      }
    });
  });
  const env = {
    OUTLOOK_TOKEN_CACHE: NO_OUTLOOK,
    OUTLOOK_TOKEN_CACHE_GRAPH: seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } }),
    OUTLOOK_TOKEN_CACHE_SUBSTRATE: seedTokenCache({ claims: { aud: 'https://substrate.office.com', oid: 'O1', tid: 'T1' } }),
    OUTLOOK_GRAPH_BASE: `${mock.url}/v1.0`,
    OUTLOOK_SUBSTRATE_BASE: mock.url,
  };
  try {
    const { code, stdout, stderr } = await runCli(['recap', 'b!D1:I1'], { env });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.equal(out.id, 'b!D1:I1');
    assert.equal(out.subject, 'Sync');
    assert.equal(out.iCalUid, 'ical');
    assert.equal(out.status, 'ready');
    assert.equal(out.actionItems[0].owner, 'Sam Lee');
    const post = seen.find((s) => s.method === 'POST');
    assert.ok(post, 'recap query was POSTed');
    assert.equal(post.url, RECAP_PATH);
    assert.equal(post.anchor, 'Oid:O1@T1');
    assert.deepEqual(JSON.parse(post.body), recapRequestBody(recapEntityId(IDS, 'I1')));
    assert.ok(seen.filter((s) => s.method !== 'POST').every((s) => s.method === 'GET'));

    const text = await runCli(['recap', 'b!D1:I1', '--format', 'text'], { env });
    assert.equal(text.code, 0, text.stderr);
    assert.match(text.stdout, /^# Sync/);
    assert.match(text.stdout, /## Follow-up tasks\n- Sam Lee: Send the store list\./);
  } finally {
    await mock.close();
  }
});
