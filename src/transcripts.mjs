// Teams meeting recordings and transcripts.
//
// Where things live (verified 2026-11-12 in a locked-down tenant):
//   - Teams saves each recording (and, for older meetings, a separate
//     "...-Meeting Transcript.mp4") in the organiser's OneDrive under
//     /Recordings. Invitees see it in /me/drive/sharedWithMe. Graph lists
//     both, with `source.iCalUid` / `source.threadId` tying the file back to
//     the calendar event and the meeting chat, and `media.mediaSource
//     .contentType` = "Transcript" when a transcript is attached.
//   - Graph will NOT return the transcript text: /onlineMeetings/{id}
//     /transcripts needs OnlineMeetingTranscript.Read.All, which no web app's
//     token carries, and /driveItem/media/transcripts is a 400 on v1.0 and
//     beta ("Unsupported segment type").
//   - SharePoint's own v2.1 API serves it:
//     GET https://<tenant>-my.sharepoint.com/_api/v2.1/drives/{d}/items/{i}/media/transcripts
//     GET .../media/transcripts/{tid}/streamContent?format=json  (or VTT)
//     This needs a SharePoint-audience token (resource "sharepoint"), which
//     the keepalive captures from its OneDrive tab.
//
// Everything here is pure (no I/O) so it can be unit-tested; cli.mjs does
// the calls.

export const MEETING_FILE_RE = /-Meeting (Recording|Transcript)\.mp4$/i;

/** "Weekly Sync-20261111_102639UTC-Meeting Recording.mp4" -> "Weekly Sync". */
export function subjectFromName(name = '') {
  return name
    .replace(/-\d{8}_\d{6}(UTC)?-Meeting (Recording|Transcript)\.mp4$/i, '')
    .replace(MEETING_FILE_RE, '');
}

/** Stable id for a recording: "<driveId>:<itemId>" (drive ids never contain ':'). */
export function encodeRecordingId(driveId, itemId) {
  return `${driveId}:${itemId}`;
}

/** Inverse of encodeRecordingId; a bare item id means "in my own drive". */
export function decodeRecordingId(id = '') {
  const i = id.lastIndexOf(':');
  if (i <= 0) return { driveId: null, itemId: id };
  return { driveId: id.slice(0, i), itemId: id.slice(i + 1) };
}

/** Graph path for a drive item, own drive when driveId is null. */
export function driveItemPath({ driveId, itemId }) {
  const item = encodeURIComponent(itemId);
  return driveId ? `/drives/${encodeURIComponent(driveId)}/items/${item}` : `/me/drive/items/${item}`;
}

export const ITEM_SELECT = 'id,name,createdDateTime,webUrl,parentReference,source,media';

/**
 * One listing row from a Graph driveItem (own Recordings folder or a
 * sharedWithMe item re-fetched with ITEM_SELECT).
 */
export function toListing(item, owner) {
  const media = item.media ?? {};
  const start = media.recordingStartDateTime ?? item.createdDateTime ?? null;
  const end = media.recordingEndDateTime ?? null;
  const minutes = start && end ? Math.round((Date.parse(end) - Date.parse(start)) / 60_000) : null;
  const driveId = item.parentReference?.driveId;
  return {
    id: driveId ? encodeRecordingId(driveId, item.id) : item.id,
    subject: subjectFromName(item.name),
    startedAt: start,
    endedAt: end,
    durationMinutes: Number.isFinite(minutes) ? minutes : null,
    hasTranscript: media.mediaSource?.contentType === 'Transcript',
    kind: /Transcript\.mp4$/i.test(item.name ?? '') ? 'transcript' : 'recording',
    owner,
    organizerId: item.source?.meetingOrganizerId ?? null,
    iCalUid: item.source?.iCalUid ?? null,
    threadId: item.source?.threadId ?? null,
    webUrl: item.webUrl ?? null,
  };
}

/** Newest first, de-duplicated by id. */
export function mergeListings(rows) {
  const seen = new Map();
  for (const r of rows) if (!seen.has(r.id)) seen.set(r.id, r);
  return [...seen.values()].sort((a, b) => Date.parse(b.startedAt ?? 0) - Date.parse(a.startedAt ?? 0));
}

/** "00:01:02.3450000" or "00:01:02.345" -> "00:01:02". */
export function shortOffset(offset) {
  if (typeof offset !== 'string') return null;
  const m = /^(\d+):(\d{2}):(\d{2})/.exec(offset);
  return m ? `${m[1].padStart(2, '0')}:${m[2]}:${m[3]}` : offset;
}

/** Entries from SharePoint's streamContent?format=json body. */
export function entriesFromJson(body) {
  const list = Array.isArray(body?.entries) ? body.entries : [];
  return list
    .map((e) => ({
      speaker: e.speakerDisplayName ?? e.speakerName ?? null,
      start: shortOffset(e.startOffset),
      end: shortOffset(e.endOffset),
      text: (e.text ?? '').trim(),
    }))
    .filter((e) => e.text);
}

/**
 * Entries from a WebVTT transcript. Teams writes cues as
 *   00:00:03.610 --> 00:00:05.050
 *   <v Jane Doe>Hello everyone.</v>
 */
export function entriesFromVtt(vtt = '') {
  const out = [];
  const blocks = String(vtt).replace(/\r/g, '').split(/\n\n+/);
  for (const block of blocks) {
    const lines = block.split('\n');
    const t = lines.findIndex((l) => l.includes('-->'));
    if (t < 0) continue;
    const [start, end] = lines[t].split('-->').map((s) => s.trim().split(/\s/)[0]);
    const raw = lines.slice(t + 1).join(' ');
    const v = /<v\s+([^>]+)>/.exec(raw);
    const text = raw.replace(/<[^>]+>/g, '').trim();
    if (!text) continue;
    out.push({ speaker: v ? v[1].trim() : null, start: shortOffset(start), end: shortOffset(end), text });
  }
  return out;
}

/** Join consecutive utterances by the same speaker into one turn. */
export function mergeTurns(entries) {
  const out = [];
  for (const e of entries) {
    const last = out[out.length - 1];
    if (last && last.speaker === e.speaker) {
      last.text = `${last.text} ${e.text}`;
      last.end = e.end ?? last.end;
    } else {
      out.push({ ...e });
    }
  }
  return out;
}

/** Plain text, one "[hh:mm:ss] Speaker: text" line per turn. */
export function toText(entries) {
  return entries.map((e) => `[${e.start ?? '--:--:--'}] ${e.speaker ?? 'Unknown'}: ${e.text}`).join('\n');
}

/** Pick a transcript: the requested id, else the default one, else the newest. */
export function pickTranscript(list, wanted) {
  if (!Array.isArray(list) || list.length === 0) return null;
  if (wanted) return list.find((t) => t.id === wanted) ?? null;
  return (
    list.find((t) => t.isDefault) ??
    [...list].sort((a, b) => Date.parse(b.createdDateTime ?? 0) - Date.parse(a.createdDateTime ?? 0))[0]
  );
}

// ---------------------------------------------------------------------------
// AI recap (Teams "intelligent recap": AI notes, follow-up tasks, mentions).
//
// Teams' Recap -> "AI summary" tab reads it from Substrate's recommendations
// API with a MeetingCatchUp query (found 2026-11-12 by sniffing the Teams
// web client). The meeting is named by its recording's SharePoint ids, and
// the call needs a Substrate-audience token from Teams web, which it only
// mints on a recap page (the keepalive keeps one open). Graph's aiInsights
// is not reachable (missing scope), and Teams' /api/mcps only lists that a
// summary exists.

export const RECAP_PATH = '/search/api/v1/recommendations/?&setflight=AiTasksV3,AiNotesV3,PeopleMentionsV2';
export const RECAP_SCENARIO = 'MeetingCatchUp.MeetingRecap.Shared';

/** "SPO_<base64(siteId,webId,listId) without padding>_<driveItemId>". */
export function recapEntityId(sharepointIds, itemId) {
  const { siteId, webId, listId } = sharepointIds ?? {};
  if (!siteId || !webId || !listId || !itemId) return null;
  const b64 = Buffer.from(`${siteId},${webId},${listId}`).toString('base64').replace(/=+$/, '');
  return `SPO_${b64}_${itemId}`;
}

export function recapRequestBody(entityId) {
  return {
    EntityRequests: [{ Context: { EntityId: entityId }, QueryParameters: [{ EntityType: 'MeetingCatchUp' }] }],
    Scenario: { Name: RECAP_SCENARIO },
  };
}

function parseMaybeJson(v) {
  if (typeof v !== 'string') return v ?? null;
  try {
    return JSON.parse(v);
  } catch {
    return null;
  }
}

function msToOffset(ms) {
  if (!Number.isFinite(ms)) return null;
  const s = Math.floor(ms / 1000);
  const pad = (n) => String(n).padStart(2, '0');
  return `${pad(Math.floor(s / 3600))}:${pad(Math.floor((s % 3600) / 60))}:${pad(s % 60)}`;
}

/**
 * The useful parts of a MeetingCatchUp response (shape verified live
 * 2026-11-13). The result lives at EntitySets[0].ResultSets[0].Results[0]
 * .Source. MeetingSummary and MeetingHighlights are JSON encoded as strings.
 * InsightProcessingStates says which insights Teams has produced; anything
 * not "ProcessedSuccessfully" is listed under `pending` so callers can tell
 * "no notes yet" from "no notes".
 */
export function parseRecap(body) {
  const src = body?.EntitySets?.[0]?.ResultSets?.[0]?.Results?.[0]?.Source;
  if (!src) return { status: 'none', topics: [], actionItems: [], highlights: [], speakers: [], mentions: [] };
  const summary = parseMaybeJson(src.MeetingSummary) ?? {};
  const hl = parseMaybeJson(src.MeetingHighlights) ?? {};
  const poi = src.PointsOfInterest ?? {};
  const topics = (Array.isArray(summary.KeyTopics) ? summary.KeyTopics : []).map((t) => ({
    headline: t.Headline ?? null,
    summary: t.Summary ?? null,
    at: shortOffset(t.Citations?.[0]?.UtteranceStartOffsetTimespan ?? null),
    details: (t.DetailedSummaries ?? []).map((d) => ({ topic: d.topic ?? null, text: d.text ?? '' })).filter((d) => d.text),
  }));
  const actionItems = (Array.isArray(poi.ActionItems) ? poi.ActionItems : [])
    .map((a) => ({ owner: a.OwnerDisplayName ?? null, text: (a.DisplayCleanContent ?? a.Headline ?? '').trim() }))
    .filter((a) => a.text);
  const highlights = (Array.isArray(hl.highlights) ? hl.highlights : []).map((h) => ({
    title: h.title ?? null,
    at: msToOffset(h.startTimeRelativeInMillisecond),
    durationSec: h.durationInSec ?? null,
  }));
  const mentions = (Array.isArray(poi.MentionEvents) ? poi.MentionEvents : []).map((m) => ({
    speaker: m.SpeakerName ?? m.SpeakerDisplayName ?? null,
    text: (m.DisplayCleanContent ?? m.Headline ?? m.Text ?? '').trim() || null,
  }));
  const states = src.InsightProcessingStates ?? {};
  const pending = Object.entries(states)
    .filter(([, v]) => v !== 'ProcessedSuccessfully' && v !== 'NotRequested')
    .map(([k, v]) => `${k}: ${v}`);
  const ready = topics.length > 0 || actionItems.length > 0;
  return {
    status: ready ? 'ready' : 'processing',
    organizer: src.Organizer?.DisplayName ?? null,
    endedAt: src.MeetingEndTime ?? src.RecordingEndTime ?? null,
    attendeesCount: src.AttendeesCount ?? null,
    speakers: (src.Speakers ?? []).map((s) => s.DisplayName).filter(Boolean),
    topics,
    actionItems,
    highlights,
    mentions,
    pending,
    recapUrl: src.TeamsRecapUrl ?? null,
  };
}

/** Plain-text recap for agents: topics, then follow-up tasks. */
export function recapToText(r) {
  const lines = [];
  if (r.subject) lines.push(`# ${r.subject}${r.startedAt ? ` (${r.startedAt})` : ''}`);
  if (r.speakers?.length) lines.push(`Speakers: ${r.speakers.join(', ')}`);
  if (r.topics?.length) {
    lines.push('', '## Notes');
    for (const t of r.topics) {
      lines.push(`- ${t.headline ?? 'Topic'}${t.at ? ` [${t.at}]` : ''}: ${t.summary ?? ''}`.trimEnd());
      for (const d of t.details) lines.push(`  - ${d.topic ? `${d.topic}: ` : ''}${d.text}`);
    }
  }
  if (r.actionItems?.length) {
    lines.push('', '## Follow-up tasks');
    for (const a of r.actionItems) lines.push(`- ${a.owner ? `${a.owner}: ` : ''}${a.text}`);
  }
  if (r.status !== 'ready') lines.push('', `(No AI notes yet: ${r.status}${r.pending?.length ? `; ${r.pending.join(', ')}` : ''})`);
  return lines.join('\n');
}
