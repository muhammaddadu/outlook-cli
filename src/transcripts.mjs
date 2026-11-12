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
