// File-attachment helpers for draft/send commands.
//
// Outlook REST v2 accepts attachments either embedded inside the Message
// resource (used by /sendmail) or POSTed to /messages/{id}/attachments
// (used after a draft is created). Both forms expect a #Microsoft.Outlook
// Services.FileAttachment with base64-encoded ContentBytes.

import { readFileSync, statSync } from 'node:fs';
import { basename, extname, resolve } from 'node:path';

import { AppError, E } from './errors.mjs';

// Per Microsoft Graph docs, the inline FileAttachment ceiling is ~3 MB;
// above that the server rejects with 413. The large-attachment upload
// session API exists but is overkill for typical screenshots / PDFs.
const MAX_BYTES = 3 * 1024 * 1024;

const MIME_BY_EXT = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.heic': 'image/heic',
  '.pdf': 'application/pdf',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
  '.md': 'text/markdown',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.zip': 'application/zip',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

function mimeFor(filename) {
  return MIME_BY_EXT[extname(filename).toLowerCase()] ?? 'application/octet-stream';
}

/**
 * Verify a file exists, is readable, and is below the inline-attachment
 * size ceiling — without loading its bytes. Used to fail fast on bad
 * paths before we call any create-draft / sendmail endpoint, so a typo
 * does not leave an orphan empty draft in the user's mailbox.
 */
export function validateAttachPath(pathArg) {
  const abs = resolve(pathArg);
  let stat;
  try {
    stat = statSync(abs);
  } catch (cause) {
    throw new AppError({
      code: E.ARGS,
      message: `Attachment not found: ${pathArg}`,
      hint: 'Pass an absolute or working-dir-relative path to a readable file.',
      cause,
    });
  }
  if (!stat.isFile()) {
    throw new AppError({
      code: E.ARGS,
      message: `Attachment is not a regular file: ${pathArg}`,
      hint: 'Directories and symlinks-to-dirs are not supported.',
    });
  }
  if (stat.size > MAX_BYTES) {
    throw new AppError({
      code: E.ARGS,
      message: `Attachment exceeds the ${MAX_BYTES} byte inline limit: ${pathArg} (${stat.size} bytes)`,
      hint: 'The Outlook inline-attachment API caps at ~3 MB. Compress or send via OneDrive instead.',
    });
  }
}

/**
 * Read a local file and return a FileAttachment resource shaped for the
 * Outlook REST v2 API. Throws AppError(E.ARGS) for missing/oversized files
 * so the top-level handler surfaces a clean message + exit 64.
 */
export function buildFileAttachment(pathArg) {
  validateAttachPath(pathArg);
  const abs = resolve(pathArg);
  const buf = readFileSync(abs);
  const name = basename(abs);
  return {
    '@odata.type': '#Microsoft.OutlookServices.FileAttachment',
    Name: name,
    ContentType: mimeFor(name),
    ContentBytes: buf.toString('base64'),
  };
}

/**
 * Attach one-or-more files to an existing draft.
 *
 * Files are read and validated upfront so a bad path fails BEFORE we
 * call the create-draft endpoint — otherwise a typo'd path would leave
 * an orphan empty draft in the user's Drafts folder.
 *
 * Each attachment is its own POST — the API does not accept a batched
 * payload at this endpoint. We POST sequentially so a 4xx on file N
 * surfaces with the path that caused it.
 *
 * `runApi` is injected so the caller's auth + error handling apply
 * uniformly; this keeps the helper free of auth concerns.
 */
export async function attachFilesToDraft(runApi, draftId, paths) {
  const results = [];
  for (const p of paths) {
    const payload = buildFileAttachment(p);
    const body = await runApi(`/messages/${encodeURIComponent(draftId)}/attachments`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    results.push({ path: p, name: payload.Name, AttachmentId: body?.Id ?? null });
  }
  return results;
}
