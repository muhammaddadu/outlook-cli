// OData query helpers for the Outlook REST API.
//
// The CLI lets users build filters via friendly flags (`--unread`,
// `--from`, `--since 7d`, …) rather than forcing them to write OData by
// hand. This module turns those flags into the wire-format query string.
//
// All values are quoted/escaped here, so command code never has to think
// about OData syntax.

import { AppError, E } from './errors.mjs';

/** Map of friendly folder aliases → Outlook well-known folder names. */
const WELL_KNOWN_FOLDERS = {
  inbox: 'Inbox',
  drafts: 'Drafts',
  sent: 'SentItems',
  sentitems: 'SentItems',
  deleted: 'DeletedItems',
  deleteditems: 'DeletedItems',
  junk: 'JunkEmail',
  junkemail: 'JunkEmail',
  outbox: 'Outbox',
  archive: 'Archive',
};

/**
 * Resolve a user-supplied folder argument to a REST path segment usable in
 * `/mailfolders/{id}/messages`.
 *
 * - Well-known names (case-insensitive, with spaces ignored) → canonical
 *   well-known string ("Inbox", "SentItems", …).
 * - Long opaque ids (≥ 30 chars of base64-ish) → passed through.
 * - Anything else throws E_ARGS; the user should pass `outlook folders`
 *   output and pick an Id manually.
 */
export function resolveFolder(input) {
  const normalized = input.toLowerCase().replace(/\s+/g, '');
  if (WELL_KNOWN_FOLDERS[normalized]) return WELL_KNOWN_FOLDERS[normalized];
  if (/^[A-Za-z0-9_=+/-]{30,}$/.test(input)) return input;
  throw new AppError({
    code: E.ARGS,
    message: `Unknown folder: ${input}`,
    hint:
      'Use Inbox / Drafts / Sent / Deleted / Junk / Outbox / Archive, or pass a folder Id from `outlook folders`.',
  });
}

/**
 * Parse a `--since`/`--until` value into an ISO timestamp string suitable
 * for use in an OData filter expression.
 *
 * Accepts:
 *   - Relative: "5m", "24h", "7d", "2w" (m=minutes, h=hours, d=days, w=weeks)
 *   - Absolute: any string `new Date()` can parse (ISO recommended).
 */
export function parseTimeArg(value) {
  const rel = value.match(/^(\d+)([mhdw])$/i);
  if (rel) {
    const n = parseInt(rel[1], 10);
    const unitMs = { m: 60_000, h: 3_600_000, d: 86_400_000, w: 7 * 86_400_000 }[
      rel[2].toLowerCase()
    ];
    return new Date(Date.now() - n * unitMs).toISOString();
  }
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) {
    throw new AppError({
      code: E.ARGS,
      message: `Invalid time value: ${value}`,
      hint: 'Use "7d" / "24h" / "2w" for relative, or an ISO date like "2026-01-15".',
    });
  }
  return d.toISOString();
}

/** Escape a value for use inside a single-quoted OData string literal. */
function escapeODataString(s) {
  return String(s).replace(/'/g, "''");
}

/**
 * Validate a count-like CLI value (`--top`, `--skip`, `--interval`) into a
 * non-negative integer. Catching this locally turns a cryptic server 400
 * into an immediate E_ARGS with the flag name in the message.
 */
export function parseCount(value, flag) {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 0) {
    throw new AppError({
      code: E.ARGS,
      message: `Invalid ${flag}: ${value}`,
      hint: `${flag} must be a non-negative integer.`,
    });
  }
  return n;
}

/**
 * Compose an OData `$filter` expression from a set of friendly flags.
 * Returns `null` when no filter clauses apply.
 *
 * @param {object} opts
 * @param {boolean} [opts.unread]
 * @param {string}  [opts.from]              email address
 * @param {string}  [opts.to]                email address — matches any ToRecipients entry
 * @param {string}  [opts.since]             time arg (relative or ISO)
 * @param {string}  [opts.until]             time arg
 * @param {boolean} [opts.hasAttachments]
 * @param {string}  [opts.importance]        Low | Normal | High
 * @param {string}  [opts.subject]           substring match against Subject
 * @param {string}  [opts.raw]               raw OData expression, ANDed in last
 */
export function buildFilter(opts = {}) {
  const clauses = [];

  if (opts.unread) clauses.push('IsRead eq false');
  if (opts.hasAttachments) clauses.push('HasAttachments eq true');

  if (opts.from) {
    clauses.push(`From/EmailAddress/Address eq '${escapeODataString(opts.from)}'`);
  }

  if (opts.to) {
    // ToRecipients is a collection; use `any` to match within it.
    clauses.push(
      `ToRecipients/any(r: r/EmailAddress/Address eq '${escapeODataString(opts.to)}')`,
    );
  }

  if (opts.since) clauses.push(`ReceivedDateTime gt ${parseTimeArg(opts.since)}`);
  if (opts.until) clauses.push(`ReceivedDateTime lt ${parseTimeArg(opts.until)}`);

  if (opts.importance) {
    const v = opts.importance;
    if (!['Low', 'Normal', 'High'].includes(v)) {
      throw new AppError({
        code: E.ARGS,
        message: `Invalid --importance: ${v}`,
        hint: 'Use one of: Low, Normal, High',
      });
    }
    clauses.push(`Importance eq '${v}'`);
  }

  if (opts.subject) {
    clauses.push(`contains(Subject, '${escapeODataString(opts.subject)}')`);
  }

  if (opts.raw) clauses.push(`(${opts.raw})`);

  return clauses.length ? clauses.join(' and ') : null;
}

/**
 * Build a complete query-string suffix for a messages collection.
 * Returns a string starting with `?` (or empty if no params).
 *
 * Note: `$search` and `$orderby` are mutually exclusive in Outlook REST —
 * searched results are scored, not date-sorted. If both are supplied, we
 * drop `$orderby` to match the server's behaviour rather than silently
 * 400-ing.
 */
export function buildQuery({
  top,
  skip,
  filter,
  orderBy,
  select,
  search,
} = {}) {
  // We construct the query string by hand rather than using URLSearchParams
  // because that helper percent-encodes `$` to `%24` — semantically correct
  // but it makes URLs unreadable and diverges from what OWA itself emits.
  // Only the values need encoding; the OData keys stay literal.
  const parts = [];
  const add = (key, value) => parts.push(`${key}=${encodeURIComponent(value)}`);

  if (top !== undefined && top !== null) add('$top', parseCount(top, '--top'));
  if (skip !== undefined && skip !== null) add('$skip', parseCount(skip, '--skip'));
  if (filter) add('$filter', filter);
  if (select) add('$select', select);
  if (search) {
    add('$search', `"${search}"`);
  } else if (orderBy) {
    add('$orderby', orderBy);
  }
  return parts.length ? `?${parts.join('&')}` : '';
}
