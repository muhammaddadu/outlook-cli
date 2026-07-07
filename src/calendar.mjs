// Calendar helpers: time-range parsing, OData filter composition for
// event queries, URL builders for the calendarView and free-busy
// endpoints.
//
// Outlook splits "show me my upcoming meetings" between two endpoints:
//
//   /me/events         — raw events (one row per recurring master series,
//                        no expansion into instances)
//   /me/calendarView   — events expanded into individual instances within
//                        a startDateTime / endDateTime window. This is
//                        what humans want when they say "show me my
//                        agenda."
//
// `agenda` uses calendarView. `events` uses /events for filter-heavy
// queries that don't need expansion.

import { AppError, E } from './errors.mjs';
import { parseCount } from './odata.mjs';

const UNIT_MS = {
  m: 60_000,
  h: 3_600_000,
  d: 86_400_000,
  w: 7 * 86_400_000,
};

/**
 * Parse a calendar-style time argument:
 *
 *   "+7d", "+24h", "+2w"   → future (now + N)
 *   "-1d", "-3h"           → past   (now - N)
 *   "7d" (no sign)         → future (calendar context defaults forward)
 *   "today"                → today 00:00 local
 *   "tomorrow"             → tomorrow 00:00 local
 *   "yesterday"            → yesterday 00:00 local
 *   ISO timestamp          → exact
 */
export function parseEventTime(value) {
  if (typeof value !== 'string') {
    throw new AppError({
      code: E.ARGS,
      message: `Invalid time value: ${value}`,
      hint: 'Use ISO date, "today"/"tomorrow", or +Nd / -Nh form.',
    });
  }

  const lower = value.trim().toLowerCase();
  if (lower === 'today') return startOfDay(new Date());
  if (lower === 'tomorrow') return startOfDay(new Date(Date.now() + 86_400_000));
  if (lower === 'yesterday') return startOfDay(new Date(Date.now() - 86_400_000));
  if (lower === 'now') return new Date();

  const rel = lower.match(/^([+-]?)(\d+)([mhdw])$/);
  if (rel) {
    const sign = rel[1] === '-' ? -1 : 1;
    const n = parseInt(rel[2], 10);
    return new Date(Date.now() + sign * n * UNIT_MS[rel[3]]);
  }

  const d = new Date(value);
  if (Number.isNaN(d.getTime())) {
    throw new AppError({
      code: E.ARGS,
      message: `Invalid time value: ${value}`,
      hint:
        'Use ISO date (e.g. 2026-05-25 or 2026-05-25T10:00:00Z), or a ' +
        'relative like +7d / -1h, or one of today/tomorrow/yesterday/now.',
    });
  }
  return d;
}

function startOfDay(d) {
  const out = new Date(d);
  out.setHours(0, 0, 0, 0);
  return out;
}

/**
 * Resolve --from / --to / --days into a concrete { start, end } window.
 * Defaults: start = now, end = now + 7 days.
 */
export function resolveEventRange({ from, to, days } = {}) {
  const start = from ? parseEventTime(from) : new Date();
  let end;
  if (to) {
    end = parseEventTime(to);
  } else {
    const ndays = days != null ? Number(days) : 7;
    if (!Number.isFinite(ndays)) {
      throw new AppError({
        code: E.ARGS,
        message: `--days must be a number, got: ${days}`,
      });
    }
    end = new Date(start.getTime() + ndays * UNIT_MS.d);
  }
  if (end <= start) {
    throw new AppError({
      code: E.ARGS,
      message: `End (${end.toISOString()}) is not after start (${start.toISOString()}).`,
      hint: 'Pass a larger --to or --days.',
    });
  }
  return { start, end };
}

/**
 * Build the path for /me/calendarView with proper query params.
 * `extra` accepts the same shape as odata.buildQuery (filter/select/top/etc.).
 */
export function calendarViewPath({ start, end }, extra = {}) {
  const params = [
    ['startDateTime', start.toISOString()],
    ['endDateTime', end.toISOString()],
  ];
  if (extra.top != null) params.push(['$top', parseCount(extra.top, '--top')]);
  if (extra.skip != null) params.push(['$skip', parseCount(extra.skip, '--skip')]);
  if (extra.select) params.push(['$select', extra.select]);
  if (extra.filter) params.push(['$filter', extra.filter]);
  if (extra.orderBy) params.push(['$orderby', extra.orderBy]);
  const qs = params.map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&');
  return `/calendarView?${qs}`;
}

/** Escape a single-quoted OData string literal. */
function escapeOData(s) {
  return String(s).replace(/'/g, "''");
}

/**
 * Compose an OData $filter expression for /me/events from friendly flags.
 * Note: time-range filtering for events is handled via calendarView (above)
 * — this is for non-time filters that work on /events itself.
 */
export function buildEventFilter(opts = {}) {
  const clauses = [];

  if (opts.isAllDay !== undefined) {
    clauses.push(`IsAllDay eq ${opts.isAllDay ? 'true' : 'false'}`);
  }
  if (opts.isCancelled !== undefined) {
    clauses.push(`IsCancelled eq ${opts.isCancelled ? 'true' : 'false'}`);
  }
  if (opts.organizer) {
    clauses.push(
      `Organizer/EmailAddress/Address eq '${escapeOData(opts.organizer)}'`,
    );
  }
  if (opts.showAs) {
    const v = opts.showAs;
    const allowed = ['Free', 'Tentative', 'Busy', 'Oof', 'WorkingElsewhere', 'Unknown'];
    if (!allowed.includes(v)) {
      throw new AppError({
        code: E.ARGS,
        message: `Invalid --show-as: ${v}`,
        hint: `Use one of: ${allowed.join(', ')}`,
      });
    }
    clauses.push(`ShowAs eq '${v}'`);
  }
  if (opts.subject) {
    clauses.push(`contains(Subject, '${escapeOData(opts.subject)}')`);
  }
  if (opts.raw) clauses.push(`(${opts.raw})`);

  return clauses.length ? clauses.join(' and ') : null;
}
