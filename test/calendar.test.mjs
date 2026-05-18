// Unit tests for the calendar helpers.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  parseEventTime,
  resolveEventRange,
  calendarViewPath,
  buildEventFilter,
} from '../src/calendar.mjs';
import { AppError, E } from '../src/errors.mjs';

// ---------------------------------------------------------------------------
// parseEventTime

test('parseEventTime accepts ISO dates', () => {
  const out = parseEventTime('2026-05-20');
  assert.match(out.toISOString(), /^2026-05-20T/);
});

test('parseEventTime understands today/tomorrow/yesterday', () => {
  const today = parseEventTime('today');
  const tomorrow = parseEventTime('tomorrow');
  const yesterday = parseEventTime('yesterday');
  assert.equal(today.getHours(), 0);
  assert.equal(tomorrow.getHours(), 0);
  assert.equal(yesterday.getHours(), 0);
  // Tomorrow is one day after today.
  assert.equal(tomorrow.getTime() - today.getTime(), 86_400_000);
  // Yesterday is one day before today.
  assert.equal(today.getTime() - yesterday.getTime(), 86_400_000);
});

test('parseEventTime handles unsigned relative as future', () => {
  const before = Date.now();
  const out = parseEventTime('7d');
  const expected = before + 7 * 86_400_000;
  // Allow ~2 seconds slack for slow CI machines.
  assert.ok(Math.abs(out.getTime() - expected) < 2000);
});

test('parseEventTime handles +N as future, -N as past', () => {
  const future = parseEventTime('+3d').getTime();
  const past = parseEventTime('-3d').getTime();
  assert.ok(future > Date.now());
  assert.ok(past < Date.now());
  assert.ok(Math.abs(future - past - 6 * 86_400_000) < 2000);
});

test('parseEventTime accepts "now"', () => {
  const out = parseEventTime('now');
  assert.ok(Math.abs(out.getTime() - Date.now()) < 200);
});

test('parseEventTime throws E_ARGS on garbage', () => {
  assert.throws(
    () => parseEventTime('next quarter'),
    (e) => e instanceof AppError && e.code === E.ARGS,
  );
});

// ---------------------------------------------------------------------------
// resolveEventRange

test('resolveEventRange defaults to now → now+7d', () => {
  const { start, end } = resolveEventRange();
  assert.ok(Math.abs(start.getTime() - Date.now()) < 200);
  assert.ok(end - start === 7 * 86_400_000);
});

test('resolveEventRange honours --days', () => {
  const { start, end } = resolveEventRange({ days: 3 });
  assert.equal(end - start, 3 * 86_400_000);
});

test('resolveEventRange accepts explicit --to', () => {
  const { start, end } = resolveEventRange({ to: '+2d' });
  assert.ok(end - start >= 2 * 86_400_000 - 1000);
  assert.ok(end - start <= 2 * 86_400_000 + 1000);
});

test('resolveEventRange throws when end is not after start', () => {
  assert.throws(
    () => resolveEventRange({ from: 'tomorrow', to: 'today' }),
    (e) => e instanceof AppError && e.code === E.ARGS,
  );
});

// ---------------------------------------------------------------------------
// calendarViewPath

test('calendarViewPath emits startDateTime/endDateTime + $select', () => {
  const start = new Date('2026-05-18T00:00:00Z');
  const end = new Date('2026-05-25T00:00:00Z');
  const path = calendarViewPath({ start, end }, { top: 50, select: 'Id,Subject' });
  assert.match(path, /^\/calendarView\?/);
  assert.match(path, /startDateTime=2026-05-18T00%3A00%3A00\.000Z/);
  assert.match(path, /endDateTime=2026-05-25T00%3A00%3A00\.000Z/);
  assert.match(path, /\$top=50/);
  assert.match(path, /\$select=Id%2CSubject/);
});

// ---------------------------------------------------------------------------
// buildEventFilter

test('buildEventFilter returns null when nothing set', () => {
  assert.equal(buildEventFilter(), null);
});

test('buildEventFilter composes individual flags', () => {
  assert.equal(buildEventFilter({ isAllDay: true }), 'IsAllDay eq true');
  assert.equal(buildEventFilter({ isCancelled: false }), 'IsCancelled eq false');
  assert.equal(
    buildEventFilter({ organizer: 'a@b.com' }),
    "Organizer/EmailAddress/Address eq 'a@b.com'",
  );
  assert.equal(buildEventFilter({ showAs: 'Busy' }), "ShowAs eq 'Busy'");
  assert.equal(buildEventFilter({ subject: 'standup' }), "contains(Subject, 'standup')");
});

test('buildEventFilter rejects invalid showAs', () => {
  assert.throws(
    () => buildEventFilter({ showAs: 'AwayFromKeyboard' }),
    (e) => e instanceof AppError && e.code === E.ARGS,
  );
});

test('buildEventFilter ANDs multiple clauses', () => {
  const expr = buildEventFilter({
    isAllDay: false,
    organizer: 'boss@example.com',
    subject: 'review',
  });
  assert.match(expr, /IsAllDay eq false/);
  assert.match(expr, /Organizer\/EmailAddress\/Address eq 'boss@example\.com'/);
  assert.match(expr, /contains\(Subject, 'review'\)/);
  assert.equal((expr.match(/ and /g) || []).length, 2);
});

test('buildEventFilter wraps raw expression in parens', () => {
  const expr = buildEventFilter({
    isAllDay: true,
    raw: "Categories/any(c: c eq 'red')",
  });
  assert.equal(
    expr,
    "IsAllDay eq true and (Categories/any(c: c eq 'red'))",
  );
});
