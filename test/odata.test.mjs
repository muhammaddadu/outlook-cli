// Unit tests for the OData filter / query builders.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  buildFilter,
  buildQuery,
  parseTimeArg,
  resolveFolder,
} from '../src/odata.mjs';
import { AppError, E } from '../src/errors.mjs';

// ---------------------------------------------------------------------------
// buildFilter

test('buildFilter returns null when nothing is set', () => {
  assert.equal(buildFilter({}), null);
  assert.equal(buildFilter(), null);
});

test('buildFilter handles each flag individually', () => {
  assert.equal(buildFilter({ unread: true }), 'IsRead eq false');
  assert.equal(buildFilter({ hasAttachments: true }), 'HasAttachments eq true');
  assert.equal(
    buildFilter({ from: 'a@b.com' }),
    "From/EmailAddress/Address eq 'a@b.com'",
  );
  assert.equal(
    buildFilter({ to: 'a@b.com' }),
    "ToRecipients/any(r: r/EmailAddress/Address eq 'a@b.com')",
  );
  assert.equal(buildFilter({ importance: 'High' }), "Importance eq 'High'");
  assert.equal(
    buildFilter({ subject: 'deploy' }),
    "contains(Subject, 'deploy')",
  );
});

test('buildFilter combines multiple flags with AND', () => {
  const expr = buildFilter({
    unread: true,
    from: 'will@example.com',
    hasAttachments: true,
  });
  assert.match(expr, /IsRead eq false/);
  assert.match(expr, /From\/EmailAddress\/Address eq 'will@example\.com'/);
  assert.match(expr, /HasAttachments eq true/);
  // Three clauses → two " and "s.
  assert.equal((expr.match(/ and /g) || []).length, 2);
});

test("buildFilter escapes single quotes in string values", () => {
  const expr = buildFilter({ from: "ev'il@example.com" });
  assert.equal(expr, "From/EmailAddress/Address eq 'ev''il@example.com'");
});

test('buildFilter wraps raw expression in parens and ANDs it last', () => {
  const expr = buildFilter({
    unread: true,
    raw: 'Categories/any(c: c eq \'red\')',
  });
  assert.equal(
    expr,
    "IsRead eq false and (Categories/any(c: c eq 'red'))",
  );
});

test('buildFilter rejects invalid importance', () => {
  assert.throws(
    () => buildFilter({ importance: 'urgent' }),
    (err) => err instanceof AppError && err.code === E.ARGS,
  );
});

test('buildFilter accepts since/until and emits ISO timestamps', () => {
  const expr = buildFilter({ since: '2026-01-15', until: '2026-01-20' });
  assert.match(expr, /ReceivedDateTime gt 2026-01-15T/);
  assert.match(expr, /ReceivedDateTime lt 2026-01-20T/);
});

// ---------------------------------------------------------------------------
// parseTimeArg

test('parseTimeArg understands minutes/hours/days/weeks', () => {
  const now = Date.now();
  const oneHourAgo = new Date(parseTimeArg('1h')).getTime();
  assert.ok(Math.abs(now - oneHourAgo - 3600_000) < 2000); // ±2s tolerance

  const oneDayAgo = new Date(parseTimeArg('1d')).getTime();
  assert.ok(Math.abs(now - oneDayAgo - 86_400_000) < 2000);

  const oneWeekAgo = new Date(parseTimeArg('1w')).getTime();
  assert.ok(Math.abs(now - oneWeekAgo - 7 * 86_400_000) < 2000);
});

test('parseTimeArg accepts ISO dates', () => {
  const out = parseTimeArg('2026-05-18');
  assert.match(out, /^2026-05-18T/);
});

test('parseTimeArg throws E_ARGS on unparseable input', () => {
  assert.throws(
    () => parseTimeArg('tomorrow'),
    (err) => err instanceof AppError && err.code === E.ARGS,
  );
});

// ---------------------------------------------------------------------------
// resolveFolder

test('resolveFolder maps friendly aliases to well-known folders', () => {
  assert.equal(resolveFolder('inbox'), 'Inbox');
  assert.equal(resolveFolder('Inbox'), 'Inbox');
  assert.equal(resolveFolder('Sent Items'), 'SentItems');
  assert.equal(resolveFolder('sent'), 'SentItems');
  assert.equal(resolveFolder('JUNK'), 'JunkEmail');
  assert.equal(resolveFolder('Junk Email'), 'JunkEmail');
  assert.equal(resolveFolder('drafts'), 'Drafts');
  assert.equal(resolveFolder('Deleted Items'), 'DeletedItems');
});

test('resolveFolder passes through long opaque ids unchanged', () => {
  const id = 'AAMkADQxZDQ1ODU0LTQ0ODUtNDVjOS04YTFiLTllYzc5ODcwNzkyMQAuAAAAAA';
  assert.equal(resolveFolder(id), id);
});

test('resolveFolder throws E_ARGS for unknown short names', () => {
  assert.throws(
    () => resolveFolder('Project Alpha'),
    (err) => err instanceof AppError && err.code === E.ARGS,
  );
});

// ---------------------------------------------------------------------------
// buildQuery

test('buildQuery returns "" when no params set', () => {
  assert.equal(buildQuery(), '');
  assert.equal(buildQuery({}), '');
});

test('buildQuery emits $top, $skip, $select, $filter, $orderby with literal $', () => {
  const qs = buildQuery({
    top: 25,
    skip: 50,
    filter: 'IsRead eq false',
    orderBy: 'ReceivedDateTime desc',
    select: 'Id,Subject',
  });
  assert.match(qs, /\$top=25/);
  assert.match(qs, /\$skip=50/);
  // encodeURIComponent uses %20 for spaces (not +).
  assert.match(qs, /\$filter=IsRead%20eq%20false/);
  assert.match(qs, /\$orderby=ReceivedDateTime%20desc/);
  assert.match(qs, /\$select=Id%2CSubject/);
  assert.ok(qs.startsWith('?'));
});

test('buildQuery drops $orderby when $search is set (server-side rule)', () => {
  const qs = buildQuery({
    search: 'quarterly review',
    orderBy: 'ReceivedDateTime desc',
    top: 10,
  });
  assert.match(qs, /\$search=/);
  assert.doesNotMatch(qs, /\$orderby/);
});

test('buildQuery wraps $search value in literal quotes (KQL convention)', () => {
  const qs = buildQuery({ search: 'quarterly review' });
  assert.match(qs, /\$search=%22quarterly%20review%22/);
});
