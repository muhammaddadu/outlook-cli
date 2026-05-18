// Unit tests for the learnings store. The module is small and pure —
// these cover the lifecycle (add, list, forget, clear) plus the dedupe
// behaviour.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { mkdtempSync, readFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import {
  addLearning,
  removeLearning,
  clearLearnings,
  loadLearnings,
  learningsFile,
} from '../src/learn.mjs';

function isolatedLearningsFile() {
  const file = join(mkdtempSync(join(tmpdir(), 'outlook-learn-')), 'learnings.md');
  process.env.OUTLOOK_LEARNINGS = file;
  return file;
}

test('loadLearnings returns [] when the file does not exist', () => {
  isolatedLearningsFile();
  assert.deepEqual(loadLearnings(), []);
});

test('addLearning persists an observation; loadLearnings reads it back', () => {
  const file = isolatedLearningsFile();
  const ok = addLearning('Signs off as Sam');
  assert.equal(ok, true);
  const items = loadLearnings();
  assert.equal(items.length, 1);
  assert.match(items[0], /\d{4}-\d{2}-\d{2} \| Signs off as Sam$/);
  // File is human-readable markdown.
  const raw = readFileSync(file, 'utf8');
  assert.match(raw, /^# outlook-cli learnings/);
  assert.match(raw, /- \d{4}-\d{2}-\d{2} \| Signs off as Sam/);
});

test('addLearning de-duplicates near-identical entries (case-insensitive)', () => {
  isolatedLearningsFile();
  assert.equal(addLearning('Prefers terse replies'), true);
  assert.equal(addLearning('Prefers terse replies'), false);
  assert.equal(addLearning('prefers terse replies'), false);
  assert.equal(loadLearnings().length, 1);
});

test('addLearning ignores empty / whitespace-only input', () => {
  isolatedLearningsFile();
  assert.equal(addLearning(''), false);
  assert.equal(addLearning('   '), false);
  assert.deepEqual(loadLearnings(), []);
});

test('removeLearning deletes every line containing the substring', () => {
  isolatedLearningsFile();
  addLearning('Signs off as Sam');
  addLearning('Prefers terse replies');
  addLearning('Boss is alice@example.com');
  const removed = removeLearning('sam'); // matches "Sam" + "alice" doesn't
  assert.equal(removed, 1);
  const items = loadLearnings();
  assert.equal(items.length, 2);
  assert.ok(items.every((i) => !i.toLowerCase().includes(' sam')));
});

test('removeLearning returns 0 when nothing matches', () => {
  isolatedLearningsFile();
  addLearning('Likes coffee');
  assert.equal(removeLearning('xyz-nomatch'), 0);
  assert.equal(loadLearnings().length, 1);
});

test('clearLearnings deletes the file entirely', () => {
  const file = isolatedLearningsFile();
  addLearning('Likes coffee');
  assert.ok(existsSync(file));
  clearLearnings();
  assert.equal(existsSync(file), false);
  assert.deepEqual(loadLearnings(), []);
});

test('learningsFile() respects OUTLOOK_LEARNINGS override', () => {
  const file = isolatedLearningsFile();
  assert.equal(learningsFile(), file);
});
