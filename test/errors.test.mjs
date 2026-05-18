// Unit tests for the error model: AppError shape + exit-code mapping.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { AppError, E, EXIT, exitCodeFor } from '../src/errors.mjs';

test('AppError carries code, message, hint, cause', () => {
  const cause = new Error('underlying');
  const err = new AppError({
    code: E.HTTP,
    message: 'API said no',
    hint: 'fix this',
    cause,
  });
  assert.equal(err.name, 'AppError');
  assert.equal(err.code, E.HTTP);
  assert.equal(err.message, 'API said no');
  assert.equal(err.hint, 'fix this');
  assert.equal(err.cause, cause);
});

test('AppError without hint or cause still constructs', () => {
  const err = new AppError({ code: E.UNEXPECTED, message: 'oops' });
  assert.equal(err.hint, undefined);
  assert.equal(err.cause, undefined);
});

test('exitCodeFor maps every E.* to a stable exit code', () => {
  const cases = [
    [E.AUTH_REQUIRED, EXIT.AUTH],
    [E.AUTH_BLOCKED, EXIT.AUTH],
    [E.HTTP, EXIT.HTTP],
    [E.ARGS, EXIT.USAGE],
    [E.UNEXPECTED, EXIT.GENERAL],
  ];
  for (const [code, expected] of cases) {
    const err = new AppError({ code, message: 'x' });
    assert.equal(exitCodeFor(err), expected, `code=${code}`);
  }
});

test('exitCodeFor returns GENERAL for non-AppError instances', () => {
  assert.equal(exitCodeFor(new Error('plain')), EXIT.GENERAL);
  assert.equal(exitCodeFor('string error'), EXIT.GENERAL);
  assert.equal(exitCodeFor(undefined), EXIT.GENERAL);
});

test('E.* and EXIT.* are frozen (no accidental mutation)', () => {
  assert.throws(() => {
    E.HTTP = 'mutated';
  }, /Cannot assign|read.only/);
  assert.throws(() => {
    EXIT.OK = 99;
  }, /Cannot assign|read.only/);
});
