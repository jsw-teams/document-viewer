import test from 'node:test';
import assert from 'node:assert/strict';
import { columnLayout } from '../src/sheets/columns.js';

test('display widths preserve source geometry and hidden columns', () => {
  const source = { columns: 3, columnWidths: [140, 80, 240], hiddenColumns: [1] };
  const before = structuredClone(source);
  const preference = { widths: new Map(), stretch: false };
  assert.deepEqual(columnLayout(source, preference, 900), [0, 140, 140, 380]);
  preference.stretch = true;
  assert.deepEqual(columnLayout(source, preference, 760), [0, 280, 280, 760]);
  assert.deepEqual(columnLayout(source, preference, 200), [0, 140, 140, 380]);
  assert.deepEqual(source, before);
});

test('display overrides are bounded, finite and never unhide source columns', () => {
  const source = { columns: 4, columnWidth: 64, hiddenColumns: [1] };
  const preference = { widths: new Map([[0, -1], [1, 100], [2, 10000], [3, NaN]]), stretch: false };
  assert.deepEqual(columnLayout(source, preference), [0, 24, 24, 2424, 2488]);
});
