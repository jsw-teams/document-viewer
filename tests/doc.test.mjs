import test from 'node:test';
import assert from 'node:assert/strict';
import CFB from 'cfb';
import { docPageLayout } from '../src/doc/layout.js';

function fixture({ invalid = false } = {}) {
  const word = Buffer.alloc(1024);
  word.writeUInt16LE(14, 32);
  word.writeUInt16LE(22, 62);
  word.writeUInt16LE(93, 152);
  word.writeUInt32LE(20, 206);
  const table = Buffer.alloc(20);
  table.writeUInt32LE(10, 4);
  table.writeInt32LE(invalid ? 1023 : 920, 10);
  const fields = [[0xb01f, 15840], [0xb020, 12240], [0xb021, 1800], [0xb022, 1200], [0x9023, -900], [0x9024, 1440]];
  word.writeUInt16LE(fields.length * 4, 920);
  fields.forEach(([operation, value], index) => { word.writeUInt16LE(operation, 922 + index * 4); word.writeInt16LE(value, 924 + index * 4); });
  const file = CFB.utils.cfb_new();
  CFB.utils.cfb_add(file, 'WordDocument', word);
  CFB.utils.cfb_add(file, '0Table', table);
  return CFB.write(file, { type: 'buffer' });
}

test('MS-DOC paper geometry comes from the FIB section directory and bounded SEPX properties', () => {
  assert.deepEqual(docPageLayout(fixture()), { width: 15840, height: 12240, left: 1800, right: 1200, top: 900, bottom: 1440 });
  assert.throws(() => docPageLayout(fixture({ invalid: true })), /section properties/);
  assert.throws(() => docPageLayout(new Uint8Array(64)));
});
