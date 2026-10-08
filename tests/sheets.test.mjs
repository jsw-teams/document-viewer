import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { utils, write } from 'xlsx';
import { workbookSource, sheetInfo, sheetWindow } from '../src/sheets/model.js';
import { sheetFixture } from './fixtures.mjs';
import { cellAddress, columnName, xmlText } from '../src/sheets/xml.js';

test('worksheet source windows retain all rows and columns without building a full cell model', async () => {
  const workbook = utils.book_new();
  const sheet = { A1: { t: 's', v: 'Heading' }, DX5001: { t: 's', v: 'Beyond old limits 中文' }, B2: { t: 'n', v: 0.25, z: '0.00%' }, C2: { t: 'n', v: 42, f: 'SUM(40,2)' }, '!ref': 'A1:DX5001', '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }] };
  utils.book_append_sheet(workbook, sheet, 'Large');
  const source = await workbookSource(write(workbook, { type: 'array', bookType: 'xlsx', bookSST: true, compression: true }), 'xlsx');
  assert.equal(source.Sheets, undefined);
  assert.deepEqual(await sheetInfo(source, 0), { rows: 5001, columns: 128, merges: sheet['!merges'] });
  const tail = await sheetWindow(source, 0, { startRow: 5000, endRow: 5000, startColumn: 127, endColumn: 127 });
  assert.equal(tail.length, 1);
  assert.equal(tail[0].text, 'Beyond old limits 中文');
  const head = await sheetWindow(source, 0, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 });
  assert.equal(head.find(cell => cell.column === 1).text, '25.00%');
  assert.equal(head.find(cell => cell.column === 2).formula, 'SUM(40,2)');
  assert.equal(source.cells, undefined);
});

test('wide rows stream individual cells, including omitted references, without buffering entire rows', async () => {
  const zip = await JSZip.loadAsync(sheetFixture());
  const cells = Array.from({ length: 512 }, (_, index) => '<c t="inlineStr"><is><t>' + (index === 511 ? 'Last wide cell' : 'x'.repeat(32767)) + '</t></is></c>').join('');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><dimension ref="A1:SR1"/><sheetData><row r="1">' + cells + '</row></sheetData></worksheet>');
  const source = await workbookSource(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), 'xlsx');
  const result = await sheetWindow(source, 0, { startRow: 0, endRow: 0, startColumn: 511, endColumn: 511 });
  assert.equal(result.length, 1);
  assert.equal(result[0].text, 'Last wide cell');
});

test('legacy worksheets are reparsed in a disposable request with no retained workbook', async () => {
  const source = await workbookSource(sheetFixture('xls'), 'xls');
  assert.deepEqual(source.names, ['Summary', 'Details']);
  assert.equal((await sheetWindow(source, 1, { startRow: 0, endRow: 10, startColumn: 0, endColumn: 5 }))[0].text, 'Second sheet');
  assert.equal(source.Sheets, undefined);
});

test('streamed XML rejects missing parts, unsafe relationships and cancelled requests', async () => {
  const source = await workbookSource(sheetFixture(), 'xlsx');
  await assert.rejects(sheetWindow(source, 0, { startRow: 0, endRow: 5, startColumn: 0, endColumn: 5 }, () => true), /Superseded/);
  const zip = await JSZip.loadAsync(sheetFixture());
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="../bad.xml"/></Relationships>');
  await assert.rejects(workbookSource(await zip.generateAsync({ type: 'uint8array' }), 'xlsx'), /Unsafe/);
  assert.deepEqual(cellAddress('DX5001'), { row: 5000, column: 127 });
  assert.equal(columnName(16383), 'XFD');
  assert.throws(() => cellAddress('XFE1'));
  assert.equal(xmlText('&lt;script&gt;&#x4e2d;&#25991;&amp;'), '<script>中文&');
});
