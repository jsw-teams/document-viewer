import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { utils, write } from 'xlsx';
import { workbookSource, sheetInfo, sheetWindow } from '../src/sheets/model.js';
import { sheetFixture, styledSheetFixture } from './fixtures.mjs';
import { spreadsheetColor } from '../src/sheets/styles.js';
import { cellAddress, columnName, xmlText } from '../src/sheets/xml.js';

test('worksheet source windows retain all rows and columns without building a full cell model', async () => {
  const workbook = utils.book_new();
  const sheet = { A1: { t: 's', v: 'Heading' }, DX5001: { t: 's', v: 'Beyond old limits 中文' }, B2: { t: 'n', v: 0.25, z: '0.00%' }, C2: { t: 'n', v: 42, f: 'SUM(40,2)' }, '!ref': 'A1:DX5001', '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }] };
  utils.book_append_sheet(workbook, sheet, 'Large');
  const source = await workbookSource(write(workbook, { type: 'array', bookType: 'xlsx', bookSST: true, compression: true }), 'xlsx');
  assert.equal(source.Sheets, undefined);
  assert.deepEqual(await sheetInfo(source, 0), { rows: 5001, columns: 128, merges: sheet['!merges'], columnWidth: 64, columnWidths: {}, rowHeight: 20, rowHeights: {}, hiddenRows: [], hiddenColumns: [], columnStyles: {}, rowStyles: {}, defaultStyle: source.styles[0], gridLines: true });
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

test('worksheet geometry preserves stored column widths and rejects unsafe dimensions', async () => {
  for (const format of ['xlsx', 'xls']) {
    const workbook = utils.book_new();
    const sheet = utils.aoa_to_sheet([['Metric', 'Value'], ['Total', 42]]);
    sheet['!cols'] = [{ wpx: 140, MDW: 7 }, { wpx: 240, MDW: 7 }];
    utils.book_append_sheet(workbook, sheet, 'Summary');
    const source = await workbookSource(write(workbook, { type: 'array', bookType: format }), format);
    const info = await sheetInfo(source, 0);
    assert.ok(Math.abs(info.columnWidths[0] - 140) <= 8);
    assert.ok(Math.abs(info.columnWidths[1] - 240) <= 8);
  }
  const zip = await JSZip.loadAsync(sheetFixture());
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><dimension ref="A1:B2"/><cols><col min="1" max="2" width="NaN"/></cols><sheetData/></worksheet>');
  const source = await workbookSource(await zip.generateAsync({ type: 'uint8array' }), 'xlsx');
  await assert.rejects(sheetInfo(source, 0), /column width/);
});

test('Open XML restores theme fonts, tint, borders, rich text, hidden geometry and inherited formats', async () => {
  const source = await workbookSource(await styledSheetFixture(), 'xlsx');
  const info = await sheetInfo(source, 0);
  assert.deepEqual(info.hiddenColumns, [1]);
  assert.deepEqual(info.hiddenRows, [1]);
  assert.equal(info.rowHeight, 24);
  assert.ok(Math.abs(info.rowHeights[0] - 32 * 96 / 72) < 0.001);
  assert.equal(info.gridLines, false);
  const cells = await sheetWindow(source, 0, { startRow: 0, endRow: 4, startColumn: 0, endColumn: 4, columnStyles: info.columnStyles });
  const cell = (row, column) => cells.find(entry => entry.row === row && entry.column === column);
  assert.equal(cell(0, 2).text, '4.50');
  assert.equal(cell(0, 2).style.font, 'Georgia');
  assert.equal(cell(0, 2).style.background, '#70a0cf');
  assert.deepEqual(cell(0, 2).style.borders.bottom, { style: 'double', color: '#ff0000' });
  assert.equal(cell(0, 0).runs[0].style.color, '#ff0000');
  assert.equal(cell(0, 0).runs[0].style.bold, true);
  assert.equal(cell(3, 2).style.bold, true);
  assert.equal(cell(3, 3).style.font, 'Arial');
  assert.equal(cell(4, 0).text, '30');
  assert.equal(cell(0, 4).text, '20');
  assert.equal(cell(1, 4).text, '40');
  const remote = await sheetWindow(source, 0, { startRow: 1, endRow: 1, startColumn: 4, endColumn: 4 });
  assert.equal(remote[0].formula, 'D2*2');
  assert.equal(remote[0].text, '40');
  assert.equal(spreadsheetColor({ rgb: 'FF204060', tint: -0.5 }), '#102030');
});

test('hidden worksheets remain calculation dependencies without exposing hidden tabs or corrupting scoped names', async () => {
  const zip = await JSZip.loadAsync(sheetFixture());
  const workbook = await zip.file('xl/workbook.xml').async('string');
  zip.file('xl/workbook.xml', workbook.replace('name="Details"', 'name="Details" state="veryHidden"'));
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><f>SUM(Details!A1,40)</f></c></row></sheetData></worksheet>');
  zip.file('xl/worksheets/sheet2.xml', '<worksheet><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>50</v></c></row></sheetData></worksheet>');
  const source = await workbookSource(await zip.generateAsync({ type: 'uint8array' }), 'xlsx');
  assert.deepEqual(source.names, ['Summary']);
  assert.equal((await sheetWindow(source, 0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }))[0].text, '90');
});
