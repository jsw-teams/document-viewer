import test from 'node:test';
import assert from 'node:assert/strict';
import { parsePpt } from '../src/ppt/parser.js';
import { documentFormat, documentUrl } from '../src/formats.js';
import { validateOfficeZip } from '../src/zip-limits.js';
import { legacyPpt, wordFixture } from './fixtures.mjs';

test('legacy PPT follows active slide order and resolves positioned outline text without a vendor engine', () => {
  const model = parsePpt(legacyPpt());
  assert.equal(model.width, 5760);
  assert.deepEqual(model.slides.map(slide => slide.id), [300, 200]);
  assert.equal(model.slides[0].texts[0].text, 'First slide: 中文');
  assert.equal(model.slides[0].shapes[0].texts[0].text, 'First slide: 中文');
  assert.equal(model.slides[1].texts[0].text, 'Second slide');
  const empty = parsePpt(legacyPpt({ emptyTextBefore: true }));
  assert.equal(empty.slides[0].texts[0].text, '');
  assert.equal(empty.slides[0].shapes[0].texts[0].text, 'First slide: 中文');
});

test('legacy PPT rejects invalid references, edit cycles, encrypted files and truncated containers', () => {
  for (const option of ['cycle', 'brokenReference', 'encrypted']) assert.throws(() => parsePpt(legacyPpt({ [option]: true })));
  assert.throws(() => parsePpt(new Uint8Array(20)), /size/);
  assert.throws(() => parsePpt(new Uint8Array(512)));
});

test('MS-PPT client rectangles differ from MS-ODRAW child anchors and local text is not lost', () => {
  for (const option of [{}, { largeAnchor: true }, { childAnchor: true }]) {
    const shape = parsePpt(legacyPpt(option)).slides[0].shapes[0];
    assert.deepEqual([shape.left, shape.top, shape.width, shape.height], [256, 128, 3744, 672]);
  }
  for (const missingAnchor of [false, true]) {
    const slide = parsePpt(legacyPpt({ inlineText: true, missingAnchor })).slides[0];
    assert.deepEqual(slide.texts.map(item => item.text), ['First slide: 中文', 'Local textbox 中文']);
  }
  for (const outlineIndex of [-1, 99]) assert.throws(() => parsePpt(legacyPpt({ outlineIndex })), /outline text reference/);
});

test('direct URLs support query strings and extensionless shares but reject unsafe sources', () => {
  assert.equal(documentFormat('/report.XLSX?download=1'), 'xlsx');
  assert.equal(documentFormat('/download/123', 'pdf'), 'pdf');
  assert.equal(documentFormat('/report.xlxs'), null);
  for (const source of ['javascript:alert(1)', '//files.example/report.pdf', 'https://user:secret@files.example/a.pdf', 'https://files.example/ a.pdf', 'http://files.example/a.pdf']) assert.throws(() => documentUrl(source));
  assert.equal(documentUrl('/a.pdf').pathname, '/a.pdf');
});

test('Office ZIP safety checks accept a document and reject expanded-size bombs', async () => {
  const data = await wordFixture();
  validateOfficeZip(data);
  const inflated = Buffer.from(data);
  const offset = inflated.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  inflated.writeUInt32LE(0x7fffffff, offset + 24);
  assert.throws(() => validateOfficeZip(inflated), /oversized/);
});
