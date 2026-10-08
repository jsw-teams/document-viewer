import test from 'node:test';
import assert from 'node:assert/strict';
import { symbolBullet } from '../src/symbols.js';
import { styledText } from '../src/ppt/text.js';

test('Symbol bullets map to Unicode while unrelated fonts and private-use characters stay intact', () => {
  assert.equal(symbolBullet('\uf0b7', 'Symbol'), '\u2022');
  assert.equal(symbolBullet('\uf0b7', '"symbol"'), '\u2022');
  assert.equal(symbolBullet('\uf0b7', 'Unrelated'), '\uf0b7');
  assert.equal(symbolBullet('\uf0b7', undefined), '\uf0b7');
  assert.equal(symbolBullet('\uf020', 'Symbol'), '\uf020');
  assert.equal(symbolBullet('\u2022', 'Symbol'), '\u2022');
});

test('PPT bullet percentages use authored paragraph text size rather than the renderer default', () => {
  const create = () => ({ style: {}, children: [], append(child) { this.children.push(child); } });
  const document = { createElement: create, createElementNS: create, createTextNode: text => ({ textContent: text }) };
  const item = {
    text: 'Example',
    paragraphs: [{ start: 0, end: 8, bullet: true, bulletChar: '\uf0b7', bulletFont: 0, bulletSize: 100 }],
    runs: [{ start: 0, end: 8, size: 16 }]
  };
  const paragraph = styledText(document, item, { fonts: ['Symbol'], size: 24 }).children[0];
  assert.equal(paragraph.style.fontSize, '128px');
  assert.equal(paragraph.children[0].textContent, '\u2022\u00a0');
  assert.equal(paragraph.children[0].style.fontFamily, 'serif');
  assert.equal(paragraph.children[0].style.fontSize, '100%');
  assert.equal(paragraph.children[1].style.fontSize, '128px');
});
