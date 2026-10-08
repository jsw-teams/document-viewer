import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { resolve, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import { utils, write } from 'xlsx';
import { pdfFixture, wordFixture, wordLayoutFixture, wordBulletFixture, sheetFixture, styledSheetFixture, slidesFixture, legacyPpt } from './fixtures.mjs';
import JSZip from 'jszip';

const output = fileURLToPath(new URL('../dist/', import.meta.url));

async function fixtureServer(provided = new Map()) {
  const fixtures = new Map([
    ['/report.pdf', pdfFixture()], ['/report.docx', await wordFixture()],
    ['/report.xlsx', sheetFixture()], ['/report.xls', sheetFixture('xls')],
    ['/report.pptx', await slidesFixture()], ['/report.ppt', legacyPpt()], ['/thin.ppt', legacyPpt({ thinAnchor: true })], ['/fit.ppt', legacyPpt({ thinAnchor: true, fitFlags: 0x40004 })],
    ['/download/123', sheetFixture()], ['/invalid.pdf', Buffer.from('Not a PDF')]
  ]);
  for (const [path, bytes] of provided) fixtures.set(path, bytes);
  const requests = [];
  const ranges = [];
  const server = createServer(async (request, response) => {
    const path = new URL(request.url, 'http://local.invalid').pathname;
    requests.push(path);
    response.setHeader('Access-Control-Allow-Origin', '*');
    response.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data: blob:; worker-src 'self' blob:; connect-src 'self'; frame-src 'self'; object-src 'none'");
    if (path === '/') {
      response.setHeader('Content-Type', 'text/html');
      response.end('<!doctype html><html lang="zh-CN"><head><link rel="stylesheet" href="/styles.css"></head><body><div id="preview"></div><script type="module" src="/fixture.js"></script></body></html>');
      return;
    }
    if (path === '/fixture.js') {
      response.setHeader('Content-Type', 'text/javascript');
      response.end('import {mountDocument} from "/index.js"; window.mount = options => { window.viewer?.destroy(); window.previewErrors=[]; window.viewer=mountDocument(document.querySelector("#preview"),{autoOpen:false,locale:"zh-CN",title:"Document",...options,onError:error=>window.previewErrors.push(error.message)}); }; window.ready=true;');
      return;
    }
    if (path === '/axe.js') {
      response.setHeader('Content-Type', 'text/javascript');
      response.end(await readFile(new URL('../node_modules/axe-core/axe.min.js', import.meta.url)));
      return;
    }
    if (path === '/safe-html.js') {
      response.setHeader('Content-Type', 'text/javascript');
      response.end(await readFile(new URL('../src/safe-html.js', import.meta.url)));
      return;
    }
    if (fixtures.has(path)) {
      const bytes = fixtures.get(path);
      const range = path === '/ranged.pdf' && request.headers.range?.match(/^bytes=(\d+)-(\d+)$/);
      if (range) {
        const begin = Number(range[1]);
        const end = Math.min(bytes.length - 1, Number(range[2]));
        ranges.push({ begin, end });
        response.writeHead(206, { 'Content-Range': 'bytes ' + begin + '-' + end + '/' + bytes.length, 'Content-Length': end - begin + 1, ETag: '"fixture"' });
        response.end(bytes.subarray(begin, end + 1));
      } else response.end(bytes);
      return;
    }
    try {
      const file = resolve(output, '.' + path);
      if (!file.startsWith(output)) throw new Error('Invalid path');
      const types = { '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.wasm': 'application/wasm' };
      response.setHeader('Content-Type', types[extname(path)] || 'application/octet-stream');
      response.end(await readFile(file));
    } catch { response.statusCode = 404; response.end('Not found'); }
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  return { server, requests, ranges, url: 'http://127.0.0.1:' + server.address().port };
}

test('two-column worksheets retain every cell without centered-page clipping and fit real file widths', { timeout: 60000 }, async () => {
  const workbook = utils.book_new();
  const sheet = utils.aoa_to_sheet([['Metric', 'Value'], ['Last row', 'Last column']]);
  sheet['!cols'] = [{ wpx: 140, MDW: 7 }, { wpx: 240, MDW: 7 }];
  utils.book_append_sheet(workbook, sheet, 'Summary');
  const fixtures = new Map(['xlsx', 'xls'].map(format => ['/summary.' + format, write(workbook, { type: 'buffer', bookType: format })]));
  const { server, url } = await fixtureServer(fixtures);
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    for (const width of [320, 1280]) {
      const page = await browser.newPage({ viewport: { width, height: 900 } });
      await page.goto(url);
      await page.waitForFunction(() => window.ready);
      for (const format of ['xlsx', 'xls']) {
        await page.evaluate(src => window.mount({ src, autoOpen: true }), '/summary.' + format);
        const frame = page.frameLocator('iframe');
        await frame.getByText('Last column', { exact: true }).waitFor();
        const geometry = await frame.getByRole('grid').evaluate(table => {
          const page = table.closest('[data-document-page]');
          return { clipped: table.getBoundingClientRect().bottom > page.getBoundingClientRect().bottom, left: page.getBoundingClientRect().left, columns: [...table.querySelectorAll('col')].map(column => column.getBoundingClientRect().width) };
        });
        assert.equal(geometry.clipped, false);
        assert.ok(geometry.left >= 0);
        assert.ok(Math.abs(geometry.columns[1] - 140) <= 2);
        assert.ok(Math.abs(geometry.columns[2] - 240) <= 2);
        await page.getByRole('button', { name: '适合宽度', exact: true }).click();
        assert.ok(await frame.getByRole('grid').evaluate(table => table.getBoundingClientRect().right <= innerWidth + 1));
        assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
      }
      await page.close();
    }
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('Word retains native paper and centered two-column table geometry instead of applying worksheet CSS', { timeout: 30000 }, async () => {
  const zip = await JSZip.loadAsync(await wordFixture());
  const xml = await zip.file('word/document.xml').async('string');
  const table = '<w:tbl><w:tblPr><w:tblW w:w="2700" w:type="dxa"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Left cell</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Right cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  zip.file('word/document.xml', xml.replace('<w:sectPr>', table + '<w:sectPr>'));
  const { server, url } = await fixtureServer(new Map([['/centered.docx', await zip.generateAsync({ type: 'nodebuffer' })]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage({ viewport: { width: 320, height: 900 } });
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/centered.docx', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await frame.getByText('Right cell', { exact: true }).waitFor();
    assert.equal(await frame.locator('table').evaluate(table => getComputedStyle(table).display), 'table');
    assert.ok(await frame.locator('table').evaluate(table => {
      const zoom = parseFloat(getComputedStyle(document.body).zoom) || 1;
      return Math.abs(table.getBoundingClientRect().width / zoom - 180) <= 2;
    }));
    assert.ok(await frame.locator('section.docx').evaluate(section => section.getBoundingClientRect().width / (parseFloat(getComputedStyle(document.body).zoom) || 1) > 800));
    await page.getByRole('button', { name: '适合宽度', exact: true }).click();
    assert.ok(await frame.locator('table').evaluate(table => table.getBoundingClientRect().left >= 0 && table.getBoundingClientRect().right <= innerWidth));
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('automatic preview respects consent, avoids focus theft and exposes no download links', { timeout: 20000 }, async () => {
  const { server, requests, url } = await fixtureServer();
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('download', () => errors.push('Unexpected download'));
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/report.xlsx', autoOpen: true, canLoad: () => false }));
    assert.ok(!requests.includes('/report.xlsx'));
    assert.equal(await page.locator('iframe').count(), 0);
    await page.evaluate(() => window.mount({ src: '/report.xlsx', autoOpen: true }));
    await page.frameLocator('iframe').getByText('First sheet', { exact: true }).waitFor();
    assert.equal(await page.locator('.document-viewer a').count(), 0);
    assert.equal(await page.evaluate(() => document.activeElement === document.body), true);
    assert.deepEqual(errors, []);
    await page.evaluate(() => window.viewer.destroy());
    assert.equal(await page.locator('iframe').count(), 0);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('worksheet windows reach remote cells, release old DOM, preserve merges and styled read-only chrome', { timeout: 60000 }, async () => {
  const workbook = utils.book_new();
  const sheet = { A1: { t: 's', v: 'Window origin' }, DX5001: { t: 's', v: 'Window tail' }, '!ref': 'A1:DX5001', '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }] };
  utils.book_append_sheet(workbook, sheet, 'Large');
  utils.book_append_sheet(workbook, utils.aoa_to_sheet([['Another sheet']]), 'Small');
  const { server, url } = await fixtureServer(new Map([['/large.xlsx', write(workbook, { type: 'buffer', bookType: 'xlsx', bookSST: true, compression: true })]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage({ viewport: { width: 320, height: 900 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/large.xlsx', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await frame.getByText('Window origin', { exact: true }).waitFor();
    assert.equal(await frame.getByText('Window origin', { exact: true }).locator('..').getAttribute('colspan'), '2');
    assert.equal(await frame.getByRole('grid').getAttribute('aria-colcount'), '129');
    assert.equal(await frame.getByRole('grid').getAttribute('aria-readonly'), 'true');
    assert.equal(await page.locator('.document-viewer-mode').textContent(), '只读预览');
    const address = page.getByRole('textbox', { name: '单元格地址' });
    await address.fill('DX5001');
    await address.press('Enter');
    await frame.getByText('Window tail', { exact: true }).waitFor();
    await frame.getByText('Window tail', { exact: true }).click();
    assert.equal(await page.locator('.document-viewer-formula output').textContent(), 'Window tail');
    await page.waitForFunction(() => !document.querySelector('iframe').contentDocument.body.textContent.includes('Window origin'));
    assert.ok(await frame.locator('td[data-row]').count() <= 5000);
    assert.equal(await frame.locator('[data-document-page]').count(), 101);
    await address.fill('A1');
    await address.press('Enter');
    await frame.getByText('Window origin', { exact: true }).waitFor();
    for (let attempt = 0; attempt < 3; attempt++) {
      await address.fill('DX5001');
      await address.press('Enter');
      await frame.getByText('Window tail', { exact: true }).waitFor();
      await address.fill('A1');
      await address.press('Enter');
      await frame.getByText('Window origin', { exact: true }).waitFor();
    }
    await page.getByRole('tab', { name: 'Small', exact: true }).click();
    await frame.getByText('Another sheet', { exact: true }).waitFor();
    assert.equal(await frame.locator('[data-document-page]').count(), 1);
    await page.getByRole('button', { name: '放大', exact: true }).click();
    assert.equal(await page.locator('.document-viewer-view-tools output').textContent(), '125%');
    await page.getByRole('button', { name: '适合宽度', exact: true }).click();
    assert.equal(await page.locator('.document-viewer-view-tools output').textContent(), '100%');
    assert.deepEqual(errors, []);
    await page.evaluate(() => window.viewer.destroy());
    assert.equal(await page.locator('.document-viewer-sheet-footer').count(), 0);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('Word preserves saved pagination, page margins, styled text, headers, footers and footnotes', { timeout: 60000 }, async () => {
  const { server, url } = await fixtureServer(new Map([['/layout.docx', await wordLayoutFixture()]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage({ viewport: { width: 320, height: 900 } });
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/layout.docx', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await frame.getByText('Authored second page', { exact: true }).waitFor();
    assert.equal(await frame.locator('section.docx').count(), 2);
    assert.equal(await frame.getByText('Authored header', { exact: true }).count(), 2);
    assert.equal(await frame.getByText('Authored footer', { exact: true }).count(), 2);
    await frame.getByText('Authored footnote', { exact: true }).waitFor();
    const layout = await frame.getByText('Authored first page', { exact: true }).evaluate(element => {
      const style = getComputedStyle(element);
      const paper = getComputedStyle(element.closest('section.docx'));
      return { weight: style.fontWeight, color: style.color, size: style.fontSize, width: paper.width, left: paper.paddingLeft, right: paper.paddingRight };
    });
    assert.ok(Math.abs(parseFloat(layout.width) - 816) < 0.1);
    delete layout.width;
    assert.deepEqual(layout, { weight: '700', color: 'rgb(128, 0, 32)', size: '24px', left: '120px', right: '120px' });
    assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('Word Symbol bullets render as Unicode without changing unrelated private-use glyphs', { timeout: 20000 }, async () => {
  const { server, url } = await fixtureServer(new Map([['/bullets.docx', await wordBulletFixture()]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/bullets.docx', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await frame.getByText('Readable Symbol bullet', { exact: true }).waitFor();
    const bullet = await frame.getByText('Readable Symbol bullet', { exact: true }).evaluate(element => getComputedStyle(element.closest('p'), '::before').content);
    assert.ok(bullet.includes('\u2022'));
    assert.equal(bullet.includes('\uf0b7'), false);
    const unrelated = await frame.getByText('Preserved unrelated symbol', { exact: true }).evaluate(element => getComputedStyle(element.closest('p'), '::before').content);
    assert.ok(unrelated.includes('\uf0b7'));
    assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('worksheet windows preserve source rich runs, rows, inherited styles, hidden cells and calculated results', { timeout: 60000 }, async () => {
  const { server, url } = await fixtureServer(new Map([['/styles.xlsx', await styledSheetFixture()]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/styles.xlsx', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await frame.getByText('Rich', { exact: true }).waitFor();
    assert.equal(await frame.getByText('Hidden column', { exact: true }).count(), 0);
    assert.equal(await frame.locator('tr[aria-rowindex="3"]').count(), 0);
    assert.equal(await frame.getByText('Merged', { exact: true }).locator('..').getAttribute('colspan'), '2');
    assert.equal(await frame.getByText('Rich', { exact: true }).evaluate(element => getComputedStyle(element).color), 'rgb(255, 0, 0)');
    const sourceStyle = await frame.getByText('4.50', { exact: true }).locator('..').evaluate(element => { const style = getComputedStyle(element); return { font: style.fontFamily, size: style.fontSize, border: style.borderBottomStyle, fill: style.backgroundColor, align: style.textAlign }; });
    assert.deepEqual(sourceStyle, { font: 'Georgia', size: '18.6667px', border: 'double', fill: 'rgb(112, 160, 207)', align: 'center' });
    assert.equal(await frame.locator('td[data-row="4"][data-column="0"]').innerText(), '30');
    await page.getByRole('textbox', { name: '单元格地址' }).fill('B1');
    await page.getByRole('textbox', { name: '单元格地址' }).press('Enter');
    assert.equal(await page.locator('.document-viewer-status').textContent(), '此单元格在原文档中已隐藏。');
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('PDF honors per-page CropBox/rotation and repaints sharp rasters on zoom and workspace resize', { timeout: 60000 }, async () => {
  const { server, url } = await fixtureServer(new Map([['/mixed.pdf', pdfFixture(0, true)]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage({ viewport: { width: 1000, height: 900 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/mixed.pdf', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await frame.locator('.pdf-page[aria-busy=false]').first().waitFor();
    const first = frame.locator('.pdf-page').first();
    const second = frame.locator('.pdf-page').nth(1);
    for (const [wrapper, expected] of [[first, 1.5], [second, 1.2]]) {
      const ratio = await wrapper.evaluate(element => parseFloat(element.style.width) / parseFloat(element.style.height));
      assert.ok(Math.abs(ratio - expected) < 0.001);
    }
    const original = await first.locator('img').evaluate(element => element.naturalWidth);
    await page.getByRole('button', { name: '放大', exact: true }).click();
    await page.waitForFunction(width => document.querySelector('iframe').contentDocument.querySelector('.pdf-page img')?.naturalWidth > width, original);
    await first.locator('.textLayer span').first().waitFor();
    await page.getByRole('button', { name: '展开工作区', exact: true }).click();
    await page.setViewportSize({ width: 600, height: 900 });
    await first.locator('img').waitFor();
    assert.equal(await frame.locator('canvas').count(), 0);
    assert.equal(await page.locator('.document-viewer-status').textContent(), '');
    assert.deepEqual(errors, []);
    assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('PPT restores authored runs and workspace expansion preserves all document frames and keyboard focus', { timeout: 120000 }, async () => {
  const { server, requests, url } = await fixtureServer(new Map([['/styled.ppt', legacyPpt({ styled: true })]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage({ viewport: { width: 780, height: 900 } });
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('download', () => errors.push('Unexpected download'));
    for (const src of ['/styled.ppt', '/report.pptx', '/report.docx', '/report.pdf', '/report.xlsx', '/report.xls']) {
      await page.evaluate(src => window.mount({ src, autoOpen: true }), src);
      await page.locator('#preview').evaluate(element => { element.style.width = 'min(500px,100%)'; });
      await page.locator('.document-viewer-viewport[aria-busy=false]').waitFor();
      if (src === '/styled.ppt') {
        const run = page.frameLocator('iframe').locator('foreignObject p span').first();
        assert.deepEqual(await run.evaluate(element => { const style = getComputedStyle(element); return { color: style.color, weight: style.fontWeight, font: style.fontFamily, size: style.fontSize, italic: style.fontStyle, underline: style.textDecorationLine, align: getComputedStyle(element.parentElement).textAlign }; }), { color: 'rgb(255, 0, 0)', weight: '700', font: 'Georgia, sans-serif', size: '144px', italic: 'italic', underline: 'underline', align: 'center' });
      }
      const count = requests.filter(path => path === src).length;
      await page.evaluate(() => { window.originalFrame = document.querySelector('iframe'); window.originalDocument = window.originalFrame.contentDocument; });
      await page.getByRole('button', { name: '展开工作区', exact: true }).click();
      assert.equal(await page.getByRole('dialog').getAttribute('aria-modal'), 'true');
      if (src === '/report.pptx') {
        await page.waitForFunction(() => document.querySelector('iframe').contentDocument.querySelector('[data-slide-index]')?.getBoundingClientRect().width > 550);
        assert.equal(await page.frameLocator('iframe').locator('[data-document-page] [data-document-page]').count(), 0);
      }
      await page.keyboard.press('Tab');
      assert.equal(await page.evaluate(() => !!document.activeElement.closest('.document-viewer')), true);
      await page.frameLocator('iframe').locator('body').evaluate(element => { element.tabIndex = 0; element.focus(); });
      await page.keyboard.press('Escape');
      assert.equal(await page.getByRole('dialog').count(), 0);
      assert.equal(await page.getByRole('button', { name: '展开工作区', exact: true }).evaluate(element => element === document.activeElement), true);
      assert.equal(await page.evaluate(() => window.originalFrame === document.querySelector('iframe') && window.originalDocument === window.originalFrame.contentDocument), true);
      assert.equal(requests.filter(path => path === src).length, count);
      await page.getByRole('button', { name: '关闭预览', exact: true }).click();
    }
    assert.deepEqual(errors, []);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('Word pages release offscreen content and rehydrate it without executing document scripts', { timeout: 60000 }, async () => {
  const { server, url } = await fixtureServer(new Map([['/long.docx', await wordFixture(20)]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/long.docx', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await page.locator('.document-viewer-viewport[aria-busy="false"]').waitFor();
    assert.equal(await frame.locator('[data-document-page]').count(), 20);
    assert.ok(await frame.locator('section.docx').count() < 5);
    await frame.locator('[data-document-page="20"]').scrollIntoViewIfNeeded();
    await frame.getByText('Word page 20', { exact: true }).waitFor();
    await page.waitForFunction(() => !document.querySelector('iframe').contentDocument.body.textContent.includes('Word preview 中文'));
    assert.ok(await frame.locator('section.docx').count() < 5);
    await frame.locator('[data-document-page="1"]').scrollIntoViewIfNeeded();
    await frame.getByText('Word preview 中文', { exact: true }).waitFor();
    assert.equal(await frame.locator('script').count(), 0);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('closing a DOC preview terminates its still-loading parser Worker and releases the frame', { timeout: 60000 }, async () => {
  const { server, url } = await fixtureServer(new Map([['/parsing.doc', Buffer.alloc(512)]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.addInitScript(() => {
      const NativeWorker = window.Worker;
      window.parserWorkers = [];
      window.Worker = class extends NativeWorker {
        constructor(url, options) { super(url, options); this.source = String(url); this.stopped = false; window.parserWorkers.push(this); }
        terminate() { this.stopped = true; return super.terminate(); }
      };
    });
    await page.route('**/doc.worker.js', async route => { await new Promise(resolve => setTimeout(resolve, 300)); try { await route.continue(); } catch {} });
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/parsing.doc', autoOpen: true }));
    await page.waitForFunction(() => window.parserWorkers.some(worker => worker.source.endsWith('/doc.worker.js')));
    await page.evaluate(() => window.viewer.close());
    assert.equal(await page.locator('iframe').count(), 0);
    assert.equal(await page.evaluate(() => window.parserWorkers.every(worker => worker.stopped)), true);
    assert.deepEqual(errors, []);
    assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('PDF range transport previews without transferring the whole file and evicts distant rasters', { timeout: 60000 }, async () => {
  const bytes = pdfFixture(8 * 1024 * 1024);
  const { server, ranges, url } = await fixtureServer(new Map([['/ranged.pdf', bytes]]));
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/ranged.pdf', autoOpen: true }));
    const frame = page.frameLocator('iframe');
    await frame.getByText('Document preview page one', { exact: true }).waitFor();
    assert.ok(ranges.length >= 2);
    assert.ok(ranges.reduce((total, range) => total + range.end - range.begin + 1, 0) < bytes.length / 4);
    await page.locator('.document-viewer-viewport').evaluate(element => { element.style.height = '180px'; });
    await frame.locator('body').evaluate(element => { element.ownerDocument.defaultView.scrollTo(0, element.scrollHeight); });
    await frame.getByText('Document preview page two', { exact: true }).waitFor();
    await page.waitForFunction(() => !document.querySelector('iframe').contentDocument.querySelector('[data-document-page="1"] img'));
    await frame.locator('[data-document-page="1"]').scrollIntoViewIfNeeded();
    await frame.getByText('Document preview page one', { exact: true }).waitFor();
    assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('document HTML loses executable content before entering the script-disabled frame', async () => {
  const { server, url } = await fixtureServer();
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('console', message => { if (/Blocked script execution/.test(message.text())) errors.push(message.text()); });
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/report.xlsx', autoOpen: true }));
    await page.frameLocator('iframe').getByText('First sheet', { exact: true }).waitFor();
    const result = await page.evaluate(async () => {
      const { inertDocumentHtml } = await import('/safe-html.js');
      const doc = document.querySelector('iframe').contentDocument;
      doc.body.replaceChildren(inertDocumentHtml(doc, '<script>parent.injected=true</script><svg onload="parent.injected=true"></svg><a href="javascript:parent.injected=true">Readable text</a><iframe srcdoc="<script>parent.injected=true</script>"></iframe>'));
      return { text: doc.body.textContent, scripts: doc.querySelectorAll('script,iframe,[onload],[href]').length, injected: window.injected || false };
    });
    assert.deepEqual(result, { text: 'Readable text', scripts: 0, injected: false });
    assert.deepEqual(errors, []);
    assert.equal(await page.locator('iframe').getAttribute('sandbox'), 'allow-same-origin');
  } finally { await browser.close(); await new Promise(resolve => server.close(resolve)); }
});

test('legacy PPT Worker renders independently without a vendor engine', { timeout: 20000 }, async () => {
  const { server, url } = await fixtureServer();
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/report.ppt' }));
    await page.getByRole('button', { name: '预览文档', exact: true }).click();
    await page.waitForFunction(() => window.previewErrors.length || document.querySelector('iframe')?.contentDocument?.querySelector('svg'));
    assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
    await page.frameLocator('iframe').getByText('First slide: 中文', { exact: true }).waitFor();
    await page.locator('.document-viewer-viewport[aria-busy="false"]').waitFor();
    assert.equal(await page.locator('.document-viewer-status').textContent(), '');
    await page.evaluate(() => window.mount({ src: '/thin.ppt', autoOpen: true }));
    const transcript = page.frameLocator('iframe').locator('.document-page-text').getByText('First slide: 中文', { exact: true });
    await transcript.waitFor();
    assert.ok(await transcript.evaluate(node => node.getBoundingClientRect().height >= 16));
    assert.equal(await page.frameLocator('iframe').locator('[data-document-page]').count(), 2);
    assert.equal(await page.getByRole('button', { name: '下一页', exact: true }).count(), 0);
    await page.evaluate(() => window.mount({ src: '/fit.ppt', autoOpen: true }));
    const positioned = page.frameLocator('iframe').locator('foreignObject').getByText('First slide: 中文', { exact: true });
    await positioned.waitFor();
    assert.ok(await positioned.evaluate(node => node.getBoundingClientRect().height >= 16));
  } finally { await browser.close(); await new Promise(resolve => server.close(resolve)); }
});

test('themed worksheet controls support keyboard, contrast, forced colors and sandbox Escape', { timeout: 60000 }, async () => {
  const { server, url } = await fixtureServer();
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    for (const width of [320, 1280]) {
      const page = await browser.newPage({ viewport: { width, height: 900 } });
      page.setDefaultTimeout(10000);
      await page.goto(url);
      await page.waitForFunction(() => window.ready);
      await page.evaluate(() => {
        document.body.style.cssText = '--paper:#f7f8ff;--panel:#fff;--ink:#171a35;--muted:#5e6480;--line:#dce0f2;--accent:#3945c6;font:16px/1.75 system-ui';
        window.mount({ src: '/report.xlsx' });
      });
      const preview = page.getByRole('button', { name: '预览文档', exact: true });
      assert.equal(await preview.getAttribute('aria-expanded'), 'false');
      await preview.click();
      await page.frameLocator('iframe').getByText('First sheet', { exact: true }).waitFor();
      assert.equal(await page.locator('select').count(), 0);
      const first = page.getByRole('tab').first();
      const last = page.getByRole('tab', { name: 'Details', exact: true });
      await first.focus();
      await page.keyboard.press('ArrowRight');
      assert.equal(await last.getAttribute('aria-selected'), 'true');
      assert.equal(await last.evaluate(element => element === document.activeElement), true);
      await page.frameLocator('iframe').getByText('Second sheet', { exact: true }).waitFor();
      await page.keyboard.press('Home');
      assert.equal(await first.getAttribute('aria-selected'), 'true');
      await page.keyboard.press('End');
      assert.equal(await last.getAttribute('aria-selected'), 'true');
      assert.equal(await page.getByRole('tabpanel').getAttribute('aria-labelledby'), await last.getAttribute('id'));
      assert.equal(await first.evaluate(element => getComputedStyle(element).appearance), 'none');
      const content = page.frames().find(frame => frame.parentFrame());
      await page.addScriptTag({ url: url + '/axe.js' });
      for (const mode of ['light', 'dark', 'custom']) {
        if (mode === 'dark') await page.evaluate(() => {
          document.body.style.cssText = '--paper:#16182a;--panel:#23263c;--ink:#eef0ff;--muted:#b9c0dc;--line:#68708b;--accent:#abb2ff;font:16px/1.75 system-ui;color-scheme:dark';
        });
        if (mode === 'custom') await page.evaluate(() => {
          document.body.style.cssText = 'background:#ffff00;color:#002a49;font:16px/1.75 Georgia';
          const link = document.createElement('a');
          link.href = '#example';
          link.textContent = 'Host theme link';
          document.body.prepend(link);
          const stylesheet = document.createElement('style');
          stylesheet.id = 'host-palette';
          stylesheet.textContent = 'body>a{color:#ffd166}';
          document.head.append(stylesheet);
        });
        const expected = { light: 'rgb(23, 26, 53)', dark: 'rgb(238, 240, 255)', custom: 'rgb(0, 42, 73)' }[mode];
        await page.waitForFunction(color => getComputedStyle(document.querySelector('iframe').contentDocument.body).color === color, expected, { timeout: 5000 });
        if (mode === 'custom') {
          assert.equal(await content.locator('body').evaluate(element => getComputedStyle(element).backgroundColor), 'rgb(255, 255, 0)');
          assert.equal(await last.evaluate(element => getComputedStyle(element).borderBottomColor), 'rgb(0, 42, 73)');
          await page.evaluate(() => { document.querySelector('#host-palette').textContent = 'body>a{color:rgba(30,30,30,.2)}'; });
          await page.waitForFunction(() => getComputedStyle(document.querySelector('.document-viewer')).getPropertyValue('--dv-accent').trim() === 'rgba(30, 30, 30, 0.2)', null, { timeout: 5000 });
        }
        const audit = async scope => scope.evaluate(async () => {
          const result = await axe.run(document.querySelector('.document-viewer'), { iframes: false, runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21aa', 'wcag22aa'] } });
          return result.violations.map(issue => ({ id: issue.id, nodes: issue.nodes.map(node => node.failureSummary) }));
        });
        const issues = await audit(page);
        assert.deepEqual(issues, [], 'Parent controls at ' + width + 'px / ' + mode);
        assert.ok(await content.locator('th[scope=col]').count());
        assert.ok(await content.locator('th[scope=row]').count());
        assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
        assert.ok(await last.evaluate(element => element.getBoundingClientRect().height >= 44));
      }
      await page.emulateMedia({ forcedColors: 'active' });
      await last.focus();
      assert.notEqual(await last.evaluate(element => getComputedStyle(element).outlineStyle), 'none');
      await content.locator('body').evaluate(element => { element.tabIndex = 0; element.focus(); });
      await page.keyboard.press('Escape');
      assert.equal(await page.locator('iframe').count(), 0);
      assert.equal(await preview.getAttribute('aria-expanded'), 'false');
      assert.equal(await preview.evaluate(element => element === document.activeElement), true);
      await page.close();
    }
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('operator-provided Office samples render without downloads', { timeout: 120000, skip: !process.env.DOCUMENT_VIEWER_SAMPLES }, async () => {
  const names = ['sample-document-medium.doc', 'sample-document-medium.docx', 'sample-document.ppt', 'sample-presentation-10-slides.pptx', 'sample-spreadsheet-100-rows.xls', 'sample-spreadsheet-100-rows.xlsx'];
  const samples = new Map(await Promise.all(names.map(async name => ['/' + name, await readFile(resolve(process.env.DOCUMENT_VIEWER_SAMPLES, name))])));
  if (process.env.DOCUMENT_VIEWER_PDF) {
    names.push('sample-document-5-pages.pdf');
    samples.set('/sample-document-5-pages.pdf', await readFile(process.env.DOCUMENT_VIEWER_PDF));
  }
  const { server, url } = await fixtureServer(samples);
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    page.setDefaultTimeout(20000);
    const errors = [];
    page.on('download', () => errors.push('Unexpected download'));
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => {
      if (/Blocked script execution|webmcp-interceptor|modelContext is not available/.test(message.text())) errors.push(message.text());
    });
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    for (const name of names) {
      await page.evaluate(src => window.mount({ src }), '/' + name);
      await page.getByRole('button', { name: '预览文档', exact: true }).click();
      await page.waitForFunction(() => window.previewErrors.length || document.querySelector('.document-viewer-viewport')?.getAttribute('aria-busy') === 'false');
      assert.deepEqual(await page.evaluate(() => window.previewErrors), [], name);
      const text = await page.frameLocator('iframe').locator('body').innerText();
      console.log(name + ': ' + text.slice(0, 180).replaceAll('\n', ' '));
      assert.ok(text.length > 20, 'Expected real document content: ' + name);
      if (name.endsWith('.pdf')) assert.equal(await page.frameLocator('iframe').locator('[data-document-page]').count(), 8);
      if (name.endsWith('.doc')) assert.equal(await page.frameLocator('iframe').locator('.msdoc-root').first().evaluate(element => parseFloat(element.style.width)), 816);
      if (name.endsWith('.xlsx')) {
        await page.getByRole('tab', { name: 'Summary', exact: true }).click();
        await page.frameLocator('iframe').getByText('595500', { exact: true }).waitFor();
        const values = await page.frameLocator('iframe').locator('td[data-column="1"] .sheet-cell-content').allTextContents();
        assert.ok(values.includes('100'));
        assert.ok(values.includes('5955'));
        assert.ok(values.includes('26400'));
      }
      if (process.env.DOCUMENT_VIEWER_SCREENSHOTS) await page.screenshot({ path: resolve(process.env.DOCUMENT_VIEWER_SCREENSHOTS, name + '.png'), fullPage: true });
      await page.getByRole('button', { name: '关闭预览', exact: true }).click();
    }
    assert.deepEqual(errors, []);
  } finally { await browser.close(); server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
});

test('browser renderers load only on preview, render real PDF/Office bytes, and work at mobile width', { timeout: 120000 }, async () => {
  const { server, requests, url } = await fixtureServer();
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    for (const width of [320, 1280]) {
      const context = await browser.newContext({ viewport: { width, height: 900 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      page.on('download', () => errors.push('Unexpected browser download instead of inline preview'));
      page.on('console', message => {
        if (/Blocked script execution|webmcp-interceptor|modelContext is not available/.test(message.text())) errors.push(message.text());
      });
      await context.route('https://blocked.example/**', route => { errors.push('Unexpected external resource'); return route.abort(); });
      await page.goto(url);
      await page.waitForFunction(() => window.ready);
      await page.evaluate(() => window.mount({ src: '/report.xlsx' }));
      assert.ok(!requests.includes('/report.xlsx'));
      assert.ok(!requests.some(path => /\/(?:sheets|slides|pdf|word|ppt)-/.test(path)));
      const originalCount = requests.filter(path => path === '/report.xlsx').length;
      await page.getByRole('button', { name: '预览文档', exact: true }).click();
      const frame = page.frameLocator('iframe');
      await frame.getByText('First sheet', { exact: true }).waitFor();
      assert.equal(requests.filter(path => path === '/report.xlsx').length, originalCount + 1);
      assert.equal(await frame.locator('img').count(), 0);
      await page.getByRole('tab', { name: 'Details', exact: true }).click();
      await frame.getByText('Second sheet', { exact: true }).waitFor();
      await page.getByRole('button', { name: '关闭预览', exact: true }).click();
      assert.equal(await page.locator('iframe').count(), 0);
      assert.equal(await page.getByRole('button', { name: '预览文档', exact: true }).evaluate(node => node === document.activeElement), true);

      for (const [source, expected] of [['/report.docx', 'Word preview 中文'], ['/report.xls', 'First sheet'], ['/report.pptx', 'PowerPoint preview 中文'], ['/report.ppt', 'First slide: 中文']]) {
        await page.evaluate(src => window.mount({ src }), source);
        await page.getByRole('button', { name: '预览文档', exact: true }).click();
        try { await page.frameLocator('iframe').getByText(expected, { exact: true }).waitFor({ timeout: 20000 }); }
        catch (error) { throw new Error(source + ': ' + JSON.stringify(await page.evaluate(() => window.previewErrors)), { cause: error }); }
        assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
        assert.equal(await page.locator('iframe').getAttribute('sandbox'), 'allow-same-origin');
        if (source.endsWith('.docx')) {
          const geometry = await page.frameLocator('iframe').locator('section.docx').evaluate(node => ({ width: node.getBoundingClientRect().width, sourceWidth: parseFloat(getComputedStyle(node).width), viewport: innerWidth }));
          assert.ok(geometry.width <= geometry.viewport, 'Word paper must fit without horizontal clipping');
          assert.ok(geometry.width >= Math.min(geometry.sourceWidth, geometry.viewport - 24) * 0.95, 'Word must retain native paper width, fitting smaller viewports without reflow');
        }
        if (source.endsWith('.ppt')) {
          await page.frameLocator('iframe').locator('[data-document-page="2"]').scrollIntoViewIfNeeded();
          await page.frameLocator('iframe').getByText('Second slide', { exact: true }).waitFor();
        }
        await page.getByRole('button', { name: '关闭预览', exact: true }).click();
      }

      await page.evaluate(() => window.mount({ src: '/report.pdf' }));
      await page.getByRole('button', { name: '预览文档', exact: true }).click();
      await page.frameLocator('iframe').getByText('Document preview page one', { exact: true }).waitFor({ timeout: 20000 });
      assert.ok(await page.frameLocator('iframe').locator('.pdf-page img').first().evaluate(image => image.naturalWidth >= Math.min(600, image.width * 2) - 1), 'Narrow PDF pages need sufficient raster density for visible text');
      assert.ok(await page.frameLocator('iframe').locator('.pdf-page img').first().evaluate(image => {
        const canvas = document.createElement('canvas');
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        canvas.getContext('2d').drawImage(image, 0, 0);
        const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
        let ink = 0;
        for (let offset = 0; offset < pixels.length; offset += 4) if (pixels[offset] < 128 && pixels[offset + 3]) ink++;
        return ink > 100;
      }), 'PDF preview must paint visible content, not only an invisible text layer');
      assert.equal(await page.getByRole('button', { name: '下一页', exact: true }).count(), 0);
      await page.frameLocator('iframe').locator('[data-document-page="2"]').scrollIntoViewIfNeeded();
      await page.frameLocator('iframe').getByText('Document preview page two', { exact: true }).waitFor();
      await page.getByRole('button', { name: '关闭预览', exact: true }).click();
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
      assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
      assert.deepEqual(errors, []);
      requests.length = 0;
      await context.close();
    }
  } finally { await browser.close(); await new Promise(resolve => server.close(resolve)); }
});

test('extensionless shares, consent denial, error recovery, size limits and close during fetch', { timeout: 40000 }, async () => {
  const { server, requests, url } = await fixtureServer();
  const browser = await chromium.launch({ headless: true, args: ['--disable-extensions'] });
  try {
    const page = await browser.newPage();
    await page.goto(url);
    await page.waitForFunction(() => window.ready);
    await page.evaluate(() => window.mount({ src: '/download/123', format: 'xlsx', canLoad: () => false }));
    await page.getByRole('button', { name: '预览文档', exact: true }).click();
    assert.ok(!requests.includes('/download/123'));
    await page.evaluate(() => window.mount({ src: '/download/123', format: 'xlsx' }));
    await page.getByRole('button', { name: '预览文档', exact: true }).click();
    await page.frameLocator('iframe').getByText('First sheet', { exact: true }).waitFor();
    for (const options of [{ src: '/invalid.pdf' }, { src: '/report.xlsx', maxBytes: 20 }]) {
      await page.evaluate(options => window.mount(options), options);
      await page.getByRole('button', { name: '预览文档', exact: true }).click();
      await page.getByRole('status').filter({ hasText: '无法预览' }).waitFor();
      assert.equal(await page.locator('iframe').count(), 0);
      assert.equal((await page.evaluate(() => window.previewErrors)).length, 1);
    }
    let release;
    await page.route('**/pending.pdf', async route => { await new Promise(resolve => { release = resolve; }); await route.fulfill({ body: pdfFixture() }).catch(() => {}); });
    await page.evaluate(() => window.mount({ src: '/pending.pdf' }));
    const requested = page.waitForRequest('**/pending.pdf');
    await page.getByRole('button', { name: '预览文档', exact: true }).click();
    await requested;
    await page.getByRole('button', { name: '关闭预览', exact: true }).click();
    release?.();
    await page.waitForTimeout(100);
    assert.equal(await page.locator('iframe').count(), 0);
    assert.deepEqual(await page.evaluate(() => window.previewErrors), []);
  } finally { await browser.close(); await new Promise(resolve => server.close(resolve)); }
});
