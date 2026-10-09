import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { readdir, readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, resolve, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const root = fileURLToPath(new URL('../', import.meta.url));

async function resourceDirectory(directory, name) {
  const files = [];
  async function walk(folder, prefix = '') {
    for (const entry of (await readdir(folder, { withFileTypes: true })).sort((left, right) => left.name.localeCompare(right.name))) {
      const path = prefix + entry.name;
      if (entry.isDirectory()) await walk(resolve(folder, entry.name), path + '/');
      else if (entry.isFile()) files.push({ path, source: resolve(folder, entry.name) });
    }
  }
  await walk(directory);
  const hash = createHash('sha256');
  for (const file of files) { hash.update(file.path); hash.update(await readFile(file.source)); }
  return { name: name + '.' + hash.digest('hex').slice(0, 16), files };
}

export async function documentViewerAssets(prefix = 'document-viewer') {
  const pdfRoot = resolve(dirname(require.resolve('pdfjs-dist/build/pdf.mjs')), '..');
  const pdfResources = [];
  for (const directory of ['cmaps', 'standard_fonts', 'wasm']) {
    const resources = await resourceDirectory(resolve(pdfRoot, directory), directory);
    pdfResources.push(...resources.files.map(file => ({ ...file, path: directory + '/' + file.path })));
  }
  const pdfHash = createHash('sha256');
  for (const file of pdfResources) { pdfHash.update(file.path); pdfHash.update(await readFile(file.source)); }
  const pdfName = 'pdf-assets.' + pdfHash.digest('hex').slice(0, 16);
  const output = resolve(root, '.bundle');
  const bundled = await build({ entryPoints: { index: resolve(root, 'src/index.js'),
    'doc.worker': resolve(root, 'src/doc/worker.js'),
    'ppt.worker': resolve(root, 'src/ppt/worker.js'),
    'sheets.worker': resolve(root, 'src/sheets/worker.js'),
    'pdf.worker': require.resolve('pdfjs-dist/build/pdf.worker.mjs') },
    outdir: output, bundle: true, splitting: true, minify: true, write: false,
    format: 'esm', target: 'es2022', chunkNames: 'chunks/[name]-[hash]', legalComments: 'eof',
    define: { __DOCUMENT_PDF_ASSETS__: JSON.stringify('../' + pdfName + '/'), 'process.env.NODE_ENV': '"production"' } });
  const assets = bundled.outputFiles.map(file => ({ path: prefix + '/' + relative(output, file.path).split(sep).join('/'),
    content: Buffer.from(file.contents), sourceName: 'document-viewer component' }));
  const styles = await build({ entryPoints: [resolve(root, 'src/styles.css')], bundle: true, minify: true, write: false });
  assets.push({ path: prefix + '/styles.css', content: Buffer.from(styles.outputFiles[0].contents), sourceName: 'document-viewer component' });
  for (const file of pdfResources) assets.push({ ...file, path: prefix + '/' + pdfName + '/' + file.path, sourceName: 'PDF.js resources' });
  const licenses = [
    ['Lucide', resolve(root, 'src/lucide-license.txt')],
    ['PDF.js', resolve(pdfRoot, 'LICENSE')],
    ['docx-preview', resolve(dirname(require.resolve('docx-preview')), '../LICENSE')],
    ['pptx-renderer', resolve(dirname(require.resolve('@aiden0z/pptx-renderer')), '../LICENSE')],
    ['SheetJS', resolve(dirname(require.resolve('xlsx')), 'LICENSE')],
    ['legacy-doc', resolve(dirname(require.resolve('@file-viewer/doc')), '../LICENSE')],
    ['cfb', resolve(dirname(require.resolve('cfb')), 'LICENSE')],
    ['fflate', resolve(dirname(require.resolve('fflate/package.json')), 'LICENSE')]
  ];
  for (const [name, source] of licenses) assets.push({ path: prefix + '/licenses/' + name + '.txt', source, sourceName: 'Renderer licenses' });
  return assets;
}
