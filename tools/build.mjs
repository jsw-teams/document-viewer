import { mkdir, writeFile, copyFile, rm } from 'node:fs/promises';
import { dirname, resolve, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { documentViewerAssets } from './assets.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const output = resolve(root, 'dist');
if (relative(root, output) !== 'dist') throw new Error('Build output must stay in the project dist directory');
const assets = await documentViewerAssets('');
await rm(output, { recursive: true, force: true });
for (const asset of assets) {
  const destination = resolve(output, asset.path.replace(/^\//, ''));
  await mkdir(dirname(destination), { recursive: true });
  if (asset.source) await copyFile(asset.source, destination);
  else await writeFile(destination, asset.content);
}
