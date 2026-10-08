import { parseMsDoc, renderMsDoc } from '@file-viewer/doc';
import { docPageLayout } from './layout.js';

self.onmessage = ({ data }) => {
  try {
    const bytes = new Uint8Array(data);
    const result = renderMsDoc(parseMsDoc(bytes), { externalLinkPolicy: 'block', externalResourcePolicy: 'block' });
    self.postMessage({ html: result.html, css: result.css, layout: docPageLayout(bytes) });
  } catch (error) { self.postMessage({ error: error.message }); }
};
