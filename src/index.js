import { documentFormat, documentUrl } from './formats.js';
import { documentLabels } from './language.js';
import { documentFrame } from './frame.js';
import { readDocument } from './read.js';
import { validateOfficeZip } from './zip-limits.js';
import { followTheme } from './theme.js';

export { documentFormat, documentUrl, documentLabels };

const renderers = {
  pdf: () => import('./renderers/pdf.js'),
  doc: () => import('./renderers/doc.js'),
  docx: () => import('./renderers/word.js'),
  ppt: () => import('./renderers/ppt.js'),
  pptx: () => import('./renderers/slides.js'),
  xlsx: () => import('./renderers/sheets.js'),
  xls: () => import('./renderers/sheets.js')
};

let viewerId = 0;

export function mountDocument(container, options) {
  const source = documentUrl(options.src, document.baseURI);
  const preview = documentUrl(options.previewSrc || options.src, document.baseURI);
  const format = documentFormat(preview.href, options.format);
  if (!format) throw new Error('Unsupported document format; supply format for extensionless URLs');
  const labels = { ...documentLabels(options.locale || document.documentElement.lang), ...options.labels };
  const title = String(options.title || decodeURIComponent(source.pathname.split('/').pop()) || format.toUpperCase());
  const maxBytes = options.maxBytes ?? 50 * 1024 * 1024;
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 1) throw new Error('maxBytes must be a positive integer');
  const root = document.createElement('section');
  root.className = 'document-viewer';
  root.setAttribute('aria-label', title);
  const heading = document.createElement('strong');
  heading.textContent = title;
  const toolbar = document.createElement('div');
  toolbar.className = 'document-viewer-toolbar';
  const status = document.createElement('p');
  status.setAttribute('role', 'status');
  status.className = 'document-viewer-status';
  const viewport = document.createElement('div');
  viewport.className = 'document-viewer-viewport';
  viewport.id = 'document-viewport-' + ++viewerId;
  viewport.hidden = true;
  const previewButton = document.createElement('button');
  previewButton.type = 'button';
  previewButton.textContent = labels.preview;
  previewButton.setAttribute('aria-controls', viewport.id);
  previewButton.setAttribute('aria-expanded', 'false');
  const closeButton = document.createElement('button');
  closeButton.type = 'button';
  closeButton.textContent = labels.close;
  closeButton.hidden = true;
  const download = document.createElement('a');
  download.href = source.href;
  download.download = '';
  download.rel = 'noopener noreferrer';
  download.referrerPolicy = 'no-referrer';
  download.textContent = labels.download;
  toolbar.append(previewButton, closeButton, download);
  root.append(heading, toolbar, status, viewport);
  container.replaceChildren(root);
  const stopTheme = followTheme(root, container);
  const controls = document.createElement('div');
  controls.className = 'document-viewer-controls';
  let active = null;
  let disposed = false;

  function close(restoreFocus = true) {
    const previous = active;
    active = null;
    previous?.controller.abort();
    previous?.cleanup?.();
    controls.replaceChildren();
    controls.remove();
    viewport.replaceChildren();
    viewport.hidden = true;
    viewport.removeAttribute('aria-busy');
    viewport.removeAttribute('role');
    viewport.removeAttribute('aria-labelledby');
    viewport.removeAttribute('tabindex');
    closeButton.hidden = true;
    previewButton.hidden = false;
    previewButton.setAttribute('aria-expanded', 'false');
    status.textContent = '';
    if (restoreFocus && !disposed) previewButton.focus();
  }

  async function open() {
    if (disposed || active || (options.canLoad && !options.canLoad())) return;
    const session = { controller: new AbortController(), cleanup: null };
    active = session;
    const signal = session.controller.signal;
    previewButton.hidden = true;
    previewButton.setAttribute('aria-expanded', 'true');
    closeButton.hidden = false;
    viewport.hidden = false;
    viewport.setAttribute('aria-busy', 'true');
    status.textContent = labels.loading;
    toolbar.after(controls);
    closeButton.focus();
    const guard = callback => (...args) => {
      if (active === session && !signal.aborted) return callback(...args);
    };
    try {
      const [data, renderer, frame] = await Promise.all([
        readDocument(preview.href, signal, maxBytes), renderers[format](), documentFrame(viewport, title, signal)
      ]);
      signal.throwIfAborted();
      frame.contentDocument.addEventListener('keydown', event => {
        if (event.key === 'Escape') { event.preventDefault(); close(); }
      }, { signal });
      if (['docx', 'pptx', 'xlsx'].includes(format)) validateOfficeZip(data);
      const cleanup = await renderer.render({ data, frame, viewport, controls, signal, labels, status,
        setCleanup: callback => { session.cleanup = callback; }, guard });
      if (active !== session) { cleanup?.(); return; }
      session.cleanup = cleanup;
      viewport.setAttribute('aria-busy', 'false');
      if (status.textContent === labels.loading) status.textContent = '';
    } catch (error) {
      if (active !== session || signal.aborted) return;
      close(false);
      status.textContent = labels.error;
      options.onError?.(error);
      previewButton.focus();
    }
  }

  previewButton.addEventListener('click', open);
  closeButton.addEventListener('click', () => close());
  root.addEventListener('keydown', event => { if (event.key === 'Escape' && active) { event.preventDefault(); close(); } });
  return { open, close, destroy() { disposed = true; close(false); stopTheme(); root.remove(); } };
}
