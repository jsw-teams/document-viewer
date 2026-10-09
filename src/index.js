import { documentFormat, documentUrl } from './formats.js';
import { documentLabels } from './language.js';
import { documentFrame } from './frame.js';
import { readDocument, readPdfDocument } from './read.js';
import { validateOfficeZip } from './zip-limits.js';
import { followTheme } from './theme.js';
import { previewChrome } from './chrome.js';
import { windowDocumentPages } from './windowed-pages.js';
import { buttonContent } from './icons.js';

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
  const maxBytes = options.maxBytes ?? (format === 'pdf' ? 1024 : 50) * 1024 * 1024;
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 1) throw new Error('maxBytes must be a positive integer');
  const root = document.createElement('dialog');
  root.className = 'document-viewer';
  root.open = true;
  root.setAttribute('role', 'region');
  root.dataset.format = format;
  root.setAttribute('aria-label', title);
  const heading = document.createElement('strong');
  heading.textContent = title;
  const header = document.createElement('div');
  header.className = 'document-viewer-header';
  const badge = document.createElement('span');
  badge.className = 'document-viewer-format';
  badge.textContent = format.toUpperCase();
  badge.hidden = title.trim().toUpperCase() === format.toUpperCase();
  const mode = document.createElement('span');
  mode.className = 'document-viewer-mode';
  mode.textContent = labels.readOnly;
  header.append(badge, heading, mode);
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
  buttonContent(previewButton, labels.preview, 'eye');
  previewButton.setAttribute('aria-controls', viewport.id);
  previewButton.setAttribute('aria-expanded', 'false');
  const closeButton = document.createElement('button');
  closeButton.type = 'button';
  buttonContent(closeButton, labels.close, 'x');
  closeButton.hidden = true;
  const expandButton = document.createElement('button');
  expandButton.type = 'button';
  expandButton.className = 'document-viewer-expand';
  buttonContent(expandButton, labels.expand, 'expand', true);
  expandButton.setAttribute('aria-controls', viewport.id);
  expandButton.setAttribute('aria-expanded', 'false');
  expandButton.hidden = true;
  toolbar.append(previewButton, closeButton);
  viewport.append(expandButton);
  root.append(header, toolbar, status, viewport);
  container.replaceChildren(root);
  const stopTheme = followTheme(root, container);
  const controls = document.createElement('div');
  controls.className = 'document-viewer-controls';
  let active = null;
  let disposed = false;
  let expanded = false;
  let previousOverflow = '';

  function collapse(restoreFocus = true) {
    if (!expanded) return;
    expanded = false;
    root.close();
    root.open = true;
    root.removeAttribute('data-expanded');
    root.setAttribute('role', 'region');
    root.removeAttribute('aria-modal');
    buttonContent(expandButton, labels.expand, 'expand', true);
    expandButton.setAttribute('aria-expanded', 'false');
    document.documentElement.style.overflow = previousOverflow;
    if (restoreFocus && !disposed) expandButton.focus({ preventScroll: true });
  }
  expandButton.addEventListener('click', () => {
    if (expanded) { collapse(); return; }
    previousOverflow = document.documentElement.style.overflow;
    root.close();
    root.setAttribute('data-expanded', '');
    root.setAttribute('role', 'dialog');
    root.setAttribute('aria-modal', 'true');
    root.showModal();
    expanded = true;
    document.documentElement.style.overflow = 'hidden';
    buttonContent(expandButton, labels.collapse, 'collapse', true);
    expandButton.setAttribute('aria-expanded', 'true');
    expandButton.focus({ preventScroll: true });
  });
  root.addEventListener('cancel', event => { event.preventDefault(); collapse(); });

  function close(restoreFocus = true) {
    collapse(false);
    const previous = active;
    active = null;
    previous?.controller.abort();
    previous?.cleanup?.();
    controls.replaceChildren();
    controls.remove();
    viewport.replaceChildren(expandButton);
    viewport.hidden = true;
    viewport.removeAttribute('aria-busy');
    viewport.removeAttribute('role');
    viewport.removeAttribute('aria-labelledby');
    viewport.removeAttribute('tabindex');
    closeButton.hidden = true;
    expandButton.hidden = true;
    previewButton.hidden = false;
    previewButton.setAttribute('aria-expanded', 'false');
    status.textContent = '';
    if (restoreFocus && !disposed) previewButton.focus();
  }

  async function open({ focus = true } = {}) {
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
    if (focus) closeButton.focus();
    const guard = callback => (...args) => {
      if (active === session && !signal.aborted) return callback(...args);
    };
    try {
      const [data, renderer, frame] = await Promise.all([
        (format === 'pdf' ? readPdfDocument : readDocument)(preview.href, signal, maxBytes), renderers[format](), documentFrame(viewport, title, signal)
      ]);
      signal.throwIfAborted();
      frame.contentDocument.addEventListener('keydown', event => {
        if (event.key === 'Escape') { event.preventDefault(); if (expanded) collapse(); else close(); }
      }, { signal });
      if (['docx', 'pptx'].includes(format)) validateOfficeZip(data);
      const cleanup = await renderer.render({ data, format, frame, viewport, controls, signal, labels, status,
        setCleanup: callback => { session.cleanup = callback; }, guard });
      if (active !== session) { cleanup?.(); return; }
      session.cleanup = cleanup;
      const disposeRenderer = cleanup;
      const disposePages = ['doc', 'docx', 'ppt'].includes(format) ? await windowDocumentPages(frame, signal) : null;
      if (active !== session) { disposePages?.(); return; }
      const disposeChrome = previewChrome({ root, toolbar, frame, signal, labels, fitWidth: ['doc', 'docx'].includes(format) });
      session.cleanup = () => { disposeChrome(); disposePages?.(); disposeRenderer?.(); };
      expandButton.hidden = false;
      viewport.setAttribute('aria-busy', 'false');
      if (status.textContent === labels.loading) status.textContent = '';
    } catch (error) {
      if (active !== session || signal.aborted) return;
      close(false);
      status.textContent = labels.error;
      options.onError?.(error);
      if (focus) previewButton.focus();
    }
  }

  previewButton.addEventListener('click', () => open());
  closeButton.addEventListener('click', () => close());
  root.addEventListener('keydown', event => { if (event.key === 'Escape' && active) { event.preventDefault(); if (expanded) collapse(); else close(); } });
  if (options.autoOpen !== false) void open({ focus: false });
  return { open, close, destroy() { disposed = true; close(false); stopTheme(); root.remove(); } };
}
