export async function documentFrame(container, title, signal) {
  signal.throwIfAborted();
  const frame = document.createElement('iframe');
  frame.className = 'document-viewer-frame';
  frame.title = title;
  frame.setAttribute('sandbox', 'allow-same-origin');
  frame.referrerPolicy = 'no-referrer';
  const loaded = new Promise((resolve, reject) => {
    const aborted = () => reject(signal.reason || new DOMException('Aborted', 'AbortError'));
    signal.addEventListener('abort', aborted, { once: true });
    frame.addEventListener('load', () => { signal.removeEventListener('abort', aborted); resolve(); }, { once: true });
  });
  frame.srcdoc = `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src data: blob:; base-uri 'none'; form-action 'none'"><style>html,body{margin:0;padding:0;font:var(--document-viewer-font,16px system-ui);background:var(--document-viewer-paper,#fbfaf7);color:var(--document-viewer-ink,#222722)}html{scrollbar-color:var(--document-viewer-muted,#667168) var(--document-viewer-paper,#fbfaf7)}body{padding:12px;overflow:auto;box-sizing:border-box}canvas{display:block;max-width:100%;height:auto;margin:auto}a{pointer-events:none}.pdf-page,.msdoc-root,.docx-wrapper>section{background:#fff;color:#222}.docx-wrapper{padding:0!important;background:transparent!important}.docx-wrapper>section{box-sizing:border-box;margin:0 0 12px!important}:focus-visible{outline:3px solid var(--document-viewer-accent,#426b57);outline-offset:3px}</style></head><body></body></html>`;
  container.append(frame);
  await loaded;
  signal.throwIfAborted();
  frame.contentDocument.documentElement.lang = document.documentElement.lang || 'en';
  const pages = frame.contentDocument.createElement('style');
  pages.textContent = '.document-page{display:block;margin:0 0 24px;break-after:page;border-bottom:1px solid var(--document-viewer-line);padding-bottom:20px;max-width:100%;overflow-wrap:normal}.document-page-label{font:inherit;font-size:.875em;color:var(--document-viewer-muted);margin:0 0 12px}.document-page-text{background:#fff;color:#222;padding:16px}.document-page svg{display:block;background:#fff}.document-page .msdoc-root{padding:16px;box-sizing:border-box}.docx-wrapper>.document-page{width:max-content;max-width:none;padding:0 0 20px}.document-page>section.docx{max-width:none;box-sizing:border-box;margin:0!important}';
  frame.contentDocument.head.append(pages);
  const synchronize = () => {
    const theme = getComputedStyle(container);
    const target = frame.contentDocument.documentElement.style;
    for (const name of ['paper', 'surface', 'ink', 'muted', 'line', 'accent']) target.setProperty('--document-viewer-' + name, theme.getPropertyValue('--dv-' + name).trim());
    target.setProperty('--document-viewer-font', theme.font);
    target.colorScheme = theme.colorScheme;
  };
  synchronize();
  const observer = new MutationObserver(synchronize);
  for (let ancestor = container; ancestor; ancestor = ancestor.parentElement) observer.observe(ancestor, { attributes: true });
  const media = matchMedia('(prefers-color-scheme: dark)');
  media.addEventListener('change', synchronize);
  signal.addEventListener('abort', () => { observer.disconnect(); media.removeEventListener('change', synchronize); }, { once: true });
  return frame;
}
