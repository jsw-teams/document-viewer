import { getDocument, GlobalWorkerOptions, TextLayer } from 'pdfjs-dist/build/pdf.mjs';

GlobalWorkerOptions.workerSrc = new URL('../pdf.worker.js', import.meta.url).href;
const resources = typeof __DOCUMENT_PDF_ASSETS__ === 'string' ? __DOCUMENT_PDF_ASSETS__ : '../pdf-assets/';

export async function render({ data, frame, controls, signal, labels, status, setCleanup, guard }) {
  const task = getDocument({ data, isEvalSupported: false, stopAtErrors: true,
    cMapUrl: new URL(resources + 'cmaps/', import.meta.url).href, cMapPacked: true,
    standardFontDataUrl: new URL(resources + 'standard_fonts/', import.meta.url).href,
    wasmUrl: new URL(resources + 'wasm/', import.meta.url).href });
  let renderTask = null;
  let textLayer = null;
  let destroyed = false;
  const cleanup = () => {
    if (destroyed) return;
    destroyed = true;
    renderTask?.cancel();
    textLayer?.cancel();
    void task.destroy().catch(() => {});
  };
  setCleanup(cleanup);
  const pdf = await task.promise;
  signal.throwIfAborted();
  let current = 1;
  let busy = false;
  const previous = document.createElement('button');
  const next = document.createElement('button');
  const position = document.createElement('span');
  previous.type = next.type = 'button';
  previous.textContent = labels.previous;
  next.textContent = labels.next;
  position.setAttribute('aria-live', 'polite');
  controls.append(previous, position, next);
  const style = frame.contentDocument.createElement('style');
  style.textContent = '.pdf-page{position:relative;margin:auto}.textLayer{position:absolute;inset:0;overflow:clip;opacity:1;line-height:1;text-align:initial;forced-color-adjust:none;transform-origin:0 0;z-index:2}.textLayer :is(span,br){color:transparent;position:absolute;white-space:pre;cursor:text;transform-origin:0 0}.textLayer span.markedContent{top:0;height:0}.textLayer ::selection{background:Highlight;color:transparent}';
  frame.contentDocument.head.append(style);
  async function draw() {
    busy = true;
    previous.disabled = next.disabled = true;
    try {
      const page = await pdf.getPage(current);
      signal.throwIfAborted();
      const available = Math.max(220, frame.clientWidth - 28);
      const viewport = page.getViewport({ scale: Math.min(2, available / page.getViewport({ scale: 1 }).width) });
      const ratio = Math.min(2, devicePixelRatio || 1);
      if (viewport.width * viewport.height * ratio * ratio > 12000000) throw new Error('PDF page exceeds rendering limits');
      const wrapper = frame.contentDocument.createElement('div');
      wrapper.className = 'pdf-page';
      wrapper.style.width = viewport.width + 'px';
      wrapper.style.height = viewport.height + 'px';
      wrapper.style.setProperty('--scale-factor', viewport.scale);
      wrapper.style.setProperty('--total-scale-factor', viewport.scale);
      const canvas = frame.contentDocument.createElement('canvas');
      canvas.width = Math.ceil(viewport.width * ratio);
      canvas.height = Math.ceil(viewport.height * ratio);
      canvas.style.width = viewport.width + 'px';
      canvas.style.height = viewport.height + 'px';
      wrapper.append(canvas);
      frame.contentDocument.body.replaceChildren(wrapper);
      renderTask = page.render({ canvasContext: canvas.getContext('2d'), viewport, transform: [ratio, 0, 0, ratio, 0, 0] });
      await renderTask.promise;
      signal.throwIfAborted();
      const text = frame.contentDocument.createElement('div');
      text.className = 'textLayer';
      wrapper.append(text);
      textLayer = new TextLayer({ textContentSource: await page.getTextContent(), container: text, viewport });
      await textLayer.render();
      signal.throwIfAborted();
      position.textContent = labels.page + ' ' + current + ' / ' + pdf.numPages;
    } finally {
      busy = false;
      if (!signal.aborted) { previous.disabled = current === 1; next.disabled = current === pdf.numPages; }
    }
  }
  const move = direction => guard(async () => {
    if (busy) return;
    current = Math.max(1, Math.min(pdf.numPages, current + direction));
    try { await draw(); } catch { if (!signal.aborted) status.textContent = labels.error; }
  });
  previous.addEventListener('click', move(-1));
  next.addEventListener('click', move(1));
  await draw();
  return cleanup;
}
