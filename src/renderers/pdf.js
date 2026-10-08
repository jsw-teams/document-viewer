import { getDocument, GlobalWorkerOptions, TextLayer, PDFDataRangeTransport } from 'pdfjs-dist/build/pdf.mjs';
import { documentPage } from '../pages.js';

GlobalWorkerOptions.workerSrc = new URL('../pdf.worker.js', import.meta.url).href;
const resources = typeof __DOCUMENT_PDF_ASSETS__ === 'string' ? __DOCUMENT_PDF_ASSETS__ : '../pdf-assets/';

export async function render({ data, frame, signal, labels, status, setCleanup }) {
  const doc = frame.contentDocument;
  let transport = null;
  if (data.readRange) {
    transport = new PDFDataRangeTransport(data.length, data.initialData);
    transport.requestDataRange = (begin, end) => {
      void data.readRange(begin, end).then(bytes => transport.onDataRange(begin, bytes)).catch(() => {
        if (!signal.aborted) { status.textContent = labels.error; void task.destroy().catch(() => {}); }
      });
    };
  }
  const task = getDocument({ ...(transport ? { range: transport, length: data.length, disableAutoFetch: true, disableStream: true, rangeChunkSize: 65536 } : { data }), ownerDocument: doc, isEvalSupported: false, stopAtErrors: true,
    cMapUrl: new URL(resources + 'cmaps/', import.meta.url).href, cMapPacked: true,
    standardFontDataUrl: new URL(resources + 'standard_fonts/', import.meta.url).href,
    wasmUrl: new URL(resources + 'wasm/', import.meta.url).href });
  let renderTask = null;
  let textLayer = null;
  let observer = null;
  let destroyed = false;
  const cache = new Map();
  const waiting = new Set();
  const nearby = new Set();
  let running = false;
  const cleanup = () => {
    if (destroyed) return;
    destroyed = true;
    observer?.disconnect();
    waiting.clear();
    renderTask?.cancel();
    textLayer?.cancel();
    for (const entry of cache.values()) URL.revokeObjectURL(entry.url);
    cache.clear();
    void task.destroy().catch(() => {});
  };
  setCleanup(cleanup);
  const pdf = await task.promise;
  signal.throwIfAborted();
  const style = doc.createElement('style');
  style.textContent = '.pdf-page{position:relative;margin:auto;max-width:100%}.textLayer{position:absolute;inset:0;overflow:clip;opacity:1;line-height:1;text-align:initial;forced-color-adjust:none;transform-origin:0 0;z-index:2}.textLayer :is(span,br){color:transparent;position:absolute;white-space:pre;cursor:text;transform-origin:0 0}.textLayer span.markedContent{top:0;height:0}.textLayer ::selection{background:Highlight;color:transparent}';
  doc.head.append(style);
  const first = await pdf.getPage(1);
  signal.throwIfAborted();
  const available = Math.max(1, frame.clientWidth - 24);
  const defaultViewport = first.getViewport({ scale: Math.min(2, available / first.getViewport({ scale: 1 }).width) });
  const pages = Array.from({ length: pdf.numPages }, (_, index) => {
    const section = documentPage(doc, labels, index, pdf.numPages);
    const wrapper = doc.createElement('div');
    wrapper.className = 'pdf-page';
    wrapper.style.width = defaultViewport.width + 'px';
    wrapper.style.height = defaultViewport.height + 'px';
    wrapper.setAttribute('aria-busy', 'true');
    section.append(wrapper);
    doc.body.append(section);
    return { section, wrapper, index };
  });
  async function draw(entry) {
    const page = await pdf.getPage(entry.index + 1);
    signal.throwIfAborted();
    const viewport = page.getViewport({ scale: Math.min(2, available / page.getViewport({ scale: 1 }).width) });
    const ratio = Math.min(2, Math.max(devicePixelRatio || 1, 600 / viewport.width), Math.sqrt(12000000 / (viewport.width * viewport.height)));
    const pixels = viewport.width * viewport.height * ratio * ratio;
    const wrapper = entry.wrapper;
    wrapper.style.width = viewport.width + 'px';
    wrapper.style.height = viewport.height + 'px';
    wrapper.style.setProperty('--scale-factor', viewport.scale);
    wrapper.style.setProperty('--total-scale-factor', viewport.scale);
    const canvas = doc.createElement('canvas');
    canvas.width = Math.ceil(viewport.width * ratio);
    canvas.height = Math.ceil(viewport.height * ratio);
    renderTask = page.render({ canvasContext: canvas.getContext('2d'), viewport, transform: [ratio, 0, 0, ratio, 0, 0] });
    await renderTask.promise;
    renderTask = null;
    signal.throwIfAborted();
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    if (!blob) throw new Error('PDF page image could not be created');
    canvas.width = 0;
    canvas.height = 0;
    signal.throwIfAborted();
    const url = URL.createObjectURL(blob);
    cache.set(entry.index, { url, pixels, busy: true });
    const painted = doc.createElement('img');
    painted.alt = '';
    painted.style.cssText = 'display:block;width:100%;height:100%';
    const loaded = new Promise((resolve, reject) => {
      const aborted = () => reject(signal.reason);
      signal.addEventListener('abort', aborted, { once: true });
      painted.onload = () => { signal.removeEventListener('abort', aborted); resolve(); };
      painted.onerror = () => { signal.removeEventListener('abort', aborted); reject(new Error('PDF page image could not be displayed')); };
    });
    painted.src = url;
    wrapper.replaceChildren(painted);
    await loaded;
    signal.throwIfAborted();
    const text = doc.createElement('div');
    text.className = 'textLayer';
    wrapper.append(text);
    textLayer = new TextLayer({ textContentSource: await page.getTextContent(), container: text, viewport });
    await textLayer.render();
    textLayer = null;
    signal.throwIfAborted();
    wrapper.setAttribute('aria-busy', 'false');
    page.cleanup();
    cache.get(entry.index).busy = false;
    if (observer && !nearby.has(entry.index)) {
      URL.revokeObjectURL(url);
      cache.delete(entry.index);
      wrapper.replaceChildren();
      wrapper.setAttribute('aria-busy', 'true');
      return;
    }
    let total = [...cache.values()].reduce((sum, item) => sum + item.pixels, 0);
    for (const [index, cached] of cache) {
      if (index === entry.index || (nearby.has(index) && cache.size <= 8 && total <= 32000000)) continue;
      URL.revokeObjectURL(cached.url);
      cache.delete(index);
      pages[index].wrapper.replaceChildren();
      pages[index].wrapper.setAttribute('aria-busy', 'true');
      total -= cached.pixels;
    }
  }
  await draw(pages[0]);
  async function drain() {
    if (running || destroyed) return;
    running = true;
    try {
      while (waiting.size && !destroyed) {
        const index = waiting.values().next().value;
        waiting.delete(index);
        if (nearby.has(index) && !cache.has(index)) await draw(pages[index]);
      }
    } catch {
      if (!signal.aborted && !destroyed) status.textContent = labels.error;
    } finally { running = false; }
  }
  observer = new IntersectionObserver(entries => {
    for (const entry of entries) {
      const index = Number(entry.target.dataset.documentPage) - 1;
      if (entry.isIntersecting) { nearby.add(index); waiting.add(index); }
      else { nearby.delete(index); waiting.delete(index); }
    }
    for (const [index, cached] of cache) if (!nearby.has(index) && !cached.busy) {
      URL.revokeObjectURL(cached.url);
      cache.delete(index);
      pages[index].wrapper.replaceChildren();
      pages[index].wrapper.setAttribute('aria-busy', 'true');
    }
    void drain();
  }, { root: doc, rootMargin: '600px 0px' });
  for (const page of pages) observer.observe(page.section);
  return cleanup;
}
