import { documentPage } from '../pages.js';
import { inertDocumentHtml } from '../safe-html.js';

export async function render({ data, frame, signal, labels, setCleanup }) {
  const worker = new Worker(new URL('../doc.worker.js', import.meta.url), { type: 'module' });
  setCleanup(() => worker.terminate());
  const result = await new Promise((resolve, reject) => {
    const aborted = () => reject(signal.reason || new DOMException('Aborted', 'AbortError'));
    signal.addEventListener('abort', aborted, { once: true });
    worker.onmessage = ({ data }) => {
      signal.removeEventListener('abort', aborted);
      if (data.error) reject(new Error(data.error));
      else resolve(data);
    };
    worker.onerror = () => { signal.removeEventListener('abort', aborted); reject(new Error('DOC parser Worker failed')); };
    worker.postMessage(data.buffer, [data.buffer]);
  });
  worker.terminate();
  const { layout } = result;
  signal.throwIfAborted();
  const styles = frame.contentDocument.createElement('style');
  styles.textContent = result.css;
  styles.textContent += '.document-page .msdoc-root{max-width:none;padding:0}.msdoc-paragraph{overflow-wrap:normal;word-break:normal}.msdoc-table{max-width:none}';
  frame.contentDocument.head.append(styles);
  const wrapper = frame.contentDocument.createElement('div');
  wrapper.append(inertDocumentHtml(frame.contentDocument, result.html));
  const breaks = [...wrapper.querySelectorAll('.msdoc-page-break')];
  const pages = [];
  for (const boundary of breaks) {
    const range = frame.contentDocument.createRange();
    range.setStart(wrapper, 0);
    range.setEndBefore(boundary);
    pages.push(range.extractContents());
    boundary.remove();
  }
  pages.push(wrapper);
  frame.contentDocument.body.replaceChildren();
  pages.forEach((content, index) => {
    const page = documentPage(frame.contentDocument, labels, index, pages.length);
    const paper = frame.contentDocument.createElement('div');
    paper.className = 'msdoc-root';
    paper.style.width = layout.width / 15 + 'px';
    paper.style.minHeight = layout.height / 15 + 'px';
    paper.style.padding = [layout.top, layout.right, layout.bottom, layout.left].map(value => value / 15 + 'px').join(' ');
    page.style.width = 'max-content';
    page.style.maxWidth = 'none';
    page.style.marginInline = 'auto';
    paper.append(content);
    page.append(paper);
    frame.contentDocument.body.append(page);
  });
}
