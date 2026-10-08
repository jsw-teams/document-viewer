import { parseMsDoc, renderMsDoc } from '@file-viewer/doc';
import { documentPage } from '../pages.js';
import { inertDocumentHtml } from '../safe-html.js';

export async function render({ data, frame, signal, labels }) {
  const result = renderMsDoc(parseMsDoc(data.buffer), { externalLinkPolicy: 'block', externalResourcePolicy: 'block' });
  signal.throwIfAborted();
  const styles = frame.contentDocument.createElement('style');
  styles.textContent = result.css;
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
    paper.append(content);
    page.append(paper);
    frame.contentDocument.body.append(page);
  });
}
