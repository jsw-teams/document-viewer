import { parseMsDoc, renderMsDoc } from '@file-viewer/doc';

export async function render({ data, frame, signal }) {
  const result = renderMsDoc(parseMsDoc(data.buffer), { externalLinkPolicy: 'block', externalResourcePolicy: 'block' });
  signal.throwIfAborted();
  const styles = frame.contentDocument.createElement('style');
  styles.textContent = result.css;
  frame.contentDocument.head.append(styles);
  const wrapper = frame.contentDocument.createElement('div');
  wrapper.className = 'msdoc-root';
  wrapper.innerHTML = result.html;
  frame.contentDocument.body.replaceChildren(wrapper);
}
