import { renderAsync } from 'docx-preview';
import { labelPages } from '../pages.js';

export async function render({ data, frame, signal, labels, setCleanup }) {
  const styles = frame.contentDocument.createElement('div');
  const layout = frame.contentDocument.createElement('style');
  layout.textContent = '.docx-wrapper{align-items:flex-start!important;width:max-content;min-width:100%}.docx-wrapper>.document-page{margin-inline:auto!important}';
  frame.contentDocument.head.append(styles);
  const markImages = node => {
    if (node.nodeType !== 1) return;
    if (node.localName === 'img') node.setAttribute('loading', 'lazy');
    for (const image of node.querySelectorAll('img')) image.setAttribute('loading', 'lazy');
  };
  const images = new MutationObserver(entries => { for (const entry of entries) for (const node of entry.addedNodes) markImages(node); });
  images.observe(frame.contentDocument.body, { childList: true, subtree: true });
  setCleanup(() => images.disconnect());
  try {
    await renderAsync(data, frame.contentDocument.body, styles, {
      renderAltChunks: false, useBase64URL: true, renderComments: false,
      renderChanges: false, breakPages: true, ignoreWidth: false,
      ignoreLastRenderedPageBreak: false, experimental: true,
      renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true
    });
    markImages(frame.contentDocument.body);
  } finally { images.disconnect(); }
  frame.contentDocument.head.append(layout);
  signal.throwIfAborted();
  for (const anchor of frame.contentDocument.querySelectorAll('a')) anchor.removeAttribute('href');
  labelPages([...frame.contentDocument.querySelectorAll('section.docx')], labels);
}
