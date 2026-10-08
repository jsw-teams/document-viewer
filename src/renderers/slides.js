import { PptxViewer, RECOMMENDED_ZIP_LIMITS } from '@aiden0z/pptx-renderer';
import { labelPages } from '../pages.js';

export async function render({ data, frame, signal, labels, setCleanup }) {
  const viewer = new PptxViewer(frame.contentDocument.body, {
    fitMode: 'contain', pdfjs: false, scrollContainer: frame.contentDocument.documentElement,
    zipLimits: RECOMMENDED_ZIP_LIMITS, lazySlides: true, lazyMedia: true
  });
  const cleanup = () => viewer.destroy();
  viewer.addEventListener('sliderendered', event => {
    for (const anchor of event.detail.element.querySelectorAll('a')) anchor.removeAttribute('href');
    labelPages([...frame.contentDocument.querySelectorAll('[data-slide-index]')], labels);
  });
  setCleanup(cleanup);
  await viewer.open(data.buffer, { signal, renderMode: 'list', listOptions: { windowed: true, batchSize: 4, initialSlides: 1, overscanViewport: 0.25, showSlideLabels: false } });
  signal.throwIfAborted();
  for (const anchor of frame.contentDocument.querySelectorAll('a')) anchor.removeAttribute('href');
  labelPages([...frame.contentDocument.querySelectorAll('[data-slide-index]')], labels);
  return cleanup;
}
