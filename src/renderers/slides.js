import { PptxViewer, RECOMMENDED_ZIP_LIMITS } from '@aiden0z/pptx-renderer';

export async function render({ data, frame, signal, setCleanup }) {
  const viewer = new PptxViewer(frame.contentDocument.body, {
    width: Math.max(240, frame.clientWidth - 24), fitMode: 'contain', pdfjs: false,
    zipLimits: RECOMMENDED_ZIP_LIMITS, lazySlides: true, lazyMedia: true
  });
  const cleanup = () => viewer.destroy();
  setCleanup(cleanup);
  await viewer.open(data.buffer, { signal, renderMode: 'list', listOptions: { windowed: true, batchSize: 4, initialSlides: 2 } });
  signal.throwIfAborted();
  for (const anchor of frame.contentDocument.querySelectorAll('a')) anchor.removeAttribute('href');
  return cleanup;
}
