import { renderAsync } from 'docx-preview';

export async function render({ data, frame, signal }) {
  const styles = frame.contentDocument.createElement('div');
  frame.contentDocument.head.append(styles);
  await renderAsync(data, frame.contentDocument.body, styles, {
    renderAltChunks: false, useBase64URL: true, renderComments: false,
    renderChanges: false, breakPages: true, ignoreWidth: true
  });
  signal.throwIfAborted();
  for (const anchor of frame.contentDocument.querySelectorAll('a')) anchor.removeAttribute('href');
}
