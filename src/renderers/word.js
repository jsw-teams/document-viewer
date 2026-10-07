import { renderAsync } from 'docx-preview';

export async function render({ data, frame, signal }) {
  const styles = frame.contentDocument.createElement('div');
  const layout = frame.contentDocument.createElement('style');
  layout.textContent = '.docx-wrapper{align-items:stretch!important}.docx-wrapper>section{width:100%!important}';
  frame.contentDocument.head.append(styles);
  await renderAsync(data, frame.contentDocument.body, styles, {
    renderAltChunks: false, useBase64URL: true, renderComments: false,
    renderChanges: false, breakPages: true, ignoreWidth: true
  });
  frame.contentDocument.head.append(layout);
  signal.throwIfAborted();
  for (const anchor of frame.contentDocument.querySelectorAll('a')) anchor.removeAttribute('href');
}
