export function previewChrome({ root, toolbar, frame, signal, labels, fitWidth = false }) {
  const tools = document.createElement('div');
  tools.className = 'document-viewer-view-tools';
  const percentage = document.createElement('output');
  percentage.textContent = '100%';
  let zoom = 1;
  const button = (label, text, action) => {
    const node = document.createElement('button');
    node.type = 'button';
    node.setAttribute('aria-label', label);
    node.textContent = text;
    node.addEventListener('click', action, { signal });
    tools.append(node);
    return node;
  };
  const update = () => {
    frame.contentDocument.body.style.zoom = String(zoom);
    percentage.textContent = Math.round(zoom * 100) + '%';
    smaller.disabled = zoom <= 0.1;
    larger.disabled = zoom >= 2;
  };
  const smaller = button(labels.zoomOut, '−', () => { zoom = Math.max(0.1, zoom - 0.25); update(); });
  tools.append(percentage);
  const larger = button(labels.zoomIn, '+', () => { zoom = Math.min(2, zoom + 0.25); update(); });
  const fit = () => {
    const body = frame.contentDocument.body;
    body.style.zoom = '1';
    const width = Math.max(frame.clientWidth, body.scrollWidth);
    zoom = Math.max(0.1, Math.min(1, frame.clientWidth / width));
    update();
    frame.contentWindow.scrollTo(0, frame.contentWindow.scrollY);
  };
  button(labels.fitWidth, labels.fitWidth, fit);
  if (fitWidth) fit();
  toolbar.append(tools);
  const position = document.createElement('p');
  position.className = 'document-viewer-position';
  position.setAttribute('role', 'status');
  root.append(position);
  const doc = frame.contentDocument;
  let scheduled = null;
  let pages = [...doc.querySelectorAll('[data-document-page]')];
  const report = () => {
    scheduled = null;
    let begin = 0;
    let end = pages.length;
    while (begin < end) {
      const middle = Math.floor((begin + end) / 2);
      if (pages[middle].getBoundingClientRect().bottom <= 0) begin = middle + 1;
      else end = middle;
    }
    const current = pages[begin];
    if (current) position.textContent = labels.page + ' ' + current.dataset.documentPage + ' / ' + pages.length;
  };
  const queue = () => { if (scheduled === null) scheduled = requestAnimationFrame(report); };
  frame.contentWindow.addEventListener('scroll', queue, { signal, passive: true });
  const observer = new MutationObserver(entries => {
    if (entries.some(entry => [...entry.addedNodes, ...entry.removedNodes].some(node => node.nodeType === 1 && node.matches('[data-document-page]')))) pages = [...doc.querySelectorAll('[data-document-page]')];
    queue();
  });
  observer.observe(doc.body, { childList: true, subtree: true });
  report();
  return () => { observer.disconnect(); if (scheduled !== null) cancelAnimationFrame(scheduled); tools.remove(); position.remove(); };
}
