export function documentPage(doc, labels, index, total) {
  const page = doc.createElement('section');
  page.className = 'document-page';
  page.dataset.documentPage = String(index + 1);
  const heading = doc.createElement('h2');
  heading.className = 'document-page-label';
  heading.textContent = `${labels.page} ${index + 1} / ${total}`;
  page.setAttribute('aria-label', heading.textContent);
  page.append(heading);
  return page;
}

export function labelPages(nodes, labels) {
  nodes.forEach((node, index) => {
    if (node.parentElement?.matches('[data-document-page]')) return;
    const page = documentPage(node.ownerDocument, labels, index, nodes.length);
    node.before(page);
    page.append(node);
  });
}
