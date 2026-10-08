export function inertDocumentHtml(doc, html) {
  const inert = document.implementation.createHTMLDocument('');
  const template = inert.createElement('template');
  template.innerHTML = html;
  for (const node of template.content.querySelectorAll('script,iframe,object,embed,meta,link,base')) node.remove();
  for (const node of template.content.querySelectorAll('*')) {
    if (node.localName === 'img') node.setAttribute('loading', 'lazy');
    for (const attribute of [...node.attributes]) {
      if (/^on/i.test(attribute.name) || attribute.name === 'srcdoc' || (/^(?:href|src|action|formaction|xlink:href)$/i.test(attribute.name) && /^\s*(?:javascript|vbscript):/i.test(attribute.value))) node.removeAttribute(attribute.name);
    }
  }
  return doc.importNode(template.content, true);
}
