const shapes = {
  eye: [['path', { d: 'M2.062 12.348a1 1 0 0 1 0-.696 10.75 10.75 0 0 1 19.876 0 1 1 0 0 1 0 .696 10.75 10.75 0 0 1-19.876 0' }], ['circle', { cx: 12, cy: 12, r: 3 }]],
  x: [['path', { d: 'M18 6 6 18' }], ['path', { d: 'm6 6 12 12' }]],
  expand: [['path', { d: 'M15 3h6v6m0-6-7 7M3 21l7-7M9 21H3v-6' }]],
  collapse: [['path', { d: 'm14 10 7-7M20 10h-6V4m-11 17 7-7M4 14h6v6' }]],
  zoomIn: [['circle', { cx: 11, cy: 11, r: 8 }], ['path', { d: 'm21 21-4.35-4.35M11 8v6M8 11h6' }]],
  zoomOut: [['circle', { cx: 11, cy: 11, r: 8 }], ['path', { d: 'm21 21-4.35-4.35M8 11h6' }]],
  fit: [['path', { d: 'm18 8 4 4-4 4M2 12h20M6 8l-4 4 4 4' }]],
  locate: [['path', { d: 'M2 12h3m14 0h3M12 2v3m0 14v3' }], ['circle', { cx: 12, cy: 12, r: 7 }], ['circle', { cx: 12, cy: 12, r: 3 }]],
  columns: [['rect', { x: 3, y: 3, width: 18, height: 18, rx: 2 }], ['path', { d: 'M12 3v18' }]],
  reset: [['path', { d: 'M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8M3 3v5h5' }]],
  table: [['path', { d: 'M3 9h18M9 3v18' }], ['rect', { x: 3, y: 3, width: 18, height: 18, rx: 2 }]],
  grip: [9, 15].flatMap(cx => [5, 12, 19].map(cy => ['circle', { cx, cy, r: 1 }]))
};

export function controlIcon(doc, name) {
  if (!Object.hasOwn(shapes, name)) throw new Error('Unknown control icon');
  const namespace = 'http://www.w3.org/2000/svg';
  const icon = doc.createElementNS(namespace, 'svg');
  for (const [attribute, value] of Object.entries({ viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true', focusable: 'false', class: 'document-viewer-icon', 'data-icon': name })) icon.setAttribute(attribute, value);
  for (const [tag, attributes] of shapes[name]) {
    const shape = doc.createElementNS(namespace, tag);
    for (const [attribute, value] of Object.entries(attributes)) shape.setAttribute(attribute, value);
    icon.append(shape);
  }
  return icon;
}

export function buttonContent(button, label, name, iconOnly = false) {
  button.setAttribute('aria-label', label);
  button.title = label;
  const content = [controlIcon(button.ownerDocument, name)];
  if (!iconOnly) {
    const text = button.ownerDocument.createElement('span');
    text.textContent = label;
    content.push(text);
  }
  button.replaceChildren(...content);
}
