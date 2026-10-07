const namespace = 'http://www.w3.org/2000/svg';

export async function render({ data, frame, controls, signal, labels, status, setCleanup, guard }) {
  const worker = new Worker(new URL('../ppt.worker.js', import.meta.url), { type: 'module' });
  const urls = [];
  const cleanup = () => { worker.terminate(); for (const url of urls) URL.revokeObjectURL(url); };
  setCleanup(cleanup);
  const presentation = await new Promise((resolve, reject) => {
    const aborted = () => reject(signal.reason || new DOMException('Aborted', 'AbortError'));
    signal.addEventListener('abort', aborted, { once: true });
    worker.onmessage = event => {
      signal.removeEventListener('abort', aborted);
      if (event.data.error) reject(new Error(event.data.error));
      else resolve(event.data.presentation);
    };
    worker.onerror = () => { signal.removeEventListener('abort', aborted); reject(new Error('PPT parser Worker failed')); };
    worker.postMessage(data.buffer, [data.buffer]);
  });
  worker.terminate();
  signal.throwIfAborted();
  const imageUrls = new Map();
  const images = presentation.pictures.map(picture => {
    if (!picture) return null;
    if (imageUrls.has(picture)) return imageUrls.get(picture);
    const url = URL.createObjectURL(new Blob([picture.bytes], { type: picture.mime }));
    urls.push(url);
    imageUrls.set(picture, url);
    return url;
  });
  let current = 0;
  const previous = document.createElement('button');
  const next = document.createElement('button');
  const position = document.createElement('span');
  previous.type = next.type = 'button';
  previous.textContent = labels.previous;
  next.textContent = labels.next;
  position.setAttribute('aria-live', 'polite');
  controls.append(previous, position, next);
  const create = (name, attributes = {}) => {
    const node = frame.contentDocument.createElementNS(namespace, name);
    for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, String(value));
    return node;
  };
  const draw = guard(() => {
    const slide = presentation.slides[current];
    const svg = create('svg', { viewBox: `0 0 ${presentation.width} ${presentation.height}`, width: '100%', role: 'img', 'aria-label': labels.page + ' ' + (current + 1) });
    svg.append(create('rect', { width: presentation.width, height: presentation.height, fill: '#fff' }));
    const represented = new Set();
    for (const shape of slide.shapes) {
      const group = create('g', { transform: `rotate(${shape.rotation} ${shape.left + shape.width / 2} ${shape.top + shape.height / 2})` });
      const common = { fill: shape.fill, stroke: shape.stroke, 'stroke-width': presentation.width / 700 };
      if (shape.image && images[shape.image - 1]) {
        group.append(create('image', { href: images[shape.image - 1], x: shape.left, y: shape.top, width: shape.width, height: shape.height, preserveAspectRatio: 'xMidYMid meet' }));
      } else if (shape.type === 3) {
        group.append(create('ellipse', { ...common, cx: shape.left + shape.width / 2, cy: shape.top + shape.height / 2, rx: shape.width / 2, ry: shape.height / 2 }));
      } else if (shape.type === 5) {
        group.append(create('polygon', { ...common, points: `${shape.left + shape.width / 2},${shape.top} ${shape.left + shape.width},${shape.top + shape.height} ${shape.left},${shape.top + shape.height}` }));
      } else if (shape.type === 20) {
        group.append(create('line', { ...common, x1: shape.left, y1: shape.top, x2: shape.left + shape.width, y2: shape.top + shape.height }));
      } else if ([1, 2, 202].includes(shape.type) || shape.background) {
        group.append(create('rect', { ...common, x: shape.left, y: shape.top, width: shape.width, height: shape.height, rx: shape.type === 2 ? Math.min(shape.width, shape.height) / 10 : 0 }));
      }
      if (shape.texts.length) {
        const box = create('foreignObject', { x: shape.left, y: shape.top, width: shape.width, height: shape.height });
        const text = frame.contentDocument.createElementNS('http://www.w3.org/1999/xhtml', 'div');
        text.style.cssText = `font:${presentation.width / 30}px system-ui;white-space:pre-wrap;overflow-wrap:anywhere;color:#222;padding:8px;box-sizing:border-box`;
        text.textContent = shape.texts.map(item => item.text).join('\n');
        box.append(text);
        group.append(box);
        for (const item of shape.texts) represented.add(item.text);
      }
      svg.append(group);
    }
    const transcript = frame.contentDocument.createElement('section');
    transcript.setAttribute('aria-label', labels.page + ' ' + (current + 1));
    for (const item of slide.texts) {
      if (!item.text || represented.has(item.text)) continue;
      const paragraph = frame.contentDocument.createElement([0, 6].includes(item.role) ? 'h2' : 'p');
      paragraph.style.whiteSpace = 'pre-wrap';
      paragraph.textContent = item.text;
      transcript.append(paragraph);
    }
    frame.contentDocument.body.replaceChildren(svg, transcript);
    position.textContent = labels.page + ' ' + (current + 1) + ' / ' + presentation.slides.length;
    previous.disabled = current === 0;
    next.disabled = current === presentation.slides.length - 1;
  });
  previous.addEventListener('click', () => { current = Math.max(0, current - 1); draw(); });
  next.addEventListener('click', () => { current = Math.min(presentation.slides.length - 1, current + 1); draw(); });
  status.textContent = labels.pptLimited;
  draw();
  return cleanup;
}
