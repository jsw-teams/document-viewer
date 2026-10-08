import { inertDocumentHtml } from './safe-html.js';

export async function windowDocumentPages(frame, signal) {
  const doc = frame.contentDocument;
  const pages = [...doc.querySelectorAll('[data-document-page]')];
  const source = new Map();
  const mounted = new Set();
  const nearby = new Set();
  let stopped = false;
  let running = false;
  let observer;
  const activateImages = page => {
    for (const image of page.querySelectorAll('image[data-document-image-url]')) {
      const address = image.getAttribute('data-document-image-url');
      if (address.startsWith('blob:') && new URL(address.slice(5)).origin === location.origin) image.setAttribute('href', address);
    }
  };
  const cleanup = () => { stopped = true; observer?.disconnect(); source.clear(); nearby.clear(); mounted.clear(); };
  signal.addEventListener('abort', cleanup, { once: true });
  try {
    for (const [index, page] of pages.entries()) {
      signal.throwIfAborted();
      const rectangle = page.getBoundingClientRect();
      const height = Math.max(1, rectangle.height);
      const html = page.innerHTML;
      const stream = new Blob([html]).stream().pipeThrough(new CompressionStream('gzip'));
      source.set(index, new Uint8Array(await new Response(stream).arrayBuffer()));
      signal.throwIfAborted();
      page.style.minHeight = height + 'px';
      page.style.boxSizing = 'border-box';
      if (rectangle.bottom >= -200 && rectangle.top <= frame.clientHeight + 200) { activateImages(page); mounted.add(index); nearby.add(index); }
      else { page.replaceChildren(); page.setAttribute('aria-busy', 'true'); }
    }
    async function update() {
      if (running || stopped) return;
      running = true;
      try {
        while (!stopped) {
          for (const index of mounted) if (!nearby.has(index)) {
            pages[index].replaceChildren();
            pages[index].setAttribute('aria-busy', 'true');
            mounted.delete(index);
          }
          const index = [...nearby].find(current => !mounted.has(current));
          if (index === undefined) break;
          const stream = new Blob([source.get(index)]).stream().pipeThrough(new DecompressionStream('gzip'));
          const html = await new Response(stream).text();
          if (stopped) break;
          if (!nearby.has(index)) continue;
          pages[index].replaceChildren(inertDocumentHtml(doc, html));
          activateImages(pages[index]);
          pages[index].setAttribute('aria-busy', 'false');
          mounted.add(index);
        }
      } finally { running = false; }
    }
    observer = new IntersectionObserver(entries => {
      for (const entry of entries) {
        const index = Number(entry.target.dataset.documentPage) - 1;
        if (entry.isIntersecting) nearby.add(index); else nearby.delete(index);
      }
      void update().catch(() => cleanup());
    }, { root: doc, rootMargin: '200px 0px' });
    for (const page of pages) observer.observe(page);
    return cleanup;
  } catch (error) { cleanup(); throw error; }
}
