export async function documentFrame(container, title, signal) {
  signal.throwIfAborted();
  const frame = document.createElement('iframe');
  frame.className = 'document-viewer-frame';
  frame.title = title;
  frame.setAttribute('sandbox', 'allow-same-origin');
  frame.referrerPolicy = 'no-referrer';
  const loaded = new Promise((resolve, reject) => {
    const aborted = () => reject(signal.reason || new DOMException('Aborted', 'AbortError'));
    signal.addEventListener('abort', aborted, { once: true });
    frame.addEventListener('load', () => { signal.removeEventListener('abort', aborted); resolve(); }, { once: true });
  });
  frame.srcdoc = `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src data: blob:; base-uri 'none'; form-action 'none'"><style>html,body{margin:0;padding:0;font:16px system-ui;background:#fff;color:#222}body{padding:12px;overflow:auto;box-sizing:border-box}canvas{display:block;max-width:100%;height:auto;margin:auto}table{border-collapse:collapse;min-width:100%}th,td{border:1px solid #ddd;padding:8px;white-space:pre-wrap;text-align:left}th{background:#eee}a{pointer-events:none}.docx-wrapper{padding:0!important;background:transparent!important}.docx-wrapper>section{max-width:100%;box-sizing:border-box;margin:0 auto 12px!important}</style></head><body></body></html>`;
  container.append(frame);
  await loaded;
  signal.throwIfAborted();
  return frame;
}
