import { workbookSource, sheetInfo, sheetWindow } from './model.js';

let source = null;
let current = 0;
let queue = Promise.resolve();
self.onmessage = ({ data }) => {
  current = data.id;
  queue = queue.catch(() => {}).then(async () => {
    const cancelled = () => data.id !== current;
    if (cancelled()) return;
    try {
      let result;
      if (data.action === 'open') { source = await workbookSource(new Uint8Array(data.buffer), data.format); result = source.names; }
      else if (data.action === 'sheet') result = await sheetInfo(source, data.index, cancelled);
      else if (data.action === 'window') result = await sheetWindow(source, data.index, data.range, cancelled);
      else throw new Error('Unknown worksheet action');
      if (!cancelled()) self.postMessage({ id: data.id, result });
    } catch (error) { if (!cancelled()) self.postMessage({ id: data.id, error: error.message }); }
  });
};
