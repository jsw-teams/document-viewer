export async function readDocument(address, signal, maxBytes) {
  const response = await fetch(address, { signal, credentials: 'omit', referrerPolicy: 'no-referrer', redirect: 'error' });
  if (!response.ok) throw new Error('Document request failed: ' + response.status);
  const length = Number(response.headers.get('content-length'));
  if (Number.isFinite(length) && length > maxBytes) throw new Error('Document exceeds size limit');
  const reader = response.body.getReader();
  const chunks = [];
  let total = 0;
  try {
    while (true) {
      signal.throwIfAborted();
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maxBytes) throw new Error('Document exceeds size limit');
      chunks.push(value);
    }
  } catch (error) { await reader.cancel().catch(() => {}); throw error; }
  finally { reader.releaseLock(); }
  if (!total) throw new Error('Document response is empty or was intercepted');
  const data = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.byteLength; }
  return data;
}
