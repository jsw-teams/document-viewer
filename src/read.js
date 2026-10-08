export async function readDocument(address, signal, maxBytes) {
  const response = await fetch(address, { signal, credentials: 'omit', referrerPolicy: 'no-referrer', redirect: 'error' });
  return readResponse(response, signal, maxBytes);
}

async function readResponse(response, signal, maxBytes) {
  if (!response.ok) { await response.body?.cancel(); throw new Error('Document request failed: ' + response.status); }
  const length = Number(response.headers.get('content-length'));
  if (Number.isFinite(length) && length > maxBytes) { await response.body?.cancel(); throw new Error('Document exceeds size limit'); }
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

export async function readPdfDocument(address, signal, maxBytes) {
  const options = { signal, credentials: 'omit', referrerPolicy: 'no-referrer', redirect: 'error' };
  const response = await fetch(address, { ...options, headers: { Range: 'bytes=0-65535' } });
  if (response.status !== 206) return readResponse(response, signal, Math.min(maxBytes, 50 * 1024 * 1024));
  const range = response.headers.get('content-range')?.match(/^bytes 0-(\d+)\/(\d+)$/);
  if (!range || Number(range[1]) >= 65536 || !Number.isSafeInteger(Number(range[2])) || Number(range[2]) > maxBytes) { await response.body.cancel(); throw new Error('Invalid PDF range response'); }
  const length = Number(range[2]);
  const initialData = await readResponse(response, signal, 65536);
  if (initialData.length !== Number(range[1]) + 1 || length < initialData.length) throw new Error('Truncated PDF range');
  const etag = response.headers.get('etag');
  return { length, initialData, async readRange(begin, end) {
    signal.throwIfAborted();
    const part = await fetch(address, { ...options, headers: { Range: 'bytes=' + begin + '-' + (end - 1) } });
    if (part.status !== 206 || part.headers.get('content-range') !== 'bytes ' + begin + '-' + (end - 1) + '/' + length || (etag && part.headers.get('etag') !== etag)) { await part.body?.cancel(); throw new Error('PDF changed or range request failed'); }
    const bytes = await readResponse(part, signal, end - begin);
    if (bytes.length !== end - begin) throw new Error('Truncated PDF range');
    return bytes;
  } };
}
