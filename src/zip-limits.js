export function validateOfficeZip(data, streaming = false) {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let footer = -1;
  for (let offset = data.byteLength - 22; offset >= Math.max(0, data.byteLength - 65557); offset--) {
    if (view.getUint32(offset, true) === 0x06054b50 && offset + 22 + view.getUint16(offset + 20, true) === data.byteLength) { footer = offset; break; }
  }
  if (footer < 0) throw new Error('Invalid Office ZIP archive');
  const count = view.getUint16(footer + 10, true);
  let offset = view.getUint32(footer + 16, true);
  if (count > 4096 || view.getUint16(footer + 4, true) || view.getUint16(footer + 6, true)) throw new Error('Office archive exceeds safety limits');
  let total = 0;
  for (let entry = 0; entry < count; entry++) {
    if (offset + 46 > footer || view.getUint32(offset, true) !== 0x02014b50) throw new Error('Invalid Office ZIP directory');
    const size = view.getUint32(offset + 24, true);
    total += size;
    if (size > (streaming ? 512 : 32) * 1024 * 1024 || total > (streaming ? 1024 : 128) * 1024 * 1024 || (view.getUint16(offset + 8, true) & 1)) {
      throw new Error('Encrypted or oversized Office archive');
    }
    offset += 46 + view.getUint16(offset + 28, true) + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
  }
}
