import CFB from 'cfb';

function view(bytes) {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return new DataView(data.buffer, data.byteOffset, data.byteLength);
}

export function docPageLayout(input) {
  const file = CFB.read(input, { type: 'array' });
  const word = CFB.find(file, 'WordDocument')?.content;
  if (!word || word.length < 34) throw new Error('DOC WordDocument stream is missing');
  const data = view(word);
  const table = CFB.find(file, data.getUint16(10, true) & 512 ? '1Table' : '0Table')?.content;
  if (!table) throw new Error('DOC Table stream is missing');
  let position = 32;
  const size = () => {
    if (position + 2 > word.length) throw new Error('Truncated DOC FIB');
    const count = data.getUint16(position, true);
    position += 2;
    return count;
  };
  const words = size();
  position += words * 2;
  const longs = size();
  position += longs * 4;
  const count = size();
  if (count < 7 || position + count * 8 > word.length) throw new Error('Invalid DOC FIB properties');
  const offset = data.getUint32(position + 48, true);
  const length = data.getUint32(position + 52, true);
  const layout = { width: 12240, height: 15840, left: 1440, right: 1440, top: 1440, bottom: 1440 };
  if (!length) return layout;
  const sections = (length - 4) / 16;
  if (!Number.isInteger(sections) || sections < 1 || sections > 4096 || offset + length > table.length) throw new Error('Invalid DOC section directory');
  const properties = view(table).getInt32(offset + (sections + 1) * 4 + 2, true);
  if (properties === -1) return layout;
  if (properties < 0 || properties + 2 > word.length) throw new Error('Invalid DOC section properties');
  const end = properties + 2 + data.getUint16(properties, true);
  if (end > word.length) throw new Error('Truncated DOC section properties');
  const fields = { 45087: 'width', 45088: 'height', 45089: 'left', 45090: 'right', 36899: 'top', 36900: 'bottom' };
  for (position = properties + 2; position < end;) {
    if (position + 2 > end) throw new Error('Truncated DOC section property');
    const operation = data.getUint16(position, true);
    position += 2;
    let bytes = [1, 1, 2, 4, 2, 2, 0, 3][operation >>> 13];
    if (!bytes) { if (position >= end) throw new Error('Truncated DOC variable property'); bytes = data.getUint8(position++); }
    if (position + bytes > end) throw new Error('Truncated DOC property operand');
    const field = fields[operation];
    if (field) layout[field] = field === 'top' || field === 'bottom' ? Math.abs(data.getInt16(position, true)) : data.getUint16(position, true);
    position += bytes;
  }
  if (layout.width < 144 || layout.width > 31680 || layout.height < 144 || layout.height > 31680 || layout.left + layout.right >= layout.width || layout.top + layout.bottom >= layout.height) throw new Error('Invalid DOC page geometry');
  return layout;
}
