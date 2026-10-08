import CFB from 'cfb';

const types = {
  document: 1000, documentAtom: 1001, slide: 1006, slidePersist: 1011, colorScheme: 2032,
  textReference: 3998, textHeader: 3999, textChars: 4000, textBytes: 4008,
  slideList: 4080, userEdit: 4085, currentUser: 4086, persist: 6002,
  shape: 0xf004, pictureStore: 0xf001, pictureEntry: 0xf007,
  shapeProperties: 0xf00a, options: 0xf00b, textbox: 0xf00d, childAnchor: 0xf00f, anchor: 0xf010
};

function binary(bytes) {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return { data, view: new DataView(data.buffer, data.byteOffset, data.byteLength) };
}

function header(stream, offset, end = stream.data.length) {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset + 8 > end) throw new Error('Truncated PPT record');
  const flags = stream.view.getUint16(offset, true);
  const size = stream.view.getUint32(offset + 4, true);
  const record = { offset, start: offset + 8, end: offset + 8 + size,
    type: stream.view.getUint16(offset + 2, true), version: flags & 15, instance: flags >>> 4, size };
  if (record.end > end || record.end < record.start) throw new Error('PPT record exceeds its container');
  return record;
}

function records(stream, start, end, state) {
  const result = [];
  for (let offset = start; offset < end;) {
    if (++state.records > 100000) throw new Error('Too many PPT records');
    const record = header(stream, offset, end);
    result.push(record);
    offset = record.end;
  }
  return result;
}

function descendants(stream, container, state, wanted, depth = 0) {
  if (depth > 32) throw new Error('PPT record nesting exceeds limit');
  const found = [];
  for (const record of records(stream, container.start, container.end, state)) {
    if (wanted.has(record.type)) found.push(record);
    else if (record.version === 15) found.push(...descendants(stream, record, state, wanted, depth + 1));
  }
  return found;
}

function textBlocks(stream, items) {
  const texts = [];
  let current = null;
  for (const record of items) {
    if (record.type === types.textHeader && record.size >= 4) {
      current = { role: stream.view.getUint32(record.start, true), text: '' };
      texts.push(current);
    }
    if (record.type === types.textChars || record.type === types.textBytes) {
      const bytes = stream.data.subarray(record.start, record.end);
      if (bytes.length > 2 * 1024 * 1024) throw new Error('PPT text exceeds limit');
      let text;
      if (record.type === types.textChars) {
        if (bytes.length % 2) throw new Error('Invalid PPT Unicode text');
        text = new TextDecoder('utf-16le').decode(bytes);
      } else {
        const parts = [];
        for (let offset = 0; offset < bytes.length; offset += 8192) parts.push(String.fromCharCode(...bytes.subarray(offset, offset + 8192)));
        text = parts.join('');
      }
      if (!current) { current = { role: 4, text: '' }; texts.push(current); }
      current.text += text.replace(/\r\n?|\v/g, '\n').replace(/\0/g, '');
    }
  }
  return texts;
}

function persistDirectory(stream, current, state) {
  const user = header(current, 0);
  if (user.type !== types.currentUser || user.size < 12) throw new Error('Invalid PPT Current User stream');
  if (current.view.getUint32(user.start + 4, true) !== 0xe391c05f) throw new Error('Encrypted PPT is not supported');
  let offset = current.view.getUint32(user.start + 8, true);
  const visited = new Set();
  const mapping = new Map();
  let documentId = null;
  while (offset) {
    if (visited.has(offset) || visited.size > 4096) throw new Error('Invalid PPT edit chain');
    visited.add(offset);
    const edit = header(stream, offset);
    if (edit.type !== types.userEdit || edit.size < 28) throw new Error('Invalid PPT UserEdit record');
    if (edit.size >= 32 && stream.view.getUint32(edit.start + 28, true)) throw new Error('Encrypted PPT is not supported');
    documentId ??= stream.view.getUint32(edit.start + 16, true);
    const directory = header(stream, stream.view.getUint32(edit.start + 12, true));
    if (directory.type !== types.persist) throw new Error('Invalid PPT persist directory');
    for (let position = directory.start; position < directory.end;) {
      if (position + 4 > directory.end) throw new Error('Truncated PPT persist entry');
      const value = stream.view.getUint32(position, true);
      const first = value & 0xfffff;
      const count = value >>> 20;
      position += 4;
      if (!count || position + count * 4 > directory.end) throw new Error('Invalid PPT persist entry');
      for (let index = 0; index < count; index++) {
        if (mapping.size > 100000) throw new Error('PPT persist directory exceeds limit');
        const target = stream.view.getUint32(position, true);
        if (!mapping.has(first + index)) mapping.set(first + index, target);
        position += 4;
      }
    }
    offset = stream.view.getUint32(edit.start + 8, true);
  }
  if (!mapping.has(documentId)) throw new Error('PPT document is missing from its persist directory');
  const document = header(stream, mapping.get(documentId));
  if (document.type !== types.document || document.version !== 15) throw new Error('Invalid PPT Document container');
  return { mapping, document };
}

function options(stream, shape, state) {
  const values = new Map();
  for (const item of records(stream, shape.start, shape.end, state).filter(record => [types.options, 0xf121, 0xf122].includes(record.type))) {
    if (item.instance * 6 > item.size) throw new Error('Invalid PPT shape properties');
    for (let index = 0; index < item.instance; index++) {
      const offset = item.start + index * 6;
      const property = stream.view.getUint16(offset, true);
      if (!(property & 0x8000)) values.set(property & 0x3fff, stream.view.getUint32(offset + 2, true));
    }
  }
  return values;
}

function picture(stream, offset) {
  const record = header(stream, offset);
  if (![0xf01d, 0xf01e, 0xf02a].includes(record.type)) return null;
  const bytes = stream.data.subarray(record.start, record.end);
  for (let index = 16; index < Math.min(64, bytes.length - 8); index++) {
    if (bytes[index] === 0x89 && bytes[index + 1] === 0x50 && bytes[index + 2] === 0x4e && bytes[index + 3] === 0x47 && bytes[index + 4] === 13 && bytes[index + 5] === 10) {
      return { mime: 'image/png', bytes: bytes.slice(index) };
    }
    if (bytes[index] === 0xff && bytes[index + 1] === 0xd8 && bytes[index + 2] === 0xff) return { mime: 'image/jpeg', bytes: bytes.slice(index) };
  }
  return null;
}

function images(stream, document, pictures, state) {
  const result = [];
  const cached = new Map();
  let totalBytes = 0;
  const read = (source, offset, key) => {
    if (cached.has(key)) return cached.get(key);
    const image = picture(source, offset);
    totalBytes += image?.bytes.length || 0;
    if (totalBytes > 128 * 1024 * 1024) throw new Error('PPT pictures exceed memory limits');
    cached.set(key, image);
    return image;
  };
  for (const store of descendants(stream, document, state, new Set([types.pictureStore]))) {
    for (const entry of records(stream, store.start, store.end, state)) {
      let image = null;
      if (entry.type === types.pictureEntry && entry.size >= 36) {
        const delayed = stream.view.getUint32(entry.start + 28, true);
        const embedded = entry.start + 36 + stream.view.getUint8(entry.start + 33);
        if (embedded + 8 <= entry.end) image = read(stream, embedded, 'embedded:' + embedded);
        else if (pictures && delayed !== 0xffffffff && delayed + 8 <= pictures.data.length) image = read(pictures, delayed, 'delayed:' + delayed);
      }
      result.push(image);
      if (result.length > 4096) throw new Error('Too many PPT pictures');
    }
  }
  return result;
}

function color(value, scheme, fallback) {
  if (value === undefined) return fallback;
  if (value & 0x08000000) value = scheme[value & 255] ?? 0;
  else if (value >>> 24) return fallback;
  return '#' + [value & 255, (value >>> 8) & 255, (value >>> 16) & 255].map(channel => channel.toString(16).padStart(2, '0')).join('');
}

function shapes(stream, slide, texts, state, scheme) {
  const result = [];
  for (const shape of descendants(stream, slide, state, new Set([types.shape]))) {
    const children = records(stream, shape.start, shape.end, state);
    const anchor = children.find(item => item.type === types.anchor || item.type === types.childAnchor);
    const properties = children.find(item => item.type === types.shapeProperties);
    if (!anchor || !properties || properties.size < 8 || ![8, 16].includes(anchor.size)) continue;
    const values = options(stream, shape, state);
    const small = anchor.size === 8;
    const coordinates = Array.from({ length: 4 }, (_, index) => small ? stream.view.getInt16(anchor.start + index * 2, true) : stream.view.getInt32(anchor.start + index * 4, true));
    const [top, left, right, bottom] = anchor.type === types.anchor ? coordinates : [coordinates[1], coordinates[0], coordinates[2], coordinates[3]];
    const inline = descendants(stream, shape, state, new Set([types.textHeader, types.textChars, types.textBytes, types.textReference]));
    const reference = inline.find(record => record.type === types.textReference && record.size >= 4);
    const referenceIndex = reference ? stream.view.getInt32(reference.start, true) : null;
    if (reference && (referenceIndex < 0 || referenceIndex >= texts.length)) throw new Error('Invalid PPT outline text reference');
    const content = reference ? [texts[referenceIndex]] : textBlocks(stream, inline);
    const flags = stream.view.getUint32(properties.start + 4, true);
    result.push({ type: properties.instance, left, top, width: Math.max(0, right - left), height: Math.max(0, bottom - top),
      fill: color(values.get(385), scheme, 'transparent'), stroke: color(values.get(448), scheme, 'transparent'),
      rotation: (values.get(4) | 0) / 65536, image: values.get(260) || null,
      fitShapeToText: ((values.get(191) || 0) & 0x40004) === 0x40004,
      background: !!(flags & 0x400), texts: content });
    if (result.length > 4096) throw new Error('Too many PPT shapes');
  }
  return result;
}

export function parsePpt(input) {
  const data = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (data.byteLength > 50 * 1024 * 1024 || data.byteLength < 512) throw new Error('Invalid PPT file size: ' + data.byteLength);
  const file = CFB.read(data, { type: 'array' });
  const streamEntry = CFB.find(file, 'PowerPoint Document');
  const currentEntry = CFB.find(file, 'Current User');
  if (!streamEntry?.content || !currentEntry?.content) throw new Error('PPT streams are missing');
  const stream = binary(streamEntry.content);
  const current = binary(currentEntry.content);
  const pictureEntry = CFB.find(file, 'Pictures');
  const state = { records: 0 };
  const { mapping, document } = persistDirectory(stream, current, state);
  const atom = records(stream, document.start, document.end, state).find(record => record.type === types.documentAtom && record.size >= 8);
  const width = atom ? stream.view.getInt32(atom.start, true) : 5760;
  const height = atom ? stream.view.getInt32(atom.start + 4, true) : 4320;
  if (width <= 0 || height <= 0 || width > 100000 || height > 100000) throw new Error('Invalid PPT slide dimensions');
  const pictures = images(stream, document, pictureEntry?.content ? binary(pictureEntry.content) : null, state);
  const lists = records(stream, document.start, document.end, state).filter(record => record.type === types.slideList && record.instance === 0);
  const slides = [];
  for (const list of lists) {
    let entry = null;
    for (const record of records(stream, list.start, list.end, state)) {
      if (record.type === types.slidePersist) {
        if (record.size < 16) throw new Error('Invalid PPT slide reference');
        entry = { persistId: stream.view.getUint32(record.start, true), id: stream.view.getUint32(record.start + 12, true), records: [] };
        slides.push(entry);
      } else if (entry) entry.records.push(record);
    }
  }
  if (!slides.length || slides.length > 1000) throw new Error('Invalid PPT slide count');
  for (const slide of slides) {
    if (!mapping.has(slide.persistId)) throw new Error('PPT slide reference is missing');
    const container = header(stream, mapping.get(slide.persistId));
    if (container.type !== types.slide || container.version !== 15) throw new Error('Invalid PPT slide container');
    slide.texts = textBlocks(stream, slide.records);
    const palette = descendants(stream, container, state, new Set([types.colorScheme])).find(record => record.size >= 32);
    const scheme = palette ? Array.from({ length: 8 }, (_, index) => stream.view.getUint32(palette.start + index * 4, true)) : [];
    slide.shapes = shapes(stream, container, slide.texts, state, scheme);
    slide.texts.push(...textBlocks(stream, descendants(stream, container, state, new Set([types.textHeader, types.textChars, types.textBytes]))));
    delete slide.records;
  }
  return { width, height, slides, pictures };
}
