function reader(stream, record) {
  let position = record.start;
  return {
    read(size, signed = false) {
      if (position + size > record.end) throw new Error('Truncated PPT text style');
      const value = size === 4 ? stream.view.getUint32(position, true) : signed ? stream.view.getInt16(position, true) : stream.view.getUint16(position, true);
      position += size;
      return value;
    },
    get remaining() { return record.end - position; }
  };
}

function paragraphStyle(input) {
  const mask = input.read(4);
  const style = {};
  if (mask & 15) { const flags = input.read(2); style.bullet = !!(flags & 1); }
  if (mask & 128) style.bulletChar = String.fromCharCode(input.read(2));
  if (mask & 16) style.bulletFont = input.read(2);
  if (mask & 64) style.bulletSize = input.read(2);
  if (mask & 32) style.bulletColor = input.read(4);
  if (mask & 2048) style.align = input.read(2);
  if (mask & 4096) style.lineSpacing = input.read(2, true);
  if (mask & 8192) style.before = input.read(2, true);
  if (mask & 16384) style.after = input.read(2, true);
  if (mask & 256) style.margin = input.read(2, true);
  if (mask & 1024) style.indent = input.read(2, true);
  if (mask & 32768) style.tabSize = input.read(2, true);
  if (mask & 1048576) {
    const count = input.read(2);
    if (count > 4096) throw new Error('Too many PPT tab stops');
    style.tabs = Array.from({ length: count }, () => ({ position: input.read(2, true), align: input.read(2) }));
  }
  if (mask & 65536) style.fontAlign = input.read(2);
  if (mask & 917504) style.wrap = input.read(2);
  if (mask & 2097152) style.direction = input.read(2);
  return style;
}

function characterStyle(input) {
  const mask = input.read(4);
  if (mask & 0xff000000) throw new Error('Invalid PPT character style mask');
  const style = {};
  if (mask & 0xffff) {
    const flags = input.read(2);
    if (mask & 1) style.bold = !!(flags & 1);
    if (mask & 2) style.italic = !!(flags & 2);
    if (mask & 4) style.underline = !!(flags & 4);
    if (mask & 16) style.shadow = !!(flags & 16);
  }
  if (mask & 65536) style.font = input.read(2);
  if (mask & 2097152) style.eastAsianFont = input.read(2);
  if (mask & 4194304) style.ansiFont = input.read(2);
  if (mask & 8388608) style.symbolFont = input.read(2);
  if (mask & 131072) {
    const size = input.read(2, true);
    if (size < 1 || size > 4000) throw new Error('Invalid PPT font size');
    style.size = size;
  }
  if (mask & 262144) style.color = input.read(4);
  if (mask & 524288) style.position = input.read(2, true);
  return style;
}

export function textStyles(stream, record, length) {
  const input = reader(stream, record);
  const paragraphs = [];
  const runs = [];
  for (let offset = 0; offset < length + 1;) {
    const count = input.read(4);
    const level = input.read(2);
    if (!count || count > length + 1 - offset || level > 8 || paragraphs.length >= 8192) throw new Error('Invalid PPT paragraph run');
    paragraphs.push({ start: offset, end: offset + count, level, ...paragraphStyle(input) });
    offset += count;
  }
  for (let offset = 0; offset < length + 1;) {
    const count = input.read(4);
    if (!count || count > length + 1 - offset || runs.length >= 8192) throw new Error('Invalid PPT character run');
    runs.push({ start: offset, end: offset + count, ...characterStyle(input) });
    offset += count;
  }
  return { paragraphs, runs };
}

export function masterTextStyles(stream, record) {
  const input = reader(stream, record);
  const count = input.read(2);
  if (count > 5) throw new Error('Invalid PPT master style count');
  const levels = [];
  for (let index = 0; index < count; index++) {
    const level = record.instance >= 5 ? input.read(2) : index;
    if (level >= count) throw new Error('Invalid PPT master style level');
    levels[level] = { paragraph: paragraphStyle(input), character: characterStyle(input) };
  }
  return levels;
}

export function inheritTextStyles(item, defaults = [], master = []) {
  const paragraphs = item.paragraphs?.length ? item.paragraphs : [{ start: 0, end: item.text.length + 1, level: 0 }];
  const runs = item.runs?.length ? item.runs : [{ start: 0, end: item.text.length + 1 }];
  item.paragraphs = paragraphs.map(entry => ({ ...defaults[entry.level]?.paragraph, ...master[entry.level]?.paragraph, ...entry }));
  item.runs = paragraphs.flatMap(paragraph => runs.filter(run => run.start < paragraph.end && run.end > paragraph.start).map(run => ({ ...defaults[paragraph.level]?.character, ...master[paragraph.level]?.character, ...run, start: Math.max(run.start, paragraph.start), end: Math.min(run.end, paragraph.end) })));
}

export function textColor(value, scheme = [], fallback = '#222222') {
  if (value === undefined || value >>> 24 === 255) return fallback;
  const index = value >>> 24;
  if (index !== 254) {
    if (index > 7 || scheme[index] === undefined) return fallback;
    value = scheme[index];
  }
  return '#' + [value & 255, value >>> 8 & 255, value >>> 16 & 255].map(channel => channel.toString(16).padStart(2, '0')).join('');
}

export function styledText(doc, item, { fonts = [], scheme = [], size = 24 } = {}) {
  const root = doc.createElementNS('http://www.w3.org/1999/xhtml', 'div');
  root.style.cssText = 'white-space:pre-wrap;overflow-wrap:normal;color:#222;box-sizing:border-box';
  root.style.fontSize = size * 8 + 'px';
  root.style.fontFamily = 'Arial, sans-serif';
  let offset = 0;
  for (const line of item.text.split('\n')) {
    const paragraph = doc.createElement('p');
    const style = item.paragraphs?.find(entry => entry.start <= offset && entry.end > offset) || {};
    paragraph.style.margin = '0';
    paragraph.style.textAlign = ['left', 'center', 'right', 'justify', 'justify'][style.align] || 'left';
    if (style.direction === 1) paragraph.dir = 'rtl';
    if (style.margin !== undefined) paragraph.style.marginInlineStart = style.margin + 'px';
    if (style.indent !== undefined) paragraph.style.textIndent = (style.indent - (style.margin || 0)) + 'px';
    if (style.lineSpacing) paragraph.style.lineHeight = style.lineSpacing > 0 ? String(style.lineSpacing / 100) : -style.lineSpacing + 'px';
    for (const [field, property] of [['before', 'marginTop'], ['after', 'marginBottom']]) if (style[field]) paragraph.style[property] = style[field] > 0 ? style[field] / 100 + 'em' : -style[field] + 'px';
    const end = offset + line.length;
    const runs = item.runs?.filter(entry => entry.start < end && entry.end > offset) || [];
    paragraph.style.fontSize = (runs.length ? Math.max(...runs.map(run => run.size || size)) : size) * 8 + 'px';
    if (style.bullet) {
      const bullet = doc.createElement('span');
      const character = style.bulletChar || '•';
      const font = fonts[style.bulletFont];
      const mapped = symbolBullet(character, font);
      bullet.textContent = mapped + '\u00a0';
      bullet.style.color = textColor(style.bulletColor, scheme);
      if (mapped !== character) bullet.style.fontFamily = 'serif';
      else if (font) bullet.style.fontFamily = JSON.stringify(font);
      if (style.bulletSize) bullet.style.fontSize = style.bulletSize <= 400 ? style.bulletSize + '%' : style.bulletSize * 8 + 'px';
      paragraph.append(bullet);
    }
    if (!runs.length) paragraph.append(doc.createTextNode(line || '\u00a0'));
    else for (const run of runs) {
      const span = doc.createElement('span');
      span.textContent = item.text.slice(Math.max(offset, run.start), Math.min(end, run.end));
      span.style.fontSize = (run.size || size) * 8 + 'px';
      const font = fonts[run.font ?? run.ansiFont ?? run.eastAsianFont ?? run.symbolFont];
      if (font) span.style.fontFamily = JSON.stringify(font) + ', sans-serif';
      if (run.bold !== undefined) span.style.fontWeight = run.bold ? '700' : '400';
      if (run.italic !== undefined) span.style.fontStyle = run.italic ? 'italic' : 'normal';
      if (run.underline !== undefined) span.style.textDecoration = run.underline ? 'underline' : 'none';
      span.style.color = textColor(run.color, scheme);
      if (run.shadow) span.style.textShadow = '8px 8px 8px #8888';
      if (run.position) span.style.verticalAlign = run.position + '%';
      paragraph.append(span);
    }
    root.append(paragraph);
    offset = end + 1;
  }
  return root;
}
import { symbolBullet } from '../symbols.js';
