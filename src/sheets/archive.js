import { Inflate } from 'fflate';
import { validateOfficeZip } from '../zip-limits.js';
import { attributes, cellAddress } from './xml.js';

export function officeArchive(data) {
  validateOfficeZip(data, true);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let footer = data.byteLength - 22;
  while (view.getUint32(footer, true) !== 0x06054b50 || footer + 22 + view.getUint16(footer + 20, true) !== data.byteLength) footer--;
  const result = new Map();
  let offset = view.getUint32(footer + 16, true);
  const count = view.getUint16(footer + 10, true);
  for (let index = 0; index < count; index++) {
    const nameSize = view.getUint16(offset + 28, true);
    const name = new TextDecoder().decode(data.subarray(offset + 46, offset + 46 + nameSize));
    const local = view.getUint32(offset + 42, true);
    if (local + 30 > offset || view.getUint32(local, true) !== 0x04034b50) throw new Error('Invalid ZIP local entry');
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const end = start + view.getUint32(offset + 20, true);
    const method = view.getUint16(offset + 10, true);
    if (end > offset || ![0, 8].includes(method) || result.has(name) || name.split('/').includes('..')) throw new Error('Invalid ZIP entry');
    result.set(name, { data: data.subarray(start, end), method, size: view.getUint32(offset + 24, true) });
    offset += 46 + nameSize + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
  }
  return result;
}

export async function scanXml(entry, consume, cancelled = () => false) {
  if (!entry) throw new Error('Missing workbook part');
  const decoder = new TextDecoder('utf-8', { fatal: true });
  let total = 0;
  let stopped = false;
  const output = (chunk, final) => {
    total += chunk.length;
    if (total > entry.size) throw new Error('Invalid expanded ZIP size');
    if (consume(decoder.decode(chunk, { stream: !final })) === false) stopped = true;
  };
  const inflater = entry.method === 8 ? new Inflate(output) : null;
  for (let offset = 0; offset < entry.data.length; offset += 2048) {
    if (cancelled()) throw new DOMException('Superseded worksheet request', 'AbortError');
    const chunk = entry.data.subarray(offset, offset + 2048);
    const final = offset + chunk.length === entry.data.length;
    if (inflater) inflater.push(chunk, final);
    else output(chunk, final);
    if (stopped) return;
    if (offset % 131072 === 0) await new Promise(resolve => setTimeout(resolve, 0));
  }
  if (total !== entry.size) throw new Error('Truncated ZIP data');
}

export async function xmlPart(entry, cancelled) {
  let result = '';
  await scanXml(entry, chunk => {
    result += chunk;
    if (result.length > 8 * 1024 * 1024) throw new Error('Workbook metadata exceeds safety budget');
  }, cancelled);
  if (/<!DOCTYPE|<!ENTITY/i.test(result)) throw new Error('XML entities are not allowed');
  return result;
}

export async function scanElements(entry, name, visit, cancelled) {
  let buffer = '';
  const opening = new RegExp('<(?:[\\w.-]+:)?' + name + '\\b[^>]*>', 'g');
  const closing = new RegExp('</(?:[\\w.-]+:)?' + name + '\\s*>');
  await scanXml(entry, chunk => {
    buffer += chunk;
    if (/<!DOCTYPE|<!ENTITY/i.test(buffer)) throw new Error('XML entities are not allowed');
    while (true) {
      opening.lastIndex = 0;
      const begin = opening.exec(buffer);
      if (!begin) { const partial = buffer.lastIndexOf('<'); buffer = partial < 0 ? buffer.slice(-256) : buffer.slice(partial); break; }
      if (begin.index) buffer = buffer.slice(begin.index);
      const tag = begin[0];
      const end = /\/>$/.test(tag) ? { index: tag.length, 0: '' } : closing.exec(buffer.slice(tag.length));
      if (!end) break;
      const length = /\/>$/.test(tag) ? tag.length : tag.length + end.index + end[0].length;
      const content = buffer.slice(tag.length, /\/>$/.test(tag) ? tag.length : tag.length + end.index);
      buffer = buffer.slice(length);
      if (visit({ attributes: attributes(tag), content }) === false) return false;
    }
    if (buffer.length > 8 * 1024 * 1024) throw new Error('XML element exceeds safety budget');
  }, cancelled);
}

export async function scanCells(entry, visit, cancelled, visitRow) {
  let buffer = '';
  let row = -1;
  let column = 0;
  let rowAttributes = {};
  const opening = /<(?:[\w.-]+:)?(row|c)\b[^>]*>/g;
  const closing = /<\/(?:[\w.-]+:)?c\s*>/;
  await scanXml(entry, chunk => {
    buffer += chunk;
    if (/<!DOCTYPE|<!ENTITY/i.test(buffer)) throw new Error('XML entities are not allowed');
    while (true) {
      opening.lastIndex = 0;
      const start = opening.exec(buffer);
      if (!start) { const partial = buffer.lastIndexOf('<'); buffer = partial < 0 ? '' : buffer.slice(partial); break; }
      buffer = buffer.slice(start.index);
      const tag = start[0];
      const properties = attributes(tag);
      if (start[1] === 'row') {
        row = properties.r ? Number(properties.r) - 1 : row + 1;
        if (!Number.isSafeInteger(row) || row < 0 || row >= 1048576) throw new Error('Invalid worksheet row');
        column = 0;
        rowAttributes = properties;
        buffer = buffer.slice(tag.length);
        if (visitRow?.({ row, attributes: properties }) === false) return false;
        continue;
      }
      const end = /\/>$/.test(tag) ? { index: 0, 0: '' } : closing.exec(buffer.slice(tag.length));
      if (!end) break;
      const content = buffer.slice(tag.length, tag.length + end.index);
      buffer = buffer.slice(tag.length + end.index + end[0].length);
      const address = properties.r ? cellAddress(properties.r) : { row, column };
      column = address.column + 1;
      if (!Number.isSafeInteger(address.row) || address.row < 0 || address.row >= 1048576 || address.column >= 16384) throw new Error('Invalid worksheet cell');
      if (visit({ attributes: properties, content, address, rowAttributes }) === false) return false;
    }
    if (buffer.length > 8 * 1024 * 1024) throw new Error('Cell exceeds safety budget');
  }, cancelled);
}
