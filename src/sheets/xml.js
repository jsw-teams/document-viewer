export function xmlText(value) {
  return value.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (match, entity) => {
    const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
    if (Object.hasOwn(named, entity)) return named[entity];
    const number = entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
    return number > 0 && number <= 0x10ffff ? String.fromCodePoint(number) : '\ufffd';
  });
}

export function attributes(value) {
  const result = {};
  for (const match of value.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) result[match[1]] = xmlText(match[2] ?? match[3]);
  return result;
}

export function elements(value, name) {
  const expression = new RegExp('<(?:[\\w.-]+:)?' + name + '\\b([^>]*?)(?:/>|>([\\s\\S]*?)</(?:[\\w.-]+:)?' + name + '\\s*>)', 'g');
  return [...value.matchAll(expression)].map(match => ({ attributes: attributes(match[1]), content: match[2] || '' }));
}

export function textRuns(value) {
  return elements(value.replace(/<(?:[\w.-]+:)?rPh\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?rPh\s*>/g, ''), 't').map(entry => xmlText(entry.content)).join('');
}

export function cellAddress(value) {
  const match = String(value).toUpperCase().match(/^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/);
  if (!match) throw new Error('Invalid cell address');
  let column = 0;
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64;
  const row = Number(match[2]) - 1;
  if (column > 16384 || row >= 1048576) throw new Error('Cell address exceeds spreadsheet dimensions');
  return { row, column: column - 1 };
}

export function columnName(column) {
  let result = '';
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result;
  return result;
}
