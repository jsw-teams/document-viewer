import { cellAddress, columnName } from './xml.js';

export class FormulaError extends Error {
  constructor(code) { super(code); this.name = 'FormulaError'; }
}
const fail = code => { throw new FormulaError(code); };
const number = value => {
  if (value === null || value === '') return 0;
  if (typeof value === 'boolean') return Number(value);
  const result = Number(value);
  if (!Number.isFinite(result)) fail('#VALUE!');
  return result;
};
const text = value => value === null ? '' : typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : String(value);
const priorities = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export function rebaseFormula(input, rows, columns) {
  return input.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'|\$?[\p{L}_][\p{L}\p{N}_.$]*/gu, (token, offset) => {
    const match = token.match(/^(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})$/i);
    if (!match || /^\s*[!(]/.test(input.slice(offset + token.length))) return token;
    try {
      const original = cellAddress(token);
      const row = original.row + (match[3] ? 0 : rows);
      const column = original.column + (match[1] ? 0 : columns);
      if (row < 0 || row >= 1048576 || column < 0 || column >= 16384) return '#REF!';
      return match[1] + columnName(column) + match[3] + (row + 1);
    } catch { return '#REF!'; }
  });
}

export function parseFormula(input) {
  const source = String(input).replace(/^=/, '');
  if (source.length > 32767) fail('#NUM!');
  const tokens = [];
  const expression = /\s+|"(?:[^"]|"")*"|'(?:[^']|'')*'|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|\$?[\p{L}_][\p{L}\p{N}_.$]*|#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)|<>|<=|>=|[+\-*/^&%=<>(),;!:]/uy;
  while (expression.lastIndex < source.length) {
    const offset = expression.lastIndex;
    const match = expression.exec(source);
    if (!match || tokens.length >= 8192) fail('#NAME?');
    if (!/^\s/.test(match[0])) tokens.push(match[0]);
    if (expression.lastIndex <= offset) fail('#VALUE!');
  }
  let position = 0;
  const reference = value => /^\$?[A-Z]{1,3}\$?[1-9]\d{0,6}$/i.test(value);
  const read = (minimum = 0, depth = 0) => {
    if (depth > 64) fail('#NUM!');
    const token = tokens[position++];
    if (token === undefined) fail('#VALUE!');
    let node;
    if (token === '+' || token === '-') node = { type: 'unary', operator: token, value: read(7, depth + 1) };
    else if (token === '(') { node = read(0, depth + 1); if (tokens[position++] !== ')') fail('#VALUE!'); }
    else if (token.startsWith('"')) node = { type: 'literal', value: token.slice(1, -1).replaceAll('""', '"') };
    else if (/^[\d.]/.test(token)) node = { type: 'literal', value: Number(token) };
    else if (token.startsWith('#')) node = { type: 'error', value: token };
    else if (tokens[position] === '(') {
      position++;
      const args = [];
      while (tokens[position] !== ')') {
        if (tokens[position] === ',' || tokens[position] === ';') args.push({ type: 'literal', value: null });
        else args.push(read(0, depth + 1));
        if (![',', ';'].includes(tokens[position])) break;
        position++;
        if (tokens[position] === ')') args.push({ type: 'literal', value: null });
      }
      if (tokens[position++] !== ')') fail('#VALUE!');
      node = { type: 'call', name: token.replace(/^_xlfn\./i, '').toUpperCase(), args };
    } else if (tokens[position] === '!' || reference(token)) {
      let sheet;
      let address = token;
      if (tokens[position] === '!') { sheet = token.startsWith("'") ? token.slice(1, -1).replaceAll("''", "'") : token; position++; address = tokens[position++]; }
      if (!reference(address)) fail('#REF!');
      const decode = address => { try { return cellAddress(address); } catch { fail('#REF!'); } };
      const start = decode(address);
      let end = start;
      if (tokens[position] === ':') { position++; const last = tokens[position++]; if (!reference(last)) fail('#REF!'); end = decode(last); }
      if (end.row < start.row || end.column < start.column) fail('#REF!');
      node = { type: 'reference', sheet, start, end };
    } else if (/^(TRUE|FALSE)$/i.test(token)) node = { type: 'literal', value: /^TRUE$/i.test(token) };
    else node = { type: 'name', name: token };
    while (tokens[position] === '%') { position++; node = { type: 'percent', value: node }; }
    while (priorities[tokens[position]] !== undefined && priorities[tokens[position]] >= minimum) {
      const operator = tokens[position++];
      node = { type: 'binary', operator, left: node, right: read(priorities[operator] + 1, depth + 1) };
    }
    return node;
  };
  const result = read();
  if (position !== tokens.length) fail('#VALUE!');
  return result;
}

function compare(left, right, operator) {
  if (typeof left === 'string') left = left.toLocaleLowerCase('en');
  if (typeof right === 'string') right = right.toLocaleLowerCase('en');
  if (left === null) left = typeof right === 'string' ? '' : 0;
  if (right === null) right = typeof left === 'string' ? '' : 0;
  return { '=': () => left === right, '<>': () => left !== right, '<': () => left < right, '>': () => left > right, '<=': () => left <= right, '>=': () => left >= right }[operator]();
}

function criterion(value, expected, tick) {
  if (typeof expected !== 'string') return compare(value, expected, '=');
  const match = expected.match(/^(<>|<=|>=|=|<|>)(.*)$/);
  const operator = match?.[1] || '=';
  const query = match?.[2] ?? expected;
  if (query.trim() !== '' && Number.isFinite(Number(query))) return compare(value, Number(query), operator);
  if (operator === '=' || operator === '<>') {
    const pattern = [];
    const characters = Array.from(query.toLocaleLowerCase('en'));
    for (let position = 0; position < characters.length; position++) {
      tick();
      const character = characters[position];
      pattern.push(character === '~' && position + 1 < characters.length ? { literal: characters[++position] } : { [character === '*' || character === '?' ? 'wildcard' : 'literal']: character });
    }
    const input = Array.from(text(value).toLocaleLowerCase('en'));
    let position = 0;
    let token = 0;
    let star = -1;
    let retry = 0;
    while (position < input.length) {
      tick();
      if (pattern[token]?.literal === input[position] || pattern[token]?.wildcard === '?') { position++; token++; }
      else if (pattern[token]?.wildcard === '*') { star = token++; retry = position; }
      else if (star >= 0) { token = star + 1; position = ++retry; }
      else break;
    }
    while (pattern[token]?.wildcard === '*') { tick(); token++; }
    const matches = position === input.length && token === pattern.length;
    return operator === '=' ? matches : !matches;
  }
  return compare(text(value), query, operator);
}

export function formulaEvaluator({ readCell, names = [], definedNames = [], cancelled = () => false, maxOperations = 1000000 }) {
  const cache = new Map();
  const visiting = new Set();
  let operations = 0;
  const tick = () => {
    if (cancelled()) throw new DOMException('Superseded worksheet request', 'AbortError');
    if (++operations > maxOperations) fail('#NUM!');
  };
  const sheetIndex = (sheet, current) => sheet === undefined ? current : names.findIndex(name => name.toLocaleLowerCase('en') === sheet.toLocaleLowerCase('en'));
  const definitionFor = (name, sheet) => definedNames.find(entry => entry.name.toLowerCase() === name.toLowerCase() && entry.sheet === sheet) || definedNames.find(entry => entry.name.toLowerCase() === name.toLowerCase() && entry.sheet === undefined);
  const scalar = async (node, sheet, depth) => {
    if (node.type === 'reference') {
      const index = sheetIndex(node.sheet, sheet);
      if (index < 0) fail('#REF!');
      return cell(index, node.start, depth + 1);
    }
    return evaluate(node, sheet, depth + 1);
  };
  async function* values(node, sheet, depth, errorsAsValues = false) {
    tick();
    if (depth > 128) fail('#NUM!');
    if (node.type === 'name') {
      const definition = definitionFor(node.name, sheet);
      if (!definition) fail('#NAME?');
      yield* values(parseFormula(definition.value), sheet, depth + 1, errorsAsValues);
      return;
    }
    if (node.type !== 'reference') {
      try { yield await scalar(node, sheet, depth); }
      catch (error) { if (!errorsAsValues || !(error instanceof FormulaError)) throw error; yield error; }
      return;
    }
    const index = sheetIndex(node.sheet, sheet);
    if (index < 0) fail('#REF!');
    for (let row = node.start.row; row <= node.end.row; row++) for (let column = node.start.column; column <= node.end.column; column++) {
      try { yield await cell(index, { row, column }, depth + 1); }
      catch (error) { if (!errorsAsValues || !(error instanceof FormulaError)) throw error; yield error; }
    }
  }
  async function cell(sheet, address, depth = 0) {
    tick();
    if (depth > 128) fail('#NUM!');
    const key = sheet + ':' + address.row + ':' + address.column;
    if (visiting.has(key)) fail('#REF!');
    if (cache.has(key)) return cache.get(key);
    const entry = await readCell(sheet, address);
    if (!entry) return null;
    if (entry.error) fail(entry.error);
    if (!entry.formula || entry.value !== null && entry.value !== undefined) return entry.value ?? null;
    visiting.add(key);
    try {
      const result = await evaluate(parseFormula(entry.formula), sheet, depth + 1);
      if (typeof result === 'number' && !Number.isFinite(result)) fail('#NUM!');
      if (cache.size >= 8192) cache.delete(cache.keys().next().value);
      cache.set(key, result);
      return result;
    } finally { visiting.delete(key); }
  }
  async function evaluate(node, sheet, depth = 0) {
    tick();
    if (depth > 128) fail('#NUM!');
    if (node.type === 'literal') return node.value;
    if (node.type === 'error') fail(node.value);
    if (node.type === 'reference') return scalar(node, sheet, depth);
    if (node.type === 'name') {
      const definition = definitionFor(node.name, sheet);
      if (!definition) fail('#NAME?');
      return evaluate(parseFormula(definition.value), sheet, depth + 1);
    }
    if (node.type === 'unary' || node.type === 'percent') {
      const value = number(await scalar(node.value, sheet, depth));
      return node.type === 'percent' ? value / 100 : node.operator === '-' ? -value : value;
    }
    if (node.type === 'binary') {
      const left = await scalar(node.left, sheet, depth);
      const right = await scalar(node.right, sheet, depth);
      if (node.operator === '&') return text(left) + text(right);
      if (priorities[node.operator] === 1) return compare(left, right, node.operator);
      const first = number(left);
      const second = number(right);
      if (node.operator === '/' && second === 0) fail('#DIV/0!');
      const result = { '+': () => first + second, '-': () => first - second, '*': () => first * second, '/': () => first / second, '^': () => first ** second }[node.operator]();
      if (!Number.isFinite(result)) fail('#NUM!');
      return result;
    }
    const args = node.args;
    const arities = { IF: [2, 3], IFERROR: [2, 2], SUMIF: [2, 3], COUNTIF: [2, 2], AVERAGEIF: [2, 3], NOT: [1, 1], LEN: [1, 1], LOWER: [1, 1], UPPER: [1, 1], TRIM: [1, 1], ABS: [1, 1], INT: [1, 1], SQRT: [1, 1], ROUND: [2, 2], MOD: [2, 2], POWER: [2, 2], TRUE: [0, 0], FALSE: [0, 0] };
    if (arities[node.name] && (args.length < arities[node.name][0] || args.length > arities[node.name][1])) fail('#VALUE!');
    if (node.name === 'TRUE' || node.name === 'FALSE') return node.name === 'TRUE';
    const argument = index => args[index] ? scalar(args[index], sheet, depth) : Promise.resolve(null);
    if (node.name === 'IF') return argument(0).then(condition => condition ? argument(1) : args.length === 2 ? false : argument(2));
    if (node.name === 'IFERROR') { try { return await argument(0); } catch (error) { if (!(error instanceof FormulaError)) throw error; return argument(1); } }
    if (['SUMIF', 'COUNTIF', 'AVERAGEIF'].includes(node.name)) {
      if (args[0]?.type !== 'reference' || args.length < 2) fail('#VALUE!');
      const expected = await argument(1);
      const range = args[0];
      const target = args[2] || range;
      if (target.type !== 'reference') fail('#VALUE!');
      let total = 0;
      let count = 0;
      for (let row = range.start.row; row <= range.end.row; row++) for (let column = range.start.column; column <= range.end.column; column++) {
        if (!criterion(await cell(sheetIndex(range.sheet, sheet), { row, column }, depth + 1), expected, tick)) continue;
        if (node.name === 'COUNTIF') { count++; continue; }
        const value = await cell(sheetIndex(target.sheet, sheet), { row: target.start.row + row - range.start.row, column: target.start.column + column - range.start.column }, depth + 1);
        if (typeof value === 'number') { total += value; count++; }
      }
      if (node.name === 'COUNTIF') return count;
      if (node.name === 'AVERAGEIF' && !count) fail('#DIV/0!');
      return node.name === 'AVERAGEIF' ? total / count : total;
    }
    if (['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'PRODUCT', 'AND', 'OR'].includes(node.name)) {
      let total = node.name === 'PRODUCT' ? 1 : 0;
      let count = 0;
      let extreme = node.name === 'MIN' ? Infinity : -Infinity;
      let logical = node.name === 'AND';
      for (const arg of args) for await (const value of values(arg, sheet, depth, node.name === 'COUNT' || node.name === 'COUNTA')) {
        if (node.name === 'COUNTA') { if (value !== null) count++; continue; }
        if (node.name === 'AND' || node.name === 'OR') { if (typeof value === 'number' || typeof value === 'boolean') { logical = node.name === 'AND' ? logical && Boolean(value) : logical || Boolean(value); count++; } continue; }
        const numeric = typeof value === 'number' ? value : !['reference', 'name'].includes(arg.type) && (typeof value === 'boolean' || typeof value === 'string') ? number(value) : undefined;
        if (numeric === undefined) continue;
        total = node.name === 'PRODUCT' ? total * numeric : total + numeric;
        count++;
        extreme = node.name === 'MIN' ? Math.min(extreme, numeric) : Math.max(extreme, numeric);
      }
      if (node.name === 'COUNT' || node.name === 'COUNTA') return count;
      if (node.name === 'AND' || node.name === 'OR') { if (!count) fail('#VALUE!'); return logical; }
      if (node.name === 'MIN' || node.name === 'MAX') return count ? extreme : 0;
      if (node.name === 'AVERAGE') { if (!count) fail('#DIV/0!'); return total / count; }
      return node.name === 'PRODUCT' && !count ? 0 : total;
    }
    const first = await argument(0);
    if (node.name === 'NOT') return !first;
    if (node.name === 'LEN') return text(first).length;
    if (node.name === 'LOWER' || node.name === 'UPPER' || node.name === 'TRIM') return node.name === 'LOWER' ? text(first).toLowerCase() : node.name === 'UPPER' ? text(first).toUpperCase() : text(first).trim().replace(/ +/g, ' ');
    if (node.name === 'ABS') return Math.abs(number(first));
    if (node.name === 'INT') return Math.floor(number(first));
    if (node.name === 'SQRT') { if (number(first) < 0) fail('#NUM!'); return Math.sqrt(number(first)); }
    if (node.name === 'ROUND') { const digits = Math.trunc(number(await argument(1))); if (Math.abs(digits) > 308) fail('#NUM!'); const scale = 10 ** digits; const amount = Math.abs(number(first)) * scale; return Math.sign(number(first)) * Math.round(amount + Number.EPSILON * amount * 4) / scale; }
    if (node.name === 'MOD') { const divisor = number(await argument(1)); if (!divisor) fail('#DIV/0!'); return number(first) - divisor * Math.floor(number(first) / divisor); }
    if (node.name === 'POWER') { const result = number(first) ** number(await argument(1)); if (!Number.isFinite(result)) fail('#NUM!'); return result; }
    fail('#NAME?');
  }
  return { cell, evaluate: async (formula, sheet = 0) => { const result = await evaluate(parseFormula(formula), sheet); if (typeof result === 'number' && !Number.isFinite(result)) fail('#NUM!'); return result; }, dispose() { cache.clear(); visiting.clear(); } };
}
