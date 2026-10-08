import { read, utils, SSF } from 'xlsx';
import { officeArchive, xmlPart, scanXml, scanElements, scanCells } from './archive.js';
import { elements, cellAddress, textRuns, xmlText } from './xml.js';
import { spreadsheetTheme, workbookStyles, richText } from './styles.js';
import { formulaEvaluator, FormulaError, rebaseFormula } from './formula.js';

function columnPixels(value) {
  const width = Number(value);
  if (!Number.isFinite(width) || width < 0 || width > 255) throw new Error('Invalid worksheet column width');
  return Math.floor((256 * width + Math.floor(128 / 7)) / 256 * 7);
}

export async function workbookSource(data, format) {
  if (data instanceof ArrayBuffer) data = new Uint8Array(data);
  if (format === 'xls') {
    const metadata = read(data, { type: 'array', bookSheets: true, bookVBA: false });
    return { data, format, names: metadata.SheetNames };
  }
  const archive = officeArchive(data);
  const workbook = await xmlPart(archive.get('xl/workbook.xml'));
  const relationships = elements(await xmlPart(archive.get('xl/_rels/workbook.xml.rels')), 'Relationship');
  const paths = new Map(relationships.filter(entry => entry.attributes.TargetMode !== 'External').map(entry => {
    const target = entry.attributes.Target;
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    if (path.split('/').includes('..')) throw new Error('Unsafe workbook relationship');
    return [entry.attributes.Id, path];
  }));
  const allSheets = elements(workbook, 'sheet').map(entry => ({ name: entry.attributes.name, path: paths.get(entry.attributes['r:id']), state: entry.attributes.state }));
  const sheets = allSheets.filter(sheet => !['hidden', 'veryHidden'].includes(sheet.state));
  if (!sheets.length || allSheets.some(sheet => !archive.has(sheet.path))) throw new Error('Invalid workbook sheets');
  const styles = archive.has('xl/styles.xml') ? await xmlPart(archive.get('xl/styles.xml')) : '';
  const theme = spreadsheetTheme(archive.has('xl/theme/theme1.xml') ? await xmlPart(archive.get('xl/theme/theme1.xml')) : '');
  const cellStyles = workbookStyles(styles, theme, SSF.get_table());
  const definedNames = elements(workbook, 'definedName').map(entry => ({ name: entry.attributes.name, value: xmlText(entry.content), sheet: entry.attributes.localSheetId === undefined ? undefined : Number(entry.attributes.localSheetId) }));
  return { data, format, archive, sheets, allSheets, names: sheets.map(sheet => sheet.name), styles: cellStyles, theme, definedNames, date1904: ['1', 'true'].includes(elements(workbook, 'workbookPr')[0]?.attributes.date1904) };
}

export async function sheetInfo(source, index, cancelled) {
  if (!source.names[index]) throw new Error('Unknown worksheet');
  if (source.format === 'xls') {
    const workbook = read(source.data, { type: 'array', sheets: [index], dense: true, cellHTML: false, cellStyles: true, bookVBA: false });
    const sheet = workbook.Sheets[source.names[index]];
    const range = utils.decode_range(sheet['!ref'] || 'A1');
    const columnWidths = {};
    const rowHeights = {};
    const hiddenRows = [];
    const hiddenColumns = [];
    for (const [position, column] of (sheet['!cols'] || []).entries()) {
      if (!column) continue;
      if (column.hidden) hiddenColumns.push(position);
      const width = column.width === undefined ? column.wpx : columnPixels(column.width);
      if (Number.isFinite(width) && width > 0 && width <= 1800) columnWidths[position] = width;
    }
    for (const [position, row] of (sheet['!rows'] || []).entries()) {
      if (!row) continue;
      if (row.hidden) hiddenRows.push(position);
      const height = row.hpx ?? (row.hpt === undefined ? undefined : row.hpt * 96 / 72);
      if (Number.isFinite(height) && height > 0 && height <= 10000) rowHeights[position] = height;
    }
    return { rows: range.e.r + 1, columns: range.e.c + 1, merges: sheet['!merges'] || [], columnWidth: 64, columnWidths, rowHeight: 20, rowHeights, hiddenRows, hiddenColumns, gridLines: true };
  }
  const entry = source.archive.get(source.sheets[index].path);
  let header = '';
  await scanXml(entry, chunk => {
    header += chunk;
    if (header.length > 8 * 1024 * 1024) throw new Error('Worksheet header exceeds safety budget');
    return !/<(?:[\w.-]+:)?sheetData\b/.test(header);
  }, cancelled);
  const dimension = elements(header, 'dimension')[0]?.attributes.ref;
  let end = dimension ? cellAddress(dimension.split(':').at(-1)) : { row: 0, column: 0 };
  if (!dimension) await scanCells(entry, cell => {
    end = { row: Math.max(end.row, cell.address.row), column: Math.max(end.column, cell.address.column) };
  }, cancelled);
  const merges = [];
  const columnWidths = {};
  const rowHeights = {};
  const hiddenRows = [];
  const hiddenColumns = [];
  const columnStyles = {};
  const rowStyles = {};
  const defaults = elements(header, 'sheetFormatPr')[0]?.attributes || {};
  const defaultWidth = defaults.defaultColWidth;
  for (const column of elements(header, 'col')) {
    const begin = Number(column.attributes.min) - 1;
    const endColumn = Number(column.attributes.max) - 1;
    if (!Number.isInteger(begin) || !Number.isInteger(endColumn) || begin < 0 || endColumn < begin || endColumn >= 16384) throw new Error('Invalid worksheet column range');
    const width = column.attributes.width === undefined ? undefined : Math.max(1, columnPixels(column.attributes.width));
    for (let position = begin; position <= Math.min(end.column, endColumn); position++) {
      if (width !== undefined) columnWidths[position] = width;
      if (['1', 'true'].includes(column.attributes.hidden)) hiddenColumns.push(position);
      if (column.attributes.style !== undefined) columnStyles[position] = source.styles[Number(column.attributes.style)];
    }
  }
  await scanCells(entry, () => {}, cancelled, ({ row, attributes }) => {
    if (['1', 'true'].includes(attributes.hidden)) hiddenRows.push(row);
    if (attributes.ht !== undefined) {
      const height = Number(attributes.ht) * 96 / 72;
      if (!Number.isFinite(height) || height < 0 || height > 10000) throw new Error('Invalid worksheet row height');
      rowHeights[row] = height;
    }
    if (attributes.s !== undefined && ['1', 'true'].includes(attributes.customFormat)) rowStyles[row] = source.styles[Number(attributes.s)];
  }, cancelled);
  await scanElements(entry, 'mergeCell', merge => {
    const [start, finish] = merge.attributes.ref.split(':').map(cellAddress);
    if (finish.row < start.row || finish.column < start.column || merges.length >= 100000) throw new Error('Invalid worksheet merges');
    merges.push({ s: { r: start.row, c: start.column }, e: { r: finish.row, c: finish.column } });
  }, cancelled);
  const height = defaults.defaultRowHeight === undefined ? 20 : Number(defaults.defaultRowHeight) * 96 / 72;
  if (!Number.isFinite(height) || height <= 0 || height > 10000) throw new Error('Invalid default worksheet row height');
  return { rows: end.row + 1, columns: end.column + 1, merges, columnWidth: defaultWidth === undefined ? 64 : Math.max(1, columnPixels(defaultWidth)), columnWidths, rowHeight: height, rowHeights, hiddenRows, hiddenColumns, columnStyles, rowStyles, defaultStyle: source.styles[0], gridLines: !['0', 'false'].includes(elements(header, 'sheetView')[0]?.attributes.showGridLines) };
}

export async function sheetWindow(source, index, range, cancelled) {
  const cells = [];
  const anchors = new Set((range.anchors || []).map(anchor => anchor.row + ':' + anchor.column));
  if (source.format === 'xls') {
    const workbook = read(source.data, { type: 'array', sheets: [index], dense: true, cellHTML: false, cellNF: true, bookVBA: false });
    const sheet = workbook.Sheets[source.names[index]];
    for (let row = range.startRow; row <= range.endRow; row++) for (let column = range.startColumn; column <= range.endColumn; column++) {
      const cell = sheet['!data']?.[row]?.[column];
      if (cell) cells.push({ row, column, text: utils.format_cell(cell), formula: cell.f || '', missingResult: !!cell.f && (cell.v === undefined || cell.v === null), numberFormat: cell.z });
    }
    for (const anchor of range.anchors || []) {
      const cell = sheet['!data']?.[anchor.row]?.[anchor.column];
      if (cell) cells.push({ ...anchor, text: utils.format_cell(cell), formula: cell.f || '', missingResult: !!cell.f && (cell.v === undefined || cell.v === null), numberFormat: cell.z });
    }
    return calculateWindow(source, index, cells, cancelled);
  }
  const strings = new Set();
  await scanCells(source.archive.get(source.sheets[index].path), cell => {
    const address = cell.address;
    if (address.row > range.endRow) return false;
    if (!anchors.has(address.row + ':' + address.column) && (address.row < range.startRow || address.column < range.startColumn || address.column > range.endColumn)) return;
    const raw = xmlText(elements(cell.content, 'v')[0]?.content || '');
    const type = cell.attributes.t;
    const inherited = cell.rowAttributes.s !== undefined && ['1', 'true'].includes(cell.rowAttributes.customFormat) ? source.styles[Number(cell.rowAttributes.s)] : range.columnStyles?.[address.column];
    const style = { ...(cell.attributes.s === undefined ? inherited || source.styles[0] : source.styles[Number(cell.attributes.s)]) };
    if (!style.align && (!type || type === 'n')) style.align = 'right';
    let text = type === 'inlineStr' ? textRuns(cell.content) : raw;
    if (type === 'b') text = raw === '1' ? 'TRUE' : 'FALSE';
    if ((!type || type === 'n') && raw !== '' && Number.isFinite(Number(raw)) && style.numberFormat) {
      try { text = SSF.format(style.numberFormat, Number(raw), { date1904: source.date1904 }); } catch { text = raw; }
    }
    const shared = type === 's' ? Number(raw) : null;
    if (shared !== null) {
      if (!Number.isSafeInteger(shared) || shared < 0) throw new Error('Invalid shared string index');
      strings.add(shared);
    }
    const definition = elements(cell.content, 'f')[0];
    const formula = xmlText(definition?.content || '');
    cells.push({ ...address, text, shared, style, runs: type === 'inlineStr' ? richText(cell.content, source.theme) : undefined, formula, sharedFormula: definition?.attributes.t === 'shared' ? definition.attributes.si : undefined, missingResult: !!definition && raw === '' && type !== 'str' });
  }, cancelled);
  await resolveSharedFormulas(source, source.sheets[index].path, cells, cancelled);
  if (strings.size) {
    const values = new Map();
    let position = 0;
    await scanElements(source.archive.get('xl/sharedStrings.xml'), 'si', entry => {
      if (strings.has(position)) { values.set(position, { text: textRuns(entry.content), runs: richText(entry.content, source.theme) }); strings.delete(position); }
      position++;
      return strings.size > 0;
    }, cancelled);
    if (strings.size) throw new Error('Missing shared string');
    for (const cell of cells) if (cell.shared !== null) Object.assign(cell, values.get(cell.shared));
  }
  return calculateWindow(source, index, cells, cancelled);
}

async function calculateWindow(source, index, cells, cancelled) {
  if (!cells.some(cell => cell.missingResult)) return cells;
  const names = source.format === 'xls' ? source.names : source.allSheets.map(sheet => sheet.name);
  const currentSheet = source.format === 'xls' ? index : source.allSheets.indexOf(source.sheets[index]);
  const blocks = new Map();
  const readCell = async (sheet, address) => {
    if (!names[sheet]) throw new FormulaError('#REF!');
    const firstRow = Math.floor(address.row / 256) * 256;
    const firstColumn = Math.floor(address.column / 32) * 32;
    const key = sheet + ':' + firstRow + ':' + firstColumn;
    if (!blocks.has(key)) {
      const values = new Map();
      const strings = new Set();
      let characters = 0;
      const keep = (row, column, value) => {
        characters += (typeof value.value === 'string' ? value.value.length : 0) + (value.formula?.length || 0);
        if (characters > 4 * 1024 * 1024) throw new FormulaError('#NUM!');
        values.set(row + ':' + column, value);
      };
      if (source.format === 'xls') {
        const workbook = read(source.data, { type: 'array', sheets: [sheet], dense: true, cellHTML: false, bookVBA: false });
        const data = workbook.Sheets[source.names[sheet]]['!data'] || [];
        for (let row = firstRow; row < firstRow + 256; row++) for (let column = firstColumn; column < firstColumn + 32; column++) {
          const cell = data[row]?.[column];
          if (cell) keep(row, column, { value: cell.v ?? null, formula: cell.f || '', error: cell.t === 'e' ? utils.format_cell(cell) : undefined });
        }
      } else {
        await scanCells(source.archive.get(source.allSheets[sheet].path), cell => {
          const position = cell.address;
          if (position.row >= firstRow + 256) return false;
          if (position.row < firstRow || position.column < firstColumn || position.column >= firstColumn + 32) return;
          const raw = xmlText(elements(cell.content, 'v')[0]?.content || '');
          const type = cell.attributes.t;
          let value = raw === '' ? null : !type || type === 'n' ? Number(raw) : type === 'b' ? raw === '1' : raw;
          if (type === 'inlineStr') value = textRuns(cell.content);
          if (type === 'str') value = raw;
          const shared = type === 's' ? Number(raw) : undefined;
          if (shared !== undefined) { if (!Number.isSafeInteger(shared) || shared < 0) throw new FormulaError('#REF!'); strings.add(shared); }
          const definition = elements(cell.content, 'f')[0];
          keep(position.row, position.column, { ...position, value, shared, formula: xmlText(definition?.content || ''), sharedFormula: definition?.attributes.t === 'shared' ? definition.attributes.si : undefined, error: type === 'e' ? raw : undefined });
        }, cancelled);
        await resolveSharedFormulas(source, source.allSheets[sheet].path, [...values.values()], cancelled);
        if (strings.size) {
          let position = 0;
          const resolved = new Map();
          await scanElements(source.archive.get('xl/sharedStrings.xml'), 'si', entry => {
            if (strings.has(position)) { const value = textRuns(entry.content); characters += value.length; if (characters > 4 * 1024 * 1024) throw new FormulaError('#NUM!'); resolved.set(position, value); strings.delete(position); }
            position++;
            return strings.size > 0;
          }, cancelled);
          if (strings.size) throw new FormulaError('#REF!');
          for (const value of values.values()) if (value.shared !== undefined) value.value = resolved.get(value.shared);
        }
      }
      if (blocks.size >= 4) blocks.delete(blocks.keys().next().value);
      blocks.set(key, values);
    }
    return blocks.get(key).get(address.row + ':' + address.column);
  };
  const evaluator = formulaEvaluator({ readCell, names, definedNames: source.definedNames, cancelled });
  try {
    for (const cell of cells) {
      if (!cell.missingResult) continue;
      try {
        const value = await evaluator.cell(currentSheet, cell);
        cell.text = typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : value === null ? '0' : typeof value === 'number' ? SSF.format(cell.style?.numberFormat || cell.numberFormat || 'General', value, { date1904: source.date1904 }) : String(value);
        cell.calculated = true;
      } catch (error) {
        if (!(error instanceof FormulaError)) throw error;
        cell.text = error.message;
        cell.calculationError = true;
      }
    }
  } finally { evaluator.dispose(); blocks.clear(); }
  return cells;
}

async function resolveSharedFormulas(source, path, cells, cancelled) {
  const unresolved = cells.filter(cell => cell.sharedFormula !== undefined && !cell.formula);
  if (!unresolved.length) return;
  const identifiers = new Set(unresolved.map(cell => cell.sharedFormula));
  const definitions = new Map();
  await scanCells(source.archive.get(path), cell => {
    const definition = elements(cell.content, 'f')[0];
    const identifier = definition?.attributes.si;
    if (definition?.content && identifiers.has(identifier)) {
      definitions.set(identifier, { ...cell.address, formula: xmlText(definition.content) });
      identifiers.delete(identifier);
    }
    return identifiers.size > 0;
  }, cancelled);
  for (const cell of unresolved) {
    const definition = definitions.get(cell.sharedFormula);
    cell.formula = definition ? rebaseFormula(definition.formula, cell.row - definition.row, cell.column - definition.column) : '#REF!';
  }
}
