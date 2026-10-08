import { read, utils, SSF } from 'xlsx';
import { officeArchive, xmlPart, scanXml, scanElements, scanCells } from './archive.js';
import { elements, cellAddress, textRuns, xmlText } from './xml.js';

const color = value => /^[a-f0-9]{6}$/i.test(value?.slice(-6) || '') ? '#' + value.slice(-6) : undefined;

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
  const sheets = elements(workbook, 'sheet').filter(entry => entry.attributes.state !== 'veryHidden').map(entry => ({ name: entry.attributes.name, path: paths.get(entry.attributes['r:id']) }));
  if (!sheets.length || sheets.some(sheet => !archive.has(sheet.path))) throw new Error('Invalid workbook sheets');
  const styles = archive.has('xl/styles.xml') ? await xmlPart(archive.get('xl/styles.xml')) : '';
  const formats = new Map(elements(styles, 'numFmt').map(entry => [Number(entry.attributes.numFmtId), entry.attributes.formatCode]));
  const fonts = elements(elements(styles, 'fonts')[0]?.content || '', 'font');
  const fills = elements(elements(styles, 'fills')[0]?.content || '', 'fill');
  const cellStyles = elements(elements(styles, 'cellXfs')[0]?.content || '', 'xf').map(entry => {
    const font = fonts[Number(entry.attributes.fontId)]?.content || '';
    const fill = fills[Number(entry.attributes.fillId)]?.content || '';
    const alignment = elements(entry.content, 'alignment')[0]?.attributes || {};
    return { numberFormat: formats.get(Number(entry.attributes.numFmtId)) || SSF.get_table()[Number(entry.attributes.numFmtId)],
      bold: /<(?:[\w.-]+:)?b(?:\s|\/>|>)/.test(font), italic: /<(?:[\w.-]+:)?i(?:\s|\/>|>)/.test(font),
      color: color(elements(font, 'color')[0]?.attributes.rgb), background: color(elements(fill, 'fgColor')[0]?.attributes.rgb),
      align: ['left', 'right', 'center'].includes(alignment.horizontal) ? alignment.horizontal : undefined };
  });
  return { data, format, archive, sheets, names: sheets.map(sheet => sheet.name), styles: cellStyles, date1904: ['1', 'true'].includes(elements(workbook, 'workbookPr')[0]?.attributes.date1904) };
}

export async function sheetInfo(source, index, cancelled) {
  if (!source.names[index]) throw new Error('Unknown worksheet');
  if (source.format === 'xls') {
    const workbook = read(source.data, { type: 'array', sheets: [index], dense: true, cellHTML: false, bookVBA: false });
    const sheet = workbook.Sheets[source.names[index]];
    const range = utils.decode_range(sheet['!ref'] || 'A1');
    return { rows: range.e.r + 1, columns: range.e.c + 1, merges: sheet['!merges'] || [] };
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
  await scanElements(entry, 'mergeCell', merge => {
    const [start, finish] = merge.attributes.ref.split(':').map(cellAddress);
    if (finish.row < start.row || finish.column < start.column || merges.length >= 100000) throw new Error('Invalid worksheet merges');
    merges.push({ s: { r: start.row, c: start.column }, e: { r: finish.row, c: finish.column } });
  }, cancelled);
  return { rows: end.row + 1, columns: end.column + 1, merges };
}

export async function sheetWindow(source, index, range, cancelled) {
  const cells = [];
  const anchors = new Set((range.anchors || []).map(anchor => anchor.row + ':' + anchor.column));
  if (source.format === 'xls') {
    const workbook = read(source.data, { type: 'array', sheets: [index], dense: true, cellHTML: false, cellNF: true, bookVBA: false });
    const sheet = workbook.Sheets[source.names[index]];
    for (let row = range.startRow; row <= range.endRow; row++) for (let column = range.startColumn; column <= range.endColumn; column++) {
      const cell = sheet['!data']?.[row]?.[column];
      if (cell) cells.push({ row, column, text: utils.format_cell(cell), formula: cell.f || '' });
    }
    for (const anchor of range.anchors || []) {
      const cell = sheet['!data']?.[anchor.row]?.[anchor.column];
      if (cell) cells.push({ ...anchor, text: utils.format_cell(cell), formula: cell.f || '' });
    }
    return cells;
  }
  const strings = new Set();
  await scanCells(source.archive.get(source.sheets[index].path), cell => {
    const address = cell.address;
    if (address.row > range.endRow) return false;
    if (!anchors.has(address.row + ':' + address.column) && (address.row < range.startRow || address.column < range.startColumn || address.column > range.endColumn)) return;
    const raw = xmlText(elements(cell.content, 'v')[0]?.content || '');
    const type = cell.attributes.t;
    const style = source.styles[Number(cell.attributes.s)] || {};
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
    cells.push({ ...address, text, shared, style, formula: xmlText(elements(cell.content, 'f')[0]?.content || '') });
  }, cancelled);
  if (strings.size) {
    const values = new Map();
    let position = 0;
    await scanElements(source.archive.get('xl/sharedStrings.xml'), 'si', entry => {
      if (strings.has(position)) { values.set(position, textRuns(entry.content)); strings.delete(position); }
      position++;
      return strings.size > 0;
    }, cancelled);
    if (strings.size) throw new Error('Missing shared string');
    for (const cell of cells) if (cell.shared !== null) cell.text = values.get(cell.shared);
  }
  return cells;
}
