import { documentPage } from '../pages.js';
import { cellAddress, columnName } from '../sheets/xml.js';
import { applyCellStyle } from '../sheets/styles.js';
import { buttonContent } from '../icons.js';
import { columnLayout } from '../sheets/columns.js';
import { columnControls } from '../sheets/column-controls.js';

export async function render({ data, format, frame, viewport, controls, signal, labels, status, setCleanup }) {
  const worker = new Worker(new URL('../sheets.worker.js', import.meta.url), { type: 'module' });
  let identifier = 0;
  const pending = new Map();
  let observer = null;
  let selected = 0;
  let revision = 0;
  let navigation = 0;
  let timer = null;
  let metadata = null;
  let columnOffsets = [];
  let hiddenRows = new Set();
  let hiddenColumns = new Set();
  let pages = [];
  let column = 0;
  let active = { row: 0, column: 0 };
  const visible = new Set();
  const mounted = new Map();
  let drawing = null;
  let repeat = false;
  const doc = frame.contentDocument;
  let widths = null;
  const cleanup = () => {
    revision++;
    clearTimeout(timer);
    observer?.disconnect();
    widths?.destroy();
    widths = null;
    worker.terminate();
    for (const request of pending.values()) request.reject(new DOMException('Worksheet closed', 'AbortError'));
    pending.clear();
    mounted.clear();
    pages = [];
    metadata = null;
  };
  worker.onmessage = ({ data: response }) => {
    const request = pending.get(response.id);
    if (!request) return;
    pending.delete(response.id);
    if (response.error) request.reject(new Error(response.error));
    else request.resolve(response.result);
  };
  worker.onerror = () => { for (const request of pending.values()) request.reject(new Error('Worksheet worker failed')); pending.clear(); };
  function request(action, fields = {}, transfer = []) {
    signal.throwIfAborted();
    for (const previous of pending.values()) previous.reject(new DOMException('Superseded worksheet request', 'AbortError'));
    pending.clear();
    const id = ++identifier;
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject });
      worker.postMessage({ action, id, ...fields }, transfer);
    });
  }
  setCleanup(cleanup);
  const names = await request('open', { buffer: data.buffer, format }, [data.buffer]);
  signal.throwIfAborted();
  if (!names.length) throw new Error('Empty workbook');
  const formulaBar = document.createElement('form');
  formulaBar.className = 'document-viewer-formula';
  formulaBar.noValidate = true;
  const address = document.createElement('input');
  address.type = 'text';
  address.value = 'A1';
  address.setAttribute('aria-label', labels.cellAddress);
  address.autocomplete = 'off';
  address.spellcheck = false;
  const jump = document.createElement('button');
  jump.type = 'submit';
  buttonContent(jump, labels.go, 'locate');
  const value = document.createElement('output');
  value.setAttribute('aria-label', labels.cellValue);
  value.tabIndex = 0;
  formulaBar.append(address, jump, value);
  controls.append(formulaBar);
  widths = columnControls({ frame, controls, signal, labels,
    current: () => ({ index: selected, metadata, offsets: columnOffsets }), changed: updateColumns });
  const tabs = document.createElement('div');
  tabs.className = 'document-viewer-sheets';
  tabs.setAttribute('role', 'tablist');
  tabs.setAttribute('aria-label', labels.sheet);
  const footer = document.createElement('div');
  footer.className = 'document-viewer-sheet-footer';
  const dimensions = document.createElement('span');
  footer.append(tabs, dimensions);
  viewport.after(footer);
  setCleanup(() => { cleanup(); footer.remove(); });
  viewport.setAttribute('role', 'tabpanel');
  viewport.tabIndex = 0;
  const buttons = names.map((name, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.id = viewport.id + '-sheet-' + index;
    button.setAttribute('role', 'tab');
    button.setAttribute('aria-controls', viewport.id);
    buttonContent(button, name, 'table');
    tabs.append(button);
    return button;
  });
  const stylesheet = doc.createElement('style');
  stylesheet.textContent = '.sheet-page{min-height:100px;overflow:visible;position:relative;padding:0;margin-inline:0}.sheet-grid{border-collapse:collapse;table-layout:fixed;font-size:14px;min-width:0;margin:0;background:#fff;color:#000}.sheet-grid th,.sheet-grid td{box-sizing:border-box;height:28px;padding:3px 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border:1px solid var(--document-viewer-line);line-height:20px;text-align:start}.sheet-grid td{cursor:cell}.sheet-grid td[aria-selected=true]{outline:2px solid var(--document-viewer-accent);outline-offset:-2px}.sheet-grid th{position:static;background:var(--document-viewer-surface);color:var(--document-viewer-ink)}.sheet-grid .sheet-spacer{padding:0;border:0}.sheet-grid caption{box-sizing:border-box;text-align:start;height:32px;padding:0;font-weight:700;color:var(--document-viewer-ink)}.sheet-page>h2{position:sticky;inset-inline-start:0;width:fit-content}';
  doc.head.append(stylesheet);
  const gridStyle = doc.createElement('style');
  gridStyle.textContent = '.sheet-grid{font-size:11pt}.sheet-grid th,.sheet-grid td{height:var(--sheet-row-height,20px);padding:1px 3px;line-height:normal}.sheet-grid thead th,.sheet-grid thead td{height:24px}.sheet-cell-content{display:block;max-height:var(--sheet-cell-height);overflow:hidden;white-space:inherit;text-overflow:ellipsis}.sheet-grid[data-grid-lines=false] td{border-color:transparent}.sheet-grid td[aria-selected=true]{outline:2px solid var(--document-viewer-accent);outline-offset:-2px}';
  doc.head.append(gridStyle);
  const rowHeight = row => hiddenRows.has(row) ? 0 : metadata.rowHeights?.[row] ?? metadata.rowHeight;
  function updateColumns(settle = false) {
    if (!metadata || !widths) return;
    const zoom = parseFloat(frame.contentWindow.getComputedStyle(doc.body).zoom) || 1;
    columnOffsets = columnLayout(metadata, widths.preference(selected), Math.max(0, doc.documentElement.clientWidth / zoom - 24 - 48));
    const total = 48 + columnOffsets[metadata.columns];
    for (const page of pages) page.style.width = total + 'px';
    for (const table of doc.querySelectorAll('.sheet-grid')) {
      table.style.width = total + 'px';
      for (const node of table.querySelectorAll('col[data-column]')) {
        const index = Number(node.dataset.column);
        node.style.width = (columnOffsets[index + 1] - columnOffsets[index]) + 'px';
      }
      for (const node of table.querySelectorAll('[data-sheet-spacer]')) node.style.width = (node.dataset.sheetSpacer === 'leading' ? columnOffsets[Number(table.dataset.startColumn)] : columnOffsets[metadata.columns] - columnOffsets[Number(table.dataset.endColumn) + 1]) + 'px';
    }
    widths.update();
    if (settle) fitCells(doc);
  }
  function fitCells(root) {
    for (const content of root.querySelectorAll('[data-shrink]')) {
      content.style.fontSize = '';
      const width = Math.max(1, content.parentElement.clientWidth - 6);
      if (content.scrollWidth > width) content.style.fontSize = Math.max(1, parseFloat(frame.contentWindow.getComputedStyle(content).fontSize) * width / content.scrollWidth) + 'px';
    }
  }
  function chooseCell(cell) {
    active = { row: Number(cell.dataset.row), column: Number(cell.dataset.column) };
    address.value = columnName(active.column) + (active.row + 1);
    value.textContent = cell.dataset.formula ? '=' + cell.dataset.formula : cell.textContent;
    for (const node of doc.querySelectorAll('[aria-selected=true]')) node.removeAttribute('aria-selected');
    cell.setAttribute('aria-selected', 'true');
    cell.tabIndex = 0;
  }
  function tableFor(index, cells) {
    const table = doc.createElement('table');
    table.className = 'sheet-grid';
    table.setAttribute('role', 'grid');
    table.setAttribute('aria-readonly', 'true');
    table.dataset.gridLines = String(metadata.gridLines);
    applyCellStyle(table, metadata.defaultStyle);
    table.setAttribute('aria-rowcount', String(metadata.rows + 1));
    table.setAttribute('aria-colcount', String(metadata.columns + 1));
    table.style.width = (48 + columnOffsets[metadata.columns]) + 'px';
    const caption = doc.createElement('caption');
    caption.textContent = names[selected];
    const start = index * 50;
    const end = Math.min(metadata.rows - 1, start + 49);
    const endColumn = Math.min(metadata.columns - 1, column + 19);
    table.dataset.startColumn = String(column);
    table.dataset.endColumn = String(endColumn);
    const merges = metadata.merges.filter(item => item.s.r <= end && item.e.r >= start && item.s.c <= endColumn && item.e.c >= column);
    const values = new Map(cells.map(cell => [cell.row + ':' + cell.column, cell]));
    const head = doc.createElement('thead');
    const headings = doc.createElement('tr');
    headings.append(doc.createElement('td'));
    const spacer = (line, width, side) => {
      if (!width) return;
      const node = doc.createElement('td');
      node.className = 'sheet-spacer';
      node.dataset.sheetSpacer = side;
      node.style.width = width + 'px';
      node.setAttribute('aria-hidden', 'true');
      line.append(node);
    };
    const columns = doc.createElement('colgroup');
    const gutter = doc.createElement('col');
    gutter.style.width = '48px';
    columns.append(gutter);
    const addColumn = (width, index, side) => { const node = doc.createElement('col'); node.style.width = width + 'px'; if (index !== null) node.dataset.column = String(index); else node.dataset.sheetSpacer = side; columns.append(node); };
    if (columnOffsets[column]) addColumn(columnOffsets[column], null, 'leading');
    spacer(headings, columnOffsets[column], 'leading');
    for (let current = column; current <= endColumn; current++) {
      if (hiddenColumns.has(current)) continue;
      addColumn(columnOffsets[current + 1] - columnOffsets[current], current);
      const heading = doc.createElement('th');
      heading.scope = 'col';
      heading.id = viewport.id + '-page-' + index + '-column-' + current;
      heading.textContent = columnName(current);
      heading.append(widths.handle(current, heading));
      headings.append(heading);
    }
    const remaining = columnOffsets[metadata.columns] - columnOffsets[endColumn + 1];
    if (remaining) addColumn(remaining, null, 'trailing');
    spacer(headings, remaining, 'trailing');
    head.append(headings);
    const body = doc.createElement('tbody');
    for (let row = start; row <= end; row++) {
      if (hiddenRows.has(row)) continue;
      const line = doc.createElement('tr');
      line.style.setProperty('--sheet-row-height', rowHeight(row) + 'px');
      line.setAttribute('aria-rowindex', String(row + 2));
      const heading = doc.createElement('th');
      heading.scope = 'row';
      heading.textContent = String(row + 1);
      line.append(heading);
      spacer(line, columnOffsets[column], 'leading');
      for (let current = column; current <= endColumn; current++) {
        if (hiddenColumns.has(current)) continue;
        const merge = merges.find(item => row >= item.s.r && row <= item.e.r && current >= item.s.c && current <= item.e.c);
        let firstRow = merge ? Math.max(start, merge.s.r) : row;
        let firstColumn = merge ? Math.max(column, merge.s.c) : current;
        while (hiddenRows.has(firstRow) && firstRow <= end) firstRow++;
        while (hiddenColumns.has(firstColumn) && firstColumn <= endColumn) firstColumn++;
        if (merge && (row !== firstRow || current !== firstColumn)) continue;
        const node = doc.createElement('td');
        node.dataset.row = String(row);
        node.dataset.column = String(current);
        node.setAttribute('aria-colindex', String(current + 2));
        const cell = values.get((merge?.s.r ?? row) + ':' + (merge?.s.c ?? current));
        const content = doc.createElement('span');
        content.className = 'sheet-cell-content';
        if (cell?.runs?.length) {
          for (const run of cell.runs) {
            const span = doc.createElement('span');
            span.textContent = run.text;
            applyCellStyle(span, run.style);
            content.append(span);
          }
        } else content.textContent = cell?.text || '';
        node.append(content);
        node.dataset.formula = cell?.formula || '';
        node.tabIndex = -1;
        const style = cell?.style || metadata.rowStyles?.[row] || metadata.columnStyles?.[current] || {};
        applyCellStyle(node, style);
        let height = rowHeight(row);
        if (merge) {
          const rows = Array.from({ length: Math.min(end, merge.e.r) - firstRow + 1 }, (_, position) => firstRow + position).filter(position => !hiddenRows.has(position));
          const columns = Array.from({ length: Math.min(endColumn, merge.e.c) - firstColumn + 1 }, (_, position) => firstColumn + position).filter(position => !hiddenColumns.has(position));
          node.rowSpan = rows.length;
          node.colSpan = columns.length;
          height = rows.reduce((total, position) => total + rowHeight(position), 0);
        }
        node.style.setProperty('--sheet-cell-height', Math.max(0, height - 3) + 'px');
        if (style.rotation === 255) content.style.writingMode = 'vertical-rl';
        else if (style.rotation > 0 && style.rotation <= 180) { content.style.transform = 'rotate(' + (style.rotation <= 90 ? -style.rotation : 180 - style.rotation) + 'deg)'; content.style.transformOrigin = 'center'; }
        if (style.shrink) content.dataset.shrink = 'true';
        node.addEventListener('click', () => { chooseCell(node); node.focus({ preventScroll: true }); });
        if (row === active.row && current === active.column) chooseCell(node);
        line.append(node);
      }
      spacer(line, remaining, 'trailing');
      body.append(line);
    }
    table.append(caption, columns, head, body);
    return table;
  }
  function draw() {
    if (!metadata || signal.aborted) return Promise.resolve();
    if (drawing) { repeat = true; return drawing; }
    drawing = paint().finally(() => {
      drawing = null;
      if (repeat) { repeat = false; return draw(); }
    });
    return drawing;
  }
  async function paint() {
    const version = revision;
    try {
      for (const [index, mountedColumn] of mounted) {
        if (!visible.has(index) || mountedColumn !== column) {
          pages[index].querySelector('table')?.remove();
          pages[index].setAttribute('aria-busy', 'true');
          mounted.delete(index);
        }
      }
      for (const index of [...visible].slice(0, 5)) {
        if (mounted.get(index) === column) continue;
        const startColumn = column;
        const startRow = index * 50;
        const endRow = Math.min(metadata.rows - 1, startRow + 49);
        const endColumn = Math.min(metadata.columns - 1, startColumn + 19);
        const anchors = metadata.merges.filter(merge => merge.s.r <= endRow && merge.e.r >= startRow && merge.s.c <= endColumn && merge.e.c >= startColumn).map(merge => ({ row: merge.s.r, column: merge.s.c }));
        const cells = await request('window', { index: selected, range: { startRow, endRow, startColumn, endColumn, anchors, columnStyles: metadata.columnStyles } });
        if (version !== revision || signal.aborted) return;
        if (startColumn !== column) { repeat = true; return; }
        if (!visible.has(index)) continue;
        pages[index].querySelector('table')?.remove();
        pages[index].append(tableFor(index, cells));
        widths.update();
        fitCells(pages[index]);
        const zoom = parseFloat(frame.contentWindow.getComputedStyle(doc.body).zoom) || 1;
        pages[index].style.minHeight = Math.ceil(pages[index].getBoundingClientRect().height / zoom) + 'px';
        pages[index].setAttribute('aria-busy', 'false');
        mounted.set(index, column);
      }
      status.textContent = '';
    } catch (error) { if (!signal.aborted && error.name !== 'AbortError') status.textContent = labels.error; }
  }
  async function select(index) {
    widths.finish();
    const version = ++revision;
    navigation++;
    clearTimeout(timer);
    selected = index;
    metadata = null;
    observer?.disconnect();
    mounted.clear();
    visible.clear();
    column = 0;
    active = { row: 0, column: 0 };
    address.value = 'A1';
    value.textContent = '';
    doc.body.replaceChildren();
    status.textContent = labels.loading;
    buttons.forEach((button, position) => { button.setAttribute('aria-selected', String(position === index)); button.tabIndex = position === index ? 0 : -1; });
    viewport.setAttribute('aria-labelledby', buttons[index].id);
    try {
      const info = await request('sheet', { index });
      if (version !== revision || signal.aborted) return;
      metadata = info;
      hiddenRows = new Set(info.hiddenRows);
      hiddenColumns = new Set(info.hiddenColumns);
      pages = [];
      updateColumns();
      dimensions.textContent = info.rows.toLocaleString() + ' × ' + info.columns.toLocaleString();
      const count = Math.ceil(info.rows / 50);
      pages = Array.from({ length: count }, (_, current) => {
        const page = documentPage(doc, labels, current, count);
        page.classList.add('sheet-page');
        page.setAttribute('aria-busy', 'true');
        page.style.width = (48 + columnOffsets[info.columns]) + 'px';
        page.style.maxWidth = 'none';
        page.style.boxSizing = 'border-box';
        let height = 100;
        for (let row = current * 50; row < Math.min(info.rows, current * 50 + 50); row++) height += rowHeight(row);
        page.style.minHeight = height + 'px';
        doc.body.append(page);
        return page;
      });
      frame.contentWindow.scrollTo(0, 0);
      visible.add(0);
      await draw();
      if (version !== revision || signal.aborted) return;
      observer = new IntersectionObserver(entries => {
        for (const entry of entries) {
          const current = Number(entry.target.dataset.documentPage) - 1;
          const rectangle = entry.target.getBoundingClientRect();
          if (rectangle.bottom >= -200 && rectangle.top <= frame.clientHeight + 200) visible.add(current); else visible.delete(current);
        }
        void draw();
      }, { root: doc, rootMargin: '200px 0px' });
      for (const page of pages) observer.observe(page);
    } catch (error) { if (!signal.aborted && version === revision && error.name !== 'AbortError') { status.textContent = labels.error; throw error; } }
  }
  async function go(target, focus = false) {
    if (!metadata || target.row >= metadata.rows || target.column >= metadata.columns) { address.setAttribute('aria-invalid', 'true'); status.textContent = labels.invalidAddress; return; }
    if (hiddenRows.has(target.row) || hiddenColumns.has(target.column)) { address.setAttribute('aria-invalid', 'true'); status.textContent = labels.hiddenCell; return; }
    address.removeAttribute('aria-invalid');
    const version = ++navigation;
    clearTimeout(timer);
    active = target;
    address.value = columnName(target.column) + (target.row + 1);
    column = Math.max(0, Math.min(metadata.columns - 1, Math.floor(target.column / 20) * 20));
    const index = Math.floor(target.row / 50);
    visible.clear();
    visible.add(index);
    const zoom = parseFloat(frame.contentWindow.getComputedStyle(doc.body).zoom) || 1;
    frame.contentWindow.scrollTo(columnOffsets[column] * zoom, pages[index].getBoundingClientRect().top + frame.contentWindow.scrollY);
    await draw();
    if (version !== navigation || signal.aborted) return;
    const cell = pages[index]?.querySelector('[data-row="' + target.row + '"][data-column="' + target.column + '"]');
    if (cell) { chooseCell(cell); cell.scrollIntoView({ block: 'nearest', inline: 'nearest' }); if (focus) cell.focus({ preventScroll: true }); }
  }
  formulaBar.addEventListener('submit', event => {
    event.preventDefault();
    try { void go(cellAddress(address.value)); }
    catch { address.setAttribute('aria-invalid', 'true'); status.textContent = labels.invalidAddress; }
  });
  address.addEventListener('input', () => address.removeAttribute('aria-invalid'));
  doc.addEventListener('keydown', event => {
    if (!event.target.matches('[data-row]')) return;
    const moves = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const move = moves[event.key];
    if (!move) return;
    event.preventDefault();
    const target = { row: Math.max(0, active.row + move[0]), column: Math.max(0, active.column + move[1]) };
    while (hiddenRows.has(target.row) && move[0] && target.row >= 0 && target.row < metadata.rows) target.row += move[0];
    while (hiddenColumns.has(target.column) && move[1] && target.column >= 0 && target.column < metadata.columns) target.column += move[1];
    if (target.row >= 0 && target.column >= 0) void go(target, true);
  }, { signal });
  frame.contentWindow.addEventListener('scroll', () => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      const zoom = parseFloat(frame.contentWindow.getComputedStyle(doc.body).zoom) || 1;
      const offset = Math.max(0, frame.contentWindow.scrollX / zoom - 48);
      let begin = 0;
      let end = columnOffsets.length - 1;
      while (begin < end) {
        const middle = Math.ceil((begin + end) / 2);
        if (columnOffsets[middle] <= offset) begin = middle; else end = middle - 1;
      }
      const next = begin;
      if (metadata && next !== column) { column = Math.min(metadata.columns - 1, next); void draw(); }
    }, 80);
  }, { signal, passive: true });
  buttons.forEach((button, index) => {
    button.addEventListener('click', () => { void select(index).catch(() => {}); });
    button.addEventListener('keydown', event => {
      const direction = getComputedStyle(tabs).direction === 'rtl' ? -1 : 1;
      const next = { ArrowRight: (index + direction + buttons.length) % buttons.length, ArrowLeft: (index - direction + buttons.length) % buttons.length, Home: 0, End: buttons.length - 1 }[event.key];
      if (next === undefined) return;
      event.preventDefault();
      void select(next).catch(() => {});
      buttons[next].focus();
    });
  });
  await select(0);
  return () => { cleanup(); footer.remove(); };
}
