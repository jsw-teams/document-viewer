import { read, utils } from 'xlsx';

export async function render({ data, frame, viewport, controls, signal, labels, status, guard }) {
  const workbook = read(data, { type: 'array', sheetRows: 1001, cellHTML: false, cellStyles: false, bookVBA: false });
  signal.throwIfAborted();
  if (!workbook.SheetNames.length) throw new Error('Empty workbook');
  const tabs = document.createElement('div');
  tabs.className = 'document-viewer-sheets';
  tabs.setAttribute('role', 'tablist');
  tabs.setAttribute('aria-label', labels.sheet);
  viewport.setAttribute('role', 'tabpanel');
  viewport.tabIndex = 0;
  controls.append(tabs);
  let selected = 0;
  const buttons = workbook.SheetNames.map((name, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.id = viewport.id + '-sheet-' + index;
    button.setAttribute('role', 'tab');
    button.setAttribute('aria-controls', viewport.id);
    button.textContent = name;
    tabs.append(button);
    return button;
  });
  const draw = guard(() => {
    const sheet = workbook.Sheets[workbook.SheetNames[selected]];
    const range = utils.decode_range(sheet['!ref'] || 'A1');
    const truncated = !!sheet['!fullref'] || range.e.r - range.s.r >= 1000 || range.e.c - range.s.c >= 100;
    range.e.r = Math.min(range.e.r, range.s.r + 999);
    range.e.c = Math.min(range.e.c, range.s.c + 99);
    const table = frame.contentDocument.createElement('table');
    const caption = frame.contentDocument.createElement('caption');
    caption.textContent = workbook.SheetNames[selected];
    const head = frame.contentDocument.createElement('thead');
    const headings = frame.contentDocument.createElement('tr');
    const corner = frame.contentDocument.createElement('td');
    headings.append(corner);
    for (let column = range.s.c; column <= range.e.c; column++) {
      const heading = frame.contentDocument.createElement('th');
      heading.scope = 'col';
      heading.textContent = utils.encode_col(column);
      headings.append(heading);
    }
    head.append(headings);
    const body = frame.contentDocument.createElement('tbody');
    for (let row = range.s.r; row <= range.e.r; row++) {
      const line = frame.contentDocument.createElement('tr');
      const heading = frame.contentDocument.createElement('th');
      heading.scope = 'row';
      heading.textContent = String(row + 1);
      line.append(heading);
      for (let column = range.s.c; column <= range.e.c; column++) {
        const cell = frame.contentDocument.createElement('td');
        const value = sheet[utils.encode_cell({ r: row, c: column })];
        cell.textContent = value ? utils.format_cell(value) : '';
        line.append(cell);
      }
      body.append(line);
    }
    table.append(caption, head, body);
    frame.contentDocument.body.replaceChildren(table);
    buttons.forEach((button, index) => {
      button.setAttribute('aria-selected', String(index === selected));
      button.tabIndex = index === selected ? 0 : -1;
    });
    viewport.setAttribute('aria-labelledby', buttons[selected].id);
    status.textContent = truncated ? labels.limited : '';
  });
  buttons.forEach((button, index) => {
    button.addEventListener('click', () => { selected = index; draw(); });
    button.addEventListener('keydown', event => {
      const direction = getComputedStyle(tabs).direction === 'rtl' ? -1 : 1;
      const next = { ArrowRight: (index + direction + buttons.length) % buttons.length, ArrowLeft: (index - direction + buttons.length) % buttons.length, Home: 0, End: buttons.length - 1 }[event.key];
      if (next === undefined) return;
      event.preventDefault();
      selected = next;
      draw();
      buttons[selected].focus();
    });
  });
  draw();
}
