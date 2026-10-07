import { read, utils } from 'xlsx';

export async function render({ data, frame, controls, signal, labels, status, guard }) {
  const workbook = read(data, { type: 'array', sheetRows: 1001, cellHTML: false, cellStyles: false, bookVBA: false });
  signal.throwIfAborted();
  const label = document.createElement('label');
  label.textContent = labels.sheet + ' ';
  const select = document.createElement('select');
  for (const name of workbook.SheetNames) {
    const option = document.createElement('option');
    option.value = option.textContent = name;
    select.append(option);
  }
  if (!workbook.SheetNames.length) throw new Error('Empty workbook');
  label.append(select);
  controls.append(label);
  const draw = guard(() => {
    const sheet = workbook.Sheets[select.value];
    const range = utils.decode_range(sheet['!ref'] || 'A1');
    const truncated = !!sheet['!fullref'] || range.e.r - range.s.r >= 1000 || range.e.c - range.s.c >= 100;
    range.e.r = Math.min(range.e.r, range.s.r + 999);
    range.e.c = Math.min(range.e.c, range.s.c + 99);
    const table = frame.contentDocument.createElement('table');
    const caption = frame.contentDocument.createElement('caption');
    caption.textContent = select.value;
    table.append(caption);
    for (let row = range.s.r; row <= range.e.r; row++) {
      const line = frame.contentDocument.createElement('tr');
      for (let column = range.s.c; column <= range.e.c; column++) {
        const cell = frame.contentDocument.createElement('td');
        const value = sheet[utils.encode_cell({ r: row, c: column })];
        cell.textContent = value ? utils.format_cell(value) : '';
        line.append(cell);
      }
      table.append(line);
    }
    frame.contentDocument.body.replaceChildren(table);
    status.textContent = truncated ? labels.limited : '';
  });
  select.addEventListener('change', draw);
  draw();
}
