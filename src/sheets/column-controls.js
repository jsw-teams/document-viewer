import { buttonContent, controlIcon } from '../icons.js';
import { columnName } from './xml.js';

export function columnControls({ frame, controls, signal, labels, current, changed }) {
  const doc = frame.contentDocument;
  const preferences = new Map();
  const preference = index => {
    if (!preferences.has(index)) preferences.set(index, { widths: new Map(), stretch: false });
    return preferences.get(index);
  };
  const tools = document.createElement('div');
  tools.className = 'document-viewer-column-tools';
  const stretch = document.createElement('button');
  stretch.type = 'button';
  buttonContent(stretch, labels.stretchColumns, 'columns');
  const reset = document.createElement('button');
  reset.type = 'button';
  buttonContent(reset, labels.resetWidths, 'reset');
  const hint = document.createElement('span');
  hint.textContent = labels.resizeColumnsHint;
  tools.append(stretch, reset, hint);
  controls.append(tools);
  let drag = null;
  let scheduled = null;
  const zoom = () => parseFloat(frame.contentWindow.getComputedStyle(doc.body).zoom) || 1;
  const queue = () => { if (scheduled === null) scheduled = requestAnimationFrame(() => { scheduled = null; if (!signal.aborted) changed(); }); };
  function update() {
    const state = current();
    stretch.disabled = reset.disabled = !state.metadata;
    stretch.setAttribute('aria-pressed', String(preference(state.index).stretch));
    if (!state.metadata) return;
    for (const handle of doc.querySelectorAll('[data-resize-column]')) {
      const index = Number(handle.dataset.resizeColumn);
      const width = Math.round(state.offsets[index + 1] - state.offsets[index]);
      handle.setAttribute('aria-valuenow', String(width));
      handle.setAttribute('aria-valuemax', String(Math.max(2400, width)));
      handle.setAttribute('aria-valuetext', labels.columnWidth.replace('{width}', width));
    }
  }
  function manual(index, width, settle = false) {
    const state = current();
    const settings = preference(state.index);
    if (settings.stretch) {
      for (let column = 0; column < state.metadata.columns; column++) {
        const displayed = state.offsets[column + 1] - state.offsets[column];
        if (displayed) settings.widths.set(column, displayed);
      }
      settings.stretch = false;
    }
    settings.widths.set(index, Math.max(24, Math.min(2400, width)));
    if (settle) changed(true); else queue();
  }
  function finish(cancel = false) {
    if (!drag) return;
    const previous = drag;
    drag = null;
    if (cancel) preferences.set(previous.sheet, previous.before);
    doc.body.removeAttribute('data-resizing-columns');
    if (previous.handle.hasPointerCapture(previous.pointer)) previous.handle.releasePointerCapture(previous.pointer);
    if (scheduled !== null) { cancelAnimationFrame(scheduled); scheduled = null; }
    changed(true);
  }
  doc.addEventListener('pointerdown', event => {
    const handle = event.target.closest?.('[data-resize-column]');
    if (!handle || event.button !== 0 || drag) return;
    event.preventDefault();
    const state = current();
    const settings = preference(state.index);
    const column = Number(handle.dataset.resizeColumn);
    handle.focus({ preventScroll: true });
    drag = { handle, pointer: event.pointerId, column, sheet: state.index, start: event.clientX, zoom: zoom(), width: state.offsets[column + 1] - state.offsets[column], before: { widths: new Map(settings.widths), stretch: settings.stretch } };
    handle.setPointerCapture(event.pointerId);
    doc.body.setAttribute('data-resizing-columns', '');
  }, { signal });
  doc.addEventListener('pointermove', event => {
    if (!drag || event.pointerId !== drag.pointer) return;
    event.preventDefault();
    manual(drag.column, drag.width + (event.clientX - drag.start) / drag.zoom);
  }, { signal });
  doc.addEventListener('pointerup', event => { if (drag?.pointer === event.pointerId) finish(); }, { signal });
  doc.addEventListener('pointercancel', event => { if (drag?.pointer === event.pointerId) finish(true); }, { signal });
  doc.addEventListener('lostpointercapture', event => { if (drag?.pointer === event.pointerId) finish(true); }, { signal });
  doc.addEventListener('keydown', event => {
    if (event.key === 'Escape' && drag) { event.preventDefault(); event.stopImmediatePropagation(); finish(true); return; }
    const handle = event.target.closest?.('[data-resize-column]');
    if (!handle || !['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const column = Number(handle.dataset.resizeColumn);
    const state = current();
    const width = state.offsets[column + 1] - state.offsets[column];
    manual(column, event.key === 'Home' ? 24 : event.key === 'End' ? 2400 : width + (event.key === 'ArrowRight' ? 1 : -1) * (event.shiftKey ? 1 : 8), true);
  }, { signal, capture: true });
  doc.addEventListener('dblclick', event => {
    const handle = event.target.closest?.('[data-resize-column]');
    if (!handle) return;
    event.preventDefault();
    const settings = preference(current().index);
    settings.stretch = false;
    settings.widths.delete(Number(handle.dataset.resizeColumn));
    changed(true);
  }, { signal });
  stretch.addEventListener('click', () => { finish(); const settings = preference(current().index); settings.stretch = !settings.stretch; changed(true); }, { signal });
  reset.addEventListener('click', () => { finish(); preferences.set(current().index, { widths: new Map(), stretch: false }); changed(true); }, { signal });
  frame.contentWindow.addEventListener('document-viewer-zoom', queue, { signal });
  const resize = new ResizeObserver(queue);
  resize.observe(frame);
  const stylesheet = doc.createElement('style');
  stylesheet.textContent = '.sheet-grid thead th{position:relative;overflow:visible;padding-inline-end:24px;height:44px}.sheet-column-resize{position:absolute;inset-block:0;inset-inline-end:0;display:flex;align-items:center;justify-content:center;width:24px;min-height:44px;cursor:col-resize;touch-action:none;color:var(--document-viewer-ink);background:var(--document-viewer-surface);border-inline-start:1px solid var(--document-viewer-line)}.sheet-column-resize svg{width:16px;height:20px;background:transparent;pointer-events:none}.sheet-column-resize:focus-visible{outline:3px solid var(--document-viewer-accent);outline-offset:-3px}body[data-resizing-columns],body[data-resizing-columns] *{cursor:col-resize!important;user-select:none!important}@media(forced-colors:active){.sheet-column-resize{color:ButtonText;background:Canvas;border-color:ButtonText}.sheet-column-resize:focus-visible{outline-color:Highlight}}';
  doc.head.append(stylesheet);
  return {
    preference, update, finish,
    handle(column, heading) {
      const handle = doc.createElement('span');
      handle.className = 'sheet-column-resize';
      handle.dataset.resizeColumn = String(column);
      handle.tabIndex = 0;
      handle.setAttribute('role', 'separator');
      handle.setAttribute('aria-orientation', 'vertical');
      handle.setAttribute('aria-valuemin', '1');
      handle.setAttribute('aria-label', labels.resizeColumn.replace('{column}', columnName(column)));
      handle.setAttribute('aria-controls', heading.id);
      handle.title = labels.resizeColumnsHint;
      handle.append(controlIcon(doc, 'grip'));
      return handle;
    },
    destroy() { finish(); resize.disconnect(); if (scheduled !== null) cancelAnimationFrame(scheduled); preferences.clear(); tools.remove(); stylesheet.remove(); }
  };
}
