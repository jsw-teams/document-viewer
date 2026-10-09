export function columnLayout(metadata, preference, available = 0) {
  const hidden = new Set(metadata.hiddenColumns);
  const widths = Array.from({ length: metadata.columns }, (_, index) => {
    if (hidden.has(index)) return 0;
    const source = metadata.columnWidths?.[index] ?? metadata.columnWidth ?? 64;
    const override = preference.widths.get(index);
    return Number.isFinite(override) ? Math.max(24, Math.min(2400, override)) : source;
  });
  const total = widths.reduce((sum, width) => sum + width, 0);
  const scale = preference.stretch && total > 0 ? Math.max(1, available / total) : 1;
  const offsets = [0];
  for (const width of widths) offsets.push(offsets.at(-1) + width * scale);
  return offsets;
}
