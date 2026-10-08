export function symbolBullet(text, font) {
  if (typeof text !== 'string' || typeof font !== 'string' || !/^["']?symbol["']?$/i.test(font.trim())) return text;
  return text.replaceAll('\uf0b7', '\u2022');
}
