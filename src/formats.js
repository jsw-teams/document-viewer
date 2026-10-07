const formats = new Set(['pdf', 'docx', 'pptx', 'xlsx', 'xls', 'doc', 'ppt']);

export function documentFormat(source, explicit) {
  if (explicit) return formats.has(String(explicit).toLowerCase()) ? String(explicit).toLowerCase() : null;
  try {
    const pathname = decodeURIComponent(new URL(source, 'https://document.invalid').pathname);
    const format = pathname.split('.').pop().toLowerCase();
    return formats.has(format) ? format : null;
  } catch { return null; }
}

export function documentUrl(source, base = 'https://document.invalid') {
  if (typeof source !== 'string' || !source.trim() || /[\x00-\x20\\]/.test(source) || source.startsWith('//')) {
    throw new Error('Document sources must be safe relative paths or HTTPS URLs');
  }
  const address = new URL(source, base);
  if (address.username || address.password || !['https:', 'http:'].includes(address.protocol)) {
    throw new Error('Document sources must be safe relative paths or HTTPS URLs');
  }
  if (address.protocol === 'http:' && address.origin !== new URL(base).origin) throw new Error('Remote documents require HTTPS');
  return address;
}
