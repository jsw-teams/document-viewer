function variable(styles, names) {
  for (const name of names) {
    const value = styles.getPropertyValue(name).trim();
    if (value) return value;
  }
  return '';
}

function background(element, fallback) {
  for (let ancestor = element; ancestor; ancestor = ancestor.parentElement) {
    const color = getComputedStyle(ancestor).backgroundColor;
    if (color !== 'transparent' && !/^rgba\((?:[^,]+,){3}\s*0(?:\.0+)?\s*\)$/.test(color) && !/\/\s*0(?:\.0+)?\s*\)$/.test(color)) return color;
  }
  return fallback;
}

export function followTheme(root, host) {
  const canvas = document.createElement('canvas');
  canvas.width = canvas.height = 1;
  const drawing = canvas.getContext('2d', { willReadFrequently: true });
  const luminance = (color, base) => {
    drawing.clearRect(0, 0, 1, 1);
    drawing.fillStyle = base;
    drawing.fillRect(0, 0, 1, 1);
    drawing.fillStyle = color;
    drawing.fillRect(0, 0, 1, 1);
    const pixels = drawing.getImageData(0, 0, 1, 1).data;
    const linear = [...pixels].slice(0, 3).map(value => value / 255 <= 0.04045 ? value / 255 / 12.92 : ((value / 255 + 0.055) / 1.055) ** 2.4);
    return linear[0] * 0.2126 + linear[1] * 0.7152 + linear[2] * 0.0722;
  };
  const contrast = (first, second, base) => (Math.max(luminance(first, base), luminance(second, base)) + 0.05) / (Math.min(luminance(first, base), luminance(second, base)) + 0.05);
  const media = matchMedia('(prefers-color-scheme: dark)');
  const synchronize = () => {
    const styles = getComputedStyle(host);
    const overrides = getComputedStyle(root);
    const dark = styles.colorScheme === 'dark' || (styles.colorScheme !== 'light' && media.matches);
    const choose = (name, names, fallback) => variable(overrides, ['--document-viewer-' + name]) || variable(styles, names) || fallback;
    const paper = choose('paper', ['--paper', '--background', '--bg', '--color-background'], background(host, dark ? '#16182a' : '#fbfaf7'));
    const ink = choose('ink', ['--ink', '--foreground', '--text', '--text-color', '--color-text'], styles.color);
    const link = [...document.querySelectorAll('a[href]')].find(element => !element.closest('.document-viewer') && getComputedStyle(element).display !== 'none');
    const accent = choose('accent', ['--accent', '--primary', '--color-primary', '--accent-color'], link ? getComputedStyle(link).color : ink);
    const surface = choose('surface', ['--surface', '--panel', '--card', '--color-surface'], paper);
    const colors = {
      paper, ink, accent, surface,
      muted: choose('muted', ['--muted', '--text-muted', '--color-muted'], 'color-mix(in srgb,' + ink + ' 78%,' + paper + ')'),
      line: choose('line', ['--border', '--line', '--border-color', '--color-border'], 'color-mix(in srgb,' + ink + ' 25%,' + paper + ')'),
      'on-accent': [ink, paper, '#fff', '#000'].sort((first, second) => contrast(second, accent, paper) - contrast(first, accent, paper))[0],
      focus: contrast(accent, surface, paper) >= 3 ? accent : ink
    };
    for (const [name, color] of Object.entries(colors)) {
      const property = '--dv-auto-' + name;
      if (root.style.getPropertyValue(property) !== color) root.style.setProperty(property, color);
    }
  };
  synchronize();
  const observer = new MutationObserver(synchronize);
  for (let ancestor = root; ancestor; ancestor = ancestor.parentElement) observer.observe(ancestor, { attributes: true });
  observer.observe(document.head, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['href', 'media', 'disabled'] });
  media.addEventListener('change', synchronize);
  const stylesheetLoaded = event => { if (event.target instanceof HTMLLinkElement && event.target.relList.contains('stylesheet')) synchronize(); };
  document.addEventListener('load', stylesheetLoaded, true);
  return () => { observer.disconnect(); media.removeEventListener('change', synchronize); document.removeEventListener('load', stylesheetLoaded, true); };
}
