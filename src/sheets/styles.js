import { elements, textRuns } from './xml.js';

const indexed = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
const themeNames = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const truth = value => value === undefined || !['0', 'false', 'off'].includes(value);
const child = (xml, name) => elements(xml, name)[0];
const rgb = value => /^(?:[a-f\d]{2})?[a-f\d]{6}$/i.test(value || '') ? '#' + value.slice(-6).toLowerCase() : undefined;

export function spreadsheetTheme(xml = '') {
  const scheme = child(xml, 'clrScheme')?.content || '';
  const fonts = child(xml, 'fontScheme')?.content || '';
  return {
    colors: themeNames.map(name => {
      const content = child(scheme, name)?.content || '';
      return rgb(child(content, 'srgbClr')?.attributes.val || child(content, 'sysClr')?.attributes.lastClr);
    }),
    major: child(child(fonts, 'majorFont')?.content || '', 'latin')?.attributes.typeface,
    minor: child(child(fonts, 'minorFont')?.content || '', 'latin')?.attributes.typeface
  };
}

function tintColor(hex, tint) {
  if (!tint) return hex;
  const channels = [1, 3, 5].map(offset => parseInt(hex.slice(offset, offset + 2), 16) / 255);
  const maximum = Math.max(...channels);
  const minimum = Math.min(...channels);
  const delta = maximum - minimum;
  const lightness = (maximum + minimum) / 2;
  const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  let hue = 0;
  if (delta) {
    if (maximum === channels[0]) hue = ((channels[1] - channels[2]) / delta + 6) % 6;
    else if (maximum === channels[1]) hue = (channels[2] - channels[0]) / delta + 2;
    else hue = (channels[0] - channels[1]) / delta + 4;
  }
  const luminance = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const chroma = (1 - Math.abs(2 * luminance - 1)) * saturation;
  const secondary = chroma * (1 - Math.abs(hue % 2 - 1));
  const base = luminance - chroma / 2;
  const parts = [[chroma, secondary, 0], [secondary, chroma, 0], [0, chroma, secondary], [0, secondary, chroma], [secondary, 0, chroma], [chroma, 0, secondary]][Math.floor(hue)];
  return '#' + parts.map(value => Math.round((value + base) * 255).toString(16).padStart(2, '0')).join('');
}

export function spreadsheetColor(attributes = {}, theme = {}) {
  const hex = rgb(attributes.rgb) || theme.colors?.[Number(attributes.theme)] || rgb(indexed[Number(attributes.indexed)]);
  const tint = Number(attributes.tint || 0);
  return hex && Number.isFinite(tint) && tint >= -1 && tint <= 1 ? tintColor(hex, tint) : undefined;
}

export function fontStyle(xml, theme = {}) {
  const flag = name => { const node = child(xml, name); return node ? truth(node.attributes.val) : undefined; };
  const scheme = child(xml, 'scheme')?.attributes.val;
  const underline = child(xml, 'u');
  return {
    bold: flag('b'), italic: flag('i'), strike: flag('strike'),
    underline: underline ? (truth(underline.attributes.val) && underline.attributes.val !== 'none' ? underline.attributes.val || 'single' : false) : undefined,
    color: spreadsheetColor(child(xml, 'color')?.attributes, theme),
    font: (child(xml, 'name')?.attributes.val || child(xml, 'rFont')?.attributes.val || theme[scheme])?.slice(0, 128),
    size: Number(child(xml, 'sz')?.attributes.val) || undefined,
    script: child(xml, 'vertAlign')?.attributes.val
  };
}

export function richText(xml, theme) {
  const runs = elements(xml.replace(/<(?:[\w.-]+:)?rPh\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?rPh\s*>/g, ''), 'r');
  return runs.length ? runs.map(run => ({ text: textRuns(run.content), style: fontStyle(child(run.content, 'rPr')?.content || '', theme) })) : undefined;
}

export function workbookStyles(xml, theme, formats) {
  const fonts = elements(child(xml, 'fonts')?.content || '', 'font').map(entry => fontStyle(entry.content, theme));
  const fills = elements(child(xml, 'fills')?.content || '', 'fill').map(entry => {
    const pattern = child(entry.content, 'patternFill');
    return pattern?.attributes.patternType === 'solid' ? spreadsheetColor(child(pattern.content, 'fgColor')?.attributes, theme) : undefined;
  });
  const borders = elements(child(xml, 'borders')?.content || '', 'border').map(entry => Object.fromEntries(['left', 'right', 'top', 'bottom'].map(side => {
    const edge = child(entry.content, side);
    return [side, edge?.attributes.style ? { style: edge.attributes.style, color: spreadsheetColor(child(edge.content, 'color')?.attributes, theme) || '#000000' } : undefined];
  })));
  const bases = elements(child(xml, 'cellStyleXfs')?.content || '', 'xf');
  const definitions = new Map(elements(xml, 'numFmt').map(entry => [Number(entry.attributes.numFmtId), entry.attributes.formatCode]));
  return elements(child(xml, 'cellXfs')?.content || '', 'xf').map(entry => {
    const base = bases[Number(entry.attributes.xfId)];
    const properties = { ...base?.attributes, ...entry.attributes };
    const alignment = { ...child(base?.content || '', 'alignment')?.attributes, ...child(entry.content, 'alignment')?.attributes };
    return {
      ...fonts[Number(properties.fontId || 0)],
      numberFormat: definitions.get(Number(properties.numFmtId)) || formats[Number(properties.numFmtId || 0)],
      background: fills[Number(properties.fillId || 0)], borders: borders[Number(properties.borderId || 0)],
      align: ['left', 'right', 'center', 'justify'].includes(alignment.horizontal) ? alignment.horizontal : undefined,
      wrap: ['1', 'true'].includes(alignment.wrapText), shrink: ['1', 'true'].includes(alignment.shrinkToFit),
      vertical: { top: 'top', center: 'middle', bottom: 'bottom' }[alignment.vertical],
      indent: Math.min(250, Math.max(0, Number(alignment.indent) || 0)),
      rotation: Number(alignment.textRotation) || 0
    };
  });
}

export function applyCellStyle(node, style = {}) {
  const target = node.style;
  if (style.bold !== undefined) target.fontWeight = style.bold ? '700' : '400';
  if (style.italic !== undefined) target.fontStyle = style.italic ? 'italic' : 'normal';
  if (style.color) target.color = style.color;
  if (style.background) target.backgroundColor = style.background;
  if (style.align) target.textAlign = style.align;
  if (style.wrap) target.whiteSpace = 'pre-wrap';
  if (style.vertical) target.verticalAlign = style.vertical;
  if (style.font) target.fontFamily = style.font;
  if (style.size > 0 && style.size <= 400) target.fontSize = style.size + 'pt';
  if (style.underline !== undefined || style.strike !== undefined) target.textDecorationLine = [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';
  if (style.underline === 'double' || style.underline === 'doubleAccounting') target.textDecorationStyle = 'double';
  if (style.script === 'superscript' || style.script === 'subscript') { target.verticalAlign = style.script === 'superscript' ? 'super' : 'sub'; target.fontSize = 'smaller'; }
  if (style.indent) target.paddingInlineStart = style.indent * 3 + 'ch';
  const lines = { hair: '0.5px solid', thin: '1px solid', medium: '2px solid', thick: '3px solid', double: '3px double', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dashed', mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed' };
  for (const [side, edge] of Object.entries(style.borders || {})) if (edge && lines[edge.style]) target['border' + side[0].toUpperCase() + side.slice(1)] = lines[edge.style] + ' ' + edge.color;
}
