import CFB from 'cfb';
import JSZip from 'jszip';
import { utils, write } from 'xlsx';

function record(type, payload, instance = 0, container = false) {
  const data = Buffer.alloc(8);
  data.writeUInt16LE((instance << 4) | (container ? 15 : 0), 0);
  data.writeUInt16LE(type, 2);
  data.writeUInt32LE(payload.length, 4);
  return Buffer.concat([data, payload]);
}

function integers(...values) {
  const result = Buffer.alloc(values.length * 4);
  values.forEach((value, index) => result.writeUInt32LE(value >>> 0, index * 4));
  return result;
}

export function legacyPpt({ cycle = false, brokenReference = false, encrypted = false, emptyTextBefore = false, largeAnchor = false, childAnchor = false, thinAnchor = false, inlineText = false, missingAnchor = false, outlineIndex, fitFlags = 0, styled = false, brokenStyle = false, masterStyle = false } = {}) {
  const text = value => {
    const style = Buffer.alloc(32);
    style.writeUInt32LE(value.length + 1, 0);
    style.writeUInt32LE(2048, 6);
    style.writeUInt16LE(1, 10);
    style.writeUInt32LE(value.length + 1, 12);
    style.writeUInt32LE(65536 | 131072 | 262144 | 7, 16);
    style.writeUInt16LE(7, 20);
    style.writeUInt16LE(0, 22);
    style.writeUInt16LE(18, 24);
    style.writeUInt32LE(0xfe0000ff, 26);
    return Buffer.concat([record(3999, integers(0)), record(4000, Buffer.from(value, 'utf16le')), ...(styled ? [record(4001, brokenStyle ? style.subarray(0, 16) : style.subarray(0, 30))] : [])]);
  };
  const shapeProperties = record(0xf00a, integers(100, 0x800), 1);
  const anchor = Buffer.alloc(largeAnchor || childAnchor ? 16 : 8);
  const coordinates = childAnchor ? [256, 128, 4000, 800] : [128, 256, 4000, thinAnchor ? 129 : 800];
  coordinates.forEach((value, index) => anchor.length === 16 ? anchor.writeInt32LE(value, index * 4) : anchor.writeInt16LE(value, index * 2));
  const option = Buffer.alloc(6);
  option.writeUInt16LE(191);
  option.writeUInt32LE(fitFlags, 2);
  const shape = record(0xf004, Buffer.concat([shapeProperties, record(0xf00b, option, 1), ...(missingAnchor ? [] : [record(childAnchor ? 0xf00f : 0xf010, anchor)]), record(0xf00d, inlineText ? text('Local textbox 中文') : record(3998, integers(outlineIndex ?? (emptyTextBefore ? 1 : 0))), 0, true)]), 0, true);
  const slideAtom = Buffer.alloc(24);
  slideAtom.writeUInt32LE(17, 12);
  slideAtom.writeUInt16LE(2, 20);
  const first = record(1006, Buffer.concat([...(masterStyle ? [record(1007, slideAtom)] : []), record(1036, record(0xf002, shape, 0, true), 0, true)]), 0, true);
  const second = record(1006, Buffer.alloc(0), 0, true);
  const list = record(4080, Buffer.concat([
    record(1011, integers(3, 4, emptyTextBefore ? 2 : 1, 300, 0)), ...(emptyTextBefore ? [record(3999, integers(4))] : []), text('First slide: 中文'),
    record(1011, integers(brokenReference ? 99 : 2, 4, 1, 200, 0)), text('Second slide')
  ]), 0, true);
  const font = Buffer.alloc(68);
  font.write('Georgia', 'utf16le');
  const masterText = Buffer.alloc(14);
  masterText.writeUInt16LE(1, 0);
  masterText.writeUInt32LE(2048, 2);
  masterText.writeUInt16LE(2, 6);
  masterText.writeUInt32LE(131072, 8);
  masterText.writeUInt16LE(30, 12);
  const master = masterStyle ? record(1016, record(4003, masterText), 0, true) : Buffer.alloc(0);
  const document = record(1000, Buffer.concat([record(1001, integers(5760, 4320)), ...(styled ? [record(4023, font)] : []), list, ...(masterStyle ? [record(4080, record(1011, integers(4, 0, 0, 17)), 1, true)] : [])]), 0, true);
  const masterOffset = document.length + first.length + second.length;
  const directoryOffset = masterOffset + master.length;
  const directory = record(6002, Buffer.concat([integers(((masterStyle ? 4 : 3) << 20) | 1), integers(0, document.length + first.length, document.length, ...(masterStyle ? [masterOffset] : []))]));
  const editOffset = directoryOffset + directory.length;
  const edit = record(4085, integers(300, 0x03000000, cycle ? editOffset : 0, directoryOffset, 1, 4, 0));
  const current = record(4086, integers(20, encrypted ? 0xf3d1c4df : 0xe391c05f, editOffset, 0, 0));
  const file = CFB.utils.cfb_new();
  CFB.utils.cfb_add(file, 'PowerPoint Document', Buffer.concat([document, first, second, master, directory, edit]));
  CFB.utils.cfb_add(file, 'Current User', current);
  return Buffer.from(CFB.write(file, { type: 'buffer' }));
}

export function pdfFixture(padding = 0, mixed = false) {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 400] /Resources << /Font << /F1 7 0 R >> >> /Contents 4 0 R >>',
    '<< /Length 55 >>\nstream\nBT /F1 24 Tf 50 300 Td (Document preview page one) Tj ET\nendstream',
    '<< /Type /Page /Parent 2 0 R /MediaBox ' + (mixed ? '[0 0 600 800] /CropBox [50 100 550 700] /Rotate 90' : '[0 0 600 400]') + ' /Resources << /Font << /F1 7 0 R >> >> /Contents 6 0 R >>',
    '<< /Length 55 >>\nstream\nBT /F1 24 Tf 50 300 Td (Document preview page two) Tj ET\nendstream',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
  ];
  if (padding) objects.push('<< /Length ' + padding + ' >>\nstream\n' + 'x'.repeat(padding) + '\nendstream');
  let data = '%PDF-1.7\n';
  const offsets = [0];
  for (const [index, object] of objects.entries()) {
    offsets.push(Buffer.byteLength(data));
    data += `${index + 1} 0 obj\n${object}\nendobj\n`;
  }
  const start = Buffer.byteLength(data);
  data += 'xref\n0 ' + (objects.length + 1) + '\n0000000000 65535 f \n' + offsets.slice(1).map(offset => String(offset).padStart(10, '0') + ' 00000 n \n').join('');
  data += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${start}\n%%EOF\n`;
  return Buffer.from(data);
}

export async function wordFixture(count = 1) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const paragraphs = Array.from({ length: count }, (_, index) => '<w:p><w:r>' + (index ? '<w:br w:type="page"/>' : '') + '<w:t>' + (index ? 'Word page ' + (index + 1) : 'Word preview 中文') + '</w:t></w:r></w:p>').join('');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + paragraphs + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

export async function wordLayoutFixture() {
  const zip = await JSZip.loadAsync(await wordFixture());
  const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const relations = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ['header', 'footer', 'footnotes'].map(name => `<Relationship Id="${name}" Type="${relations}/${name}" Target="${name}.xml"/>`).join('') + '</Relationships>');
  zip.file('word/header.xml', `<w:hdr xmlns:w="${word}"><w:p><w:r><w:t>Authored header</w:t></w:r></w:p></w:hdr>`);
  zip.file('word/footer.xml', `<w:ftr xmlns:w="${word}"><w:p><w:r><w:t>Authored footer</w:t></w:r></w:p></w:ftr>`);
  zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${word}"><w:footnote w:id="1"><w:p><w:r><w:t>Authored footnote</w:t></w:r></w:p></w:footnote></w:footnotes>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${word}" xmlns:r="${relations}"><w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="800020"/><w:sz w:val="36"/></w:rPr><w:t>Authored first page</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p><w:r><w:lastRenderedPageBreak/><w:t>Authored second page</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

export async function wordBulletFixture() {
  const zip = await JSZip.loadAsync(await wordFixture());
  const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const relation = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="${relation}" Target="numbering.xml"/></Relationships>`);
  zip.file('word/numbering.xml', `<w:numbering xmlns:w="${word}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\uf0b7"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\uf0b7"/><w:rPr><w:rFonts w:ascii="Unrelated" w:hAnsi="Unrelated"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${word}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Readable Symbol bullet</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Preserved unrelated symbol</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

export function sheetFixture(format = 'xlsx') {
  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, utils.aoa_to_sheet([['First sheet', '<img src=https://blocked.example/x onerror=alert(1)>'], [42, '中文']]), 'Summary');
  utils.book_append_sheet(workbook, utils.aoa_to_sheet([['Second sheet']]), 'Details');
  return write(workbook, { type: 'buffer', bookType: format });
}

export async function styledSheetFixture() {
  const zip = await JSZip.loadAsync(sheetFixture());
  zip.file('xl/theme/theme1.xml', '<a:theme><a:themeElements><a:clrScheme><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk1><a:srgbClr val="000000"/></a:dk1><a:accent1><a:srgbClr val="204060"/></a:accent1></a:clrScheme><a:fontScheme><a:minorFont><a:latin typeface="Georgia"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  zip.file('xl/styles.xml', '<styleSheet><fonts count="2"><font><name val="Arial"/><sz val="11"/></font><font><scheme val="minor"/><sz val="14"/><b/><i/><u val="double"/><color theme="4"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.5"/></patternFill></fill></fills><borders count="2"><border/><border><bottom style="double"><color indexed="2"/></bottom></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="2" fontId="1" fillId="1" borderId="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><dimension ref="A1:E5"/><sheetViews><sheetView showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="2" max="2" hidden="1"/><col min="3" max="3" width="24" style="1"/></cols><sheetData><row r="1" ht="32"><c r="A1" t="inlineStr"><is><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Rich</t></r><r><t> text</t></r></is></c><c r="B1" t="inlineStr"><is><t>Hidden column</t></is></c><c r="C1"><v>4.5</v></c><c r="D1"><v>10</v></c><c r="E1"><f t="shared" si="0" ref="E1:E2">D1*2</f></c></row><row r="2" hidden="1"><c r="D2"><v>20</v></c><c r="E2"><f t="shared" si="0"/></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Merged</t></is></c></row><row r="4" s="1" customFormat="1"><c r="C4" t="inlineStr"><is><t>Inherited row style</t></is></c><c r="D4" s="0" t="inlineStr"><is><t>Explicit normal</t></is></c></row><row r="5"><c r="A5"><f>SUM(D1:D2)</f></c></row></sheetData><mergeCells><mergeCell ref="A3:C3"/></mergeCells></worksheet>');
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function slidesFixture() {
  const zip = new JSZip();
  const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${office}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="8000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>PowerPoint preview 中文</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  return zip.generateAsync({ type: 'nodebuffer' });
}
