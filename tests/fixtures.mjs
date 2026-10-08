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

export function legacyPpt({ cycle = false, brokenReference = false, encrypted = false, emptyTextBefore = false, largeAnchor = false, childAnchor = false, thinAnchor = false, inlineText = false, missingAnchor = false, outlineIndex, fitFlags = 0 } = {}) {
  const text = value => Buffer.concat([record(3999, integers(0)), record(4000, Buffer.from(value, 'utf16le'))]);
  const shapeProperties = record(0xf00a, integers(100, 0x800), 1);
  const anchor = Buffer.alloc(largeAnchor || childAnchor ? 16 : 8);
  const coordinates = childAnchor ? [256, 128, 4000, 800] : [128, 256, 4000, thinAnchor ? 129 : 800];
  coordinates.forEach((value, index) => anchor.length === 16 ? anchor.writeInt32LE(value, index * 4) : anchor.writeInt16LE(value, index * 2));
  const option = Buffer.alloc(6);
  option.writeUInt16LE(191);
  option.writeUInt32LE(fitFlags, 2);
  const shape = record(0xf004, Buffer.concat([shapeProperties, record(0xf00b, option, 1), ...(missingAnchor ? [] : [record(childAnchor ? 0xf00f : 0xf010, anchor)]), record(0xf00d, inlineText ? text('Local textbox 中文') : record(3998, integers(outlineIndex ?? (emptyTextBefore ? 1 : 0))), 0, true)]), 0, true);
  const first = record(1006, record(1036, record(0xf002, shape, 0, true), 0, true), 0, true);
  const second = record(1006, Buffer.alloc(0), 0, true);
  const list = record(4080, Buffer.concat([
    record(1011, integers(3, 4, emptyTextBefore ? 2 : 1, 300, 0)), ...(emptyTextBefore ? [record(3999, integers(4))] : []), text('First slide: 中文'),
    record(1011, integers(brokenReference ? 99 : 2, 4, 1, 200, 0)), text('Second slide')
  ]), 0, true);
  const document = record(1000, Buffer.concat([record(1001, integers(5760, 4320)), list]), 0, true);
  const directoryOffset = document.length + first.length + second.length;
  const directory = record(6002, Buffer.concat([integers((3 << 20) | 1), integers(0, document.length + first.length, document.length)]));
  const editOffset = directoryOffset + directory.length;
  const edit = record(4085, integers(300, 0x03000000, cycle ? editOffset : 0, directoryOffset, 1, 4, 0));
  const current = record(4086, integers(20, encrypted ? 0xf3d1c4df : 0xe391c05f, editOffset, 0, 0));
  const file = CFB.utils.cfb_new();
  CFB.utils.cfb_add(file, 'PowerPoint Document', Buffer.concat([document, first, second, directory, edit]));
  CFB.utils.cfb_add(file, 'Current User', current);
  return Buffer.from(CFB.write(file, { type: 'buffer' }));
}

export function pdfFixture() {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 400] /Resources << /Font << /F1 7 0 R >> >> /Contents 4 0 R >>',
    '<< /Length 55 >>\nstream\nBT /F1 24 Tf 50 300 Td (Document preview page one) Tj ET\nendstream',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 400] /Resources << /Font << /F1 7 0 R >> >> /Contents 6 0 R >>',
    '<< /Length 55 >>\nstream\nBT /F1 24 Tf 50 300 Td (Document preview page two) Tj ET\nendstream',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
  ];
  let data = '%PDF-1.7\n';
  const offsets = [0];
  for (const [index, object] of objects.entries()) {
    offsets.push(Buffer.byteLength(data));
    data += `${index + 1} 0 obj\n${object}\nendobj\n`;
  }
  const start = Buffer.byteLength(data);
  data += 'xref\n0 8\n0000000000 65535 f \n' + offsets.slice(1).map(offset => String(offset).padStart(10, '0') + ' 00000 n \n').join('');
  data += `trailer\n<< /Size 8 /Root 1 0 R >>\nstartxref\n${start}\n%%EOF\n`;
  return Buffer.from(data);
}

export async function wordFixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Word preview 中文</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

export function sheetFixture(format = 'xlsx') {
  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, utils.aoa_to_sheet([['First sheet', '<img src=https://blocked.example/x onerror=alert(1)>'], [42, '中文']]), 'Summary');
  utils.book_append_sheet(workbook, utils.aoa_to_sheet([['Second sheet']]), 'Details');
  return write(workbook, { type: 'buffer', bookType: format });
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
