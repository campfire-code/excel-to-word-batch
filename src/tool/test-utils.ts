// 测试与生成示例用：拼一个最小的 .docx，以及读出 .docx 正文文字（不进入正式运行）
import PizZip from 'pizzip';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** 每个段落是若干「文字片段」：Word 编辑时经常把一个词拆进多个片段，这里可以模拟 */
export function makeDocx(paragraphs: (string | string[])[], opts: { header?: string } = {}): Uint8Array {
  const para = (p: string | string[]) => `<w:p>${(Array.isArray(p) ? p : [p]).map((r) => `<w:r><w:t xml:space="preserve">${esc(r)}</w:t></w:r>`).join('')}</w:p>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${opts.header ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const sect = opts.header ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>' : '<w:sectPr/>';
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${paragraphs.map(para).join('')}${sect}</w:body></w:document>`);
  if (opts.header) {
    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`);
    zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${para(opts.header)}</w:hdr>`);
  }
  return zip.generate({ type: 'uint8array', compression: 'DEFLATE' });
}

/** 读出某个部件里的文字，段落之间用换行，<w:br/> 记为 ⏎ */
export function docxText(bytes: Uint8Array, part = 'word/document.xml'): string {
  const xml = new PizZip(bytes).file(part)!.asText();
  return xml
    .split(/<\/w:p>/)
    .map((p) => p.replace(/<w:br\/>/g, '⏎').replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, '$1').replace(/<[^>]+>/g, ''))
    .filter(Boolean)
    .join('\n')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
    .trim();
}
