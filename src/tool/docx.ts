// Word 模板填充：基于 docxtemplater + pizzip（均为 MIT / GPL-3.0 双许可，本工具按 MIT 使用）。
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { tagKey } from './core';

interface DxtError {
  properties?: { id?: string; xtag?: string; explanation?: string; errors?: DxtError[]; file?: string };
  message?: string;
}

/** 把 docxtemplater 的英文报错翻译成能照着改的中文 */
export function explainTemplateError(e: unknown): string {
  const err = e as DxtError;
  const list = err.properties?.errors?.length ? err.properties.errors : [err];
  const msgs = list.slice(0, 5).map((x) => {
    const id = x.properties?.id;
    const tag = (x.properties?.xtag ?? '').slice(0, 20);
    switch (id) {
      case 'unclosed_tag': return `「{${tag}」少了右括号 }`;
      case 'unopened_tag': return `「${tag}}」少了左括号 {`;
      case 'duplicate_open_tag': return `「${tag}」里有两个连续的左括号 {{，只留一个`;
      case 'duplicate_close_tag': return `「${tag}」里有两个连续的右括号 }}，只留一个`;
      case 'unclosed_loop': case 'unopened_loop': case 'unbalanced_loop_tags': case 'closing_tag_does_not_match_opening_tag':
        return `占位符不能以 # / ^ 开头（这些是循环写法），请改成普通的 {列名}`;
      case 'raw_xml_tag_should_be_only_text_in_paragraph':
        return `占位符不能以 @ 开头，请改成普通的 {列名}`;
      case 'file_has_invalid_xml': case 'malformed_xml': case 'filetype_not_identified':
        return '模板文件损坏，请用 Word 或 WPS 打开后重新「另存为 .docx」';
      default: return x.properties?.explanation ?? x.message ?? String(e);
    }
  });
  const more = list.length > 5 ? `等 ${list.length} 处` : '';
  return `模板里的占位符写法有问题：${msgs.join('；')}${more}。请在 Word 里改好模板后重新拖进来`;
}

function open(bytes: Uint8Array) {
  let zip: PizZip;
  try {
    zip = new PizZip(bytes);
  } catch {
    throw new Error('这不是有效的 .docx 文件。如果是旧版 .doc，请用 Word 或 WPS 另存为 .docx');
  }
  if (!zip.file('word/document.xml')) throw new Error('这不是 Word 文档（.docx），请检查文件');
  try {
    return new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true, // 单元格里的换行在 Word 里也换行
      nullGetter: () => '',
      parser: (tag: string) => ({ get: (scope: Record<string, unknown>) => (tag === '.' ? scope : scope[tagKey(tag)]) }),
    });
  } catch (e) {
    throw new Error(explainTemplateError(e));
  }
}

/** 检查模板，返回里面的占位符名（含页眉页脚，去重，按出现顺序） */
export function inspectTemplate(bytes: Uint8Array): { tags: string[]; text: string } {
  const doc = open(bytes);
  const parts = Object.keys(doc.getZip().files).filter((f) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(f)).sort((a, b) => Number(b.includes('document')) - Number(a.includes('document')));
  const text = parts.map((f) => doc.getFullText(f)).join('\n');
  const tags = [...text.matchAll(/\{([^{}]+)\}/g)].map((m) => tagKey(m[1])).filter(Boolean);
  return { tags: [...new Set(tags)], text };
}

/** 用一行数据填模板，返回新的 .docx */
export function fillTemplate(bytes: Uint8Array, data: Record<string, string>): Uint8Array {
  const doc = open(bytes);
  try {
    doc.render(data);
  } catch (e) {
    throw new Error(explainTemplateError(e));
  }
  return doc.getZip().generate({ type: 'uint8array', compression: 'DEFLATE' });
}
