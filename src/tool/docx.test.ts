import { describe, expect, it } from 'vitest';
import { fillTemplate, inspectTemplate } from './docx';
import { docxText, makeDocx } from './test-utils';

describe('Word 模板', () => {
  it('占位符被 Word 拆成多个片段也能识别和替换，页眉里的也算', () => {
    const tpl = makeDocx([['尊敬的 {', '姓', '名}：'], '您的月薪为 {金额} 元。'], { header: '编号：{编号}' });
    expect(inspectTemplate(tpl).tags.sort()).toEqual(['编号', '姓名', '金额'].sort());
    const out = fillTemplate(tpl, { 姓名: '张三', 金额: '8000', 编号: 'HR-001' });
    expect(docxText(out)).toBe('尊敬的 张三：\n您的月薪为 8000 元。');
    expect(docxText(out, 'word/header1.xml')).toBe('编号：HR-001');
  });
  it('缺少的值留空，换行保留，特殊字符不破坏文档', () => {
    const out = fillTemplate(makeDocx(['{姓名}|{地址}|{备注}']), { 姓名: 'A&B <公司>', 地址: '第一行\n第二行' });
    expect(docxText(out)).toBe('A&B <公司>|第一行⏎第二行|');
  });
  it('括号不配对时用中文说明哪里错了', () => {
    expect(() => inspectTemplate(makeDocx(['尊敬的 {姓名：您好'])) ).toThrow('少了右括号');
    expect(() => inspectTemplate(makeDocx(['尊敬的 姓名}：您好']))).toThrow('少了左括号');
  });
  it('不是 docx 时提示另存', () => {
    expect(() => inspectTemplate(new TextEncoder().encode('not a zip'))).toThrow('另存为 .docx');
  });
});
