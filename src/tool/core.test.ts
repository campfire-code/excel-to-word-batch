import { describe, expect, it } from 'vitest';
import { cellText, fileNameFor, missingTags, parseCsv, patternProblems, rowData, suspiciousBraces, toTable, uniqueNamer } from './core';

describe('单元格转文字', () => {
  it('日期、小数、布尔、空值', () => {
    expect(cellText(new Date(Date.UTC(2026, 9, 3)))).toBe('2026-10-03');
    expect(cellText(new Date(Date.UTC(2026, 9, 3, 14, 5)))).toBe('2026-10-03 14:05');
    expect(cellText(0.1 + 0.2)).toBe('0.3');
    expect(cellText(8000)).toBe('8000');
    expect(cellText(true)).toBe('是');
    expect(cellText(null)).toBe('');
  });
  it('一行转成表头到值的对应', () => {
    expect(rowData(['姓名', '金额', '备注'], ['张三', 12.5])).toEqual({ 姓名: '张三', 金额: '12.5', 备注: '' });
  });
});

describe('文件名', () => {
  const data = { 姓名: '张三', 部门: '市场部' };
  it('按占位符规则', () => {
    expect(fileNameFor('{姓名}-{部门}-录用通知', data, 0)).toBe('张三-市场部-录用通知');
    expect(fileNameFor('{ 姓名 }', data, 0)).toBe('张三');
  });
  it('只写列名也行', () => expect(fileNameFor('姓名', data, 0)).toBe('张三'));
  it('不填或取到空值时用序号', () => {
    expect(fileNameFor('', data, 4)).toBe('第5份');
    expect(fileNameFor('{工号}', data, 1)).toBe('第2份');
  });
  it('不是列名就当固定文字加序号', () => expect(fileNameFor('通知', data, 2)).toBe('通知-3'));
  it('非法字符替换，重名加序号', () => {
    expect(fileNameFor('{姓名}', { 姓名: 'A/B' }, 0)).toBe('A_B');
    const n = uniqueNamer();
    expect([n('张三', '.docx'), n('张三', '.docx')]).toEqual(['张三.docx', '张三(2).docx']);
  });
  it('规则里引用了不存在的列', () => expect(patternProblems('{姓名}-{工号}', ['姓名'])).toEqual(['工号']));
});

describe('核对', () => {
  it('模板里有、表格里没有的占位符', () => expect(missingTags(['姓名', '金额', '日期'], ['姓名', ' 金额 '])).toEqual(['日期']));
  it('找出全角括号', () => {
    expect(suspiciousBraces('尊敬的｛姓名｝：您好 {金额}')).toEqual(['｛姓名｝']);
    expect(suspiciousBraces('{姓名}')).toEqual([]);
  });
  it('表头行可调，空行去掉', () => {
    const t = toTable('a', [['录用名单'], ['姓名', '金额'], ['张三', 1], [null, ''], ['李四', 2]], 2);
    expect(t.header).toEqual(['姓名', '金额']);
    expect(t.rows).toHaveLength(2);
  });
  it('CSV 引号与换行', () => expect(parseCsv('姓名,备注\n张三,"第一行\n第二行"\n')).toEqual([['姓名', '备注'], ['张三', '第一行\n第二行']]));
});
