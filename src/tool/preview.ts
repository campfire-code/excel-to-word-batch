// 截图与预览用的示例数据（不会进入正式运行）
import type { Row } from '../kit/runner';

const files: [string, number, Row['status'], string?][] = [
  ['录用通知模板.docx', 18_000, 'ok', '模板：找到 5 个占位符 {姓名}{部门}{岗位}{月薪}{入职日期}'],
  ['十一月录用名单.xlsx', 12_000, 'ok', '表格：36 行，列有 姓名、部门、岗位、月薪、入职日期'],
  ['十二月录用名单.xlsx', 9_000, 'run'],
  ['实习生名单.csv', 3_000, 'wait'],
  ['旧版合同模板.doc', 41_000, 'err', '旧版 .doc 模板不支持，请用 Word 或 WPS 另存为 .docx 后再拖进来'],
];

export const rows = (state: string): Row[] =>
  state === 'empty' ? [] : files.map(([name, size, status, text], i) => {
    const st = state === 'done' && status !== 'err' ? 'ok' : status;
    return {
      item: { id: `p${i}`, name, path: `/示例/${name}`, size },
      status: st,
      progress: st === 'run' ? 0.55 : st === 'ok' ? 1 : 0,
      error: st === 'err' ? text : undefined,
      result: st === 'ok' ? { summary: text ?? '表格：12 行，列有 姓名、部门、岗位、月薪、入职日期' } : undefined,
    };
  });

export const params = { nameRule: '{姓名}-录用通知' };
