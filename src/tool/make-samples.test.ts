// 生成 samples/ 里的示例文件（只在设置了 MAKE_SAMPLES=1 时运行）
import { it } from 'vitest';
import { mkdirSync, writeFileSync } from 'node:fs';
import { writeXlsx } from './xlsx';
import { makeDocx } from './test-utils';

it.runIf(process.env.MAKE_SAMPLES === '1')('生成示例文件', async () => {
  mkdirSync('samples', { recursive: true });
  writeFileSync('samples/录用通知模板.docx', makeDocx([
    '录用通知书',
    '',
    '尊敬的 {姓名}：',
    '感谢您对本公司的信任。经过面试评估，我们很高兴地通知您：您已被录用为{部门}{岗位}，月薪 {月薪} 元（税前）。',
    '请于 {入职日期} 上午 9:00 携带身份证原件到公司前台报到。',
    '',
    '示例商贸有限公司 人力资源部',
  ]));
  const d = (m: number, day: number) => new Date(Date.UTC(2026, m - 1, day));
  writeFileSync('samples/录用名单.xlsx', await writeXlsx([{
    name: '十一月入职',
    header: ['姓名', '部门', '岗位', '月薪', '入职日期'],
    rows: [
      ['张晓明', '市场部', '活动策划', 8000, d(11, 2)],
      ['李佳怡', '技术部', '前端工程师', 15000, d(11, 2)],
      ['王磊', '财务部', '出纳', 6500, d(11, 9)],
      ['赵敏', '客服部', '客服专员', 5500, d(11, 9)],
      ['陈思远', '技术部', '测试工程师', 11000, d(11, 16)],
    ],
  }]));
});
