import { describe, expect, it } from 'vitest';
import tool from './index';
import { writeXlsx } from './xlsx';
import { docxText, makeDocx } from './test-utils';
import type { InputItem, ItemResult, Params, RunContext } from '../kit/types';

// 内存文件系统，模拟本机读写
function fakeCtx(files: Map<string, Uint8Array>, params: Params): RunContext {
  return {
    params, outputDir: '/out', signal: new AbortController().signal, progress: () => {},
    readFile: async (p) => files.get(p)!,
    readText: async (p) => new TextDecoder().decode(files.get(p)!),
    writeFile: async (p, d) => void files.set(p, d),
    writeText: async (p, t) => void files.set(p, new TextEncoder().encode(t)),
    mkdir: async () => {},
    join: async (...parts) => parts.join('/'),
    llm: async () => '', sidecar: async () => '', ffmpeg: async () => {}, duration: async () => 0,
  };
}
const defaults = (over: Params): Params => ({ ...Object.fromEntries(tool.params.map((f) => [f.key, f.default])), ...over });
const item = (name: string): InputItem => ({ id: name, name, path: '/in/' + name });

async function runAll(files: Map<string, Uint8Array>, names: string[], params: Params) {
  const ctx = fakeCtx(files, defaults(params));
  const results: { item: InputItem; result: ItemResult }[] = [];
  const errors: string[] = [];
  for (const n of names) {
    const bad = tool.check?.(item(n));
    if (bad) { errors.push(bad); continue; }
    try { results.push({ item: item(n), result: await tool.run(item(n), ctx) }); } catch (e) { errors.push((e as Error).message); }
  }
  return { results, errors, ctx };
}

const template = () => makeDocx(['录用通知', ['尊敬的 {姓', '名}：'], '欢迎加入{部门}，月薪 {月薪} 元，入职日期 {入职日期}。']);

describe('完整流程', () => {
  it('模板和表格一起拖进来，每行一份，按规则命名', async () => {
    const files = new Map<string, Uint8Array>([
      ['/in/录用通知模板.docx', template()],
      ['/in/名单.xlsx', await writeXlsx([{ name: 'S', header: ['姓名', '部门', '月薪', '入职日期'], rows: [['张三', '市场部', 8000, new Date(Date.UTC(2026, 10, 1))], ['李四', '技术部', 12000.5, new Date(Date.UTC(2026, 10, 3))], ['张三', '财务部', 7000, null]] }])],
    ]);
    // 表格在前、模板在后也行
    const { results, errors, ctx } = await runAll(files, ['名单.xlsx', '录用通知模板.docx'], { nameRule: '{姓名}-录用通知' });
    expect(errors).toEqual([]);
    expect(results[0].result.summary).toContain('表格：3 行');
    expect(results[1].result.summary).toBe('模板：找到 4 个占位符 {姓名}{部门}{月薪}{入职日期}');
    const msg = await tool.afterAll!(results, ctx);
    expect(msg).toBe('用模板「录用通知模板.docx」生成了 3 份 Word，在 /out/名单');
    expect(docxText(files.get('/out/名单/张三-录用通知.docx')!)).toBe('录用通知\n尊敬的 张三：\n欢迎加入市场部，月薪 8000 元，入职日期 2026-11-01。');
    expect(docxText(files.get('/out/名单/李四-录用通知.docx')!)).toContain('月薪 12000.5 元');
    expect(docxText(files.get('/out/名单/张三-录用通知(2).docx')!)).toContain('入职日期 。');
  });

  it('表格缺列时照样生成，并提醒哪些位置是空的', async () => {
    const files = new Map<string, Uint8Array>([
      ['/in/t.docx', template()],
      ['/in/名单.csv', new TextEncoder().encode('﻿姓名,部门\n王五,行政部\n')],
    ]);
    const { results, ctx } = await runAll(files, ['t.docx', '名单.csv'], {});
    const msg = await tool.afterAll!(results, ctx);
    expect(msg).toContain('生成了 1 份');
    expect(msg).toContain('「名单.csv」里没有 {月薪}{入职日期} 这几列');
    expect(files.has('/out/名单/第1份.docx')).toBe(true);
  });

  it('只拖了表格、拖了两个模板、只拖了模板，都给出怎么办', async () => {
    const files = new Map<string, Uint8Array>([
      ['/in/a.docx', template()], ['/in/b.docx', template()],
      ['/in/名单.csv', new TextEncoder().encode('姓名\n张三\n')],
    ]);
    let r = await runAll(files, ['名单.csv'], {});
    await expect(tool.afterAll!(r.results, r.ctx)).rejects.toThrow('没有可用的 Word 模板');
    r = await runAll(files, ['a.docx', 'b.docx', '名单.csv'], {});
    await expect(tool.afterAll!(r.results, r.ctx)).rejects.toThrow('拖进来了 2 个 Word 模板');
    r = await runAll(files, ['a.docx'], {});
    await expect(tool.afterAll!(r.results, r.ctx)).rejects.toThrow('没有可用的表格');
  });

  it('模板问题在列表里直接说明', async () => {
    const files = new Map<string, Uint8Array>([
      ['/in/全角.docx', makeDocx(['尊敬的｛姓名｝'])],
      ['/in/无占位符.docx', makeDocx(['普通文档'])],
    ]);
    const { errors } = await runAll(files, ['全角.docx', '无占位符.docx', '旧模板.doc', '旧表.xls'], {});
    expect(errors[0]).toContain('｛姓名｝ 用了全角括号');
    expect(errors[1]).toContain('没有找到占位符');
    expect(errors[2]).toContain('另存为 .docx');
    expect(errors[3]).toContain('另存为 .xlsx');
  });
});
