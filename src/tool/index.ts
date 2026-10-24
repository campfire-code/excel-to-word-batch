// 工具声明：Excel 批量套 Word 模板
// 用法：把 Word 模板（.docx）和表格（.xlsx / .csv）一起拖进来。逐个文件只做检查，
// 全部检查完后由 afterAll 用模板给表格的每一行生成一份 Word。
import { FileText as FileTextAsIcon } from 'lucide-react';
import type { BatchTool, InputItem, ItemResult } from '../kit/types';
import { baseName, fileNameFor, missingTags, patternProblems, rowData, safeName, suspiciousBraces, toTable, uniqueNamer, type Table } from './core';
import { fillTemplate, inspectTemplate } from './docx';
import { readSheets } from './xlsx';
import { Empty } from './Empty';
import { ResultBody } from './Result';

type Data = { kind: 'template'; tags: string[] } | { kind: 'table'; table: Table };

const isTemplate = (name: string) => /\.docx$/i.test(name);
const list = (xs: string[], n = 8) => xs.slice(0, n).map((x) => `{${x}}`).join('') + (xs.length > n ? ` 等 ${xs.length} 个` : '');

const tool: BatchTool = {
  id: 'o08',
  name: 'Excel 批量套 Word 模板',
  tagline: '把 Word 模板和 Excel 一起拖进来：模板里写 {姓名}、{金额} 这样的占位符，表格每一行生成一份 Word，用来批量做合同、证书、通知。',
  icon: FileTextAsIcon,
  category: '办公提效',
  type: 'data',
  mode: 'batch',
  resultView: 'list',
  concurrency: 1,
  Empty,
  ResultBody,
  accept: { label: 'Word 模板和表格', extensions: ['docx', 'xlsx', 'csv', 'doc', 'xls'] },
  params: [
    {
      key: 'nameRule', label: '文件名怎么起', type: 'text', default: '', placeholder: '例如 {姓名}-录用通知',
      help: '用花括号写列名，会换成每一行的值；不填就叫「第1份」「第2份」',
    },
    { key: 'headerRow', label: '表头在第几行', type: 'number', default: 1, min: 1, max: 50, help: '表格上方有标题行时改这里。表头的文字要和模板里花括号中的字完全一样' },
    { key: 'out', label: '保存到', type: 'folder', default: '' },
  ],
  check: (item) => {
    if (/\.doc$/i.test(item.name)) return '旧版 .doc 模板不支持，请用 Word 或 WPS 另存为 .docx 后再拖进来';
    if (/\.xls$/i.test(item.name)) return '旧版 .xls 表格不支持，请用 Excel 或 WPS 另存为 .xlsx 后再拖进来';
    return null;
  },

  async run(item, ctx) {
    const bytes = await ctx.readFile(item.path!);
    if (isTemplate(item.name)) {
      const { tags, text } = inspectTemplate(bytes);
      const bad = suspiciousBraces(text);
      if (bad.length) throw new Error(`模板里的 ${bad.slice(0, 3).join('、')} 用了全角括号，请改成英文半角的 { }，例如 {姓名}`);
      if (!tags.length) throw new Error('模板里没有找到占位符。请在 Word 里需要替换的位置写上 {列名}，例如 {姓名}，括号用英文半角');
      return { summary: `模板：找到 ${tags.length} 个占位符 ${list(tags)}`, data: { kind: 'template', tags } satisfies Data };
    }
    const sheets = await readSheets(bytes, item.name);
    const sheet = sheets.find((s) => s.grid.some((r) => r.some((c) => c != null && String(c).trim() !== '')));
    if (!sheet) throw new Error('表格是空的');
    const table = toTable(item.name, sheet.grid, Number(ctx.params.headerRow) || 1);
    if (!table.rows.length) throw new Error('表头下面没有数据行。如果表头不在第 1 行，请改「表头在第几行」');
    if (table.rows.length > 5000) throw new Error(`有 ${table.rows.length} 行，一次最多生成 5000 份，请把表格拆小一些`);
    const sheetNote = sheets.length > 1 ? `（只用第一个有数据的工作表「${sheet.name}」）` : '';
    return { summary: `表格：${table.rows.length} 行，列有 ${table.header.slice(0, 6).join('、')}${table.header.length > 6 ? '…' : ''}${sheetNote}`, data: { kind: 'table', table } satisfies Data };
  },

  async afterAll(results, ctx) {
    const of = (k: Data['kind']) => results.filter((r) => (r.result.data as Data | undefined)?.kind === k);
    const templates = of('template');
    const tables = of('table') as { item: InputItem; result: ItemResult & { data: Data & { kind: 'table' } } }[];
    if (!templates.length) throw new Error('没有可用的 Word 模板：请把写好占位符的模板（.docx）和表格一起拖进来');
    if (templates.length > 1) throw new Error(`拖进来了 ${templates.length} 个 Word 模板，一次只能用一个，请在列表里删掉多余的再开始`);
    if (!tables.length) throw new Error('没有可用的表格：请把 Excel（.xlsx）或 CSV 和模板一起拖进来');

    const tpl = templates[0];
    const tplBytes = await ctx.readFile(tpl.item.path!);
    const tags = (tpl.result.data as Data & { kind: 'template' }).tags;
    const rule = String(ctx.params.nameRule ?? '');
    const notes: string[] = [];
    let total = 0;
    let firstDir = '';
    for (const { data: { table } } of tables.map((t) => t.result)) {
      const missing = missingTags(tags, table.header);
      if (missing.length) notes.push(`「${table.name}」里没有 ${list(missing, 5)} 这几列，生成的文档里这些位置是空的`);
      const badRule = patternProblems(rule, table.header);
      if (badRule.length) notes.push(`文件名里的 ${list(badRule, 3)} 在「${table.name}」里找不到对应的列`);
      else if (rule.trim() && !rule.includes('{') && !table.header.includes(rule.trim())) notes.push(`「${rule.trim()}」不是表头，文件名按固定文字加序号`);
      const dir = await ctx.join(ctx.outputDir, safeName(baseName(table.name)));
      firstDir ||= dir;
      const name = uniqueNamer();
      for (let i = 0; i < table.rows.length; i++) {
        if (ctx.signal.aborted) return `已停止，生成了 ${total} 份`;
        const data = rowData(table.header, table.rows[i]);
        await ctx.writeFile(await ctx.join(dir, name(fileNameFor(rule, data, i), '.docx')), fillTemplate(tplBytes, data));
        total++;
      }
    }
    return `用模板「${tpl.item.name}」生成了 ${total} 份 Word，在 ${tables.length > 1 ? ctx.outputDir : firstDir}${notes.length ? '。注意：' + notes.join('；') : ''}`;
  },
};

export default tool;
