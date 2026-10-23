// 纯函数：表格整理、单元格转文字、文件名生成、占位符核对，不依赖界面和本机能力。

export type Cell = string | number | boolean | Date | null;
export interface Table {
  /** 来源文件名 */
  name: string;
  header: string[];
  rows: Cell[][];
}

/** 文本解码：优先 UTF-8（去掉 BOM），不是合法 UTF-8 时按 GBK 解（Excel 中文版导出的 CSV 常见） */
export function decodeText(bytes: Uint8Array): string {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
  } catch {
    return new TextDecoder('gbk').decode(bytes);
  }
}

/** 解析 CSV：支持引号、引号内换行与逗号、两个引号转义，自动识别逗号 / 制表符 / 分号分隔 */
export function parseCsv(text: string): string[][] {
  const firstLine = text.slice(0, text.indexOf('\n') === -1 ? text.length : text.indexOf('\n'));
  const sep = [',', '\t', ';'].map((s) => [s, firstLine.split(s).length] as const).sort((a, b) => b[1] - a[1])[0][0];
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') {
        if (text[i + 1] === '"') { cell += '"'; i++; } else quoted = false;
      } else cell += c;
    } else if (c === '"' && cell === '') quoted = true;
    else if (c === sep) { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((c) => c.trim() !== ''));
}

/** 把二维数组的第 headerRow 行（从 1 开始）当表头，空表头补成「列1」「列2」，去掉整行为空的行 */
export function toTable(name: string, grid: Cell[][], headerRow: number): Table {
  const hi = Math.max(0, headerRow - 1);
  const raw = grid[hi] ?? [];
  const width = Math.max(raw.length, ...grid.slice(hi + 1).map((r) => r.length), 0);
  const header = Array.from({ length: width }, (_, i) => {
    const v = raw[i];
    return v == null || String(v).trim() === '' ? `列${i + 1}` : String(v).trim();
  });
  const rows = grid.slice(hi + 1).filter((r) => r.some((c) => c != null && String(c).trim() !== ''));
  return { name, header, rows };
}

/** 单元格转成填进 Word 的文字：日期写成 2026-10-23，小数去掉浮点误差，空格子是空字符串 */
export function cellText(v: Cell): string {
  if (v == null) return '';
  if (v instanceof Date) {
    const p = (n: number) => String(n).padStart(2, '0');
    const date = `${v.getUTCFullYear()}-${p(v.getUTCMonth() + 1)}-${p(v.getUTCDate())}`;
    const time = v.getUTCHours() || v.getUTCMinutes() ? ` ${p(v.getUTCHours())}:${p(v.getUTCMinutes())}` : '';
    return date + time;
  }
  if (typeof v === 'number') return String(Number(v.toPrecision(12)));
  if (typeof v === 'boolean') return v ? '是' : '否';
  return String(v);
}

/** 一行数据转成「表头 → 文字」 */
export function rowData(header: string[], row: Cell[]): Record<string, string> {
  return Object.fromEntries(header.map((h, i) => [h, cellText(row[i] ?? null)]));
}

/** 占位符里的名字：去掉首尾空格，便于和表头对上 */
export const tagKey = (tag: string) => tag.trim();

/** 模板里有、表格里没有的占位符 */
export function missingTags(tags: string[], header: string[]): string[] {
  const have = new Set(header.map((h) => h.trim()));
  return tags.filter((t) => !have.has(tagKey(t)));
}

/** 检查模板文字里有没有用错括号：全角｛｝、中文【】里像占位符的写法 */
export function suspiciousBraces(text: string): string[] {
  const found = new Set<string>();
  for (const m of text.matchAll(/[｛{]\s*[^｛｝{}\s]{1,20}\s*[｝}]/g)) if (/[｛｝]/.test(m[0])) found.add(m[0]);
  return [...found];
}

/** 文件名里不能出现的字符替换掉，并限制长度 */
export function safeName(s: string): string {
  const cleaned = s.replace(/[\\/:*?"<>|\r\n\t]/g, '_').replace(/^\.+/, '_').trim();
  return (cleaned || '未命名').slice(0, 80);
}

/**
 * 按「文件名怎么起」生成文件名（不含扩展名）：
 * - 含 {列名} 时按占位符替换，例如「{姓名}-录用通知」；
 * - 只写了一个列名（如「姓名」）时用这一列的值；写的不是列名就当固定文字，后面加序号；
 * - 不填，或替换后为空时用「第 N 份」。
 */
export function fileNameFor(pattern: string, data: Record<string, string>, index: number): string {
  const p = pattern.trim();
  let name = '';
  if (p.includes('{')) name = p.replace(/\{([^{}]*)\}/g, (_, k: string) => data[tagKey(k)] ?? '');
  else if (p && p in data) name = data[p];
  else if (p) name = `${p}-${index + 1}`;
  name = name.trim();
  return safeName(name && name.replace(/[-_\s]/g, '') ? name : `第${index + 1}份`);
}

/** 文件名规则里引用了哪些表格里没有的列 */
export function patternProblems(pattern: string, header: string[]): string[] {
  const p = pattern.trim();
  const have = new Set(header.map((h) => h.trim()));
  if (p.includes('{')) return [...p.matchAll(/\{([^{}]*)\}/g)].map((m) => tagKey(m[1])).filter((k) => !have.has(k));
  return [];
}

/** 同一批里文件名重复时加 (2)、(3) */
export function uniqueNamer() {
  const used = new Set<string>();
  return (name: string, ext: string) => {
    let n = `${name}${ext}`;
    for (let i = 2; used.has(n.toLowerCase()); i++) n = `${name}(${i})${ext}`;
    used.add(n.toLowerCase());
    return n;
  };
}

/** 去掉扩展名 */
export const baseName = (file: string) => file.replace(/\.[^.]+$/, '');

/** CSV 里的文本数字转成数值；以 0 开头的编号和超过 15 位的长数字保持文本 */
export function numeric(v: string): Cell {
  const t = v.trim();
  return /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(t) && t.replace(/[-.]/g, '').length <= 15 ? Number(t) : v;
}
