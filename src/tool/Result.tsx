// 单个文件的检查结果：模板画成一张纸，占位符高亮；名单用大字写出将生成几份，并摆出前几行数据
import { ArrowRight, FileText, Table2 } from 'lucide-react';
import type { InputItem, ItemResult } from '../kit/types';
import { cellText, type Cell } from './core';
import './tool.css';

type Shown = { kind: 'template'; tags: string[] } | { kind: 'table'; table: { name: string; header: string[]; rows: Cell[][] } };

const MAX_COLS = 5;
const MAX_ROWS = 4;

export function ResultBody({ result }: { item: InputItem; result: ItemResult }) {
  const d = result.data as Shown | undefined;
  if (d?.kind === 'template') {
    const tags = d.tags.map((t) => t.trim());
    return (
      <div className="xw-body">
        <div className="xw-tpl">
          <div className="xw-page" aria-hidden="true">
            <i className="xw-bar w50" />
            <div className="xw-tags">{tags.slice(0, 8).map((t) => <b key={t}>{`{${t}}`}</b>)}</div>
            <i className="xw-bar" /><i className="xw-bar w80" /><i className="xw-bar w60" />
          </div>
          <div className="xw-side">
            <div className="xw-big"><b className="k-num">{tags.length}</b><span>处要替换</span></div>
            <p className="xw-sub"><FileText size={14} />Word 模板，字体、表格、页眉页脚照原样保留</p>
          </div>
        </div>
        <div className="xw-chips">{tags.map((t) => <span key={t}>{`{${t}}`}</span>)}</div>
      </div>
    );
  }
  if (d?.kind === 'table') {
    const { header, rows } = d.table;
    const cols = header.slice(0, MAX_COLS);
    const shown = rows.slice(0, MAX_ROWS);
    return (
      <div className="xw-body">
        <div className="xw-big"><b className="k-num">{rows.length}</b><span>份 Word 将生成</span></div>
        <p className="xw-sub"><Table2 size={14} />一行数据一份文档<ArrowRight size={13} />共 <span className="k-num">{header.length}</span> 列，表头要和模板里花括号中的字一样</p>
        <div className="xw-grid-wrap">
          <table className="xw-grid">
            <thead><tr>{cols.map((h) => <th key={h}>{h}</th>)}{header.length > MAX_COLS && <th className="xw-more">…</th>}</tr></thead>
            <tbody>
              {shown.map((r, i) => <tr key={i}>{cols.map((_, c) => <td key={c} title={cellText(r[c] ?? null)}>{cellText(r[c] ?? null)}</td>)}{header.length > MAX_COLS && <td className="xw-more" />}</tr>)}
            </tbody>
          </table>
          {rows.length > MAX_ROWS && <p className="xw-rest">还有 {rows.length - MAX_ROWS} 行</p>}
        </div>
      </div>
    );
  }
  return <p className="xw-sub">{result.summary}</p>;
}
