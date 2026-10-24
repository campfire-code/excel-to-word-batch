// 空状态：画出「一份模板 + 一张名单 → 每人一份 Word」，占位符、表头和文件名都是真实感的中文
import { Upload } from 'lucide-react';
import './tool.css';

export function Empty({ hover }: { hover: boolean }) {
  return (
    <div className="xw-empty">
      <svg viewBox="0 0 520 210" role="img" aria-label="一份 Word 模板加一张名单，生成每人一份的 Word 文件">
        {/* 模板 */}
        <g>
          <rect x="10" y="14" width="112" height="146" rx="5" className="xw-paper" />
          <rect x="24" y="28" width="52" height="7" rx="3" className="xw-ink" />
          <rect x="24" y="46" width="30" height="6" rx="3" className="xw-tag" />
          <text x="28" y="51" className="xw-tagt">{'{姓名}'}</text>
          <rect x="24" y="64" width="84" height="4" rx="2" className="xw-ln" />
          <rect x="24" y="76" width="34" height="6" rx="3" className="xw-tag" />
          <text x="28" y="81" className="xw-tagt">{'{部门}'}</text>
          <rect x="62" y="77" width="46" height="4" rx="2" className="xw-ln" />
          <rect x="24" y="94" width="84" height="4" rx="2" className="xw-ln" />
          <rect x="24" y="106" width="34" height="6" rx="3" className="xw-tag" />
          <text x="28" y="111" className="xw-tagt">{'{月薪}'}</text>
          <rect x="24" y="124" width="60" height="4" rx="2" className="xw-ln" />
          <text x="66" y="182" textAnchor="middle" className="xw-cap">Word 模板</text>
        </g>
        <text x="136" y="92" textAnchor="middle" className="xw-plus">+</text>
        {/* 名单 */}
        <g>
          <rect x="150" y="30" width="136" height="114" rx="5" className="xw-sheet" />
          <rect x="150" y="30" width="136" height="20" rx="5" className="xw-head" />
          <text x="162" y="44" className="xw-ht">姓名</text>
          <text x="204" y="44" className="xw-ht">部门</text>
          <text x="248" y="44" className="xw-ht">月薪</text>
          {[['张晓明', '市场部', '8000'], ['李佳怡', '技术部', '15000'], ['王磊', '财务部', '6500'], ['赵敏', '客服部', '5500']].map(([a, b, c], i) => (
            <g key={a}>
              <line x1="150" x2="286" y1={50 + (i + 1) * 23} y2={50 + (i + 1) * 23} className="xw-rule" />
              <text x="162" y={66 + i * 23} className="xw-ct">{a}</text>
              <text x="204" y={66 + i * 23} className="xw-ct">{b}</text>
              <text x="248" y={66 + i * 23} className="xw-ct">{c}</text>
            </g>
          ))}
          <text x="218" y="182" textAnchor="middle" className="xw-cap">Excel 名单，一行一个人</text>
        </g>
        {/* 箭头 */}
        <path d="M296 87 H334 M326 79 L336 87 L326 95" className="xw-arrow" />
        {/* 结果 */}
        <g>
          {[['张晓明', 0], ['李佳怡', 1], ['王磊', 2]].map(([n, i]) => (
            <g key={n as string}>
              <rect x="346" y={22 + Number(i) * 44} width="150" height="34" rx="5" className={Number(i) === 0 ? 'xw-doc hot' : 'xw-doc'} />
              <rect x="356" y={31 + Number(i) * 44} width="14" height="16" rx="2" className="xw-w" />
              <text x="363" y={43 + Number(i) * 44} textAnchor="middle" className="xw-wt">W</text>
              <text x="378" y={43 + Number(i) * 44} className="xw-ft">{n}-录用通知.docx</text>
            </g>
          ))}
          <text x="421" y="182" textAnchor="middle" className="xw-cap">每人一份，内容已填好</text>
        </g>
      </svg>
      <h2 className="xw-title">{hover ? '松开鼠标，开始添加' : '把模板和名单一起拖进来'}</h2>
      <p className="xw-act"><Upload size={15} />一个 Word 模板（docx）加一张 Excel 或 CSV 名单，模板里用 {'{列名}'} 标出要替换的地方</p>
    </div>
  );
}
