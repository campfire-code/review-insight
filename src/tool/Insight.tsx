// 差评归因的结果展示：把汇总文本画成「类别占比条 + 高频问题排行」，以及还没导入时的空状态示意。
import { FileSpreadsheet } from 'lucide-react';
import type { ItemResult } from '../kit/types';
import './tool.css';

const HUES = ['#D37017', '#C0392B', '#8E6BB8', '#2F8F83', '#3F7CC4', '#B8962E', '#7C8794'];

interface Parsed {
  head: string;
  cats: { name: string; n: number; pct: number }[];
  noIssue: number;
  top: { issue: string; n: number }[];
  notes: string[];
}

function parse(text: string): Parsed {
  const p: Parsed = { head: '', cats: [], noIssue: 0, top: [], notes: [] };
  text.split('\n').forEach((line, i) => {
    const cat = /^(.+?)：(\d+) 条（([\d.]+)%）$/.exec(line);
    const none = /^无具体问题：(\d+) 条$/.exec(line);
    const top = /^高频问题：(.+)$/.exec(line);
    if (i === 0) p.head = line;
    else if (cat) p.cats.push({ name: cat[1], n: Number(cat[2]), pct: Number(cat[3]) });
    else if (none) p.noIssue = Number(none[1]);
    else if (top) top[1].split('、').forEach((t) => { const m = /^(.+)（(\d+)）$/.exec(t); if (m) p.top.push({ issue: m[1], n: Number(m[2]) }); });
    else if (line.trim()) p.notes.push(line);
  });
  return p;
}

export function InsightBody({ result }: { result: ItemResult }) {
  const p = parse(result.text ?? '');
  const total = Number(/^(\d+) 条评论/.exec(p.head)?.[1] ?? 0);
  const analyzed = Number(/分析 (\d+) 条/.exec(p.head)?.[1] ?? 0);
  const calls = /调用大模型 (\d+) 次/.exec(p.head)?.[1];
  if (!p.cats.length) return <p className="ri-head">{p.head || result.summary}</p>;
  const max = Math.max(1, ...p.top.map((t) => t.n));
  return (
    <div className="ri-body">
      <div className="ri-nums">
        <div><b className="k-num">{analyzed}</b><span>条差评 / 中评</span></div>
        {total > 0 && <div><b className="k-num">{total}</b><span>评论总数</span></div>}
        {calls && <div><b className="k-num">{calls}</b><span>次调用大模型</span></div>}
      </div>
      <div className="ri-bar" role="img" aria-label={p.cats.map((c) => `${c.name} ${c.pct}%`).join('，')}>
        {p.cats.map((c, i) => <i key={c.name} style={{ flex: c.n, background: HUES[i % HUES.length] }} />)}
      </div>
      <ul className="ri-cats">
        {p.cats.map((c, i) => <li key={c.name}><i style={{ background: HUES[i % HUES.length] }} />{c.name}<b className="k-num">{c.pct}%</b><small className="k-num">{c.n} 条</small></li>)}
      </ul>
      {p.top.length > 0 && (
        <ol className="ri-top">
          {p.top.map((t, i) => (
            <li key={t.issue}><span className="ri-rank k-num">{i + 1}</span><span className="ri-issue">{t.issue}</span><span className="ri-track"><i style={{ width: `${(t.n / max) * 100}%` }} /></span><b className="k-num">{t.n}</b></li>
          ))}
        </ol>
      )}
      {(p.noIssue > 0 || p.notes.length > 0) && <p className="ri-note">{[p.noIssue ? `另有 ${p.noIssue} 条没有具体问题` : '', ...p.notes].filter(Boolean).join('；')}</p>}
    </div>
  );
}

export function InsightEmpty({ hover }: { hover: boolean }) {
  return (
    <div className={`ri-empty${hover ? ' hover' : ''}`}>
      <svg viewBox="0 0 420 120" aria-hidden="true">
        <g className="ri-s-doc">
          <rect x="8" y="12" width="104" height="96" rx="8" />
          {[0, 1, 2, 3].map((i) => <rect key={i} className="ri-s-line" x="20" y={26 + i * 20} width={i === 3 ? 48 : 80} height="8" rx="3" />)}
          <circle cx="94" cy="28" r="5" className="ri-s-star" />
        </g>
        <path className="ri-s-ar" d="M130 60h60m-9-9 9 9-9 9" />
        <g>
          <rect className="ri-s-seg" x="212" y="20" width="96" height="14" rx="3" fill="#D37017" />
          <rect className="ri-s-seg" x="308" y="20" width="60" height="14" rx="3" fill="#C0392B" />
          <rect className="ri-s-seg" x="368" y="20" width="40" height="14" rx="3" fill="#8E6BB8" />
          {[78, 58, 40].map((w, i) => <g key={i}><rect className="ri-s-line" x="212" y={52 + i * 22} width="190" height="12" rx="3" /><rect x="212" y={52 + i * 22} width={w * 2.4} height="12" rx="3" fill="#D37017" opacity={1 - i * 0.2} /></g>)}
        </g>
      </svg>
      <strong className="ri-h">把一堆差评，归成几个能改的问题</strong>
      <p>拖入店铺后台导出的评论表（Excel 或 CSV），AI 把差评和中评归到质量、物流、描述不符等类别，给出占比条、高频问题排行和改进建议。</p>
      <p className="ri-fmt"><FileSpreadsheet size={14} />支持 XLSX、CSV，可一次拖入多个文件</p>
    </div>
  );
}
