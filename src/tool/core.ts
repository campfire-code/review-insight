// 纯函数：评分识别、差评筛选、分批、提示词构造与回复解析、统计汇总、报告表格。
import { extractJson, pick, splitList, stripBullet, type Cell } from './common';
import type { OutSheet } from './xlsx';

export type ChatMessage = { role: 'system' | 'user' | 'assistant'; content: string };

export const DEFAULT_CATEGORIES = ['质量问题', '物流快递', '描述不符', '客服服务', '包装破损', '价格', '其他'];
/** 模型判断这条其实不是差评（没有评分列时常见） */
export const NO_ISSUE = '无问题';
export const BATCH_SIZE = 25;

/** 用户填的类别：按顿号、逗号、换行拆开，去重，保证有「其他」 */
export function parseCategories(text: string): string[] {
  const list = [...new Set(splitList(text).map((s) => s.slice(0, 20)))].filter((c) => c !== NO_ISSUE);
  const out = list.length ? list : [...DEFAULT_CATEGORIES];
  if (!out.includes('其他')) out.push('其他');
  return out;
}

const CN_NUM: Record<string, number> = { 一: 1, 二: 2, 两: 2, 三: 3, 四: 4, 五: 5 };

/**
 * 识别一个评分单元格：数字、「4星」「三星」「★★☆☆☆」「差评/中评/好评」。识别不了返回 null。
 * 返回原始分值（可能是 10 分制或 100 分制，由 normalizeRatings 统一）。
 */
export function parseRating(v: Cell): number | null {
  if (v == null) return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const s = String(v).trim();
  if (!s) return null;
  if (/差评/.test(s)) return 1;
  if (/中评/.test(s)) return 3;
  if (/好评/.test(s)) return 5;
  const stars = (s.match(/[★⭐]/g) ?? []).length;
  if (stars) return stars;
  const m = s.match(/(\d+(?:\.\d+)?)/);
  if (m) return Number(m[1]);
  const c = s.match(/([一二两三四五])\s*[星分]/);
  if (c) return CN_NUM[c[1]];
  return null;
}

/** 按整列的最大值判断分制，统一换算成 5 分制 */
export function normalizeRatings(values: (number | null)[]): (number | null)[] {
  const max = Math.max(0, ...values.filter((v): v is number => v != null));
  const scale = max > 10 ? 20 : max > 5 ? 2 : 1;
  return values.map((v) => (v == null ? null : Math.round((v / scale) * 10) / 10));
}

export type Scope = 'bad_mid' | 'bad';
/** 是否要分析：差评 ≤2 分，中评 3 分；没有评分时交给模型判断 */
export const shouldAnalyze = (rating: number | null, scope: Scope) => rating == null || rating <= (scope === 'bad' ? 2 : 3);

export interface Review {
  id: number;
  file: string;
  /** 表格行号（含表头那一行，从 2 开始） */
  line: number;
  text: string;
  rating: number | null;
}

export interface Judged {
  category: string;
  issue: string;
  reason: string;
}

export const chunk = <T,>(arr: T[], size: number): T[][] => Array.from({ length: Math.ceil(arr.length / size) }, (_, i) => arr.slice(i * size, i * size + size));

export const SYSTEM_PROMPT = `你是电商售后分析师，负责给买家的差评、中评归类。规则：
1. 类别只能从给定列表里选一个；一条评论有多个问题时选最主要的。
2. 问题点用 4～10 个字的短语概括，例如「保温效果差」「快递太慢」「尺码偏小」。已经有的说法请直接复用，同一个问题不要换说法。
3. 原因摘要不超过 30 个字，忠于原文，不编造评论里没有的内容。
4. 如果评论其实是好评、或者没有说出任何具体问题，类别填「${NO_ISSUE}」。
5. 只输出 JSON，不要任何解释。`;

/** 构造一批评论的归类请求 */
export function buildBatchMessages(batch: Review[], categories: string[], knownIssues: string[]): ChatMessage[] {
  const lines = batch.map((r, i) => `[${i + 1}]${r.rating != null ? `（${r.rating} 分）` : ''} ${r.text.replace(/\s+/g, ' ').slice(0, 300)}`).join('\n');
  const known = knownIssues.length ? `\n已有的问题点（尽量复用）：${knownIssues.slice(0, 40).join('、')}` : '';
  const user = `类别：${categories.join('、')}${known}

评论共 ${batch.length} 条：
${lines}

按这个格式输出，每条评论都要有一项，id 是方括号里的编号：
{"results": [{"id": 1, "category": "类别", "issue": "问题点", "reason": "原因摘要"}]}`;
  return [{ role: 'system', content: SYSTEM_PROMPT }, { role: 'user', content: user }];
}

/** 把模型写的类别对到用户的类别上：完全一致 → 互相包含 → 「其他」 */
export function matchCategory(raw: string, categories: string[]): string {
  const s = stripBullet(String(raw ?? '')).replace(/[「」"']/g, '').trim();
  if (!s) return '其他';
  if (s === NO_ISSUE || /^(无|没有)(问题|差评)?$|^好评$|不是差评/.test(s)) return NO_ISSUE;
  const exact = categories.find((c) => c === s);
  if (exact) return exact;
  const fuzzy = categories.find((c) => c !== '其他' && (s.includes(c) || c.includes(s)));
  if (fuzzy) return fuzzy;
  // 再按前两个字粗配（「物流」对「物流快递」、「质量」对「质量问题」）
  const head = categories.find((c) => c !== '其他' && c.length >= 2 && s.includes(c.slice(0, 2)));
  return head ?? (categories.includes('其他') ? '其他' : categories[categories.length - 1]);
}

/**
 * 解析一批归类结果：兼容代码块、多余文字、顶层数组、按编号做键的对象、中文键名、缺 id（按顺序对应）。
 * 返回 编号(从 1 开始) → 结果；模型漏掉的编号不在里面。
 */
export function parseBatchReply(text: string, size: number, categories: string[]): Map<number, Judged> {
  const data = extractJson(text);
  let list: unknown[] = [];
  if (Array.isArray(data)) list = data;
  else {
    const inner = pick(data, ['results', 'result', 'items', 'data', 'list', '结果', 'reviews']);
    if (Array.isArray(inner)) list = inner;
    else if (data && typeof data === 'object') {
      // {"1": {...}, "2": {...}}
      const entries = Object.entries(data as Record<string, unknown>).filter(([k, v]) => /^\d+$/.test(k) && v && typeof v === 'object');
      list = entries.map(([k, v]) => ({ id: Number(k), ...(v as object) }));
    }
  }
  const out = new Map<number, Judged>();
  list.forEach((x, pos) => {
    if (!x || typeof x !== 'object') return;
    const rawId = pick(x, ['id', '编号', '序号', 'no', 'index']);
    const id = rawId != null && /^\d+$/.test(String(rawId).replace(/[[\]]/g, '')) ? Number(String(rawId).replace(/[[\]]/g, '')) : pos + 1;
    if (id < 1 || id > size || out.has(id)) return;
    const category = matchCategory(String(pick(x, ['category', '类别', '分类', 'type', 'label']) ?? ''), categories);
    const issue = String(pick(x, ['issue', '问题点', '问题', 'keyword', 'problem']) ?? '').trim().slice(0, 20);
    const reason = String(pick(x, ['reason', '原因摘要', '原因', '摘要', 'summary']) ?? '').trim().slice(0, 80);
    out.set(id, { category, issue: issue || (category === NO_ISSUE ? '' : category), reason });
  });
  if (!out.size) throw new Error('大模型返回的结果里没有归类');
  return out;
}

export interface Detail extends Review, Judged {}

export interface Summary {
  analyzed: number;
  noIssue: number;
  categories: { name: string; count: number; ratio: number }[];
  top: { issue: string; count: number; category: string; example: string }[];
}

const normIssue = (s: string) => s.replace(/[\s，。,.!！、]/g, '').toLowerCase();

/** 统计：各类占比（不含「无问题」）和高频问题 TOP N */
export function summarize(details: Detail[], categories: string[], topN = 10): Summary {
  const real = details.filter((d) => d.category !== NO_ISSUE);
  const counts = new Map(categories.map((c) => [c, 0]));
  for (const d of real) counts.set(d.category, (counts.get(d.category) ?? 0) + 1);
  const cats = [...counts.entries()].map(([name, count]) => ({ name, count, ratio: real.length ? count / real.length : 0 })).sort((a, b) => b.count - a.count || categories.indexOf(a.name) - categories.indexOf(b.name));
  const issues = new Map<string, { issue: string; count: number; cats: Map<string, number>; example: string }>();
  for (const d of real) {
    const key = normIssue(d.issue || d.category);
    let g = issues.get(key);
    if (!g) issues.set(key, (g = { issue: d.issue || d.category, count: 0, cats: new Map(), example: d.text.replace(/\s+/g, ' ').slice(0, 60) }));
    g.count++;
    g.cats.set(d.category, (g.cats.get(d.category) ?? 0) + 1);
  }
  const top = [...issues.values()].sort((a, b) => b.count - a.count).slice(0, topN).map((g) => ({
    issue: g.issue, count: g.count, example: g.example,
    category: [...g.cats.entries()].sort((a, b) => b[1] - a[1])[0][0],
  }));
  return { analyzed: details.length, noIssue: details.length - real.length, categories: cats, top };
}

export const pct = (r: number) => `${(r * 100).toFixed(1).replace(/\.0$/, '')}%`;

/** 构造改进建议请求：只给统计结果和少量原因，不重复发送全部评论 */
export function buildSuggestMessages(s: Summary, details: Detail[]): ChatMessage[] {
  const cats = s.categories.filter((c) => c.count).map((c) => `${c.name} ${c.count} 条（${pct(c.ratio)}）`).join('；');
  const top = s.top.map((t, i) => {
    const reasons = details.filter((d) => normIssue(d.issue) === normIssue(t.issue) && d.reason).slice(0, 3).map((d) => d.reason);
    return `${i + 1}. ${t.issue}（${t.count} 次，${t.category}）${reasons.length ? `：${reasons.join('；')}` : ''}`;
  }).join('\n');
  return [
    { role: 'system', content: '你是电商运营顾问。根据差评统计给出具体、可执行的改进建议，优先解决次数多的问题。不编造数据，不写空话。只输出 JSON。' },
    { role: 'user', content: `问题类别分布：${cats}\n\n高频问题：\n${top}\n\n给出 3～6 条改进建议，每条对应一个问题，建议不超过 60 个字。\n格式：{"suggestions": [{"problem": "针对的问题", "advice": "改进建议"}]}` },
  ];
}

export function parseSuggestReply(text: string): { problem: string; advice: string }[] {
  let data: unknown;
  try {
    data = extractJson(text);
  } catch {
    const lines = String(text).split('\n').map(stripBullet).filter((l) => l && !/[{}[\]]/.test(l));
    return lines.slice(0, 8).map((advice) => ({ problem: '', advice }));
  }
  const list = Array.isArray(data) ? data : pick(data, ['suggestions', 'suggestion', 'advice', '建议', '改进建议', 'items', 'list']);
  if (!Array.isArray(list)) return [];
  return list.map((x) => (typeof x === 'string'
    ? { problem: '', advice: stripBullet(x) }
    : { problem: String(pick(x, ['problem', 'issue', '问题', '针对的问题']) ?? '').trim(), advice: String(pick(x, ['advice', 'suggestion', '建议', '改进建议', 'content']) ?? '').trim() })).filter((x) => x.advice);
}

export interface Totals {
  files: number;
  total: number;
  skipped: number;
  calls: number;
}

/** 报告的两张表：汇总 + 明细 */
export function reportSheets(details: Detail[], s: Summary, suggestions: { problem: string; advice: string }[] | string, t: Totals): OutSheet[] {
  const rows: Cell[][] = [];
  const bold: number[] = [];
  const title = (r: Cell[]) => { bold.push(rows.length); rows.push(r); };
  rows.push(['评论总数', t.total], ['分析的差评 / 中评', s.analyzed], ['跳过（好评或空评论）', t.skipped], ['判定为无具体问题', s.noIssue], ['大模型调用次数', t.calls]);
  if (t.files > 1) rows.push(['合并的文件数', t.files]);
  rows.push([]);
  title(['问题类别', '条数', '占比']);
  for (const c of s.categories) rows.push([c.name, c.count, pct(c.ratio)]);
  rows.push([]);
  title(['高频问题 TOP10', '次数', '所属类别', '典型评论']);
  s.top.forEach((x) => rows.push([x.issue, x.count, x.category, x.example]));
  rows.push([]);
  title(['改进建议', '针对的问题']);
  if (typeof suggestions === 'string') rows.push([suggestions]);
  else suggestions.forEach((x, i) => rows.push([`${i + 1}. ${x.advice}`, x.problem]));

  const multi = t.files > 1;
  const header = [...(multi ? ['来源文件'] : []), '行号', '评分', '评论内容', '问题类别', '问题点', '原因摘要'];
  const detailRows = details.map((d) => [...(multi ? [d.file] : []), d.line, d.rating, d.text, d.category, d.issue, d.reason] as Cell[]);
  const textCol = multi ? 3 : 2;
  return [
    { name: '汇总', header: ['概况', '数量'], rows, plain: true, boldRows: bold, wrap: [0, 3] },
    { name: '明细', header, rows: detailRows, wrap: [textCol, textCol + 3] },
  ];
}
