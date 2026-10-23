// 通用纯函数：表格解析（CSV 编码识别、表头、找列）、大模型回复里提取 JSON、限流并发。
// 不依赖界面和本机能力，便于单元测试。

export type Cell = string | number | boolean | Date | null;
export interface Table {
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
  const nl = text.indexOf('\n');
  const firstLine = text.slice(0, nl === -1 ? text.length : nl);
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

/** CSV 里的文本数字转成数值；以 0 开头的编号和超过 15 位的长数字保持文本 */
export function numeric(v: string): Cell {
  const t = v.trim();
  return /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(t) && t.replace(/[-.]/g, '').length <= 15 ? Number(t) : v;
}

/** 第一行非空行当表头，空表头补成「列1」「列2」，去掉全空行 */
export function toTable(grid: Cell[][]): Table {
  const blank = (c: Cell) => c == null || String(c).trim() === '';
  const hi = grid.findIndex((r) => r && r.some((c) => !blank(c)));
  if (hi < 0) return { header: [], rows: [] };
  const raw = grid[hi];
  const width = Math.max(raw.length, ...grid.slice(hi + 1).map((r) => r?.length ?? 0));
  const header = Array.from({ length: width }, (_, i) => (blank(raw[i] ?? null) ? `列${i + 1}` : String(raw[i]).trim()));
  const rows = grid.slice(hi + 1).filter((r) => r && r.some((c) => !blank(c))).map((r) => header.map((_, i) => r[i] ?? null));
  return { header, rows };
}

/** 列号字母转序号：A→0、Z→25、AA→26 */
export function letterToIndex(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** 找列：表头名精确匹配 → 去空格匹配 → 列字母（A、B…）→ 序号（1、2…） */
export function resolveColumn(header: string[], spec: string): number {
  const want = spec.trim();
  if (!want) return -1;
  let i = header.findIndex((h) => h === want);
  if (i >= 0) return i;
  i = header.findIndex((h) => h.replace(/\s/g, '') === want.replace(/\s/g, ''));
  if (i >= 0) return i;
  if (/^[A-Za-z]{1,3}$/.test(want)) {
    const k = letterToIndex(want);
    return k < header.length ? k : -1;
  }
  if (/^\d+$/.test(want)) {
    const k = Number(want) - 1;
    return k >= 0 && k < header.length ? k : -1;
  }
  return -1;
}

/** 按关键词自动找列：按关键词顺序优先，表头包含关键词即算 */
export function guessColumn(header: string[], keywords: RegExp[]): number {
  for (const re of keywords) {
    const i = header.findIndex((h) => re.test(h));
    if (i >= 0) return i;
  }
  return -1;
}

/** 把用户填的「列名，列名」拆开 */
export const splitList = (s: string) => s.split(/[,，、;；\n]+/).map((x) => x.trim()).filter(Boolean);

export const cellText = (v: Cell) => (v == null ? '' : v instanceof Date ? v.toISOString().slice(0, 10) : String(v).trim());

/** 去掉扩展名 */
export const baseName = (file: string) => file.replace(/\.[^.]+$/, '');

/** 文件名里不能出现的字符替换掉，并限制长度 */
export function safeName(s: string): string {
  const cleaned = s.replace(/[\\/:*?"<>|\r\n\t]/g, '_').replace(/^\.+/, '_').trim();
  return (cleaned || '未命名').slice(0, 80);
}

/**
 * 从大模型回复里取出 JSON：兼容 ```json 包裹、前后有说明文字、结尾多余逗号、中文引号包裹键名。
 * 取第一个完整的 { } 或 [ ]；都取不到时抛出错误。
 */
export function extractJson(text: string): unknown {
  const s = String(text ?? '').trim();
  const candidates: string[] = [];
  const fence = s.match(/```(?:json|JSON)?\s*([\s\S]*?)```/);
  if (fence) candidates.push(fence[1].trim());
  candidates.push(s);
  for (const c of candidates) {
    // 先试位置靠前的括号：顶层是数组时 [ 在 { 前面
    const starts = ['{', '['].map((o) => c.indexOf(o)).filter((i) => i >= 0).sort((a, b) => a - b);
    for (const start of starts) {
      const body = balanced(c, start);
      if (!body) continue;
      for (const variant of [body, body.replace(/,\s*([}\]])/g, '$1'), body.replace(/[“”]/g, '"').replace(/,\s*([}\]])/g, '$1')]) {
        try {
          return JSON.parse(variant);
        } catch { /* 换下一种修正方式 */ }
      }
    }
  }
  throw new Error('大模型返回的内容不是有效的 JSON');
}

/** 从 start 处的括号开始，找到与之配对的结束括号（跳过字符串里的括号） */
function balanced(s: string, start: number): string | null {
  const stack: string[] = [];
  let inStr = false;
  for (let i = start; i < s.length; i++) {
    const c = s[i];
    if (inStr) {
      if (c === '\\') i++;
      else if (c === '"') inStr = false;
      continue;
    }
    if (c === '"') inStr = true;
    else if (c === '{' || c === '[') stack.push(c === '{' ? '}' : ']');
    else if (c === '}' || c === ']') {
      if (stack.pop() !== c) return null;
      if (!stack.length) return s.slice(start, i + 1);
    }
  }
  return null;
}

/** 从对象里按多个候选键名取值（大模型可能用中文或别的英文键名） */
export function pick(obj: unknown, keys: string[]): unknown {
  if (!obj || typeof obj !== 'object') return undefined;
  const o = obj as Record<string, unknown>;
  for (const k of keys) if (o[k] !== undefined && o[k] !== null) return o[k];
  const lower = Object.fromEntries(Object.entries(o).map(([k, v]) => [k.toLowerCase(), v]));
  for (const k of keys) if (lower[k.toLowerCase()] !== undefined && lower[k.toLowerCase()] !== null) return lower[k.toLowerCase()];
  return undefined;
}

/** 把字符串 / 数组 / 对象数组统一成字符串数组，去掉编号前缀和包裹的引号 */
export function toStringList(v: unknown): string[] {
  let arr: unknown[];
  if (Array.isArray(v)) arr = v;
  else if (typeof v === 'string') arr = v.split(/\n+/);
  else return [];
  return arr
    .map((x) => (typeof x === 'string' || typeof x === 'number' ? String(x) : x && typeof x === 'object' ? String(pick(x, ['text', 'content', 'title', '标题', '内容', 'value']) ?? '') : ''))
    .map(stripBullet)
    .filter(Boolean);
}

export const stripBullet = (s: string) => s.trim().replace(/^(?:\d{1,2}[.、)）:：]|[-*•·]|[（(]\d{1,2}[)）])\s*/, '').replace(/^["“「『]+|["”」』]+$/g, '').trim();

/** 限流并发：最多 n 个同时进行，结果按原顺序返回 */
export async function mapPool<T, R>(items: T[], n: number, fn: (item: T, index: number) => Promise<R>, signal?: AbortSignal): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      if (signal?.aborted) throw new Error('已暂停');
      const i = next++;
      out[i] = await fn(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, Math.min(n, items.length)) }, worker));
  return out;
}

/** 这类错误重试也没用，应该停下整批：没配置大模型、密钥不对、地址或模型名不对、额度用完 */
export const isFatalLlmError = (e: unknown) => /还没有设置大模型|（401|（403|（404|（429|已暂停/.test(e instanceof Error ? e.message : String(e));

export const errText = (e: unknown) => (e instanceof Error ? e.message : String(e));
