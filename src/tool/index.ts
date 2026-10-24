// 工具声明：差评归因分析
import { MessageSquareWarning as MessageSquareWarningAsIcon } from 'lucide-react';
import { createElement } from 'react';
import type { BatchTool, RunContext } from '../kit/types';
import { baseName, cellText, errText, guessColumn, isFatalLlmError, mapPool, resolveColumn, safeName, toTable } from './common';
import {
  BATCH_SIZE, DEFAULT_CATEGORIES, NO_ISSUE, buildBatchMessages, buildSuggestMessages, chunk, normalizeRatings, parseBatchReply, parseCategories,
  parseRating, parseSuggestReply, pct, reportSheets, shouldAnalyze, summarize, type Detail, type Judged, type Review, type Scope,
} from './core';
import { readSheets, writeXlsx } from './xlsx';
import { InsightBody, InsightEmpty } from './Insight';

const COMMENT_KEYWORDS = [/^(评论|评价)(内容)?$/, /评论内容|评价内容|追评|买家评论|买家评价/, /评论|评价|留言/, /^内容$/, /comment|review|content/i];
const RATING_KEYWORDS = [/评分|星级|打分|分数|评级|满意度/, /^(星|分)$/, /score|rating|star/i, /好评|差评/];

export interface FileData {
  file: string;
  total: number;
  skipped: number;
  calls: number;
  details: Detail[];
}

/** 一批评论：调用一次；漏掉的那几条再单独补一次；仍缺的记为「其他」 */
async function judgeBatch(batch: Review[], categories: string[], known: string[], ctx: RunContext, count: () => void): Promise<Detail[]> {
  const got = new Map<number, Judged>();
  let pending = batch;
  let lastErr: unknown;
  for (let attempt = 0; attempt < 2 && pending.length; attempt++) {
    try {
      count();
      const reply = await ctx.llm(buildBatchMessages(pending, categories, known), { json: true, temperature: 0.2 });
      const m = parseBatchReply(reply, pending.length, categories);
      pending.forEach((r, i) => { const j = m.get(i + 1); if (j) got.set(r.id, j); });
      pending = pending.filter((r) => !got.has(r.id));
    } catch (e) {
      if (isFatalLlmError(e)) throw e;
      lastErr = e;
    }
  }
  if (!got.size && lastErr) throw lastErr;
  for (const j of got.values()) if (j.issue && j.category !== NO_ISSUE && !known.includes(j.issue)) known.push(j.issue);
  return batch.map((r) => ({ ...r, ...(got.get(r.id) ?? { category: '其他', issue: '', reason: '（大模型没有返回这条的结果，可重新处理）' }) }));
}

const tool: BatchTool = {
  id: 'e20',
  name: '差评归因分析',
  tagline: '导入评论导出表，大模型把差评、中评归到质量、物流、描述不符等类别，写出每条原因，汇总占比、高频问题前十 和改进建议，导出 Excel 报告。',
  icon: MessageSquareWarningAsIcon,
  category: '电商运营',
  type: 'data',
  mode: 'batch',
  resultView: 'text',
  Empty: InsightEmpty,
  ResultBody: ({ result }) => createElement(InsightBody, { result }),
  columns: [['total', '评论'], ['analyzed', '差/中评'], ['calls', '调用次数'], ['top', '最多的问题']],
  needsLLM: true,
  concurrency: 1,
  accept: { label: '评论表格', extensions: ['xlsx', 'csv', 'xls'] },
  params: [
    { key: 'scope', label: '分析哪些评论', type: 'segment', options: [['bad_mid', '差评 + 中评'], ['bad', '只看差评']], default: 'bad_mid', help: '按 5 分制：差评 1～2 分，中评 3 分。没有评分列时全部交给 AI 判断' },
    { key: 'categories', label: '问题类别', type: 'textarea', rows: 3, default: DEFAULT_CATEGORIES.join('、'), help: '用顿号或换行隔开，可以按自己的品类改，例如加上「尺码偏差」' },
    { key: 'commentCol', label: '评论在哪一列', type: 'text', default: '', placeholder: '不填自动识别「评论内容」等列；也可以写列字母' },
    { key: 'ratingCol', label: '评分在哪一列', type: 'text', default: '', placeholder: '不填自动识别；没有评分列就填「无」' },
    { key: 'maxCount', label: '最多分析多少条', type: 'number', default: 1000, min: 25, max: 5000, step: 25, unit: '条', help: `每 ${BATCH_SIZE} 条调用一次大模型，条数越多费用越高` },
    { key: 'reportName', label: '报告文件名', type: 'text', default: '差评归因报告' },
    { key: 'out', label: '保存到', type: 'folder', default: '' },
  ],
  check: (item) => (/\.xls$/i.test(item.name) ? '旧版 .xls 格式不支持，请用 Excel 或 WPS 另存为 .xlsx 后再拖进来' : null),

  async run(item, ctx) {
    const p = ctx.params;
    const sheets = await readSheets(await ctx.readFile(item.path!), item.name);
    const sheet = sheets.find((s) => s.grid.some((r) => r.length));
    if (!sheet) throw new Error('表格是空的');
    const table = toTable(sheet.grid);
    if (!table.rows.length) throw new Error('表格里只有表头，没有评论');
    const heads = `${table.header.slice(0, 12).join('、')}${table.header.length > 12 ? '…' : ''}`;

    const cSpec = String(p.commentCol ?? '').trim();
    const cCol = cSpec ? resolveColumn(table.header, cSpec) : guessColumn(table.header, COMMENT_KEYWORDS);
    if (cCol < 0) throw new Error(cSpec ? `找不到列「${cSpec}」，表头是：${heads}` : `没认出哪一列是评论，请在「评论在哪一列」里填写。表头是：${heads}`);
    const rSpec = String(p.ratingCol ?? '').trim();
    let rCol = -1;
    if (rSpec && rSpec !== '无') {
      rCol = resolveColumn(table.header, rSpec);
      if (rCol < 0) throw new Error(`找不到评分列「${rSpec}」，表头是：${heads}`);
    } else if (!rSpec) {
      rCol = guessColumn(table.header.map((h, i) => (i === cCol ? '' : h)), RATING_KEYWORDS);
    }

    const ratings = normalizeRatings(table.rows.map((r) => (rCol >= 0 ? parseRating(r[rCol]) : null)));
    const scope = (p.scope === 'bad' ? 'bad' : 'bad_mid') as Scope;
    const withText = table.rows.map((r, i) => ({ text: cellText(r[cCol]), rating: ratings[i], line: i + 2 })).filter((x) => x.text);
    const picked = withText.filter((x) => shouldAnalyze(x.rating, scope));
    const max = Math.max(1, Number(p.maxCount) || 1000);
    const reviews: Review[] = picked.slice(0, max).map((x, i) => ({ id: i, file: item.name, line: x.line, text: x.text, rating: x.rating }));
    const total = table.rows.length;
    const skipped = total - reviews.length;
    if (!reviews.length) {
      const data: FileData = { file: item.name, total, skipped, calls: 0, details: [] };
      return { summary: `${total} 条评论里没有${scope === 'bad' ? '差评' : '差评或中评'}`, row: { total, analyzed: 0, calls: 0, top: '—' }, data };
    }

    const categories = parseCategories(String(p.categories ?? ''));
    const known: string[] = [];
    let calls = 0;
    let done = 0;
    const batches = chunk(reviews, BATCH_SIZE);
    const failed: string[] = [];
    const parts = await mapPool(batches, 2, async (b) => {
      try {
        return await judgeBatch(b, categories, known, ctx, () => calls++);
      } catch (e) {
        if (isFatalLlmError(e)) throw e;
        failed.push(errText(e));
        return b.map((r) => ({ ...r, category: '其他', issue: '', reason: `（这一批分析失败：${errText(e).slice(0, 40)}）` }));
      } finally {
        ctx.progress(0.05 + (0.95 * ++done) / batches.length);
      }
    }, ctx.signal);
    if (failed.length === batches.length) throw new Error(`全部分析失败：${failed[0]}`);

    const details = parts.flat();
    const s = summarize(details, categories);
    const first = s.categories[0];
    const data: FileData = { file: item.name, total, skipped, calls, details };
    const lines = s.categories.filter((c) => c.count).map((c) => `${c.name}：${c.count} 条（${pct(c.ratio)}）`);
    const text = [
      `${total} 条评论，分析 ${reviews.length} 条${picked.length > max ? `（超过上限，只分析了前 ${max} 条）` : ''}，调用大模型 ${calls} 次${rCol < 0 ? '；没有评分列，由 AI 判断是否差评' : ''}`,
      ...lines,
      s.noIssue ? `无具体问题：${s.noIssue} 条` : '',
      s.top.length ? `高频问题：${s.top.slice(0, 5).map((t) => `${t.issue}（${t.count}）`).join('、')}` : '',
      failed.length ? `${failed.length} 批分析失败，相关评论记为「其他」` : '',
    ].filter(Boolean).join('\n');
    return {
      summary: `分析 ${reviews.length} 条，调用 ${calls} 次`,
      row: { total, analyzed: reviews.length, calls, top: first?.count ? `${first.name} ${pct(first.ratio)}` : '—' },
      text,
      data,
    };
  },

  async afterAll(results, ctx) {
    const files = results.map((r) => r.result.data as FileData | undefined).filter((d): d is FileData => !!d);
    if (!files.length) return;
    const details = files.flatMap((f) => f.details);
    const categories = parseCategories(String(ctx.params.categories ?? ''));
    const s = summarize(details, categories);
    let calls = files.reduce((a, f) => a + f.calls, 0);
    let suggestions: { problem: string; advice: string }[] | string = '没有需要改进的问题';
    if (s.analyzed - s.noIssue > 0) {
      try {
        calls++;
        suggestions = parseSuggestReply(await ctx.llm(buildSuggestMessages(s, details), { json: true, temperature: 0.5 }));
        if (!suggestions.length) suggestions = '大模型没有给出建议，可重新处理';
      } catch (e) {
        suggestions = `改进建议生成失败：${errText(e)}`;
      }
    }
    const totals = { files: files.length, total: files.reduce((a, f) => a + f.total, 0), skipped: files.reduce((a, f) => a + f.skipped, 0), calls };
    const name = files.length === 1 ? `${baseName(files[0].file)}-${String(ctx.params.reportName || '差评归因报告')}` : String(ctx.params.reportName || '差评归因报告');
    const path = await ctx.join(ctx.outputDir, `${safeName(name)}.xlsx`);
    await ctx.writeFile(path, await writeXlsx(reportSheets(details, s, suggestions, totals), '篝火 · 差评归因分析'));
    return `共分析 ${s.analyzed} 条差评 / 中评，调用大模型 ${calls} 次${s.analyzed - s.noIssue > 0 ? '（含 1 次生成改进建议）' : ''}。报告：${path}`;
  },
};

export default tool;
