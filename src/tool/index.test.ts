import { describe, expect, it } from 'vitest';
import tool from './index';
import { readSheets, writeXlsx } from './xlsx';
import type { InputItem, ItemResult, Params, RunContext } from '../kit/types';

type Llm = RunContext['llm'];
function fakeCtx(files: Map<string, Uint8Array>, params: Params, llm: Llm): RunContext {
  return {
    params, outputDir: '/out', signal: new AbortController().signal, progress: () => {},
    readFile: async (p) => files.get(p)!,
    readText: async (p) => new TextDecoder().decode(files.get(p)!),
    writeFile: async (p, d) => void files.set(p, d),
    writeText: async (p, t) => void files.set(p, new TextEncoder().encode(t)),
    mkdir: async () => {},
    join: async (...parts) => parts.join('/'),
    llm, sidecar: async () => '',
  };
}
const defaults = (over: Params): Params => ({ ...Object.fromEntries(tool.params.map((f) => [f.key, f.default])), ...over });
const item = (name: string): InputItem => ({ id: name, name, path: '/in/' + name });

/** 假大模型：按评论里的关键词归类；建议请求返回固定建议 */
function fakeLlm(log: string[]): Llm {
  return async (m) => {
    const u = m[1].content;
    log.push(u);
    if (u.includes('suggestions')) return '```json\n{"suggestions":[{"problem":"快递慢","advice":"换用时效更稳定的快递"}]}\n```';
    const items = [...u.matchAll(/^\[(\d+)\][^\n]*?\s(\S[^\n]*)$/gm)].map(([, id, text]) => {
      const cat = /快递|物流|发货/.test(text) ? '物流' : /漏|坏|质量/.test(text) ? '质量问题' : /不错|喜欢/.test(text) ? '无问题' : '其他';
      const issue = cat === '物流' ? '快递慢' : cat === '质量问题' ? '漏水' : '';
      return { id: Number(id), category: cat, issue, reason: text.slice(0, 10) };
    });
    return '好的：' + JSON.stringify({ results: items });
  };
}

async function runAll(files: Map<string, Uint8Array>, names: string[], params: Params, llm: Llm) {
  const ctx = fakeCtx(files, defaults(params), llm);
  const results: { item: InputItem; result: ItemResult }[] = [];
  for (const n of names) results.push({ item: item(n), result: await tool.run(item(n), ctx) });
  const msg = await tool.afterAll!(results, ctx);
  return { results, msg };
}

describe('完整流程', () => {
  it('按评分筛出差评中评，分批归类，导出汇总 + 明细', async () => {
    const rows = Array.from({ length: 60 }, (_, i) => [`订单${i}`, i % 3 === 0 ? 5 : i % 3 === 1 ? 1 : 3, i % 2 ? '快递太慢了，等了一周' : '杯盖漏水']);
    const files = new Map([['/in/评论.xlsx', await writeXlsx([{ name: 'S', header: ['订单号', '评分', '评论内容'], rows }])]]);
    const log: string[] = [];
    const { results, msg } = await runAll(files, ['评论.xlsx'], {}, fakeLlm(log));
    // 60 条里 40 条 ≤3 分，每 25 条一批 → 2 次，再加 1 次建议；两类各 20 条并列时按类别顺序
    expect(log).toHaveLength(3);
    expect(results[0].result.row).toEqual({ total: 60, analyzed: 40, calls: 2, top: '质量问题 50%' });
    expect(msg).toBe('共分析 40 条差评 / 中评，调用大模型 3 次（含 1 次生成改进建议）。报告：/out/评论-差评归因报告.xlsx');
    const [sum, det] = await readSheets(files.get('/out/评论-差评归因报告.xlsx')!, 'x.xlsx');
    expect(sum.name).toBe('汇总');
    const flat = sum.grid.map((r) => r.filter((c) => c != null && c !== ''));
    expect(flat).toContainEqual(['大模型调用次数', 3]);
    expect(flat).toContainEqual(['物流快递', 20, '50%']);
    expect(flat).toContainEqual(['1. 换用时效更稳定的快递', '快递慢']);
    expect(det.grid[0]).toEqual(['行号', '评分', '评论内容', '问题类别', '问题点', '原因摘要']);
    expect(det.grid).toHaveLength(41);
    expect(det.grid[1].slice(0, 4)).toEqual([3, 1, '快递太慢了，等了一周', '物流快递']);
  });

  it('没有评分列时全部交给 AI，好评记为无问题，不计入占比；多个文件合并成一份报告', async () => {
    const files = new Map([
      ['/in/a.csv', new TextEncoder().encode('评价\n物流太慢\n很喜欢\n')],
      ['/in/b.csv', new TextEncoder().encode('买家评论,备注\n杯子漏水,x\n')],
    ]);
    const { results, msg } = await runAll(files, ['a.csv', 'b.csv'], { reportName: '十月差评' }, fakeLlm([]));
    expect(results[0].result.text).toContain('没有评分列，由 AI 判断是否差评');
    expect(results[0].result.text).toContain('无具体问题：1 条');
    expect(msg).toContain('报告：/out/十月差评.xlsx');
    const [sum, det] = await readSheets(files.get('/out/十月差评.xlsx')!, 'x.xlsx');
    expect(det.grid[0][0]).toBe('来源文件');
    expect(det.grid).toHaveLength(4);
    const flat = sum.grid.map((r) => r.filter((c) => c != null && c !== ''));
    expect(flat).toContainEqual(['物流快递', 1, '50%']);
    expect(flat).toContainEqual(['合并的文件数', 2]);
  });

  it('模型漏掉的评论会补问一次', async () => {
    const files = new Map([['/in/a.csv', new TextEncoder().encode('评论,评分\n快递慢,1\n漏水,1\n')]]);
    let n = 0;
    const llm: Llm = async (m) => {
      if (m[1].content.includes('suggestions')) return '{"suggestions":[]}';
      n++;
      return n === 1 ? '{"results":[{"id":1,"category":"物流快递","issue":"快递慢"}]}' : '{"results":[{"id":1,"category":"质量问题","issue":"漏水"}]}';
    };
    const { results } = await runAll(files, ['a.csv'], {}, llm);
    expect(n).toBe(2);
    expect(results[0].result.row?.calls).toBe(2);
    const det = (results[0].result.data as { details: { category: string }[] }).details;
    expect(det.map((x) => x.category)).toEqual(['物流快递', '质量问题']);
  });

  it('改进建议失败时报告照样生成，写明原因', async () => {
    const files = new Map([['/in/a.csv', new TextEncoder().encode('评论,评分\n快递慢,1\n')]]);
    const base = fakeLlm([]);
    const llm: Llm = async (m, o) => { if (m[1].content.includes('suggestions')) throw new Error('大模型调用失败（500，接口返回错误）'); return base(m, o); };
    await runAll(files, ['a.csv'], {}, llm);
    const [sum] = await readSheets(files.get('/out/a-差评归因报告.xlsx')!, 'x.xlsx');
    expect(sum.grid.flat()).toContain('改进建议生成失败：大模型调用失败（500，接口返回错误）');
  });

  it('只看差评时 3 分不分析；全是好评时不调用大模型', async () => {
    const files = new Map([['/in/a.csv', new TextEncoder().encode('评论,星级\n还行,3\n很好,5\n')]]);
    const log: string[] = [];
    const { results } = await runAll(files, ['a.csv'], { scope: 'bad' }, fakeLlm(log));
    expect(log).toHaveLength(0);
    expect(results[0].result.summary).toBe('2 条评论里没有差评');
  });

  it('没配置大模型时整个文件失败，提示去设置页', async () => {
    const files = new Map([['/in/a.csv', new TextEncoder().encode('评论,评分\n差,1\n')]]);
    const llm: Llm = async () => { throw new Error('还没有设置大模型：请到「设置」里填写接口地址、密钥和模型名'); };
    await expect(tool.run(item('a.csv'), fakeCtx(files, defaults({}), llm))).rejects.toThrow('还没有设置大模型');
  });

  it('认不出评论列时提示表头', async () => {
    const files = new Map([['/in/a.csv', new TextEncoder().encode('订单号,金额\n1,2\n')]]);
    await expect(tool.run(item('a.csv'), fakeCtx(files, defaults({}), async () => ''))).rejects.toThrow('表头是：订单号、金额');
  });
});
