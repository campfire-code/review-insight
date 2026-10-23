import { describe, expect, it } from 'vitest';
import {
  NO_ISSUE, buildBatchMessages, buildSuggestMessages, chunk, matchCategory, normalizeRatings, parseBatchReply, parseCategories, parseRating,
  parseSuggestReply, reportSheets, shouldAnalyze, summarize, type Detail,
} from './core';

const CATS = parseCategories('质量问题、物流快递、描述不符、客服服务、包装破损、价格');

describe('类别', () => {
  it('拆分、去重、自动补「其他」、空时用默认', () => {
    expect(CATS).toEqual(['质量问题', '物流快递', '描述不符', '客服服务', '包装破损', '价格', '其他']);
    expect(parseCategories('尺码\n尺码\n做工')).toEqual(['尺码', '做工', '其他']);
    expect(parseCategories('  ')).toContain('物流快递');
  });
  it('模型写的类别对到用户类别', () => {
    expect(matchCategory('物流快递', CATS)).toBe('物流快递');
    expect(matchCategory('物流', CATS)).toBe('物流快递');
    expect(matchCategory('产品质量问题', CATS)).toBe('质量问题');
    expect(matchCategory('「客服服务」', CATS)).toBe('客服服务');
    expect(matchCategory('售后政策', CATS)).toBe('其他');
    expect(matchCategory('无问题', CATS)).toBe(NO_ISSUE);
    expect(matchCategory('好评', CATS)).toBe(NO_ISSUE);
    expect(matchCategory('', CATS)).toBe('其他');
  });
});

describe('评分', () => {
  it('识别各种写法', () => {
    expect(parseRating(2)).toBe(2);
    expect(parseRating('4星')).toBe(4);
    expect(parseRating('三星')).toBe(3);
    expect(parseRating('★★☆☆☆')).toBe(2);
    expect(parseRating('差评')).toBe(1);
    expect(parseRating('中评')).toBe(3);
    expect(parseRating('好评')).toBe(5);
    expect(parseRating('')).toBeNull();
    expect(parseRating('未评分')).toBeNull();
  });
  it('10 分制、100 分制换算成 5 分制', () => {
    expect(normalizeRatings([2, 10, null])).toEqual([1, 5, null]);
    expect(normalizeRatings([40, 100])).toEqual([2, 5]);
    expect(normalizeRatings([1, 5])).toEqual([1, 5]);
  });
  it('筛选差评 / 中评', () => {
    expect(shouldAnalyze(3, 'bad_mid')).toBe(true);
    expect(shouldAnalyze(3, 'bad')).toBe(false);
    expect(shouldAnalyze(4, 'bad_mid')).toBe(false);
    expect(shouldAnalyze(null, 'bad')).toBe(true);
  });
  it('分批', () => expect(chunk([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]]));
});

const rev = (id: number, text: string, rating: number | null = 1) => ({ id, file: 'a.xlsx', line: id + 2, text, rating });

describe('提示词', () => {
  it('带类别、已有问题点、编号和评分，长评论截断', () => {
    const m = buildBatchMessages([rev(0, '快递太慢了'), rev(1, '长'.repeat(400), null)], CATS, ['快递太慢']);
    expect(m[0].content).toContain(NO_ISSUE);
    expect(m[1].content).toContain('类别：质量问题、物流快递');
    expect(m[1].content).toContain('已有的问题点（尽量复用）：快递太慢');
    expect(m[1].content).toContain('[1]（1 分） 快递太慢了');
    expect(m[1].content).toContain('[2] ' + '长'.repeat(300) + '\n');
    expect(m[1].content).toContain('"results"');
  });
});

describe('解析归类结果（容错）', () => {
  it('标准格式', () => {
    const m = parseBatchReply('{"results":[{"id":1,"category":"物流快递","issue":"快递慢","reason":"5天才到"},{"id":2,"category":"质量问题","issue":"漏水","reason":"杯盖漏水"}]}', 2, CATS);
    expect(m.get(1)).toEqual({ category: '物流快递', issue: '快递慢', reason: '5天才到' });
    expect(m.get(2)?.category).toBe('质量问题');
  });
  it('代码块 + 顶层数组 + 中文键名 + 字符串编号', () => {
    const m = parseBatchReply('结果如下\n```json\n[{"编号":"2","类别":"物流","问题点":"发货慢","原因":"拍下三天才发"}]\n```', 2, CATS);
    expect(m.get(2)).toEqual({ category: '物流快递', issue: '发货慢', reason: '拍下三天才发' });
    expect(m.has(1)).toBe(false);
  });
  it('缺 id 时按顺序对应；缺问题点时用类别', () => {
    const m = parseBatchReply('{"items":[{"category":"价格"},{"category":"包装"}]}', 2, CATS);
    expect(m.get(1)).toEqual({ category: '价格', issue: '价格', reason: '' });
    expect(m.get(2)?.category).toBe('包装破损');
  });
  it('以编号为键的对象', () => {
    const m = parseBatchReply('{"1":{"category":"客服","issue":"不回消息"}}', 1, CATS);
    expect(m.get(1)?.category).toBe('客服服务');
  });
  it('越界编号和重复编号忽略；无问题保留', () => {
    const m = parseBatchReply('[{"id":1,"category":"无问题"},{"id":1,"category":"价格"},{"id":9,"category":"价格"}]', 2, CATS);
    expect(m.size).toBe(1);
    expect(m.get(1)).toEqual({ category: NO_ISSUE, issue: '', reason: '' });
  });
  it('没有任何归类时报错', () => {
    expect(() => parseBatchReply('{"msg":"ok"}', 2, CATS)).toThrow('没有归类');
    expect(() => parseBatchReply('我无法处理', 2, CATS)).toThrow();
  });
});

const d = (id: number, category: string, issue: string, text = '评论' + id): Detail => ({ ...rev(id, text), category, issue, reason: issue + '的原因' });

describe('汇总', () => {
  const details = [d(0, '物流快递', '快递慢'), d(1, '物流快递', '快递 慢'), d(2, '质量问题', '漏水'), d(3, NO_ISSUE, ''), d(4, '物流快递', '发错地址')];
  const s = summarize(details, CATS);
  it('占比不含「无问题」，按条数排序', () => {
    expect(s.analyzed).toBe(5);
    expect(s.noIssue).toBe(1);
    expect(s.categories[0]).toEqual({ name: '物流快递', count: 3, ratio: 0.75 });
    expect(s.categories[1]).toEqual({ name: '质量问题', count: 1, ratio: 0.25 });
    expect(s.categories.find((c) => c.name === '价格')?.count).toBe(0);
  });
  it('高频问题合并空格差异，带所属类别和典型评论', () => {
    expect(s.top[0]).toEqual({ issue: '快递慢', count: 2, category: '物流快递', example: '评论0' });
    expect(s.top).toHaveLength(3);
  });
  it('TOP10 只取前 10', () => {
    const many = Array.from({ length: 15 }, (_, i) => d(i, '其他', '问题' + i));
    expect(summarize(many, CATS).top).toHaveLength(10);
  });
  it('建议请求带统计和原因', () => {
    const m = buildSuggestMessages(s, details);
    expect(m[1].content).toContain('物流快递 3 条（75%）');
    expect(m[1].content).toContain('1. 快递慢（2 次，物流快递）：快递慢的原因；快递 慢的原因');
  });
  it('解析建议：对象数组、字符串数组、纯文本', () => {
    expect(parseSuggestReply('{"suggestions":[{"problem":"快递慢","advice":"换快递"}]}')).toEqual([{ problem: '快递慢', advice: '换快递' }]);
    expect(parseSuggestReply('```json\n{"建议":["1. 换快递","2. 加固包装"]}\n```')).toEqual([{ problem: '', advice: '换快递' }, { problem: '', advice: '加固包装' }]);
    expect(parseSuggestReply('1. 换快递\n2. 加固包装')).toHaveLength(2);
    expect(parseSuggestReply('{"x":1}')).toEqual([]);
  });
  it('报告：汇总页含调用次数、占比、TOP10、建议；明细页多文件时带来源列', () => {
    const [sum, det] = reportSheets(details, s, [{ problem: '快递慢', advice: '换快递' }], { files: 2, total: 20, skipped: 15, calls: 2 });
    expect(sum.rows).toContainEqual(['大模型调用次数', 2]);
    expect(sum.rows).toContainEqual(['物流快递', 3, '75%']);
    expect(sum.rows).toContainEqual(['快递慢', 2, '物流快递', '评论0']);
    expect(sum.rows).toContainEqual(['1. 换快递', '快递慢']);
    expect(det.header).toEqual(['来源文件', '行号', '评分', '评论内容', '问题类别', '问题点', '原因摘要']);
    expect(det.rows[0]).toEqual(['a.xlsx', 2, 1, '评论0', '物流快递', '快递慢', '快递慢的原因']);
    const [, single] = reportSheets(details, s, '建议生成失败', { files: 1, total: 5, skipped: 0, calls: 1 });
    expect(single.header[0]).toBe('行号');
  });
});
