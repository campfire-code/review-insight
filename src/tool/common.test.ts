import { describe, expect, it } from 'vitest';
import { decodeText, extractJson, guessColumn, isFatalLlmError, mapPool, parseCsv, pick, resolveColumn, toStringList, toTable } from './common';

describe('extractJson 容错', () => {
  it('纯 JSON', () => expect(extractJson('{"a":1}')).toEqual({ a: 1 }));
  it('```json 包裹', () => expect(extractJson('好的，结果如下：\n```json\n{"a": [1, 2]}\n```\n希望有帮助')).toEqual({ a: [1, 2] }));
  it('前后有多余文字、没有代码块', () => expect(extractJson('结果：{"t":"含 } 的字符串"} 以上')).toEqual({ t: '含 } 的字符串' }));
  it('结尾多余逗号', () => expect(extractJson('{"a":[1,2,],}')).toEqual({ a: [1, 2] }));
  it('中文引号', () => expect(extractJson('{“a”: “b”}')).toEqual({ a: 'b' }));
  it('顶层是数组', () => expect(extractJson('[{"id":1}]')).toEqual([{ id: 1 }]));
  it('完全不是 JSON 时报错', () => expect(() => extractJson('抱歉，我无法完成')).toThrow('不是有效的 JSON'));
  it('括号不完整时报错', () => expect(() => extractJson('{"a": [1, 2')).toThrow());
});

describe('字段容错', () => {
  it('pick 支持多个键名和大小写', () => {
    expect(pick({ Titles: ['x'] }, ['titles'])).toEqual(['x']);
    expect(pick({ 标题: 'y' }, ['titles', '标题'])).toBe('y');
    expect(pick(null, ['a'])).toBeUndefined();
  });
  it('toStringList 去编号、去引号、兼容对象数组和换行字符串', () => {
    expect(toStringList(['1. 甲', '2、乙', '“丙”'])).toEqual(['甲', '乙', '丙']);
    expect(toStringList('- 甲\n- 乙\n')).toEqual(['甲', '乙']);
    expect(toStringList([{ text: '甲' }, { title: '乙' }])).toEqual(['甲', '乙']);
    expect(toStringList(undefined)).toEqual([]);
  });
});

describe('表格', () => {
  it('CSV 引号、换行、GBK', () => {
    expect(parseCsv('a,b\n"x,1","多\n行"\n')).toEqual([['a', 'b'], ['x,1', '多\n行']]);
    const gbk = new Uint8Array([0xc9, 0xcc, 0xc6, 0xb7]); // 「商品」的 GBK 编码
    expect(decodeText(gbk)).toBe('商品');
  });
  it('toTable 跳过开头空行、补表头、补齐列', () => {
    const t = toTable([[], ['名称', null], ['杯子'], [null, null]]);
    expect(t.header).toEqual(['名称', '列2']);
    expect(t.rows).toEqual([['杯子', null]]);
  });
  it('找列', () => {
    const h = ['商品名称', '品类', '卖点'];
    expect(resolveColumn(h, 'B')).toBe(1);
    expect(resolveColumn(h, '3')).toBe(2);
    expect(resolveColumn(h, ' 品 类 ')).toBe(1);
    expect(resolveColumn(h, '价格')).toBe(-1);
    expect(guessColumn(h, [/标题/, /名称/])).toBe(0);
  });
});

describe('并发与错误', () => {
  it('mapPool 不超过并发上限且保持顺序', async () => {
    let live = 0;
    let max = 0;
    const out = await mapPool([1, 2, 3, 4, 5], 2, async (x) => {
      live++; max = Math.max(max, live);
      await new Promise((r) => setTimeout(r, 5));
      live--;
      return x * 10;
    });
    expect(out).toEqual([10, 20, 30, 40, 50]);
    expect(max).toBe(2);
  });
  it('致命错误识别', () => {
    expect(isFatalLlmError(new Error('还没有设置大模型：请到「设置」里填写'))).toBe(true);
    expect(isFatalLlmError(new Error('大模型调用失败（401，密钥不对或已过期）'))).toBe(true);
    expect(isFatalLlmError(new Error('大模型调用失败（500，接口返回错误）'))).toBe(false);
  });
});
