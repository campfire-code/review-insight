// 截图与预览用的示例数据（不会进入正式运行）
import type { Row } from '../kit/runner';

const text1 = `412 条评论，分析 138 条，调用大模型 6 次
质量问题：49 条（36.6%）
物流快递：37 条（27.6%）
描述不符：21 条（15.7%）
包装破损：14 条（10.4%）
客服服务：11 条（8.2%）
其他：2 条（1.5%）
无具体问题：4 条
高频问题：杯盖漏水（26）、发货慢（19）、保温效果差（15）、颜色与图片不符（12）、外箱压坏（9）`;

const files: [string, number, Row['status'], Record<string, string | number>?, string?][] = [
  ['保温杯-十月评论导出.xlsx', 96_000, 'ok', { total: 412, analyzed: 138, calls: 6, top: '质量问题 36.6%' }, text1],
  ['帆布包-十月评论导出.xlsx', 58_000, 'ok', { total: 236, analyzed: 61, calls: 3, top: '描述不符 31.1%' }, '236 条评论，分析 61 条，调用大模型 3 次'],
  ['蓝牙耳机-评论.csv', 41_000, 'ok', { total: 180, analyzed: 52, calls: 3, top: '质量问题 40.4%' }, '180 条评论，分析 52 条，调用大模型 3 次'],
  ['香薰礼盒-评论.xlsx', 33_000, 'run'],
  ['退款原因.xls', 12_000, 'err', undefined, '旧版 .xls 格式不支持，请用 Excel 或 WPS 另存为 .xlsx 后再拖进来'],
];

export const rows = (state: string): Row[] =>
  state === 'empty' ? [] : files.map(([name, size, status, row, text], i) => {
    const st = state === 'done' && status === 'run' ? 'ok' : status;
    return {
      item: { id: `p${i}`, name, path: `/示例/${name}`, size },
      status: st,
      progress: st === 'run' ? 0.48 : st === 'ok' ? 1 : 0,
      error: st === 'err' ? text : undefined,
      result: st === 'ok' ? { summary: `分析 ${row?.analyzed ?? 40} 条`, row: row ?? { total: 120, analyzed: 40, calls: 2, top: '包装破损 30%' }, text: text ?? '' } : undefined,
    };
  });

export const params = { scope: 'bad_mid' };
