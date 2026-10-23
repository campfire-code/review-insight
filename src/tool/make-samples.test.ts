// 生成 samples/ 里的示例文件（只在设置了 MAKE_SAMPLES=1 时运行）
import { it } from 'vitest';
import { writeFileSync } from 'node:fs';
import { writeXlsx } from './xlsx';

const BAD = [
  '快递太慢了，下单五天才到', '物流一直不更新，催了客服也没用', '杯盖拧紧了还是漏水，包里全湿了', '保温效果很差，早上倒的热水中午就凉了',
  '颜色和图片差很多，实物发灰', '页面写的 500ml，实际装不满 450', '客服回复很慢，问了半天才回一句', '外包装压坏了，杯身有个凹痕',
  '收到就有划痕，像是退回来的', '比别家贵了二十块，质量还不如', '用了一周杯底开始掉漆', '发货慢，说好48小时发结果拖了4天',
  '盒子破了，里面连个气泡膜都没有', '退换货流程太麻烦，客服推来推去', '有很重的塑料味，洗了好几遍还有',
];
const MID = ['还行吧，就是有点重', '保温一般，价格还算可以', '东西凑合，物流有点慢', '和描述基本一致，就是盖子有点紧'];
const GOOD = ['很好用，保温效果不错', '颜色好看，送人很合适', '第二次回购了，质量稳定', '物流很快，第二天就到了'];

it.runIf(process.env.MAKE_SAMPLES === '1')('生成示例文件', async () => {
  const rows = Array.from({ length: 60 }, (_, i) => {
    const kind = i % 5 < 2 ? 'bad' : i % 5 === 2 ? 'mid' : 'good';
    const text = kind === 'bad' ? BAD[i % BAD.length] : kind === 'mid' ? MID[i % MID.length] : GOOD[i % GOOD.length];
    const rating = kind === 'bad' ? 1 + (i % 2) : kind === 'mid' ? 3 : 5;
    return [`2026101${String(100000 + i)}`, '316不锈钢保温杯', rating, text, `2026-10-${String((i % 20) + 1).padStart(2, '0')}`];
  });
  writeFileSync('samples/评论导出.xlsx', await writeXlsx([{ name: '评论', header: ['订单号', '商品', '评分', '评论内容', '评价时间'], rows }]));
  const csv = ['买家评价,评价时间', '快递太慢了，下单五天才到,2026-10-01', '很好用，保温效果不错,2026-10-02', '杯盖漏水，包里全湿了,2026-10-02', '客服回复很慢,2026-10-03', '颜色和图片差很多,2026-10-04', '第二次回购了,2026-10-05'];
  writeFileSync('samples/评论-无评分列.csv', '﻿' + csv.join('\r\n') + '\r\n');
});
