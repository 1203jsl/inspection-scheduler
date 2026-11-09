/** samples/ 폴더에 시험용 엑셀을 만듭니다: npm run samples */
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { buildTemplate } from '../src/main/core/excel/template';
import { DEFAULT_INSPECTION_ITEMS, TEMPLATE_V1 } from '../src/main/core/schema';
import { MIXED_ROWS, NORMAL_ROWS, REVISION_ROWS, vendorReply } from '../tests/fixtures';

const out = path.resolve(__dirname, '../samples');
await fs.mkdir(out, { recursive: true });
const files: [string, Buffer][] = [
  ['0_빈양식_v1.xlsx', await buildTemplate(TEMPLATE_V1, DEFAULT_INSPECTION_ITEMS)],
  ['1_정상_신규3건.xlsx', await vendorReply(NORMAL_ROWS)],
  ['2_오류섞임_11행.xlsx', await vendorReply(MIXED_ROWS)],
  ['3_수정본_1번가져온뒤.xlsx', await vendorReply(REVISION_ROWS)],
];
for (const [n, b] of files) await fs.writeFile(path.join(out, n), b);
console.log('만든 파일:', files.map((f) => f[0]).join(', '));
