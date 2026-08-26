import fs from 'node:fs/promises';
import { strFromU8, unzipSync } from 'fflate';

const pageSource = await fs.readFile(new URL('../app/page.tsx', import.meta.url), 'utf8');
const emptyInitializerCount = (
  pageSource.match(/Array\.from\(\{ length: SLOT_COUNT \}, \(\) => ''\)/g) ?? []
).length;

if (emptyInitializerCount !== 3) {
  throw new Error('私隱檢查失敗：三組預設名單並非全部空白。');
}

const templateBytes = new Uint8Array(await fs.readFile(
  new URL('../public/ot-case-file-index-template.xlsx', import.meta.url),
));
const files = unzipSync(templateBytes);
const indexXml = strFromU8(files['xl/worksheets/sheet1.xml']);
const lookupXml = strFromU8(files['xl/worksheets/sheet2.xml']);
const isEmptyCell = (xml, ref) => new RegExp(`<x:c\\b(?=[^>]*\\br="${ref}")[^>]*/>`).test(xml);

const indexColumns = ['B', 'C', 'D', 'E', 'F', 'G', 'J', 'K', 'L', 'M', 'N', 'O', 'R', 'S', 'T', 'U', 'V', 'W'];
for (const column of indexColumns) {
  for (let row = 4; row <= 34; row += 1) {
    const ref = `${column}${row}`;
    if (!isEmptyCell(indexXml, ref)) {
      throw new Error(`私隱檢查失敗：Excel 範本 ${ref} 並非空白。`);
    }
  }
}

for (const column of ['H', 'I', 'J']) {
  for (let row = 4; row <= 96; row += 1) {
    const ref = `${column}${row}`;
    if (!isEmptyCell(lookupXml, ref)) {
      throw new Error(`私隱檢查失敗：Excel 查詢清單 ${ref} 並非空白。`);
    }
  }
}

console.log('私隱檢查通過：程式預設名單及 Excel 範本均為空白。');
