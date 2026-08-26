import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const PUBLIC_BASE_PATH = process.env.NEXT_PUBLIC_BASE_PATH ?? '';

type ExportSlot = {
  color: 'blue' | 'red' | 'orange';
  number: number;
  name: string;
  c: boolean;
  h: boolean;
  e: boolean;
  d: boolean;
  v: boolean;
};

const COLUMN_MAP = {
  red: { name: 'B', c: 'C', h: 'D', e: 'E', d: 'F', v: 'G' },
  orange: { name: 'J', c: 'K', h: 'L', e: 'M', d: 'N', v: 'O' },
  blue: { name: 'R', c: 'S', h: 'T', e: 'U', d: 'V', v: 'W' },
} as const;

const COLOR_LABEL = {
  blue: '藍色',
  red: '紅色',
  orange: '橙色',
} as const;

function escapeXml(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

export function replaceCell(xml: string, reference: string, value: string) {
  const pattern = new RegExp(
    `<((?:[A-Za-z_][\\w.-]*:)?c)\\b([^>]*\\br="${reference}"[^>]*?)(?:\\/>|>[\\s\\S]*?<\\/\\1>)`,
  );
  const match = xml.match(pattern);
  const current = match?.[0];
  const cellTag = match?.[1];
  if (!current || !cellTag) throw new Error(`找不到 Excel 儲存格 ${reference}`);
  const namespace = cellTag.slice(0, -1);
  const style = current.match(/\ss="(\d+)"/)?.[1];
  const styleAttribute = style ? ` s="${style}"` : '';
  const next = value
    ? `<${cellTag} r="${reference}"${styleAttribute} t="str"><${namespace}v>${escapeXml(value)}</${namespace}v></${cellTag}>`
    : `<${cellTag} r="${reference}"${styleAttribute}/>`;
  return xml.replace(pattern, next);
}

export async function exportExcel(slots: ExportSlot[]) {
  const response = await fetch(`${PUBLIC_BASE_PATH}/ot-case-file-index-template.xlsx`);
  if (!response.ok) throw new Error('未能讀取 Excel 範本');

  const files = unzipSync(new Uint8Array(await response.arrayBuffer()));
  const sheetKey = 'xl/worksheets/sheet1.xml';
  const searchSheetKey = 'xl/worksheets/sheet2.xml';
  let sheetXml = strFromU8(files[sheetKey]);
  let searchSheetXml = strFromU8(files[searchSheetKey]);

  for (const slot of slots) {
    const columns = COLUMN_MAP[slot.color];
    const row = slot.number + 3;
    sheetXml = replaceCell(sheetXml, `${columns.name}${row}`, slot.name);
    sheetXml = replaceCell(sheetXml, `${columns.c}${row}`, slot.c ? '√' : '');
    sheetXml = replaceCell(sheetXml, `${columns.h}${row}`, slot.h ? '√' : '');
    sheetXml = replaceCell(sheetXml, `${columns.e}${row}`, slot.e ? '√' : '');
    sheetXml = replaceCell(sheetXml, `${columns.d}${row}`, slot.d ? '√' : '');
    sheetXml = replaceCell(sheetXml, `${columns.v}${row}`, slot.v ? '√' : '');
  }

  const searchableSlots = slots.filter((slot) => slot.name.trim());
  for (let index = 0; index < 93; index += 1) {
    const row = index + 4;
    const slot = searchableSlots[index];
    searchSheetXml = replaceCell(searchSheetXml, `H${row}`, slot?.name.trim() ?? '');
    searchSheetXml = replaceCell(searchSheetXml, `I${row}`, slot ? COLOR_LABEL[slot.color] : '');
    searchSheetXml = replaceCell(searchSheetXml, `J${row}`, slot ? String(slot.number) : '');
  }

  files[sheetKey] = strToU8(sheetXml);
  files[searchSheetKey] = strToU8(searchSheetXml);

  const output = zipSync(files, { level: 6 });
  const blob = new Blob([output.buffer as ArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const date = new Intl.DateTimeFormat('en-CA', {
    year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Hong_Kong',
  }).format(new Date());
  const link = document.createElement('a');
  link.href = url;
  link.download = `LPY職業治療_個案記錄索引_${date}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
