import type { Cell, CellValue, Workbook, Worksheet } from 'exceljs';
import { loadExcelFile, writeWorkbook } from '../../excel/workbook';

export const SALES_COLUMNS = [
  '売上日', '店舗コード', '店舗名', '商品コード', '商品名', 'カテゴリ', '販売数量', '単価', '売上金額', '支払方法',
] as const;

export const SALES_LIMITS = {
  files: 10,
  totalFileBytes: 50 * 1024 * 1024,
  rows: 100_000,
} as const;

export type SalesRecord = {
  salesDate: string;
  storeCode: string;
  storeName: string;
  productCode: string;
  productName: string;
  category: string;
  quantity: number;
  unitPrice: number;
  salesAmount: number;
  paymentMethod: string;
};

export type StoreSales = {
  storeCode: string;
  storeName: string;
  salesAmount: number;
  quantity: number;
  lines: number;
};

export type CategorySales = {
  category: string;
  salesAmount: number;
  quantity: number;
};

export type SalesSummary = {
  salesDate: string;
  records: SalesRecord[];
  stores: StoreSales[];
  categories: CategorySales[];
  totalSales: number;
  totalQuantity: number;
};

function fail(message: string): never {
  throw new Error(message);
}

function textValue(cell: Cell, field: string): string {
  const value = cell.value;
  if (value == null) return fail(`${field}が空欄です。店舗のExcelを確認してください。`);
  if (value instanceof Date) return fail(`${field}の形式が正しくありません。`);
  if (typeof value === 'object') return fail(`${field}に未対応の値があります。値に置き換えてください。`);
  const text = (cell.text || String(value)).trim();
  if (!text) return fail(`${field}が空欄です。店舗のExcelを確認してください。`);
  return text;
}

function numberValue(value: CellValue | undefined, field: string, row: number, allowZero = false): number {
  if (value == null || (typeof value === 'string' && !value.trim())) return fail(`${row}行目の「${field}」が空欄です。`);
  const parsed = typeof value === 'number' ? value : Number(String(value).trim().replaceAll(',', ''));
  if (!Number.isFinite(parsed) || parsed < 0 || (!allowZero && parsed === 0)) {
    return fail(`${row}行目の「${field}」は${allowZero ? '0以上の数値' : '1以上の数値'}で入力してください。`);
  }
  return parsed;
}

function dateValue(value: CellValue | undefined, row: number): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const day = String(value.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const dayNumber = Math.floor(value);
    const date = new Date(Date.UTC(1899, 11, 30 + dayNumber));
    return date.toISOString().slice(0, 10);
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    const iso = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (iso) {
      const date = new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
      if (date.getUTCFullYear() === Number(iso[1]) && date.getUTCMonth() === Number(iso[2]) - 1 && date.getUTCDate() === Number(iso[3])) return date.toISOString().slice(0, 10);
    }
  }
  return fail(`${row}行目の「売上日」が日付として読み取れません。`);
}

function blankRow(worksheet: Worksheet, rowNumber: number, width: number): boolean {
  for (let column = 1; column <= width; column += 1) {
    const value = worksheet.getRow(rowNumber).getCell(column).value;
    if (value != null && !(typeof value === 'string' && value.trim() === '')) return false;
  }
  return true;
}

function headerMap(worksheet: Worksheet): Map<(typeof SALES_COLUMNS)[number], number> {
  const header = worksheet.getRow(1);
  const found = new Map<string, number>();
  for (let column = 1; column <= Math.max(worksheet.columnCount, SALES_COLUMNS.length); column += 1) {
    const value = header.getCell(column).value;
    if (value == null || String(value).trim() === '') continue;
    const name = String(value).trim();
    if (found.has(name)) return fail(`見出し「${name}」が重複しています。列名を確認してください。`);
    found.set(name, column);
  }
  const missing = SALES_COLUMNS.filter((name) => !found.has(name));
  if (missing.length) return fail(`必要な列が見つかりません（${missing.join('、')}）。見出し行を確認してください。`);
  const extra = [...found.keys()].filter((name) => !SALES_COLUMNS.includes(name as (typeof SALES_COLUMNS)[number]));
  if (extra.length) return fail(`未対応の列「${extra[0]}」があります。集計対象の10項目だけのExcelを選んでください。`);
  return new Map(SALES_COLUMNS.map((name) => [name, found.get(name)!]));
}

function recordsFromSheet(worksheet: Worksheet, columns: Map<(typeof SALES_COLUMNS)[number], number>): SalesRecord[] {
  const records: SalesRecord[] = [];
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (blankRow(worksheet, rowNumber, worksheet.columnCount)) continue;
    const row = worksheet.getRow(rowNumber);
    const get = (name: (typeof SALES_COLUMNS)[number]) => row.getCell(columns.get(name)!);
    const quantity = numberValue(get('販売数量').value, '販売数量', rowNumber);
    if (!Number.isInteger(quantity)) return fail(`${rowNumber}行目の「販売数量」は整数で入力してください。`);
    records.push({
      salesDate: dateValue(get('売上日').value, rowNumber),
      storeCode: textValue(get('店舗コード'), '店舗コード'),
      storeName: textValue(get('店舗名'), '店舗名'),
      productCode: textValue(get('商品コード'), '商品コード'),
      productName: textValue(get('商品名'), '商品名'),
      category: textValue(get('カテゴリ'), 'カテゴリ'),
      quantity,
      unitPrice: numberValue(get('単価').value, '単価', rowNumber, true),
      salesAmount: numberValue(get('売上金額').value, '売上金額', rowNumber, true),
      paymentMethod: textValue(get('支払方法'), '支払方法'),
    });
  }
  if (!records.length) return fail(`「${worksheet.name}」に売上データがありません。`);
  return records;
}

function findSalesSheet(workbook: Workbook): { worksheet: Worksheet; columns: Map<(typeof SALES_COLUMNS)[number], number> } {
  let validationError = '';
  for (const worksheet of workbook.worksheets) {
    if (worksheet.rowCount === 0) continue;
    const firstRowValues = Array.from({ length: worksheet.columnCount }, (_, index) => String(worksheet.getRow(1).getCell(index + 1).value ?? '').trim());
    if (!firstRowValues.some((value) => SALES_COLUMNS.includes(value as (typeof SALES_COLUMNS)[number]))) continue;
    try {
      return { worksheet, columns: headerMap(worksheet) };
    } catch (error) {
      validationError = error instanceof Error ? error.message : '';
    }
  }
  if (validationError) return fail(validationError);
  return fail('売上データの見出し行が見つかりません。1行目に指定の10項目があるExcelを選んでください。');
}

export function summarizeSales(records: SalesRecord[]): SalesSummary {
  if (!records.length) return fail('売上データがありません。');
  const dates = new Set(records.map((record) => record.salesDate));
  if (dates.size !== 1) return fail('売上日が複数含まれています。同じ日の店舗ファイルをまとめてください。');

  const stores = new Map<string, StoreSales>();
  const categories = new Map<string, CategorySales>();
  let totalSales = 0;
  let totalQuantity = 0;
  for (const record of records) {
    const store = stores.get(record.storeCode) ?? { storeCode: record.storeCode, storeName: record.storeName, salesAmount: 0, quantity: 0, lines: 0 };
    if (store.storeName !== record.storeName) return fail(`店舗コード「${record.storeCode}」に複数の店舗名があります。`);
    store.salesAmount += record.salesAmount;
    store.quantity += record.quantity;
    store.lines += 1;
    stores.set(record.storeCode, store);

    const category = categories.get(record.category) ?? { category: record.category, salesAmount: 0, quantity: 0 };
    category.salesAmount += record.salesAmount;
    category.quantity += record.quantity;
    categories.set(record.category, category);
    totalSales += record.salesAmount;
    totalQuantity += record.quantity;
  }
  return {
    salesDate: [...dates][0],
    records,
    stores: [...stores.values()],
    categories: [...categories.values()].sort((left, right) => right.salesAmount - left.salesAmount),
    totalSales,
    totalQuantity,
  };
}

export async function aggregateSalesFiles(files: File[]): Promise<SalesSummary> {
  if (!files.length) return fail('店舗のExcelを選択してください。');
  if (files.length > SALES_LIMITS.files) return fail(`一度に選べるファイルは${SALES_LIMITS.files}個までです。`);
  if (files.reduce((sum, file) => sum + file.size, 0) > SALES_LIMITS.totalFileBytes) return fail('選択したファイルの合計は50 MiB以下にしてください。');

  const records: SalesRecord[] = [];
  const seenStores = new Map<string, string>();
  for (const file of files) {
    const workbook = await loadExcelFile(file);
    const { worksheet, columns } = findSalesSheet(workbook);
    const fileRecords = recordsFromSheet(worksheet, columns);
    const storeKeys = new Set(fileRecords.map((record) => `${record.storeCode}\u0000${record.storeName}`));
    if (storeKeys.size !== 1) return fail(`「${file.name}」には複数店舗のデータがあります。店舗ごとにファイルを分けてください。`);
    const [{ storeCode, storeName }] = fileRecords;
    const priorName = seenStores.get(storeCode);
    if (priorName) return fail(`「${storeName}」のファイルが複数選ばれています。店舗ごとに1ファイルを選んでください。`);
    seenStores.set(storeCode, storeName);
    records.push(...fileRecords);
    if (records.length > SALES_LIMITS.rows) return fail(`集計行数は${SALES_LIMITS.rows.toLocaleString()}行以下にしてください。`);
  }
  return summarizeSales(records);
}

function styleSheet(worksheet: Worksheet, widths: number[]) {
  worksheet.columns = widths.map((width) => ({ width }));
  const header = worksheet.getRow(1);
  header.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9A512D' } };
  header.height = 24;
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  worksheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, worksheet.rowCount), column: widths.length } };
}

function moneyFormat(worksheet: Worksheet, column: number) {
  worksheet.getColumn(column).numFmt = '#,##0"円"';
}

export async function createSalesWorkbook(summary: SalesSummary): Promise<ArrayBuffer> {
  const { default: ExcelJS } = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'まちのパン工房 集計デモ';
  workbook.subject = `${summary.salesDate} 店舗別売上集計`;

  const detail = workbook.addWorksheet('全店舗統合データ');
  detail.addRow([...SALES_COLUMNS]);
  summary.records.forEach((record) => detail.addRow([
    record.salesDate, record.storeCode, record.storeName, record.productCode, record.productName,
    record.category, record.quantity, record.unitPrice, record.salesAmount, record.paymentMethod,
  ]));
  styleSheet(detail, [14, 15, 16, 14, 23, 17, 13, 13, 14, 14]);
  detail.getColumn(1).numFmt = 'yyyy-mm-dd';
  moneyFormat(detail, 8);
  moneyFormat(detail, 9);

  const byStore = workbook.addWorksheet('店舗別集計');
  byStore.addRow(['店舗コード', '店舗名', '売上日', '売上金額合計', '販売数量合計', '商品行数']);
  summary.stores.forEach((store) => byStore.addRow([store.storeCode, store.storeName, summary.salesDate, store.salesAmount, store.quantity, store.lines]));
  byStore.addRow(['', '全店舗合計', summary.salesDate, summary.totalSales, summary.totalQuantity, summary.records.length]);
  styleSheet(byStore, [16, 20, 16, 20, 18, 14]);
  moneyFormat(byStore, 4);
  byStore.getRow(byStore.rowCount).font = { bold: true, color: { argb: 'FF71391F' } };
  byStore.getRow(byStore.rowCount).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5EADD' } };

  const byCategory = workbook.addWorksheet('カテゴリ別集計');
  byCategory.addRow(['カテゴリ', '売上金額合計', '販売数量合計', '構成比']);
  summary.categories.forEach((category) => byCategory.addRow([category.category, category.salesAmount, category.quantity, category.salesAmount / summary.totalSales]));
  styleSheet(byCategory, [22, 20, 18, 14]);
  moneyFormat(byCategory, 2);
  byCategory.getColumn(4).numFmt = '0.0%';

  return writeWorkbook(workbook);
}
