import type { CellValue, Row, Workbook, Worksheet } from 'exceljs';

export const LIMITS = {
  fileBytes: 10 * 1024 * 1024,
  expandedBytes: 100 * 1024 * 1024,
  archiveEntries: 2_000,
  worksheets: 100,
  rows: 100_000,
  columns: 200,
  splitSheets: 50,
} as const;

export class WorkbookInputError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'WorkbookInputError';
  }
}

export type WorksheetInfo = {
  name: string;
  id: number;
  rowCount: number;
  columnCount: number;
};

export type GroupPreview = {
  key: string;
  label: string;
  sheetName: string;
  count: number;
};

export type SplitSummary = {
  sourceSheet: string;
  splitSheets: GroupPreview[];
  processedRows: number;
  skippedBlankRows: number;
};

type Group = {
  key: string;
  label: string;
  rows: Row[];
  sheetName: string;
};

const forbiddenPackageParts = [
  /^xl\/vbaProject\.bin$/i,
  /^xl\/externalLinks\//i,
  /^xl\/(drawings|charts|pivotTables|pivotCache|media)\//i,
];

function inputError(message: string): never {
  throw new WorkbookInputError(message);
}

function readU16(view: DataView, offset: number): number {
  if (offset < 0 || offset + 2 > view.byteLength) inputError('Excelファイルの構造を確認できません。別の .xlsx ファイルをお試しください。');
  return view.getUint16(offset, true);
}

function readU32(view: DataView, offset: number): number {
  if (offset < 0 || offset + 4 > view.byteLength) inputError('Excelファイルの構造を確認できません。別の .xlsx ファイルをお試しください。');
  return view.getUint32(offset, true);
}

function decodeEntryName(bytes: Uint8Array): string {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder().decode(bytes);
  }
}

/** ZIPの中央ディレクトリを確認してからExcelJSへ渡す。 */
export function validateXlsxArchive(input: ArrayBuffer | Uint8Array): void {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (bytes.byteLength < 22) inputError('ファイルが小さすぎるか、.xlsx 形式ではありません。');
  if (bytes.byteLength > LIMITS.fileBytes) inputError('ファイルは10 MiB以下にしてください。');

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const searchStart = Math.max(0, view.byteLength - 65_557);
  let endRecord = -1;
  for (let offset = view.byteLength - 22; offset >= searchStart; offset -= 1) {
    if (readU32(view, offset) === 0x06054b50 && offset + 22 + readU16(view, offset + 20) === view.byteLength) {
      endRecord = offset;
      break;
    }
  }
  if (endRecord < 0) inputError('有効な .xlsx のZIP構造が見つかりません。');

  const diskNumber = readU16(view, endRecord + 4);
  const directoryDisk = readU16(view, endRecord + 6);
  const diskEntries = readU16(view, endRecord + 8);
  const entryCount = readU16(view, endRecord + 10);
  const directoryBytes = readU32(view, endRecord + 12);
  const directoryOffset = readU32(view, endRecord + 16);
  const commentLength = readU16(view, endRecord + 20);

  if (endRecord + 22 + commentLength !== view.byteLength) inputError('ExcelファイルのZIP情報が壊れています。');
  if (diskNumber !== 0 || directoryDisk !== 0 || diskEntries !== entryCount) inputError('分割されたZIP形式には対応していません。');
  if (entryCount === 0xffff || directoryBytes === 0xffffffff || directoryOffset === 0xffffffff) inputError('非常に大きいZIP64形式には対応していません。');
  if (entryCount === 0 || entryCount > LIMITS.archiveEntries) inputError('ファイル内の項目数が上限を超えています。');
  if (directoryOffset + directoryBytes > endRecord) inputError('ExcelファイルのZIP情報が壊れています。');

  let offset = directoryOffset;
  let expandedBytes = 0;
  let hasContentTypes = false;
  let hasWorkbook = false;
  for (let index = 0; index < entryCount; index += 1) {
    if (readU32(view, offset) !== 0x02014b50 || offset + 46 > directoryOffset + directoryBytes) {
      inputError('ExcelファイルのZIP情報が壊れています。');
    }
    const flags = readU16(view, offset + 8);
    const compression = readU16(view, offset + 10);
    const compressedSize = readU32(view, offset + 20);
    const uncompressedSize = readU32(view, offset + 24);
    const nameLength = readU16(view, offset + 28);
    const extraLength = readU16(view, offset + 30);
    const entryCommentLength = readU16(view, offset + 32);
    const recordBytes = 46 + nameLength + extraLength + entryCommentLength;
    if (offset + recordBytes > directoryOffset + directoryBytes) inputError('ExcelファイルのZIP情報が壊れています。');

    if ((flags & 1) !== 0) inputError('暗号化されたExcelファイルには対応していません。');
    if (compression !== 0 && compression !== 8) inputError('この圧縮形式のExcelファイルには対応していません。');
    if (uncompressedSize === 0xffffffff || compressedSize === 0xffffffff) inputError('ZIP64形式には対応していません。');
    if (uncompressedSize > LIMITS.expandedBytes) inputError('展開後のファイルサイズが上限を超えています。');

    expandedBytes += uncompressedSize;
    if (expandedBytes > LIMITS.expandedBytes) inputError('展開後の合計サイズが100 MiBを超えています。');

    const name = decodeEntryName(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (name === '[Content_Types].xml') hasContentTypes = true;
    if (name.toLowerCase() === 'xl/workbook.xml') hasWorkbook = true;
    if (forbiddenPackageParts.some((pattern) => pattern.test(name))) {
      inputError('マクロ、図形、グラフ、ピボット、外部リンクを含むブックには対応していません。元ファイルは変更されていません。');
    }
    offset += recordBytes;
  }

  if (!hasContentTypes || !hasWorkbook) inputError('Excelの .xlsx ブックとして認識できません。');
}

export async function loadWorkbook(input: ArrayBuffer | Uint8Array): Promise<Workbook> {
  validateXlsxArchive(input);
  const { default: ExcelJS } = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  try {
    // ExcelJS's browser API accepts ArrayBuffer/Uint8Array at runtime; its declarations still expose Node Buffer.
    await workbook.xlsx.load(input as unknown as Parameters<typeof workbook.xlsx.load>[0]);
  } catch {
    inputError('Excelファイルを読み込めません。壊れているか、未対応の形式の可能性があります。');
  }
  validateSupportedWorkbook(workbook);
  return workbook;
}

export async function loadExcelFile(file: File): Promise<Workbook> {
  if (!file.name.toLowerCase().endsWith('.xlsx')) inputError('`.xlsx` ファイルを選択してください。');
  if (file.size > LIMITS.fileBytes) inputError('ファイルは10 MiB以下にしてください。');
  return loadWorkbook(await file.arrayBuffer());
}

export function listWorksheets(workbook: Workbook): WorksheetInfo[] {
  return workbook.worksheets.map((worksheet) => ({
    name: worksheet.name,
    id: worksheet.id,
    rowCount: worksheet.rowCount,
    columnCount: worksheet.columnCount,
  }));
}

export function displayCellValue(value: CellValue | undefined): string {
  if (value == null || (typeof value === 'string' && value.trim() === '')) return '(空欄)';
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? '(日付)' : value.toISOString().slice(0, 10);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'object') return '(未対応の値)';
  return String(value);
}

export function getHeaderOptions(worksheet: Worksheet, headerRowNumber: number) {
  if (!Number.isInteger(headerRowNumber) || headerRowNumber < 1 || headerRowNumber > worksheet.rowCount) {
    inputError('見出し行はシート内の行番号を指定してください。');
  }
  return Array.from({ length: worksheet.columnCount }, (_, index) => {
    const columnNumber = index + 1;
    const value = worksheet.getRow(headerRowNumber).getCell(columnNumber).value;
    const label = value == null || displayCellValue(value) === '(空欄)' ? `列 ${columnNumber}` : displayCellValue(value);
    return { columnNumber, label, optionLabel: `${label}（${columnNumber}列目）` };
  });
}

function isEmptyRow(row: Row, width: number): boolean {
  for (let column = 1; column <= width; column += 1) {
    const value = row.getCell(column).value;
    if (value != null && !(typeof value === 'string' && value.trim() === '')) return false;
  }
  return true;
}

function groupValue(value: CellValue | undefined): { key: string; label: string } {
  if (value == null || (typeof value === 'string' && value.trim() === '')) return { key: 'empty:', label: '空欄' };
  if (value instanceof Date) return { key: `date:${value.getTime()}`, label: displayCellValue(value) };
  if (typeof value === 'object') inputError('文字列・数値・日付・真偽値以外のセル値には対応していません。');
  return { key: `${typeof value}:${String(value)}`, label: String(value) };
}

function takeCharacters(value: string, max: number): string {
  return Array.from(value).slice(0, max).join('');
}

function createUniqueSheetName(raw: string, existingNames: Set<string>): string {
  let base = raw
    .replace(/[\\/?*:[\]\u0000-\u001f]/g, '_')
    .replace(/^'+|'+$/g, '')
    .trim();
  if (!base || base.toLocaleLowerCase('en-US') === 'history') base = base ? `${base}_` : '空欄';
  base = takeCharacters(base, 31);

  let name = base;
  let suffixNumber = 2;
  while (existingNames.has(name.toLocaleLowerCase('en-US'))) {
    const suffix = ` (${suffixNumber})`;
    suffixNumber += 1;
    name = `${takeCharacters(base, 31 - Array.from(suffix).length)}${suffix}`;
  }
  existingNames.add(name.toLocaleLowerCase('en-US'));
  return name;
}

function makeGroups(worksheet: Worksheet, headerRowNumber: number, columnNumber: number): { groups: Group[]; skippedBlankRows: number } {
  const width = worksheet.columnCount;
  const headerRow = worksheet.getRow(headerRowNumber);
  if (isEmptyRow(headerRow, width)) inputError('見出し行に列名やデータがありません。別の見出し行を選択してください.');
  if (!Number.isInteger(columnNumber) || columnNumber < 1 || columnNumber > width) inputError('仕分け列を選択してください。');

  const grouped = new Map<string, { label: string; rows: Row[] }>();
  let skippedBlankRows = 0;
  for (let rowNumber = headerRowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    if (isEmptyRow(row, width)) {
      skippedBlankRows += 1;
      continue;
    }
    const { key, label } = groupValue(row.getCell(columnNumber).value);
    const group = grouped.get(key) ?? { label, rows: [] };
    group.rows.push(row);
    grouped.set(key, group);
  }
  if (grouped.size === 0) inputError('仕分けるデータ行がありません。見出し行を確認してください。');
  if (grouped.size > LIMITS.splitSheets) inputError(`仕分け先が${LIMITS.splitSheets}シートを超えています。列を確認するか、データを分けてお試しください。`);

  const usedNames = new Set(worksheet.workbook.worksheets.map((item) => item.name.toLocaleLowerCase('en-US')));
  const groups = [...grouped.entries()].map(([key, group]) => ({
    key,
    label: group.label,
    rows: group.rows,
    sheetName: createUniqueSheetName(group.label, usedNames),
  }));
  return { groups, skippedBlankRows };
}

export function previewGroups(worksheet: Worksheet, headerRowNumber: number, columnNumber: number): SplitSummary {
  const { groups, skippedBlankRows } = makeGroups(worksheet, headerRowNumber, columnNumber);
  return {
    sourceSheet: worksheet.name,
    splitSheets: groups.map(({ key, label, sheetName, rows }) => ({ key, label, sheetName, count: rows.length })),
    processedRows: groups.reduce((sum, group) => sum + group.rows.length, 0),
    skippedBlankRows,
  };
}

function copyRow(source: Row, target: Row, width: number) {
  for (let column = 1; column <= width; column += 1) {
    const sourceCell = source.getCell(column);
    const targetCell = target.getCell(column);
    targetCell.value = sourceCell.value;
    if (sourceCell.style && Object.keys(sourceCell.style).length > 0) targetCell.style = { ...sourceCell.style };
  }
  if (source.height) target.height = source.height;
}

export function splitWorkbook(
  workbook: Workbook,
  worksheetName: string,
  headerRowNumber: number,
  columnNumber: number,
): { workbook: Workbook; summary: SplitSummary } {
  validateSupportedWorkbook(workbook);
  const worksheet = workbook.getWorksheet(worksheetName);
  if (!worksheet) inputError('選択したシートが見つかりません。ファイルを読み込み直してください。');
  const { groups, skippedBlankRows } = makeGroups(worksheet, headerRowNumber, columnNumber);
  const width = worksheet.columnCount;

  for (const group of groups) {
    const output = workbook.addWorksheet(group.sheetName);
    for (let column = 1; column <= width; column += 1) {
      const sourceColumn = worksheet.getColumn(column);
      if (sourceColumn.width) output.getColumn(column).width = sourceColumn.width;
      if (sourceColumn.hidden) output.getColumn(column).hidden = true;
    }
    copyRow(worksheet.getRow(headerRowNumber), output.addRow([]), width);
    group.rows.forEach((sourceRow) => copyRow(sourceRow, output.addRow([]), width));
    output.views = [{ state: 'frozen', ySplit: 1 }];
    output.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: group.rows.length + 1, column: width },
    };
  }

  return {
    workbook,
    summary: {
      sourceSheet: worksheet.name,
      splitSheets: groups.map(({ key, label, sheetName, rows }) => ({ key, label, sheetName, count: rows.length })),
      processedRows: groups.reduce((sum, group) => sum + group.rows.length, 0),
      skippedBlankRows,
    },
  };
}

export async function writeWorkbook(workbook: Workbook): Promise<ArrayBuffer> {
  const output = await workbook.xlsx.writeBuffer();
  return new Uint8Array(output).buffer;
}

function validateSupportedWorkbook(workbook: Workbook): void {
  if (workbook.worksheets.length === 0) inputError('ブックにシートがありません。');
  if (workbook.worksheets.length > LIMITS.worksheets) inputError(`シート数は${LIMITS.worksheets}以下にしてください。`);
  if (workbook.model.media?.length) inputError('画像などのメディアを含むブックには対応していません。元ファイルは変更されていません。');

  for (const worksheet of workbook.worksheets) {
    if (worksheet.rowCount > LIMITS.rows) inputError(`「${worksheet.name}」が${LIMITS.rows.toLocaleString()}行を超えています。`);
    if (worksheet.columnCount > LIMITS.columns) inputError(`「${worksheet.name}」が${LIMITS.columns}列を超えています。`);
    if (worksheet.model.merges.length > 0) inputError(`「${worksheet.name}」に結合セルがあります。結合を解除してからお試しください。元ファイルは変更されていません。`);
    const sheetProtection = (worksheet as Worksheet & { sheetProtection?: unknown }).sheetProtection;
    if (sheetProtection) inputError(`「${worksheet.name}」は保護されています。保護を解除してからお試しください。`);

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const value = cell.value;
        if (value && typeof value === 'object' && !(value instanceof Date)) {
          inputError(`「${worksheet.name}」の${cell.address}に数式または未対応のセル値があります。値に置き換えてからお試しください。元ファイルは変更されていません。`);
        }
      });
    });
  }
}
