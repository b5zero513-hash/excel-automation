import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import {
  WorkbookInputError,
  getHeaderOptions,
  loadWorkbook,
  previewGroups,
  splitWorkbook,
  writeWorkbook,
} from './workbook';

async function fixtureWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('日本語の元データ');
  sheet.addRow(['受付番号', '部門', '金額']);
  sheet.addRow(['A-001', '東京/営業', 1200]);
  sheet.addRow(['A-002', '大阪営業', 800]);
  sheet.addRow(['A-003', '東京/営業', 300]);
  sheet.addRow(['A-004', null, 99]);
  sheet.addRow([null, null, null]);
  sheet.getRow(1).font = { bold: true };
  sheet.getColumn(1).width = 21;
  workbook.addWorksheet('説明').addRow(['元の別シートも残す']);
  return workbook;
}

async function serializedFixture() {
  return writeWorkbook(await fixtureWorkbook());
}

describe('xlsx splitting', () => {
  it('TC-FILE-001 / TC-SPLIT-001 / TC-DATA-001 / TC-EMPTY-001: Japanese workbook reads, splits, writes, and reopens', async () => {
    const original = await serializedFixture();
    const workbook = await loadWorkbook(original);
    const worksheet = workbook.getWorksheet('日本語の元データ');

    expect(worksheet).toBeDefined();
    expect(getHeaderOptions(worksheet!, 1).map(({ label }) => label)).toEqual(['受付番号', '部門', '金額']);
    expect(previewGroups(worksheet!, 1, 2)).toMatchObject({ processedRows: 4, skippedBlankRows: 1, splitSheets: [
      { label: '東京/営業', count: 2, sheetName: '東京_営業' },
      { label: '大阪営業', count: 1, sheetName: '大阪営業' },
      { label: '空欄', count: 1, sheetName: '空欄' },
    ] });

    const { workbook: split } = splitWorkbook(workbook, '日本語の元データ', 1, 2);
    const reopened = await loadWorkbook(await writeWorkbook(split));
    expect(reopened.getWorksheet('説明')?.getCell('A1').value).toBe('元の別シートも残す');
    expect(reopened.getWorksheet('日本語の元データ')?.rowCount).toBe(6);
    expect(reopened.getWorksheet('東京_営業')?.getCell('A2').value).toBe('A-001');
    expect(reopened.getWorksheet('東京_営業')?.getCell('C3').value).toBe(300);
    expect(reopened.getWorksheet('東京_営業')?.getColumn(1).width).toBe(21);
    expect(reopened.getWorksheet('東京_営業')?.getCell('A1').font.bold).toBe(true);
    expect(reopened.getWorksheet('大阪営業')?.getCell('C2').value).toBe(800);
    expect(reopened.getWorksheet('空欄')?.getCell('A2').value).toBe('A-004');
  });

  it('TC-NAME-001: sanitizes invalid, long, reserved, and colliding worksheet names', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('元');
    sheet.addRow(['部門']);
    sheet.addRow(['東京/営業']);
    sheet.addRow(['東京_営業']);
    sheet.addRow(['x'.repeat(40)]);
    sheet.addRow(['History']);

    const result = previewGroups(sheet, 1, 1);
    const names = result.splitSheets.map(({ sheetName }) => sheetName);
    expect(names).toEqual(['東京_営業', '東京_営業 (2)', 'x'.repeat(31), 'History_']);
    expect(names.every((name) => Array.from(name).length <= 31)).toBe(true);
    expect(new Set(names.map((name) => name.toLowerCase())).size).toBe(names.length);
  });

  it('TC-SAFETY-001: refuses formula cells without writing an output', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('計算');
    sheet.addRow(['区分', '値']);
    sheet.addRow(['東京', { formula: '1+1', result: 2 }]);
    const bytes = await writeWorkbook(workbook);

    await expect(loadWorkbook(bytes)).rejects.toBeInstanceOf(WorkbookInputError);
  });

  it('TC-SAFETY-001: refuses files without a valid xlsx archive', async () => {
    await expect(loadWorkbook(new Uint8Array([0x50, 0x4b, 0x03, 0x04])))
      .rejects.toBeInstanceOf(WorkbookInputError);
  });
});
