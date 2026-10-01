import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';

const source = new ExcelJS.Workbook();
const original = source.addWorksheet('元データ_日本語');
original.addRow(['受付番号', '担当部署', '金額']);
original.addRow(['A-001', '東京/営業', 1200]);
original.addRow(['A-002', '大阪営業', 800]);
original.addRow(['A-003', '東京/営業', 300]);
original.addRow(['A-004', null, 99]);
source.addWorksheet('説明').addRow(['このシートも保持']);

const inputBytes = await source.xlsx.writeBuffer();
const loaded = new ExcelJS.Workbook();
await loaded.xlsx.load(new Uint8Array(inputBytes));

const sheet = loaded.getWorksheet('元データ_日本語');
assert.ok(sheet, '日本語のシート名を読み込める');
assert.deepEqual(sheet.getRow(1).values.slice(1), ['受付番号', '担当部署', '金額']);

const groups = new Map();
for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
  const row = sheet.getRow(rowNumber);
  const value = row.getCell(2).value;
  const key = value == null || String(value).trim() === '' ? '(空欄)' : String(value);
  const group = groups.get(key) ?? [];
  group.push(row.values.slice(1));
  groups.set(key, group);
}

const invalidSheetCharacters = /[\\/?*:[\]]/g;
const usedNames = new Set(loaded.worksheets.map((worksheet) => worksheet.name.toLocaleLowerCase()));
function uniqueSheetName(rawName) {
  const base = (rawName || '空欄').replace(invalidSheetCharacters, '_').replace(/^'+|'+$/g, '').slice(0, 31) || '空欄';
  let name = base;
  let suffix = 2;
  while (usedNames.has(name.toLocaleLowerCase())) {
    const tail = ` (${suffix++})`;
    name = `${base.slice(0, 31 - tail.length)}${tail}`;
  }
  usedNames.add(name.toLocaleLowerCase());
  return name;
}

for (const [key, rows] of groups) {
  const outputSheet = loaded.addWorksheet(uniqueSheetName(key === '(空欄)' ? '空欄' : key));
  outputSheet.addRow(sheet.getRow(1).values.slice(1));
  rows.forEach((values) => outputSheet.addRow(values));
}

const outputBytes = await loaded.xlsx.writeBuffer();
const reopened = new ExcelJS.Workbook();
await reopened.xlsx.load(new Uint8Array(outputBytes));

assert.ok(reopened.getWorksheet('説明'), '元の別シートを保持できる');
assert.equal(reopened.getWorksheet('東京_営業').rowCount, 3);
assert.equal(reopened.getWorksheet('大阪営業').rowCount, 2);
assert.equal(reopened.getWorksheet('空欄').rowCount, 2);
assert.deepEqual(reopened.getWorksheet('東京_営業').getRow(2).values.slice(1), ['A-001', '東京/営業', 1200]);
assert.equal([...groups.values()].reduce((sum, rows) => sum + rows.length, 0), 4);

console.log('ExcelJS smoke check passed: Japanese names, XLSX load, split, multi-sheet write, and reload.');
