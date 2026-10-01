import { describe, expect, it } from 'vitest';
import { aggregateSalesFiles, createSalesWorkbook, SALES_COLUMNS, summarizeSales } from './sales';
import { loadWorkbook } from '../../excel/workbook';

const stores = [
  ['OS-KB', '京橋店', 1.12], ['OS-TM', '天満店', 0.96], ['OS-MY', '都島店', 1.03],
  ['OS-MG', '守口店', 0.91], ['OS-NY', '寝屋川店', 1.08],
] as const;
const products = [
  ['B001', '湯種食パン', 'パン', 360, 30], ['B002', 'クロワッサン', 'パン', 230, 52],
  ['B003', '塩バターロール', 'パン', 190, 48], ['B004', 'つぶあんぱん', 'パン', 210, 35],
  ['B005', 'メロンパン', 'パン', 240, 32], ['S001', 'たまごサンド', 'サンドイッチ', 430, 23],
  ['S002', 'ミックスサンド', 'サンドイッチ', 490, 17], ['S003', '照り焼きチキンサンド', 'サンドイッチ', 540, 13],
  ['C001', 'カスタードプリン', '菓子', 330, 18], ['C002', '季節の焼き菓子', '菓子', 280, 25],
  ['D001', 'ブレンドコーヒー', 'ドリンク', 250, 29], ['D002', 'カフェオレ', 'ドリンク', 300, 21],
] as const;

async function sampleFiles() {
  const { default: ExcelJS } = await import('exceljs');
  return Promise.all(stores.map(async ([code, name, factor], storeIndex) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('日次売上');
    worksheet.addRow(['売上日', '店舗コード', '店舗名', '商品コード', '商品名', 'カテゴリ', '販売数量', '単価', '売上金額', '支払方法']);
    products.forEach(([productCode, productName, category, unitPrice, baseQuantity], productIndex) => {
      const quantity = Math.max(5, Math.round(baseQuantity * factor) + ((productIndex * 7 + storeIndex * 11) % 9) - 4);
      worksheet.addRow(['2026-09-30', code, name, productCode, productName, category, quantity, unitPrice, quantity * unitPrice, '現金']);
    });
    const bytes = await workbook.xlsx.writeBuffer();
    return new File([bytes], `${name}_売上.xlsx`, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }));
}

describe('店舗売上の集計', () => {
  it('5店舗のサンプルを統合し、店舗とカテゴリの合計が全体売上に一致する', async () => {
    const summary = await aggregateSalesFiles(await sampleFiles());
    expect(summary.stores).toHaveLength(5);
    expect(summary.records).toHaveLength(60);
    expect(summary.categories.map(({ category }) => category).sort()).toEqual(['サンドイッチ', 'ドリンク', 'パン', '菓子']);
    expect(summary.stores.reduce((sum, store) => sum + store.salesAmount, 0)).toBe(summary.totalSales);
    expect(summary.categories.reduce((sum, category) => sum + category.salesAmount, 0)).toBe(summary.totalSales);
    expect(summary.totalSales).toBeGreaterThan(400_000);
    expect(summary.totalSales).toBeLessThan(700_000);
  });

  it('出力ブックは指定された3シートを持ち、統合行と店舗別合計を含む', async () => {
    const summary = await aggregateSalesFiles(await sampleFiles());
    const output = await createSalesWorkbook(summary);
    const workbook = await loadWorkbook(output);

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(['全店舗統合データ', '店舗別集計', 'カテゴリ別集計']);
    expect(workbook.getWorksheet('全店舗統合データ')?.rowCount).toBe(61);
    expect(workbook.getWorksheet('店舗別集計')?.getRow(7).getCell(4).value).toBe(summary.totalSales);
  });

  it('異なる売上日のデータを同じ日に混ぜない', () => {
    const record = {
      salesDate: '2026-09-30', storeCode: 'A', storeName: '京橋店', productCode: 'P1', productName: 'パン',
      category: 'パン', quantity: 1, unitPrice: 200, salesAmount: 200, paymentMethod: '現金',
    };
    expect(() => summarizeSales([record, { ...record, salesDate: '2026-10-01', storeCode: 'B', storeName: '天満店' }]))
      .toThrow('売上日が複数含まれています');
  });

  it('重複した見出しのファイルを理由付きで拒否する', async () => {
    const { default: ExcelJS } = await import('exceljs');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('日次売上');
    worksheet.addRow([...SALES_COLUMNS, '売上日']);
    worksheet.addRow(['2026-09-30', 'OS-KB', '京橋店', 'B001', '湯種食パン', 'パン', 1, 360, 360, '現金', '2026-09-30']);
    const bytes = await workbook.xlsx.writeBuffer();
    const file = new File([bytes], '重複見出し.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await expect(aggregateSalesFiles([file])).rejects.toThrow('見出し「売上日」が重複しています');
  });
});
