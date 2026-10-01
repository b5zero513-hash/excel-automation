import ExcelJS from 'exceljs';
import { mkdir } from 'node:fs/promises';

const storeRows = [
  { slug: 'kyobashi', code: 'OS-KB', name: '京橋店', factor: 1.12 },
  { slug: 'temma', code: 'OS-TM', name: '天満店', factor: 0.96 },
  { slug: 'miyakojima', code: 'OS-MY', name: '都島店', factor: 1.03 },
  { slug: 'moriguchi', code: 'OS-MG', name: '守口店', factor: 0.91 },
  { slug: 'neyagawa', code: 'OS-NY', name: '寝屋川店', factor: 1.08 },
];
const products = [
  ['B001', '湯種食パン', 'パン', 360, 30],
  ['B002', 'クロワッサン', 'パン', 230, 52],
  ['B003', '塩バターロール', 'パン', 190, 48],
  ['B004', 'つぶあんぱん', 'パン', 210, 35],
  ['B005', 'メロンパン', 'パン', 240, 32],
  ['S001', 'たまごサンド', 'サンドイッチ', 430, 23],
  ['S002', 'ミックスサンド', 'サンドイッチ', 490, 17],
  ['S003', '照り焼きチキンサンド', 'サンドイッチ', 540, 13],
  ['C001', 'カスタードプリン', '菓子', 330, 18],
  ['C002', '季節の焼き菓子', '菓子', 280, 25],
  ['D001', 'ブレンドコーヒー', 'ドリンク', 250, 29],
  ['D002', 'カフェオレ', 'ドリンク', 300, 21],
];
const paymentMethods = ['現金', 'クレジットカード', '交通系IC'];

await mkdir('public/samples/sales', { recursive: true });
for (const [storeIndex, store] of storeRows.entries()) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'まちのパン工房（架空データ）';
  workbook.subject = '店舗別の日次売上サンプル';
  const sheet = workbook.addWorksheet('日次売上');
  sheet.columns = [
    { width: 14 }, { width: 15 }, { width: 16 }, { width: 14 }, { width: 25 },
    { width: 18 }, { width: 14 }, { width: 12 }, { width: 14 }, { width: 18 },
  ];
  sheet.addRow(['売上日', '店舗コード', '店舗名', '商品コード', '商品名', 'カテゴリ', '販売数量', '単価', '売上金額', '支払方法']);
  products.forEach(([productCode, productName, category, unitPrice, baseQuantity], productIndex) => {
    const variation = ((productIndex * 7 + storeIndex * 11) % 9) - 4;
    const quantity = Math.max(5, Math.round(baseQuantity * store.factor) + variation);
    const paymentMethod = paymentMethods[(productIndex + storeIndex * 2) % paymentMethods.length];
    sheet.addRow(['2026-09-30', store.code, store.name, productCode, productName, category, quantity, unitPrice, quantity * unitPrice, paymentMethod]);
  });
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9A512D' } };
  sheet.getRow(1).height = 24;
  sheet.autoFilter = { from: 'A1', to: 'J13' };
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.getColumn(8).numFmt = '#,##0"円"';
  sheet.getColumn(9).numFmt = '#,##0"円"';
  await workbook.xlsx.writeFile(`public/samples/sales/${store.slug}.xlsx`);
}
