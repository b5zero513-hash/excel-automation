import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
workbook.creator = 'ミライラボ';
workbook.subject = 'Excel仕分けのサンプル';

const data = workbook.addWorksheet('案件一覧');
data.columns = [
  { width: 14 },
  { width: 12 },
  { width: 19 },
  { width: 15 },
  { width: 34 },
];
data.addRow(['受付番号', '担当者', '部門', '見積金額', 'メモ']);
data.addRows([
  ['A-001', '佐藤', '東京/営業', 125000, '新規サイト制作'],
  ['A-002', '田中', '大阪営業', 88000, 'バナー追加'],
  ['A-003', '鈴木', '東京/営業', 45000, '写真差し替え'],
  ['A-004', '高橋', null, 32000, '部門未入力'],
  ['A-005', '伊藤', '東京_営業', 64000, '同じシート名になる例'],
  ['A-006', '山本', '製造・開発・カスタマーサポート合同チームを中心に関係者確認チーム', 210000, '長い部門名の例'],
]);
data.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
data.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
data.autoFilter = { from: 'A1', to: 'E7' };
data.views = [{ state: 'frozen', ySplit: 1 }];

const help = workbook.addWorksheet('使い方');
help.addRows([
  ['このサンプルについて'],
  ['「部門」列で仕分けると、部門ごとのシートが追加されます。'],
  ['空欄、使用できないシート名の文字、同じ名前になる値、長い値を含みます。'],
  ['入力例はすべて架空のデータです。'],
]);
help.getColumn(1).width = 78;

await workbook.xlsx.writeFile('public/samples/excel-demo.xlsx');
