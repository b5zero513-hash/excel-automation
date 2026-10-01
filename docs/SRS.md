SRS_MODE: LITE

# 店舗売上Excel 集計デモ 第1版 要件

## 1. 目的と範囲

架空のベーカリー「まちのパン工房」を題材に、各店舗の日次売上Excelを一括統合し、売上状況の確認と集計済みExcelの取得を体験できる。Excel仕分け機能はサブ機能として維持する。ユーザー依頼を `E-01`、既存合意と制約を `E-02` とする。

## 2. Must要件

| ID | 根拠 | Given / When / Then | 受入テスト |
| --- | --- | --- | --- |
| FR-SALES-FILE-001 | E-01, E-02 | Given 同じ売上日の店舗別 `.xlsx` を1〜10個選んだとき、When 読み込むと、Then 必須見出しと入力境界を検証し、店舗・商品データを取得する | TC-SALES-001 |
| FR-SALES-AGG-001 | E-01 | Given 有効な店舗ファイルが複数あるとき、When 集計すると、Then 全店合計、店舗別売上、カテゴリ別売上と販売数量を表示する | TC-SALES-002 |
| FR-SALES-XLSX-001 | E-01 | Given 集計が成功したとき、When ダウンロードすると、Then 統合データ・店舗別集計・カテゴリ別集計の3シートを含む `.xlsx` を取得できる | TC-SALES-003 |
| FR-SALES-SAMPLE-001 | E-01 | Given サンプルで試すを選んだとき、When サンプル処理が完了すると、Then 5店舗の架空データと実際の集計結果を表示する | TC-SALES-004 |
| FR-SALES-LOCAL-001 | E-01, E-02 | Given 利用者が店舗Excelを選んだとき、When 読込・集計すると、Then ファイル本文をサーバー送信・ブラウザーストレージ保存しない | TC-SALES-005 |
| FR-SPLIT-001 | E-02 | Given 仕分け画面で対象シート・見出し行・列を指定したとき、When 仕分けると、Then 選択列の値ごとにシートを追加する | TC-SPLIT-001 |
| FR-SPLIT-DATA-001 | E-02 | Given 有効な仕分け対象があるとき、When 出力すると、Then 元シートを保ったExcelをダウンロードできる | TC-DATA-001, TC-DOWNLOAD-001 |
| FR-SAFETY-001 | E-02 | Given 非対応形式・不正データ・上限超過を入力したとき、When 処理すると、Then 出力せず日本語で理由を表示する | TC-SAFETY-001 |
| FR-UI-001 | E-01 | Given 販売HPからデモを閲覧するとき、When 画面を操作すると、Then 主な説明・入力・集計結果を読みやすい文字サイズで確認できる | TC-UI-001 |

## 3. 入出力

- 店舗ファイルは1行目に「売上日、店舗コード、店舗名、商品コード、商品名、カテゴリ、販売数量、単価、売上金額、支払方法」を含む。列順は任意。
- 1ファイル内は1店舗、店舗コードはファイル間で一意、すべて同じ売上日とする。
- 空行は読み飛ばし、必須項目や値形式に問題がある場合は停止する。
- 集計出力には「全店舗統合データ」「店舗別集計」「カテゴリ別集計」を含む。
- 売上金額を店舗別・カテゴリ別に合算し、店舗・カテゴリの合計が全店合計と一致する。

## 4. 非機能要件と制約

- Excel処理はブラウザー内で行い、サーバー、API、LocalStorage、IndexedDBへ利用者ファイルを送信・保存しない。
- 上限と対象外ファイルは `CONSTRAINTS.md` に従う。
- PC・スマートフォン幅で操作でき、ボタン・入力ラベル・エラーを備える。
- Cloudflare Pages向けの静的配信を維持する。

## 5. トレーサビリティ

| 要件 | 実装 | 確認 |
| --- | --- | --- |
| FR-SALES-FILE-001, FR-SALES-AGG-001 | `src/features/sales/sales.ts`, `SalesApp.tsx` | TC-SALES-001, TC-SALES-002 |
| FR-SALES-XLSX-001 | `src/features/sales/sales.ts` | TC-SALES-003 |
| FR-SALES-SAMPLE-001 | `scripts/create-sales-samples.mjs`, `public/samples/sales/`, `SalesApp.tsx` | TC-SALES-004 |
| FR-SALES-LOCAL-001 | `SalesApp.tsx`, `sales.ts` | TC-SALES-005 |
| FR-SPLIT-001, FR-SPLIT-DATA-001 | `src/excel/workbook.ts`, `src/features/split/SplitApp.tsx` | TC-SPLIT-001, TC-DATA-001, TC-DOWNLOAD-001 |
| FR-SAFETY-001 | `src/excel/workbook.ts`, `src/features/sales/sales.ts` | TC-SAFETY-001 |
| FR-UI-001 | `src/features/sales/SalesApp.tsx`, `src/styles.css` | TC-UI-001 |
