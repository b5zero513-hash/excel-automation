SRS_MODE: LITE

# Excel業務自動化ツール 第1版 要件

## 1. 目的と範囲

販売HPから利用するデモとして、一般ユーザーが `.xlsx` をブラウザ内で仕分け、加工結果をダウンロードできるようにする。サーバー保存・外部送信・ログイン・履歴・HP改修は対象外。出所はユーザー依頼を `E-01`、合意済み `PLAN.md` を `E-02` とする。

## 2. Must要件（Given / When / Then）

| ID | 根拠 | Given / When / Then | 受入テスト |
| --- | --- | --- | --- |
| FR-FILE-001 | E-01 | Given `.xlsx` を選ぶとき、When 読み込むと、Then ブラウザ内でブックとシート一覧を表示する | TC-FILE-001 |
| FR-SPLIT-001 | E-01 | Given 対象シート・見出し行・列を指定したとき、When 仕分けると、Then 選択列の値ごとに行を分けたシートを追加する | TC-SPLIT-001 |
| FR-DATA-001 | E-01, E-02 | Given 有効な表データを仕分けたとき、When 出力すると、Then 元シートを保ち、出力ファイルを再読込できる | TC-DATA-001 |
| FR-EMPTY-001 | E-01 | Given 選択列に空欄があるとき、When 仕分けると、Then 空欄グループのシートに行をまとめる | TC-EMPTY-001 |
| FR-NAME-001 | E-01 | Given 値にシート名制限文字・長さ超過・重複があるとき、When シートを作ると、Then Excelで有効かつブック内一意な名前にする | TC-NAME-001 |
| FR-DOWNLOAD-001 | E-01 | Given 仕分けに成功したとき、When ダウンロードを選ぶと、Then 加工済み `.xlsx` を取得できる | TC-DOWNLOAD-001 |
| FR-SAFETY-001 | E-01, E-02 | Given 非対応形式・上限超過・対象外構造のファイルを選んだとき、When 読込または処理すると、Then 書き出さず日本語の理由を表示する | TC-SAFETY-001 |

## 3. 非機能要件

- ファイル本文をネットワーク送信せず、ブラウザー内だけで扱う。`TC-SAFETY-001`。
- 上限は `CONSTRAINTS.md` のとおり。超過は処理前に拒否する。`TC-SAFETY-001`。
- 対応ブラウザーの幅で基本操作ができ、ラベル・キーボード操作・エラー文を備える。`TC-UI-001`。

## 4. データと外部連携

入力はFile APIで選択する `.xlsx`。処理中のメモリー以外に保存しない。出力は同じブックへ元データシートを残し、分割シートを追加した `.xlsx`。外部サービスやAPI連携なし。

## 5. 未決事項・制限

ExcelJSによる書式などの完全な保持は保証しない。数式、結合セル、図形等の対象外構造は読み込んだ後に検出し停止する。ファイル上限は初版の固定値とし、実利用で必要性が出た場合に変更を別途判断する。

## 6. トレーサビリティ

| 要件 | 実装 | 確認 |
| --- | --- | --- |
| FR-FILE-001 | `src/excel/workbook.ts`, `src/features/split/App.tsx` | TC-FILE-001 |
| FR-SPLIT-001, FR-EMPTY-001, FR-NAME-001 | `src/excel/workbook.ts` | TC-SPLIT-001, TC-EMPTY-001, TC-NAME-001 |
| FR-DATA-001, FR-DOWNLOAD-001 | `src/features/split/App.tsx` | TC-DATA-001, TC-DOWNLOAD-001 |
| FR-SAFETY-001 | `src/excel/workbook.ts` | TC-SAFETY-001 |
