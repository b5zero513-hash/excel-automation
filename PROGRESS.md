# 進捗

更新日: 2026-10-01
状態: 第1版実装済み。公開・販売HPへの接続は未着手。

## Nagame原則の適用

- `docs/SRS.md` に第1版Must要件、根拠E-ID、受入TC-ID、実装先を記録。
- `CONSTRAINTS.md` に対象範囲、入力上限、対象外条件、処理境界を記録。
- 実装範囲を `PLAN.md` の第1版に限定。サーバー保存、認証、履歴、Web Worker、HP更新、公開は含めていない。
- 入力を処理できない場合は書き出しを止め、理由を表示する。

## 実装・確認

- ExcelJS 4.4.0で日本語シート・列、`.xlsx` 読込、列値ごとの仕分け、複数シート書出し、再読込を確認。
- React・TypeScript・Viteの静的アプリを作成。元ブックと元シートを残し、分割シートを追加。
- 日本語の架空データを使ったサンプル `.xlsx` を追加。
- `npm test`: 4件 PASS。
- `npm run verify:exceljs`: PASS。
- `npm run build`: PASS。ExcelJSの分割チャンクは929.55 kB（gzip 256.42 kB）。初期画面からは遅延読込。
- Chromeでサンプル選択→5グループ表示→仕分け→ダウンロードを確認。ダウンロード名 `excel-demo_仕分け.xlsx`。
- 390px幅で横スクロールなし（viewport/content とも390px）。ブラウザーエラー・外部通信なし。
- `npm audit`: 既知の脆弱性0件。ExcelJSの `uuid` 依存を互換確認済みの修正版へ上書き。
- 静的プレビューでトップページとサンプルをHTTP 200で取得。

## 残件

- デプロイと販売HPへのリンク追加は未実施。
- ExcelJSの全機能・表示の完全保持は保証しない。詳細は `README.md` と `CONSTRAINTS.md` を参照。
