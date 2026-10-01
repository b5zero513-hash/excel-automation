import { useRef, useState } from 'react';
import type { ChangeEvent } from 'react';
import { aggregateSalesFiles, createSalesWorkbook } from './sales';
import type { SalesSummary } from './sales';

const SAMPLE_FILES = ['kyobashi', 'temma', 'miyakojima', 'moriguchi', 'neyagawa'] as const;
const STORE_NAMES: Record<(typeof SAMPLE_FILES)[number], string> = {
  kyobashi: '京橋店',
  temma: '天満店',
  miyakojima: '都島店',
  moriguchi: '守口店',
  neyagawa: '寝屋川店',
};
const yen = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY', maximumFractionDigits: 0 });
const integer = new Intl.NumberFormat('ja-JP');

function dateLabel(date: string) {
  return new Intl.DateTimeFormat('ja-JP', { year: 'numeric', month: 'long', day: 'numeric', timeZone: 'UTC' }).format(new Date(`${date}T00:00:00Z`));
}

export function SalesApp({ onOpenSplit }: { onOpenSplit: () => void }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<File[]>([]);
  const [summary, setSummary] = useState<SalesSummary | null>(null);
  const [output, setOutput] = useState<ArrayBuffer | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  async function aggregate(filesToAggregate: File[]) {
    setFiles(filesToAggregate);
    setSummary(null);
    setOutput(null);
    setError('');
    if (!filesToAggregate.length) return;
    setBusy(true);
    try {
      const nextSummary = await aggregateSalesFiles(filesToAggregate);
      const bytes = await createSalesWorkbook(nextSummary);
      setSummary(nextSummary);
      setOutput(bytes);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Excelを集計できませんでした。ファイルの内容を確認してください。');
    } finally {
      setBusy(false);
    }
  }

  async function onFilesSelected(event: ChangeEvent<HTMLInputElement>) {
    const selected = Array.from(event.currentTarget.files ?? []);
    event.currentTarget.value = '';
    await aggregate(selected);
  }

  async function trySample() {
    setBusy(true);
    setFiles([]);
    setSummary(null);
    setOutput(null);
    setError('');
    try {
      const sampleFiles = await Promise.all(SAMPLE_FILES.map(async (store) => {
        const response = await fetch(`${import.meta.env.BASE_URL}samples/sales/${store}.xlsx`);
        if (!response.ok) throw new Error('サンプルファイルを読み込めませんでした。少し時間をおいてお試しください。');
        const bytes = await response.arrayBuffer();
        return new File([bytes], `${STORE_NAMES[store]}_売上.xlsx`, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      }));
      setFiles(sampleFiles);
      const nextSummary = await aggregateSalesFiles(sampleFiles);
      const bytes = await createSalesWorkbook(nextSummary);
      setSummary(nextSummary);
      setOutput(bytes);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'サンプルを集計できませんでした。');
    } finally {
      setBusy(false);
    }
  }

  function download() {
    if (!output || !summary) return;
    const blob = new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `まちのパン工房_${summary.salesDate}_全店舗売上集計.xlsx`;
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
  }

  function reset() {
    setFiles([]);
    setSummary(null);
    setOutput(null);
    setError('');
  }

  const maxCategory = Math.max(1, ...(summary?.categories.map((item) => item.salesAmount) ?? []));

  return (
    <main className="sales-page">
      <header className="sales-topbar">
        <a className="bakery-brand" href="#top" aria-label="まちのパン工房 売上集計デモ">
          <span className="bakery-mark" aria-hidden="true">麦</span>
          <span><strong>まちのパン工房</strong><small>店舗売上 集計デモ</small></span>
        </a>
        <div className="sales-top-actions">
          <span className="local-badge"><span aria-hidden="true">●</span> ファイルはブラウザー内で処理</span>
          <button type="button" className="secondary-action" onClick={onOpenSplit}>Excel仕分けも試す</button>
        </div>
      </header>

      <section className="sales-hero" id="top">
        <div className="hero-text">
          <p className="sales-eyebrow">毎日の店舗集計を、まとめて簡単に</p>
          <h1>5店舗の売上Excel。<br />毎日ひとつずつ開く作業を、<br className="desktop-line" />まとめて集計。</h1>
          <p className="sales-lead">各店から届く売上ファイルを一度に選ぶだけ。<br className="desktop-line" />店舗別・カテゴリ別の結果をすぐに確認できます。</p>
          <div className="value-points"><span>✓ 店舗別の売上</span><span>✓ 全店舗の合計</span><span>✓ カテゴリ別の内訳</span></div>
        </div>
        <div className="aggregation-visual" aria-label="5店舗のExcelを1つの集計表にまとめるイメージ">
          <div className="store-stack">
            {['京橋', '天満', '都島', '守口', '寝屋川'].map((name, index) => <div className="mini-file" key={name}><span className="mini-file-icon">X</span><span>{name}店</span><small>売上.xlsx</small><b>{String(index + 1).padStart(2, '0')}</b></div>)}
          </div>
          <div className="flow-arrow" aria-hidden="true">→</div>
          <div className="report-illustration"><div className="report-cap"><span>DAILY REPORT</span><b>本日の売上</b></div><strong>集計完了</strong><div className="report-bars"><i /><i /><i /><i /></div><small>店舗別 ＋ カテゴリ別</small></div>
        </div>
      </section>

      <section className="sales-workspace" aria-labelledby="upload-heading">
        <div className="section-heading">
          <div><p className="sales-eyebrow">3ステップで完了</p><h2 id="upload-heading">5店舗の売上をまとめる</h2></div>
          {files.length > 0 && <button type="button" className="link-action" onClick={reset}>最初からやり直す</button>}
        </div>
        <ol className="sales-steps" aria-label="集計の流れ">
          <li className={files.length ? 'done' : 'active'}><b>1</b><span>5店舗のファイルを選ぶ</span></li>
          <li className={summary ? 'done' : files.length ? 'active' : ''}><b>2</b><span>自動でまとめて集計</span></li>
          <li className={output ? 'active' : ''}><b>3</b><span>結果を確認して保存</span></li>
        </ol>

        <div className="upload-area">
          <div className="upload-copy"><div className="upload-symbol" aria-hidden="true">＋</div><div><h3>店舗の売上Excelをまとめて選択</h3><p>各店の .xlsx ファイルを一度に選べます。最大10ファイル・合計50 MiBまで。</p></div></div>
          <div className="upload-buttons">
            <button type="button" className="sample-button" onClick={trySample} disabled={busy}>{busy ? '集計しています…' : 'サンプル5店舗で集計を試す'}</button>
            <button type="button" className="browse-button" onClick={() => inputRef.current?.click()} disabled={busy}>ファイルを選ぶ</button>
            <input ref={inputRef} className="visually-hidden" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" multiple onChange={onFilesSelected} aria-label="店舗の売上Excelを複数選択" />
          </div>
        </div>

        {files.length > 0 && <div className="selected-files" aria-live="polite"><strong>{files.length}店舗分のファイル</strong><div className="file-chips">{files.map((file) => <span className="file-chip" key={`${file.name}-${file.size}`}>▧ {file.name}</span>)}</div></div>}
        {busy && <div className="loading-message" role="status"><span className="loading-dot" />Excelを読み込み、店舗別とカテゴリ別に集計しています…</div>}
        {error && <div className="sales-error" role="alert"><strong>集計できませんでした</strong><span>{error}</span></div>}

        {summary && <section className="results" aria-labelledby="result-heading">
          <div className="result-head"><div><p className="sales-eyebrow">集計結果</p><h2 id="result-heading">{`${dateLabel(summary.salesDate)}の売上`}</h2><p>{summary.stores.length}店舗・{integer.format(summary.records.length)}商品行をまとめました</p></div><button type="button" className="download-action" onClick={download}>↓　集計Excelをダウンロード</button></div>

          <div className="total-card"><div><span>全店舗の売上合計</span><strong>{yen.format(summary.totalSales)}</strong><small>{integer.format(summary.totalQuantity)}点 販売</small></div><div className="total-card-mark" aria-hidden="true">¥</div></div>

          <div className="result-grid">
            <section className="result-card store-result"><div className="card-heading"><span className="card-icon">店</span><div><h3>店舗別売上</h3><p>各店の売上と販売点数</p></div></div><div className="store-results">{summary.stores.map((store) => <div className="store-result-row" key={store.storeCode}><div className="store-row-title"><strong>{store.storeName}</strong><span>{integer.format(store.quantity)}点</span></div><b>{yen.format(store.salesAmount)}</b><div className="bar-track"><i style={{ width: `${Math.max(4, (store.salesAmount / Math.max(...summary.stores.map((item) => item.salesAmount))) * 100)}%` }} /></div></div>)}</div></section>
            <section className="result-card category-result"><div className="card-heading"><span className="card-icon category-icon">菓</span><div><h3>カテゴリ別売上</h3><p>何が売上を支えているか一目で</p></div></div><div className="category-results">{summary.categories.map((category) => <div className="category-row" key={category.category}><div className="category-title"><strong>{category.category}</strong><b>{yen.format(category.salesAmount)}</b></div><div className="bar-track category-track"><i style={{ width: `${Math.max(4, (category.salesAmount / maxCategory) * 100)}%` }} /></div><small>{integer.format(category.quantity)}点</small></div>)}</div></section>
          </div>

          <div className="download-detail"><div><span className="spreadsheet-icon">XLSX</span><p><strong>集計済みExcel</strong><small>3シート：全店舗統合データ・店舗別集計・カテゴリ別集計</small></p></div><button type="button" className="download-action" onClick={download}>ダウンロード</button></div>
        </section>}
      </section>

      <section className="how-it-works"><div><span className="step-number">01</span><h3>5店舗から届く</h3><p>京橋・天満・都島・守口・寝屋川。各店の売上ファイルをまとめて選択。</p></div><span className="how-arrow">→</span><div><span className="step-number">02</span><h3>自動でひとつに</h3><p>同じ項目の売上データを自動で統合。店舗ごと、カテゴリごとに集計。</p></div><span className="how-arrow">→</span><div><span className="step-number">03</span><h3>すぐに確認・保存</h3><p>全体の売上と内訳を画面で確認し、集計済みExcelをダウンロード。</p></div></section>

      <footer className="sales-footer"><p><span>🔒</span> 選んだ売上ファイルはこのブラウザー内で処理します。集計のためにサーバーへ送信しません。</p><small>サンプルはすべて架空の店舗・商品・売上データです。対応形式は .xlsx（1ファイル10 MiB以下）です。</small><div><span>まちのパン工房｜売上集計デモ</span><button type="button" onClick={onOpenSplit}>Excel仕分け機能はこちら</button></div></footer>
    </main>
  );
}
