import { useEffect, useState } from 'react';
import type { KeyboardEvent } from 'react';
import { SalesApp } from '../sales/SalesApp';
import { SplitApp } from './SplitApp';
import { aggregateSalesFiles } from '../sales/sales';
import type { SalesSummary } from '../sales/sales';

type ViewId = 'demo' | 'dashboard' | 'specs' | 'diagram';
const views: { id: ViewId; label: string; eyebrow: string }[] = [
  { id: 'demo', label: '動くデモ', eyebrow: 'TRY THE DEMO' },
  { id: 'dashboard', label: 'ダッシュボード', eyebrow: 'SAMPLE RESULTS' },
  { id: 'specs', label: '仕様書', eyebrow: 'PRODUCT DETAILS' },
  { id: 'diagram', label: '図解', eyebrow: 'HOW IT WORKS' },
];
const sampleStores = [
  ['kyobashi', '京橋店'], ['temma', '天満店'], ['miyakojima', '都島店'], ['moriguchi', '守口店'], ['neyagawa', '寝屋川店'],
] as const;
const yen = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY', maximumFractionDigits: 0 });
const integer = new Intl.NumberFormat('ja-JP');

function currentView(): ViewId {
  const id = window.location.hash.replace(/^#\/?/, '').split('/')[0];
  return views.some((view) => view.id === id) ? id as ViewId : 'demo';
}

function Dashboard({ summary, loading }: { summary: SalesSummary | null; loading: boolean }) {
  if (loading || !summary) return <section className="product-content"><p className="product-eyebrow">SAMPLE RESULTS</p><h1>集計結果を準備しています</h1><p>架空の5店舗サンプルから、売上・店舗別・カテゴリ別の結果を読み込んでいます。</p></section>;
  const maxStore = Math.max(1, ...summary.stores.map((store) => store.salesAmount));
  const maxCategory = Math.max(1, ...summary.categories.map((item) => item.salesAmount));
  return <section className="product-content dashboard-content">
    <div className="content-intro"><p className="product-eyebrow">SAMPLE RESULTS · 架空データ</p><h1>集計結果をひと目で確認</h1><p>5店舗・{integer.format(summary.records.length)}商品行のサンプル結果です。動くデモで集計した結果がある場合は、その結果に切り替わります。</p></div>
    <div className="dashboard-kpis"><article className="dashboard-total"><span>全店舗の売上合計</span><strong>{yen.format(summary.totalSales)}</strong><small>{integer.format(summary.totalQuantity)}点 · {summary.stores.length}店舗</small></article><article><span>集計日</span><strong>{new Intl.DateTimeFormat('ja-JP', { dateStyle: 'long', timeZone: 'UTC' }).format(new Date(`${summary.salesDate}T00:00:00Z`))}</strong><small>店舗別Excelの売上日</small></article><article><span>出力ブック</span><strong>3シート</strong><small>統合データ・店舗別・カテゴリ別</small></article></div>
    <div className="dashboard-grid">
      <section className="dashboard-card"><div className="dashboard-card-head"><div><p className="product-eyebrow">BY STORE</p><h2>店舗別売上</h2></div><span>売上金額と販売点数</span></div><div className="dashboard-bars">{summary.stores.map((store) => <div className="dashboard-bar-row" key={store.storeCode}><div><strong>{store.storeName}</strong><b>{yen.format(store.salesAmount)}</b></div><div className="dashboard-track"><i style={{ width: `${Math.max(3, store.salesAmount / maxStore * 100)}%` }} /></div><small>{integer.format(store.quantity)}点</small></div>)}</div></section>
      <section className="dashboard-card"><div className="dashboard-card-head"><div><p className="product-eyebrow">BY CATEGORY</p><h2>カテゴリ別売上</h2></div><span>売上構成の内訳</span></div><div className="dashboard-bars">{summary.categories.map((item) => <div className="dashboard-bar-row" key={item.category}><div><strong>{item.category}</strong><b>{yen.format(item.salesAmount)}</b></div><div className="dashboard-track category-track"><i style={{ width: `${Math.max(3, item.salesAmount / maxCategory * 100)}%` }} /></div><small>{integer.format(item.quantity)}点</small></div>)}</div></section>
    </div>
    <div className="dashboard-note"><span aria-hidden="true">i</span><p><strong>サンプル表示</strong> この数値は架空データを集計した結果です。実店舗の実績や将来の効果を示すものではありません。</p></div>
  </section>;
}

function Specs() {
  return <section className="product-content specs-content"><div className="content-intro"><p className="product-eyebrow">PRODUCT DETAILS</p><h1>利用条件と処理内容</h1><p>売上集計デモで対応する入力、処理、出力の範囲をまとめています。</p></div>
    <div className="spec-section"><h2>入力条件</h2><dl className="spec-list"><div><dt>ファイル</dt><dd>店舗ごとのExcelブック。1ファイルにつき1店舗の売上データ</dd></div><div><dt>必須見出し</dt><dd>売上日、店舗コード、店舗名、商品コード、商品名、カテゴリ、販売数量、単価、売上金額、支払方法</dd></div><div><dt>列の並び</dt><dd>見出し名で項目を判定するため列順は変更できます。見出しは1行目、データは2行目以降です。</dd></div><div><dt>データ条件</dt><dd>全ファイルで売上日を統一し、店舗コードが重複しないようにしてください。</dd></div><div><dt>上限</dt><dd>最大10ファイル、1ファイル10 MiB以下、合計50 MiB以下、統合後100,000行以下</dd></div></dl></div>
    <div className="spec-two-col"><section className="spec-section"><h2>対応形式</h2><p>通常の値が並ぶ `.xlsx` ファイルに対応します。`.xls`、暗号化ブック、マクロブックは対象外です。</p></section><section className="spec-section"><h2>処理とデータの扱い</h2><p>ブラウザー内で読み込み・検証・集計します。選択した売上ファイルをサーバーへ送信したり、ローカルストレージへ保存したりしません。</p></section></div>
    <div className="spec-two-col"><section className="spec-section"><h2>集計内容</h2><ul><li>全店舗の売上合計・販売数量</li><li>店舗ごとの売上と販売数量</li><li>カテゴリごとの売上と販売数量</li></ul></section><section className="spec-section"><h2>出力内容</h2><p>新しい `.xlsx` ブックに3シートを作成します。</p><ol><li>全店舗統合データ</li><li>店舗別集計</li><li>カテゴリ別集計</li></ol></section></div>
    <div className="spec-caveat"><strong>対象外・ご注意</strong><p>数式、結合セル、画像・グラフ・ピボット、外部リンクなどを含むブックは処理対象外です。出力ブックがExcelの高度な機能や表示をすべて保持するものではありません。料金は現在未確定です。</p></div>
  </section>;
}

const flowSteps = [
  { number: '01', icon: '▤', title: '店舗別Excel', detail: '各店舗から届いた売上ファイルをまとめて選択', output: '複数の .xlsx' },
  { number: '02', icon: '✓', title: '検証・統合', detail: '見出し・日付・店舗・数値を確認して売上行を統合', output: '統合データ' },
  { number: '03', icon: '▥', title: '店舗別／カテゴリ別集計', detail: '売上金額と販売数量をそれぞれの切り口で集計', output: '2種類の集計' },
  { number: '04', icon: '◉', title: '結果確認', detail: '全体合計と内訳を画面で確認', output: '集計サマリー' },
  { number: '05', icon: 'X', title: '3シートExcel出力', detail: '集計結果を新しいExcelブックにまとめて保存', output: '.xlsx ダウンロード' },
];

function Diagram() {
  return <section className="product-content diagram-content"><div className="content-intro"><p className="product-eyebrow">HOW IT WORKS</p><h1>店舗のExcelから、集計済みブックまで</h1><p>入力データがどのように検証・集計され、結果ファイルになるかを示します。</p></div>
    <div className="flow-diagram" aria-label="店舗別Excelから検証と統合、店舗別とカテゴリ別の集計、結果確認、3シートExcel出力の順に進む処理フロー">{flowSteps.map((step, index) => <div className="flow-item-wrap" key={step.number}><article className={`flow-step flow-step-${step.number}`}><div className="flow-step-top"><span>{step.number}</span><b aria-hidden="true">{step.icon}</b></div><h2>{step.title}</h2><p>{step.detail}</p><small>{step.output}</small></article>{index < flowSteps.length - 1 && <span className="flow-connector" aria-hidden="true">→</span>}</div>)}</div>
    <div className="diagram-branches"><div><span className="branch-mark">A</span><div><strong>店舗別集計</strong><p>どの店舗の売上・販売数量かを一覧で比較</p></div></div><div><span className="branch-mark branch-green">B</span><div><strong>カテゴリ別集計</strong><p>商品カテゴリごとの売上・販売数量を確認</p></div></div></div>
    <p className="diagram-footnote">入力ファイルは利用者のブラウザー内で処理します。サンプルファイルのみ、デモ表示のためアプリから読み込みます。</p>
  </section>;
}

export function App() {
  const [view, setView] = useState<ViewId>(currentView);
  const [splitOpen, setSplitOpen] = useState(false);
  const [demoSummary, setDemoSummary] = useState<SalesSummary | null>(null);
  const [sampleSummary, setSampleSummary] = useState<SalesSummary | null>(null);
  const [sampleLoading, setSampleLoading] = useState(true);

  useEffect(() => {
    const syncView = () => setView(currentView());
    window.addEventListener('hashchange', syncView);
    window.addEventListener('popstate', syncView);
    const files = Promise.all(sampleStores.map(async ([id, name]) => {
      const response = await fetch(`${import.meta.env.BASE_URL}samples/sales/${id}.xlsx`);
      if (!response.ok) throw new Error('サンプルを読み込めませんでした');
      return new File([await response.arrayBuffer()], `${name}_売上.xlsx`, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }));
    void files.then(aggregateSalesFiles).then(setSampleSummary).catch(() => setSampleSummary(null)).finally(() => setSampleLoading(false));
    return () => {
      window.removeEventListener('hashchange', syncView);
      window.removeEventListener('popstate', syncView);
    };
  }, []);

  function selectView(nextView: ViewId) {
    if (currentView() !== nextView) window.history.pushState(null, '', `#${nextView}`);
    setView(nextView);
    setSplitOpen(false);
  }

  function handleTabKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight' && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const currentIndex = views.findIndex((item) => item.id === view);
    const nextIndex = event.key === 'Home' ? 0 : event.key === 'End' ? views.length - 1 : (currentIndex + (event.key === 'ArrowRight' ? 1 : views.length - 1)) % views.length;
    const nextView = views[nextIndex].id;
    selectView(nextView);
    window.requestAnimationFrame(() => document.getElementById(`tab-${nextView}`)?.focus());
  }

  return <div className="product-page">
    <header className="product-header">
      <a className="catalog-link" href="#demo" onClick={() => selectView('demo')}><span aria-hidden="true">←</span> 商品ページ</a>
      <div className="product-heading"><span className="product-mark" aria-hidden="true">X</span><div><p>店舗業務を、Excelから効率化</p><h1>店舗売上Excel 集計</h1></div></div>
      <div className="product-price"><span>料金</span><strong>未定</strong><small>提供条件を確認中</small></div>
    </header>
    <nav className="product-nav" aria-label="商品ページ"><div role="tablist" aria-label="商品情報の画面" onKeyDown={handleTabKeyDown}>{views.map((item) => <button key={item.id} type="button" role="tab" id={`tab-${item.id}`} aria-selected={view === item.id} aria-controls={`panel-${item.id}`} tabIndex={view === item.id ? 0 : -1} className={view === item.id ? 'active' : ''} onClick={() => selectView(item.id)}>{item.label}</button>)}</div><span className="nav-caption">商品ページ</span></nav>
    <main>
      <section id="panel-demo" role="tabpanel" aria-labelledby="tab-demo" hidden={view !== 'demo'} className="product-panel demo-view"><div className="demo-section-label"><span className="product-eyebrow">{splitOpen ? 'EXCEL SPLIT DEMO' : 'LIVE DEMO'}</span><span>ブラウザー内で処理</span></div><div hidden={splitOpen}><SalesApp onOpenSplit={() => setSplitOpen(true)} onSummaryChange={setDemoSummary} /></div><div hidden={!splitOpen}><SplitApp onBack={() => setSplitOpen(false)} /></div></section>
      <section id="panel-dashboard" role="tabpanel" aria-labelledby="tab-dashboard" hidden={view !== 'dashboard'} className="product-panel"><Dashboard summary={demoSummary ?? sampleSummary} loading={!demoSummary && sampleLoading} /></section>
      <section id="panel-specs" role="tabpanel" aria-labelledby="tab-specs" hidden={view !== 'specs'} className="product-panel"><Specs /></section>
      <section id="panel-diagram" role="tabpanel" aria-labelledby="tab-diagram" hidden={view !== 'diagram'} className="product-panel"><Diagram /></section>
    </main>
    <footer className="product-footer"><span>店舗売上Excel 集計</span><span>サンプルデータはすべて架空です · ファイルはブラウザー内で処理</span></footer>
  </div>;
}
