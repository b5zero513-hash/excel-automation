import { useMemo, useRef, useState } from 'react';
import type { ChangeEvent } from 'react';
import type { Workbook } from 'exceljs';
import {
  displayCellValue,
  getHeaderOptions,
  listWorksheets,
  loadExcelFile,
  loadWorkbook,
  previewGroups,
  splitWorkbook,
  writeWorkbook,
} from '../../excel/workbook';
import type { SplitSummary, WorksheetInfo } from '../../excel/workbook';

type LoadedFile = {
  file: File;
  workbook: Workbook;
  sheets: WorksheetInfo[];
};

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : 'ファイルを処理できませんでした。別の .xlsx ファイルをお試しください。';
}

function formatBytes(bytes: number): string {
  return `${(bytes / 1024 / 1024).toFixed(2)} MiB`;
}

function makeDownloadName(filename: string): string {
  const base = filename.replace(/\.xlsx$/i, '').replace(/[\\/:*?"<>|]/g, '_').slice(0, 100) || 'Excel';
  return `${base}_仕分け.xlsx`;
}

export function App() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [loaded, setLoaded] = useState<LoadedFile | null>(null);
  const [sheetName, setSheetName] = useState('');
  const [headerRow, setHeaderRow] = useState(1);
  const [columnNumber, setColumnNumber] = useState(1);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState<{ bytes: ArrayBuffer; summary: SplitSummary; filename: string } | null>(null);

  const worksheet = loaded?.workbook.getWorksheet(sheetName);
  const columnOptions = useMemo(() => {
    if (!worksheet || headerRow < 1 || headerRow > worksheet.rowCount) return [];
    try {
      return getHeaderOptions(worksheet, headerRow);
    } catch {
      return [];
    }
  }, [worksheet, headerRow]);
  const preview = useMemo(() => {
    if (!worksheet || !columnNumber || headerRow < 1 || headerRow > worksheet.rowCount) return null;
    try {
      return previewGroups(worksheet, headerRow, columnNumber);
    } catch (previewError) {
      return { error: messageOf(previewError) } as const;
    }
  }, [worksheet, headerRow, columnNumber]);

  async function acceptFile(file: File) {
    setBusy(true);
    setError('');
    setResult(null);
    try {
      const workbook = await loadExcelFile(file);
      const sheets = listWorksheets(workbook);
      setLoaded({ file, workbook, sheets });
      setSheetName(sheets[0]?.name ?? '');
      setHeaderRow(1);
      setColumnNumber(1);
    } catch (loadError) {
      setLoaded(null);
      setError(messageOf(loadError));
    } finally {
      setBusy(false);
    }
  }

  async function onFileChange(event: ChangeEvent<HTMLInputElement>) {
    const input = event.currentTarget;
    const file = input.files?.[0];
    if (file) await acceptFile(file);
    input.value = '';
  }

  async function loadSample() {
    setBusy(true);
    setError('');
    setResult(null);
    try {
      const response = await fetch(`${import.meta.env.BASE_URL}samples/excel-demo.xlsx`);
      if (!response.ok) throw new Error('サンプルファイルを読み込めませんでした。時間をおいて再度お試しください。');
      const file = new File([await response.arrayBuffer()], 'excel-demo.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const workbook = await loadExcelFile(file);
      const sheets = listWorksheets(workbook);
      setLoaded({ file, workbook, sheets });
      setSheetName(sheets[0]?.name ?? '');
      setHeaderRow(1);
      const sampleSheet = workbook.getWorksheet(sheets[0]?.name ?? '');
      const departmentColumn = sampleSheet
        ? Array.from({ length: sampleSheet.columnCount }, (_, index) => index + 1)
            .find((column) => sampleSheet.getRow(1).getCell(column).value === '部門')
        : undefined;
      setColumnNumber(departmentColumn ?? 1);
    } catch (sampleError) {
      setError(messageOf(sampleError));
    } finally {
      setBusy(false);
    }
  }

  async function runSplit() {
    if (!loaded) return;
    setBusy(true);
    setError('');
    setResult(null);
    try {
      const originalBytes = await loaded.file.arrayBuffer();
      const outputWorkbook = await loadWorkbook(originalBytes);
      const split = splitWorkbook(outputWorkbook, sheetName, headerRow, columnNumber);
      const bytes = await writeWorkbook(split.workbook);
      setResult({ bytes, summary: split.summary, filename: makeDownloadName(loaded.file.name) });
    } catch (splitError) {
      setError(messageOf(splitError));
    } finally {
      setBusy(false);
    }
  }

  function downloadResult() {
    if (!result) return;
    const blob = new Blob([result.bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = result.filename;
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
  }

  function reset() {
    setLoaded(null);
    setSheetName('');
    setHeaderRow(1);
    setColumnNumber(1);
    setResult(null);
    setError('');
    if (inputRef.current) inputRef.current.value = '';
  }

  const canSplit = Boolean(loaded && preview && !('error' in preview) && preview.processedRows > 0 && !busy);

  return (
    <main className="page-shell">
      <header className="topbar">
        <a className="brand" href="#top" aria-label="ミライラボ Excel仕分けデモ">
          <span className="brand-mark" aria-hidden="true">M</span>
          <span>ミライラボ <small>業務自動化デモ</small></span>
        </a>
        <span className="demo-pill"><span /> ブラウザ内で処理</span>
      </header>

      <section className="hero" id="top">
        <p className="eyebrow">EXCEL WORKFLOW DEMO</p>
        <h1>Excelの仕分けを、<br className="mobile-break" />かんたんに。</h1>
        <p className="hero-copy">列を選ぶだけで、値ごとにシートを分割。<br className="desktop-break" />元データを残したExcelをダウンロードできます。</p>
        <p className="privacy-note"><span aria-hidden="true">▣</span> ファイルはこのブラウザ内で処理され、サーバーへ送信されません。</p>
      </section>

      <section className="workspace" aria-labelledby="workspace-title">
        <div className="workspace-heading">
          <div>
            <p className="eyebrow">TRY IT NOW</p>
            <h2 id="workspace-title">Excel仕分けを試す</h2>
          </div>
          {loaded && <button className="text-button" type="button" onClick={reset}>最初からやり直す</button>}
        </div>

        <ol className="steps" aria-label="操作の流れ">
          <li className={!loaded ? 'current' : 'complete'}><span>1</span>ファイルを選ぶ</li>
          <li className={loaded && !result ? 'current' : result ? 'complete' : ''}><span>2</span>シートと列を選ぶ</li>
          <li className={result ? 'current' : ''}><span>3</span>仕分けて保存</li>
        </ol>

        {!loaded ? (
          <div className="upload-panel">
            <div className="upload-icon" aria-hidden="true">↑</div>
            <h3>仕分けたいExcelファイルを選択</h3>
            <p>.xlsx形式・10 MiB以下のファイルに対応しています</p>
            <div className="upload-actions">
              <button className="button primary" type="button" onClick={() => inputRef.current?.click()} disabled={busy}>
                {busy ? '読み込み中…' : 'ファイルを選択'}
              </button>
              <button className="button secondary" type="button" onClick={loadSample} disabled={busy}>サンプルで試す</button>
            </div>
            <input ref={inputRef} className="visually-hidden" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={onFileChange} aria-label="Excelファイルを選択" />
            <p className="upload-footnote">個人情報を含まないデータでお試しください。</p>
          </div>
        ) : (
          <div className="config-grid">
            <div className="settings-column">
              <div className="file-card">
                <div className="file-icon" aria-hidden="true">XLSX</div>
                <div className="file-details"><strong>{loaded.file.name}</strong><span>{formatBytes(loaded.file.size)} ・ {loaded.sheets.length}シート</span></div>
                <button type="button" className="text-button" onClick={() => inputRef.current?.click()} disabled={busy}>変更</button>
                <input ref={inputRef} className="visually-hidden" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={onFileChange} aria-label="Excelファイルを変更" />
              </div>

              <div className="field-block">
                <label htmlFor="sheet-select">対象シート</label>
                <select id="sheet-select" value={sheetName} onChange={(event) => { setSheetName(event.target.value); setHeaderRow(1); setColumnNumber(1); setResult(null); }}>
                  {loaded.sheets.map((sheet) => <option key={sheet.id} value={sheet.name}>{sheet.name}（{sheet.rowCount.toLocaleString()}行）</option>)}
                </select>
              </div>

              <div className="field-block">
                <label htmlFor="header-row">見出し行</label>
                <div className="number-field"><input id="header-row" type="number" min={1} max={worksheet?.rowCount ?? 1} value={headerRow} onChange={(event) => { const maxRows = worksheet?.rowCount ?? 1; setHeaderRow(Math.min(maxRows, Math.max(1, Number(event.target.value) || 1))); setResult(null); }} /><span>行目</span></div>
                <small>列名が並んでいる行を指定してください。</small>
              </div>

              <div className="field-block">
                <label htmlFor="column-select">仕分けに使う列</label>
                <select id="column-select" value={columnNumber} onChange={(event) => { setColumnNumber(Number(event.target.value)); setResult(null); }} disabled={!columnOptions.length}>
                  {columnOptions.map((column) => <option key={column.columnNumber} value={column.columnNumber}>{column.optionLabel}</option>)}
                </select>
              </div>

              <div className="original-note"><span aria-hidden="true">✓</span><p>元のシートはそのまま残し、仕分けシートを追加します。</p></div>
            </div>

            <div className="preview-column">
              <div className="preview-title"><div><h3>仕分け結果のプレビュー</h3><p>選択した列の値ごとにシートを作成します。</p></div></div>
              {worksheet && (
                <div className="table-wrap">
                  <table>
                    <thead><tr><th>行</th>{Array.from({ length: Math.min(worksheet.columnCount, 5) }, (_, index) => <th key={index}>{displayCellValue(worksheet.getRow(headerRow).getCell(index + 1).value)}</th>)}{worksheet.columnCount > 5 && <th>…</th>}</tr></thead>
                    <tbody>
                      {Array.from({ length: Math.min(4, Math.max(0, worksheet.rowCount - headerRow)) }, (_, index) => {
                        const rowNumber = headerRow + index + 1;
                        const row = worksheet.getRow(rowNumber);
                        return <tr key={rowNumber}><td>{rowNumber}</td>{Array.from({ length: Math.min(worksheet.columnCount, 5) }, (_, columnIndex) => <td key={columnIndex}>{displayCellValue(row.getCell(columnIndex + 1).value)}</td>)}{worksheet.columnCount > 5 && <td>…</td>}</tr>;
                      })}
                    </tbody>
                  </table>
                </div>
              )}
              {preview && 'error' in preview ? <div className="inline-warning">{preview.error}</div> : preview && (
                <div className="group-list">
                  <div className="group-summary"><strong>{preview.processedRows.toLocaleString()}行</strong><span>{preview.splitSheets.length}シートに分割</span>{preview.skippedBlankRows > 0 && <span>空行{preview.skippedBlankRows}行を除外</span>}</div>
                  {preview.splitSheets.map((group) => <div className="group-row" key={group.key}><span className="sheet-dot" aria-hidden="true" /><span className="group-name"><strong>{group.sheetName}</strong>{group.label === '空欄' && <small>仕分け値が空欄</small>}</span><span className="row-count">{group.count.toLocaleString()}行</span></div>)}
                </div>
              )}
            </div>
          </div>
        )}

        {error && <div className="alert error" role="alert">{error}</div>}

        {loaded && !result && (
          <div className="run-row"><button className="button primary run-button" type="button" onClick={runSplit} disabled={!canSplit}>{busy ? '仕分け中…' : '仕分けを実行する'}</button><span>元ファイルは変更されません</span></div>
        )}

        {result && (
          <div className="result-panel" role="status">
            <div className="result-heading"><span className="success-mark" aria-hidden="true">✓</span><div><h3>仕分けが完了しました</h3><p>{result.summary.processedRows.toLocaleString()}行を{result.summary.splitSheets.length}シートに分けました。</p></div></div>
            <button className="button primary download-button" type="button" onClick={downloadResult}>加工済みファイルをダウンロード</button>
            <p className="result-file-name">{result.filename}</p>
          </div>
        )}
      </section>

      <section className="notice-card">
        <h2>安心してお試しいただくために</h2>
        <ul>
          <li>ファイルはお使いのブラウザ内で処理します。サーバーへアップロードしません。</li>
          <li>数式、結合セル、画像、グラフ、マクロを含むファイルには対応していません。</li>
          <li>元のファイルを残し、加工結果を別ファイルとして保存します。</li>
        </ul>
      </section>

      <footer className="footer"><span>© ミライラボ</span><span>Excel業務自動化デモ</span></footer>
    </main>
  );
}
