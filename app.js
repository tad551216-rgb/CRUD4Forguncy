/* app.js — UIグルー。重い処理はすべて Web Worker で実行し、画面を固まらせない。 */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);

  const state = {
    fgcp: null, fgdoc: null, blob: null, filename: null, mermaid: null, mdName: null,
    worker: null, workerBlobURL: null, busy: false,
  };

  // ---- ドロップゾーン配線 ----
  function wireSlot(slotId, inputId, fileLabelId, kind, accept) {
    const slot = $(slotId), input = $(inputId), label = $(fileLabelId);

    const setFile = (file) => {
      if (!file) return;
      const name = file.name.toLowerCase();
      const ok = accept.some((ext) => name.endsWith(ext));
      if (!ok) { showError(`${kind === 'fgcp' ? '①' : '②'} は ${accept.join(' / ')} を選んでください（選択: ${file.name}）`); return; }
      clearError();
      state[kind] = file;
      label.textContent = file.name + '  (' + fmtSize(file.size) + ')';
      slot.classList.add('filled');
      refreshGo();
    };

    slot.addEventListener('click', (e) => { if (!e.target.closest('.clear')) input.click(); });
    slot.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', () => setFile(input.files[0]));

    ['dragenter', 'dragover'].forEach((ev) => slot.addEventListener(ev, (e) => { e.preventDefault(); slot.classList.add('drag'); }));
    ['dragleave', 'drop'].forEach((ev) => slot.addEventListener(ev, (e) => { e.preventDefault(); if (ev === 'dragleave' && slot.contains(e.relatedTarget)) return; slot.classList.remove('drag'); }));
    slot.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; setFile(f); });
  }

  document.querySelectorAll('.clear').forEach((btn) => {
    btn.addEventListener('click', () => {
      const kind = btn.dataset.clear;
      state[kind] = null;
      const slot = kind === 'fgcp' ? $('slotFgcp') : $('slotFgdoc');
      const input = kind === 'fgcp' ? $('inputFgcp') : $('inputFgdoc');
      slot.classList.remove('filled');
      input.value = '';
      refreshGo();
    });
  });

  function fmtSize(b) {
    if (b > 1048576) return (b / 1048576).toFixed(1) + ' MB';
    if (b > 1024) return (b / 1024).toFixed(0) + ' KB';
    return b + ' B';
  }
  function refreshGo() {
    $('go').disabled = !state.fgcp || state.busy;
    $('diag').disabled = !state.fgcp || state.busy;
  }
  function showError(msg) { const e = $('error'); e.textContent = '⚠ ' + msg; e.classList.add('show'); }
  function clearError() { $('error').classList.remove('show'); }
  function showNotice(msg) { const n = $('notice'); n.textContent = msg; n.classList.add('show'); }
  function clearNotice() { $('notice').classList.remove('show'); }
  function escapeHtml(s) { return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }

  // ---- 進捗 ----
  function setProgress(label, pct) {
    $('progress').classList.add('show');
    $('progLabel').textContent = label;
    $('progPct').textContent = pct == null ? '' : pct + '%';
    $('progBar').style.width = (pct == null ? 8 : pct) + '%';
  }
  function hideProgress() { $('progress').classList.remove('show'); }

  // ---- ビジー状態の切り替え ----
  function setBusy(isBusy) {
    state.busy = isBusy;
    refreshGo();
    $('cancelBtn').classList.toggle('show', isBusy);
  }

  // =====================================================================
  // Worker管理: 単一版は埋め込みソースから、複数ファイル版は fetch で構築
  // =====================================================================
  async function getWorkerBlobURL() {
    if (state.workerBlobURL) return state.workerBlobURL;
    let src;
    const embedded = document.getElementById('workerSrc');
    if (embedded && embedded.textContent && embedded.textContent.trim().length > 1000) {
      // 単一版HTML: ワーカーのソースがDOMに埋め込まれている
      src = embedded.textContent;
    } else {
      // 複数ファイル版: 同一オリジンの相対ファイルを取得して連結
      const paths = ['vendor/jszip.min.js', 'vendor/exceljs.min.js', 'crud-core.js', 'worker-core.js'];
      const texts = await Promise.all(paths.map((p) => fetch(p).then((r) => {
        if (!r.ok) throw new Error(`${p} の読み込みに失敗しました (${r.status})`);
        return r.text();
      })));
      src = texts.join('\n;\n');
    }
    const blob = new Blob([src], { type: 'application/javascript' });
    state.workerBlobURL = URL.createObjectURL(blob);
    return state.workerBlobURL;
  }

  async function getWorker() {
    if (state.worker) return state.worker;
    const url = await getWorkerBlobURL();
    state.worker = new Worker(url);
    return state.worker;
  }

  function killWorker() {
    if (state.worker) {
      try { state.worker.terminate(); } catch (e) { /* ignore */ }
      state.worker = null;
    }
  }

  // 指定コマンドをWorkerへ送り、完了/エラー/進捗を処理する共通ヘルパ
  function runInWorker(msg, transfer, onProgress) {
    return new Promise(async (resolve, reject) => {
      let worker;
      try {
        worker = await getWorker();
      } catch (err) {
        reject(new Error('処理の準備に失敗しました: ' + (err && err.message ? err.message : String(err))));
        return;
      }
      const handleMessage = (e) => {
        const data = e.data || {};
        if (data.type === 'progress') {
          if (onProgress) onProgress(data.ev);
        } else if (data.type === 'error') {
          cleanup();
          reject(new Error(data.message));
        } else if (data.type === 'generate-done' || data.type === 'diagnose-done') {
          cleanup();
          resolve(data);
        }
      };
      const handleError = (e) => {
        cleanup();
        reject(new Error('処理中に問題が発生しました: ' + (e && e.message ? e.message : '不明なエラー')));
      };
      const cleanup = () => {
        worker.removeEventListener('message', handleMessage);
        worker.removeEventListener('error', handleError);
      };
      worker.addEventListener('message', handleMessage);
      worker.addEventListener('error', handleError);
      worker.postMessage(msg, transfer || []);
    });
  }

  // ---- 中止 ----
  $('cancelBtn').addEventListener('click', () => {
    if (!state.busy) return;
    killWorker();
    setBusy(false);
    hideProgress();
    showNotice('処理を中止しました。ファイルはそのままなので、もう一度実行できます。');
  });

  // ---- 実行: CRUD一覧生成 ----
  $('go').addEventListener('click', async () => {
    if (!state.fgcp || state.busy) return;
    clearError(); clearNotice();
    $('result').classList.remove('show');
    setBusy(true);
    state.blob = null;
    try {
      setProgress('ファイルを読み込み中…', null);
      const fgcpBuf = await state.fgcp.arrayBuffer();
      const fgdocBuf = state.fgdoc ? await state.fgdoc.arrayBuffer() : null;
      const transfer = [fgcpBuf]; if (fgdocBuf) transfer.push(fgdocBuf);

      const data = await runInWorker(
        { cmd: 'generate', fgcpBuffer: fgcpBuf, fgdocBuffer: fgdocBuf },
        transfer,
        (ev) => {
          if (ev.stage === 'parse') {
            const pct = Math.round((ev.done / ev.total) * 85);
            setProgress(`ページ解析中… ${ev.done} / ${ev.total}`, pct);
          } else if (ev.stage === 'message') {
            setProgress(ev.message, ev.message.indexOf('Excel') >= 0 ? 92 : null);
          }
        }
      );

      state.blob = new Blob([data.xlsx], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      state.filename = `ページ別CRUD一覧_${today}.xlsx`;
      state.mermaid = data.mermaid || '';
      state.mdName = `データ関連図_${today}.md`;

      setProgress('完了', 100);
      setTimeout(hideProgress, 400);
      showResult(data.stats);
    } catch (err) {
      hideProgress();
      showError(err && err.message ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  });

  function showResult(stats) {
    $('stTotal').textContent = stats.total;
    $('stWrite').textContent = stats.write;
    $('stDel').textContent = stats.del;
    const areas = Object.keys(stats.areas).sort((a, b) => a.localeCompare(b, 'ja'));
    let extra = '';
    if (stats.numbered) extra += `<div class="row" style="color:var(--ink-soft)"><span>No.付与</span><span>${stats.numbered}件</span></div>`;
    if (stats.definedTables) {
      extra += `<div class="row" style="color:var(--ink-soft)"><span>TABLES定義</span><span>${stats.definedTables}件</span></div>`;
      if (stats.tableNamesAlign) extra += `<div class="row" style="color:var(--ink-soft)"><span>未使用テーブル候補</span><span>${stats.orphanTables}件</span></div>`;
    }
    $('areas').innerHTML = areas.map((a) =>
      `<div class="row"><span>${escapeHtml(a)}</span><span>${stats.areas[a]}</span></div>`
    ).join('') + extra;
    $('result').classList.add('show');
    $('result').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  $('download').addEventListener('click', () => {
    if (!state.blob) return;
    const url = URL.createObjectURL(state.blob);
    const a = document.createElement('a');
    a.href = url; a.download = state.filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  });

  $('downloadMd').addEventListener('click', () => {
    if (!state.mermaid) return;
    const blob = new Blob([state.mermaid], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = state.mdName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  });

  // ---- 対応診断 ----
  $('diag').addEventListener('click', async () => {
    if (!state.fgcp || state.busy) return;
    clearError(); clearNotice();
    $('diagPanel').classList.remove('show');
    setBusy(true);
    try {
      setProgress('診断中…', null);
      const buf = await state.fgcp.arrayBuffer();
      const data = await runInWorker(
        { cmd: 'diagnose', fgcpBuffer: buf },
        [buf],
        (ev) => {
          if (ev.stage === 'parse' && ev.total) {
            setProgress(`診断中… ${ev.done} / ${ev.total}`, Math.round((ev.done / ev.total) * 100));
          }
        }
      );
      setTimeout(hideProgress, 300);
      renderDiag(data.result);
    } catch (err) {
      hideProgress();
      showError(err && err.message ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  });

  function chips(obj, bad) {
    const keys = Object.keys(obj).sort();
    if (!keys.length) return '<span class="k" style="color:var(--ink-soft)">なし</span>';
    return keys.map((t) => `<span class="tag2 ${bad ? 'bad' : 'ok'}">${escapeHtml(t)} ×${obj[t]}</span>`).join('');
  }

  function renderDiag(d) {
    const v = $('diagVerdict');
    v.className = 'verdict ' + d.verdict;
    v.textContent = d.verdict === 'ok' ? '○ 使えます'
      : d.verdict === 'warn' ? '△ 使えますが一部漏れの可能性'
      : '× 構造が異なる可能性';
    $('diagText').textContent = d.verdictText;

    const rows = [
      ['ページJSON 検出', d.pagesFound + ' 件'],
      ['パース成功', d.parseOk + ' / ' + d.pagesFound + (d.parseErrors.length ? '（失敗 ' + d.parseErrors.length + '）' : '')],
      ['TABLES 定義', (d.tablesFound || 0) + ' 件'],
      ['AttachInfos 無し', d.noAttach + ' ページ'],
      ['末尾署名あり', d.withSignature + ' / ' + d.pagesFound],
    ];
    const folderStr = Object.keys(d.folders || {}).sort().map((k) => `${k}:${d.folders[k]}`).join('  ');
    if (folderStr) rows.push(['JSONフォルダ構成', folderStr]);
    $('diagGrid').innerHTML = rows.map((r) =>
      `<div class="grow"><span class="k">${r[0]}</span><span class="v">${escapeHtml(String(r[1]))}</span></div>`).join('');

    const pick = (keys, counts) => { const o = {}; keys.forEach((t) => o[t] = counts[t]); return o; };
    const utRows = Object.keys(d.updateTypes).sort().map((k) => `${k}:${d.updateTypes[k]}`).join('  ') || 'なし';
    const keyChips = (arr) => (arr && arr.length) ? arr.map((k) => `<span class="tag2">${escapeHtml(k)}</span>`).join('') : '<span class="k" style="color:var(--ink-soft)">—</span>';

    $('diagTypes').innerHTML =
      '<h4>認識できるコマンド</h4>' + chips(pick(d.knownCmds, d.cmdCounts), false) +
      '<h4>未対応のコマンド（拾えない可能性）</h4>' + chips(pick(d.unknownCmds, d.cmdCounts), true) +
      (d.unknownCells.length ? '<h4>未対応のセル種別</h4>' + chips(pick(d.unknownCells, d.cellCounts), true) : '') +
      '<h4>UpdateType の内訳（「(未指定)」が多いと C↔U の取り違え注意）</h4><span class="tag2">' + escapeHtml(utRows) + '</span>' +
      '<h4>ページJSONの最上位キー</h4>' + keyChips(d.samplePageKeys) +
      '<h4>テーブルJSONの最上位キー</h4>' + keyChips(d.sampleTableKeys);

    $('diagPanel').classList.add('show');
    $('diagPanel').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // ---- 初期化 ----
  wireSlot('slotFgcp', 'inputFgcp', 'fileFgcp', 'fgcp', ['.fgcp', '.zip']);
  wireSlot('slotFgdoc', 'inputFgdoc', 'fileFgdoc', 'fgdoc', ['.xlsx']);

  // ページ全体へのドロップで誤ってブラウザがファイルを開かないように
  ['dragover', 'drop'].forEach((ev) => window.addEventListener(ev, (e) => {
    if (!e.target.closest('.slot')) e.preventDefault();
  }));

  window.addEventListener('beforeunload', killWorker);

  // ---- Service Worker ----
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch(() => {
        const b = $('offlineBadge'); if (b) b.style.display = 'none';
      });
    });
  } else {
    const b = $('offlineBadge'); if (b) b.style.display = 'none';
  }
})();
