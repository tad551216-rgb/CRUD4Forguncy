/* app.js — UIグルー（crud-core.js を呼ぶ） */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);

  const state = { fgcp: null, fgdoc: null, blob: null, filename: null };

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
  function refreshGo() { $('go').disabled = !state.fgcp; $('diag').disabled = !state.fgcp; }
  function showError(msg) { const e = $('error'); e.textContent = '⚠ ' + msg; e.classList.add('show'); }
  function clearError() { $('error').classList.remove('show'); }

  // ---- 進捗 ----
  function setProgress(label, pct) {
    $('progress').classList.add('show');
    $('progLabel').textContent = label;
    $('progPct').textContent = pct == null ? '' : pct + '%';
    $('progBar').style.width = (pct == null ? 8 : pct) + '%';
  }
  function hideProgress() { $('progress').classList.remove('show'); }

  // ---- 実行 ----
  $('go').addEventListener('click', async () => {
    if (!state.fgcp) return;
    clearError();
    $('result').classList.remove('show');
    $('go').disabled = true;
    state.blob = null;
    try {
      setProgress('ファイルを読み込み中…', null);
      const fgcpBuf = await state.fgcp.arrayBuffer();
      const fgdocBuf = state.fgdoc ? await state.fgdoc.arrayBuffer() : null;

      const result = await window.CrudCore.generate({
        fgcpBuffer: fgcpBuf,
        fgdocBuffer: fgdocBuf,
        onProgress: (e) => {
          if (e.stage === 'parse') {
            const pct = Math.round((e.done / e.total) * 90);
            setProgress(`ページ解析中… ${e.done} / ${e.total}`, pct);
          } else if (e.stage === 'message') {
            setProgress(e.message, e.message.indexOf('Excel') >= 0 ? 95 : null);
          }
        }
      });

      setProgress('Excelを書き出し中…', 98);
      const buf = await result.workbook.xlsx.writeBuffer();
      state.blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      state.filename = `ページ別CRUD一覧_${today}.xlsx`;

      setProgress('完了', 100);
      setTimeout(hideProgress, 400);
      showResult(result.stats);
    } catch (err) {
      hideProgress();
      showError(err && err.message ? err.message : String(err));
    } finally {
      $('go').disabled = !state.fgcp;
    }
  });

  function showResult(stats) {
    $('stTotal').textContent = stats.total;
    $('stWrite').textContent = stats.write;
    $('stDel').textContent = stats.del;
    const areas = Object.keys(stats.areas).sort((a, b) => a.localeCompare(b, 'ja'));
    $('areas').innerHTML = areas.map((a) =>
      `<div class="row"><span>${escapeHtml(a)}</span><span>${stats.areas[a]}</span></div>`
    ).join('') + (stats.numbered ? `<div class="row" style="color:var(--ink-soft)"><span>No.付与</span><span>${stats.numbered}件</span></div>` : '');
    $('result').classList.add('show');
    $('result').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function escapeHtml(s) { return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }

  $('download').addEventListener('click', () => {
    if (!state.blob) return;
    const url = URL.createObjectURL(state.blob);
    const a = document.createElement('a');
    a.href = url; a.download = state.filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  });

  // ---- 対応診断 ----
  $('diag').addEventListener('click', async () => {
    if (!state.fgcp) return;
    clearError();
    $('diagPanel').classList.remove('show');
    $('diag').disabled = true;
    try {
      setProgress('診断中…', null);
      const buf = await state.fgcp.arrayBuffer();
      const d = await window.CrudCore.diagnose(buf, (done, total) => {
        setProgress(`診断中… ${done} / ${total}`, Math.round((done / total) * 100));
      });
      setTimeout(hideProgress, 300);
      renderDiag(d);
    } catch (err) {
      hideProgress();
      showError(err && err.message ? err.message : String(err));
    } finally {
      $('diag').disabled = !state.fgcp;
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
      ['AttachInfos 無し', d.noAttach + ' ページ'],
      ['末尾署名あり', d.withSignature + ' / ' + d.pagesFound],
    ];
    $('diagGrid').innerHTML = rows.map((r) =>
      `<div class="grow"><span class="k">${r[0]}</span><span class="v">${escapeHtml(String(r[1]))}</span></div>`).join('');

    const pick = (keys, counts) => { const o = {}; keys.forEach((t) => o[t] = counts[t]); return o; };
    const utRows = Object.keys(d.updateTypes).sort().map((k) => `${k}:${d.updateTypes[k]}`).join('  ') || 'なし';

    $('diagTypes').innerHTML =
      '<h4>認識できるコマンド</h4>' + chips(pick(d.knownCmds, d.cmdCounts), false) +
      '<h4>未対応のコマンド（拾えない可能性）</h4>' + chips(pick(d.unknownCmds, d.cmdCounts), true) +
      (d.unknownCells.length ? '<h4>未対応のセル種別</h4>' + chips(pick(d.unknownCells, d.cellCounts), true) : '') +
      '<h4>UpdateType の内訳（「(未指定)」が多いと C↔U の取り違え注意）</h4><span class="tag2">' + escapeHtml(utRows) + '</span>';

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
