/* =====================================================================
 * crud-core.js — ページ別CRUD一覧 生成コア (generate_crud.py のJS移植)
 * ---------------------------------------------------------------------
 * ブラウザでもNode(テスト)でも動くよう、JSZip / ExcelJS はグローバル参照。
 * ブラウザ : <script src="vendor/jszip.min.js"> 等で window に載る
 * Node     : global.JSZip = require('jszip') 等を事前に設定
 * ===================================================================== */
(function (root, factory) {
  const JSZipRef  = root.JSZip   || (typeof require !== 'undefined' ? require('jszip')   : null);
  const ExcelRef  = root.ExcelJS || (typeof require !== 'undefined' ? require('exceljs') : null);
  const api = factory(JSZipRef, ExcelRef);
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.CrudCore = api;
})(typeof self !== 'undefined' ? self : this, function (JSZip, ExcelJS) {
  'use strict';

  const DISCLAIMER = '免責: 本ツールおよび生成物は無保証で提供されます。出力結果の正確性・完全性は保証されません。ご利用は自己責任で行ってください。本ツールの使用または使用不能、出力結果の利用によって生じたいかなる損害についても、作者・提供者は一切の責任を負いません。';

  // =====================================================================
  // 1. ページJSON 解析 (fgparse.py 由来)
  // =====================================================================

  function loadJsonText(raw) {
    // 末尾の }//署名 を除去してパース
    const idx = raw.lastIndexOf('}//');
    const clean = idx > 0 ? raw.slice(0, idx + 1) : raw;
    return JSON.parse(clean);
  }

  function shortType(t) {
    if (!t) return '';
    return t.split(',')[0].split('.').pop();
  }

  function describeCmd(c, depth) {
    depth = depth || 0;
    const ind = '  '.repeat(depth);
    const t = shortType(c['$type'] || '');
    let out = [];
    if (t === 'ConditionCommand') {
      for (const pair of (c.ConditionAndCommandPairList || [])) {
        for (const sub of (pair.CommandList || [])) out = out.concat(describeCmd(sub, depth + 1));
      }
      for (const sub of (c.ElseCommandList || [])) out = out.concat(describeCmd(sub, depth + 1));
    } else if (t === 'UpdateDataTableCommand') {
      const ut = c.UpdateType;
      const map = { add: '追加', update: '更新', delete: '削除', None: '更新' };
      let op = map[String(ut)] || '更新';
      if (ut === null || ut === undefined) op = '更新';
      const conf = c.ShowConfirm ? ' (確認ダイアログあり)' : '';
      out.push(`${ind}・DB${op}: [${c.TableName}]${conf}`);
    } else if (t === 'NavigateCommand') {
      out.push(`${ind}・画面遷移→ [${c.PageName}]`);
    } else if (t === 'ShowPopupCommand') {
      out.push(`${ind}・ポップアップ表示→ [${c.PageName || ''}]`);
    } else if (t === 'QueryCommand') {
      out.push(`${ind}・データ検索: [${c.TableName || ''}]`);
    } else if (t === 'LoopCommand') {
      for (const sub of (c.CommandList || [])) out = out.concat(describeCmd(sub, depth + 1));
    }
    return out;
  }

  function parsePageJson(jsonText, pageName) {
    let d;
    try {
      d = loadJsonText(jsonText);
    } catch (e) {
      return { C: [], R: [], U: [], D: [], transitions: [], error: String(e) };
    }

    const cSet = new Set(), uSet = new Set(), dSet = new Set(), rSet = new Set(), trSet = new Set(), dncSet = new Set();

    const attaches = d.AttachInfos || {};
    // ── リストビュー・バインディング・選択肢から R ──
    for (const att of Object.values(attaches)) {
      if (att.ListViewInfo && att.ListViewInfo.TableName) rSet.add(att.ListViewInfo.TableName);
      if (att.BindingInfo && att.BindingInfo.TableName) rSet.add(att.BindingInfo.TableName);
      const ct = att.CellType || {};
      const ctype = shortType(ct['$type'] || '');
      if (ctype === 'DropDownListCellType' || ctype === 'ComboBoxCellType') {
        if (ct.DataSource) rSet.add(ct.DataSource);
      }
    }

    // ── ボタンコマンド + ページロードコマンドから C/U/D/遷移 ──
    let allCmds = [];
    for (const att of Object.values(attaches)) {
      const ct = att.CellType || {};
      if (shortType(ct['$type'] || '') === 'ButtonCellType') {
        for (const c of (ct.CommandList || [])) allCmds = allCmds.concat(describeCmd(c));
      }
    }
    const pi = d.PageInfo || {};
    for (const key of ['LoadCommand', 'PageLoadCommandList', 'PageLoadedCommandList', 'CommandList']) {
      for (const c of (pi[key] || [])) allCmds = allCmds.concat(describeCmd(c));
    }

    for (const line of allCmds) {
      let m;
      if ((m = line.match(/・DB追加: \[([^\]]+)\]/))) cSet.add(m[1].trim());
      if ((m = line.match(/・DB更新: \[([^\]]+)\]/))) uSet.add(m[1].trim());
      if ((m = line.match(/・DB削除: \[([^\]]+)\]/))) {
        const tbl = m[1].trim(); dSet.add(tbl);
        if (!/確認ダイアログあり/.test(line)) dncSet.add(tbl);
      }
      if ((m = line.match(/・画面遷移→ \[([^\]]+)\]/))) trSet.add(m[1].trim());
      if ((m = line.match(/・ポップアップ表示→ \[([^\]]+)\]/))) trSet.add(m[1].trim());
      if ((m = line.match(/・データ検索: \[([^\]]+)\]/))) rSet.add(m[1].trim());
    }

    // C/U/D で操作するテーブルも R に含める
    for (const x of cSet) rSet.add(x);
    for (const x of uSet) rSet.add(x);
    for (const x of dSet) rSet.add(x);

    const cleanName = (n) => String(n).trim().replace(/^\[+|\]+$/g, '').trim();
    const sortTables = (s) => Array.from(s).map(cleanName).filter(Boolean).sort((a, b) => a.localeCompare(b, 'ja'));

    return {
      C: sortTables(cSet),
      R: sortTables(rSet),
      U: sortTables(uSet),
      D: sortTables(dSet),
      Dnoconfirm: sortTables(dncSet),
      transitions: Array.from(trSet).sort((a, b) => a.localeCompare(b, 'ja')),
    };
  }

  // =====================================================================
  // 2. 業務領域分類 (classify_area の移植・優先順位を厳守)
  // =====================================================================

  function classifyArea(pageName, allTablesStr) {
    const p = String(pageName);
    const t = String(allTablesStr);
    const has = (arr) => arr.some((x) => p.indexOf(x) !== -1);
    const hasT = (arr) => arr.some((x) => t.indexOf(x) !== -1);

    if (p.startsWith('D_')) return 'G 代理店';

    if (has(['施工実績'])) return 'J 施工実績';
    if (has(['情報共有', 'お知らせ', '共有フォルダ', '共有情報登録', 'ダウンロード履歴', 'POP-添付ファイル'])) return 'K 情報共有';
    if (has(['粗利', '物件マスタ（合成）', '分析用配車Excel', '分析用配車データ'])) return 'F 粗利分析';
    if (has(['代理店スタートページ', '代理店ＯＤ・ＣＡＳ', '代理店入力物件から', '代理店最終番号', '代理店飛込み',
      'M代理店配車', '代理店配車', '代理店電帳法', 'オーダー開始 (代理店)', 'オーダー入力物件から (代理店)',
      '飛込み発注処理 (代理店)', '飛込み発注処理 (物件代理店)', '新規物件番号取得 (代理店', '商事',
      '代理店標準価格', '販社名称', '代理店ＯＤ', '代理店選択'])) return 'G 代理店';
    if (has(['実績差異', '期計画ファイル', '拠点期別計画', '拠点期別品種別', '年度参照マスタ', '期参照マスタ',
      '月参照マスタ', 'TOP集計', '販売計画保存', '比較一覧', '保存記録', 'メインメニュー', 'お知らせ拡大表示',
      '生産計画ＴＯＰ', '販売計画へ書き込み', '年度期判定'])) return 'H 計画・予算';
    if (has(['運賃', '後追加', '請求書管理', '入金管理', '入金済', 'POP請求', 'POP未請求', '営業通知メール確認',
      'ＣＷ輸送', '追加運賃', '運賃入力', '一律追加運賃', '標準追加運賃', '運賃管理', 'POP分析用DP'])) return 'D 運賃';
    if (has(['物件マスタ_', '新規物件番号取得', '複写物件番号', '物件活動記録', '参照ー物件マスタ', '◎参照販売行',
      '新規行登録', '変更登録', '変更参照', '翌期失注', '一覧 (元)', '刈取担当者', '不整合データ削除',
      '添付拡大物件管理', 'POPUP_物件番号', '重要物件', '◎添付拡大物件管理', '物件番号行検索', '1行目Ｒ売更新'])) return 'A 物件管理';
    if (has(['販売計画', 'オーダー', '飛込み発注', '受注管理', '生産依頼', '確定後変更', 'オーダー変更依頼', '電帳法',
      '☆商社オーダー', '☆オーダー変更', 'オーダー表', '配車（振り分け）', '配車（参照）', '配車（業務',
      'ステータス管理', '業務排他', 'フルサイズ単価修正', '製品Ｓ数補正', '新Nスペーサー対応', 'オーダー別数量'])) return 'B 販売計画・オーダー';
    if (has(['配車', 'DP配車', 'Ａ番', '積込み', '出荷', 'ピッキング', '生産管理一覧', 'POP-ドライバー', '車両・ドライバー',
      'DP車両', 'ＤＰ受取', '朝_7時', '朝7時', '一斉配車', '配車一覧', 'urlsendn', 'urlsendt', '読取り履歴',
      '配車情報詳細', 'DP複数選択', '配車依頼', '配車カレンダー', 'M配車', 'ＣＳＶ出力（ＰＩＣ）'])) return 'C 配車・物流';
    if (has(['ランクマスタ', 'P0P_大ランク', '背番号入力'])) return 'I ランク・失注';
    if (has(['生産品番', 'スペーサー', '品種別数量換算', 'アングル品番', '品種マスタメンテ', 'Neへの一括変換',
      'pop-スペーサー', '品種マスタ登録確認'])) return 'E 生産品番・スペーサー';

    const Lkw = ['都道府県', '市区町村', '郵便番号', '祝日', '祝祭日', '住所検索', '荷受人マスタ', '荷受人メンテ',
      '得意先マスタ', '代理店マスタ_', '場所マスタ_', '担当マスタ_', '拠点マスタ_', '土被りマスタ_', '特別ＣＤマスタ_',
      '物件分類マスタ_', '元図面マスタ_', '個人期計画_', '子ランクマスタ_', '失注マスタ_', '車種マスタメンテ',
      '色マスター', 'ランクマスタ_', '品種マスタ_', 'オーダー行情報_一覧', '配車情報_一覧', '販売計画データ_',
      '共有ツリーメニューデータ_', '施工実績テーブル_一覧', '担当者マスタ登録', 'オーダー状況マスター', 'ユーザー新規登録',
      '物件マスタ_一覧ページ1', '物件マスタメンテナンス', 'アングル品番メンテ', '事業形態登録', '事業形態ツリー',
      '荷受人マスタ_一覧', '荷受人マスタ_登録', '運送会社マスタメンテ', 'マスタメンテナンス', 'マスタメンテ', '車両参照', '修正'];
    if (has(Lkw)) return 'L マスタ管理';

    // テーブル名フォールバック
    if (t.indexOf('施工実績テーブル') !== -1) return 'J 施工実績';
    if (hasT(['情報共有テーブル', '連絡事項_お知らせ_', '共有ツリーメニューデータ', '情報共有データDL_ログ'])) return 'K 情報共有';
    if (hasT(['粗利明細保存テーブル', '分析用車両情報ビュー', 'オーダー別数量合計'])) return 'F 粗利分析';
    if (hasT(['後追加運賃明細', '後追加請求オーダー', '追加運賃テーブル', 'CW輸送費分析ビュー'])) return 'D 運賃';
    if (t.indexOf('OD物件マスタ') !== -1 || t.indexOf('OD販売計画データ') !== -1) return 'G 代理店';
    if (hasT(['期計画ファイル_拠点', '個人期計画', 'ＴＯＰ計画', '保存記録ヘッダ', '拠点期別品種別計画単価マスタ'])) return 'H 計画・予算';
    if (hasT(['生産品番マスタ', 'スペーサー品種別数量変換マスタ', '本体品種別数量換算マスタ', 'アングル品番マスタ'])) return 'E 生産品番・スペーサー';
    if (hasT(['配車情報', 'Ａ番読取り履歴', 'DP別車両マスタ', '車両変更履歴', 'Ｗ複数日選択'])) return 'C 配車・物流';
    if (hasT(['オーダー行情報', 'オーダー変更依頼履歴'])) return 'B 販売計画・オーダー';
    if (t.indexOf('販売計画データ') !== -1) return 'B 販売計画・オーダー';
    if (t.indexOf('物件マスタ') !== -1 && t.indexOf('D_物件') === -1) return 'A 物件管理';
    if (hasT(['ランクマスタ', '子ランクマスタ', '失注マスタ', 'アクションマスタ'])) return 'I ランク・失注';
    if (hasT(['担当マスタ', '拠点マスタ', '代理店マスタ', '品種マスタ', '車種マスタ', '荷受人マスタ',
      '年度参照マスタ', '期参照マスタ', '場所マスタ', '祝日休日'])) return 'L マスタ管理';

    return 'M システム・補助';
  }

  // =====================================================================
  // 3. SSC注記
  // =====================================================================

  const SSC_NOTES = {
    '粗利明細Excel出力': 'SSC: 粗利明細保存テーブル(Cmd10-13)',
    '粗利明細Escel出力_年月指定': 'SSC: 粗利明細保存テーブル(Cmd10-13)',
    '販売計画保存': 'SSC: 保存用生産品番マスタ(Cmd7-9)',
    '拠点期別計画単価作成': 'SSC: 拠点期別品種別計画単価マスタ修正(Cmd3)',
    '実績差異ＴＯＰ': 'SSC: 拠点期別品種別計画単価マスタ修正(Cmd3)',
  };

  function getNote(pageName, cList, uList, dList) {
    const notes = [];
    for (const key of Object.keys(SSC_NOTES)) {
      if (pageName.indexOf(key) !== -1) { notes.push(SSC_NOTES[key]); break; }
    }
    const totalWrite = cList.length + uList.length + dList.length;
    if (totalWrite >= 5) notes.push(`要注意: ${totalWrite}テーブルへ書込み`);
    return notes.join(' / ');
  }

  // =====================================================================
  // 4. .fgcp 解凍 → ページ抽出（フォルダ名の揺れに対応）
  //    PAGE / PAGES / Page / Pages いずれでも拾う。TABLES も収集。
  // =====================================================================

  function pathParts(p) { return String(p).split('/').filter(Boolean); }
  function baseName(p) { const a = pathParts(p); return (a[a.length - 1] || '').replace(/\.json$/i, ''); }
  function isPageJson(p) { const a = pathParts(p); return a.length >= 2 && /^pages?$/i.test(a[0]) && /\.json$/i.test(a[a.length - 1]); }
  function isTableJson(p) { const a = pathParts(p); return a.length >= 2 && /^tables?$/i.test(a[0]) && /\.json$/i.test(a[a.length - 1]); }

  function collectPageEntries(zip) {
    const entries = [];
    zip.forEach((relPath, file) => { if (!file.dir && isPageJson(relPath)) entries.push(file); });
    entries.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
    return entries;
  }

  function collectTableNames(zip) {
    const set = new Set();
    zip.forEach((relPath, file) => { if (!file.dir && isTableJson(relPath)) set.add(baseName(relPath)); });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ja'));
  }

  async function pagesFromZip(zip, onProgress) {
    const entries = collectPageEntries(zip);
    if (entries.length === 0) {
      throw new Error('ページのJSONが見つかりません（PAGE / PAGES / Pages フォルダ）。「対応診断」でフォルダ構成を確認してください。');
    }
    const pages = [];
    let i = 0;
    for (const file of entries) {
      const pageName = baseName(file.name);
      const text = await file.async('string');
      const crud = parsePageJson(text, pageName);
      const allTables = [].concat(crud.C, crud.R, crud.U, crud.D).join(' ');
      pages.push({
        no: null,
        name: pageName,
        C: crud.C, R: crud.R, U: crud.U, D: crud.D,
        Dnoconfirm: crud.Dnoconfirm,
        transitions: crud.transitions,
        area: classifyArea(pageName, allTables),
        note: getNote(pageName, crud.C, crud.U, crud.D),
      });
      i++;
      if (onProgress && (i % 10 === 0 || i === entries.length)) {
        onProgress(i, entries.length);
        await new Promise((r) => setTimeout(r, 0)); // UIスレッドを解放
      }
    }
    return pages;
  }

  async function extractPagesFromFgcp(arrayBuffer, onProgress) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    return pagesFromZip(zip, onProgress);
  }

  // =====================================================================
  // 5. Forguncyドキュメント Excel からページ番号マップ（オプション）
  //    [content_types].xml 小文字問題を JSZip で補修してから ExcelJS で読む
  // =====================================================================

  async function loadPageNumbers(arrayBuffer) {
    let buf = arrayBuffer;
    // XML補修: [content_types].xml(小文字) → [Content_Types].xml
    try {
      const probe = await JSZip.loadAsync(arrayBuffer);
      const hasLower = probe.file('[content_types].xml');
      const hasProper = probe.file('[Content_Types].xml');
      if (hasLower && !hasProper) {
        const data = await hasLower.async('uint8array');
        probe.remove('[content_types].xml');
        probe.file('[Content_Types].xml', data);
        buf = await probe.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      }
    } catch (e) { /* 補修失敗時はそのまま試す */ }

    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(buf);
    } catch (e) {
      return {}; // 読めなければ番号なしで続行
    }

    const map = {};
    wb.eachSheet((ws) => {
      if (Object.keys(map).length > 0) return;
      let headerRow = null, nameCol = null, noCol = null;
      const maxScan = Math.min(ws.rowCount, 20);
      for (let r = 1; r <= maxScan; r++) {
        const row = ws.getRow(r);
        row.eachCell({ includeEmpty: false }, (cell, col) => {
          const val = String(cell.value == null ? '' : cell.value).trim();
          if (val.indexOf('ページ名') !== -1 || val.indexOf('Page') !== -1) { headerRow = r; nameCol = col; }
          if (val === 'No' || val === 'ID' || val.indexOf('番号') !== -1) { noCol = col; }
        });
      }
      if (headerRow && nameCol) {
        for (let r = headerRow + 1; r <= ws.rowCount; r++) {
          const row = ws.getRow(r);
          const nameCell = row.getCell(nameCol).value;
          const name = String(nameCell == null ? '' : nameCell).trim();
          let noVal = noCol ? row.getCell(noCol).value : null;
          if (noVal && typeof noVal === 'object' && 'result' in noVal) noVal = noVal.result;
          if (name && noVal != null && !isNaN(parseInt(noVal, 10))) {
            map[name] = parseInt(noVal, 10);
          }
        }
      }
    });
    return map;
  }

  // =====================================================================
  // 6. xlsx 生成 (ExcelJS) — 4シート構成・色分けを Python版と一致
  // =====================================================================

  const C = {
    HEADER: 'FF1F4E78', AREA: 'FFE2EFDA', NOTE: 'FFFCE4D6',
    C: 'FFC6EFCE', U: 'FFFFEB9C', D: 'FFFFC7CE', DELROW: 'FFFFF2CC',
  };
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const FONT = 'Meiryo UI';
  const F = {
    WHITE_BOLD: { name: FONT, size: 9, bold: true, color: { argb: 'FFFFFFFF' } },
    AREA: { name: FONT, size: 9, bold: true, color: { argb: 'FF375623' } },
    NOTE: { name: FONT, size: 8, color: { argb: 'FF833C00' } },
    DATA: { name: FONT, size: 9 },
    BOLD: { name: FONT, size: 9, bold: true },
    ITALIC_GRAY: { name: FONT, size: 8, italic: true, color: { argb: 'FF595959' } },
  };
  const WRAP_TOP = { wrapText: true, vertical: 'top' };
  const CENTER = { horizontal: 'center', vertical: 'center', wrapText: true };

  function hdr(ws, r, c, text) {
    const cell = ws.getCell(r, c);
    cell.value = text;
    cell.fill = fill(C.HEADER);
    cell.font = F.WHITE_BOLD;
    cell.alignment = CENTER;
  }

  function buildWorkbook(pages) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'CRUD一覧ジェネレーター';
    buildSummary(wb, pages);
    buildMatrix(wb, pages);
    buildTableIndex(wb, pages);
    buildImpact(wb, pages);
    buildDeletePages(wb, pages);
    return wb;
  }

  function areaCounts(pages) {
    const m = {};
    for (const p of pages) m[p.area] = (m[p.area] || 0) + 1;
    return m;
  }

  function buildSummary(wb, pages) {
    const ws = wb.addWorksheet('統計サマリ');
    ws.getColumn(1).width = 38;
    ws.getColumn(2).width = 8;

    ws.getCell(1, 1).value = 'ページ別CRUD一覧'; ws.getCell(1, 1).font = F.BOLD;
    const today = new Date().toISOString().slice(0, 10);
    ws.getCell(2, 1).value = `生成日: ${today}  (CRUD一覧ジェネレーター)`; ws.getCell(2, 1).font = F.ITALIC_GRAY;

    const writePages = pages.filter((p) => p.C.length || p.U.length || p.D.length);
    const rOnly = pages.filter((p) => p.R.length && !p.C.length && !p.U.length && !p.D.length);
    const noOp = pages.filter((p) => !p.C.length && !p.U.length && !p.D.length && !p.R.length);
    const sum = (sel) => pages.reduce((a, p) => a + p[sel].length, 0);

    const stats = [
      ['指標', '値'],
      ['総ページ数', pages.length],
      ['書き込み(C/U/D)を行うページ', writePages.length],
      ['参照のみ(R)のページ', rOnly.length],
      ['表示専用（テーブル操作なし）', noOp.length],
      ['C操作（のべ）', sum('C')],
      ['U操作（のべ）', sum('U')],
      ['D操作（のべ）', sum('D')],
    ];
    stats.forEach((kv, i) => {
      const r = 4 + i;
      ws.getCell(r, 1).value = kv[0]; ws.getCell(r, 1).font = F.DATA;
      ws.getCell(r, 2).value = kv[1]; ws.getCell(r, 2).font = F.DATA;
    });
    ws.getCell(4, 1).font = F.BOLD;
    ws.getCell(4, 2).font = F.BOLD;

    ws.getCell(13, 1).value = 'シート構成'; ws.getCell(13, 1).font = F.BOLD;
    const desc = [
      '・統計サマリ … 本シート',
      '・CRUDマトリクス … 全ページのCRUD一覧（業務領域・注記付き）',
      '・テーブル別逆引き … テーブルから操作ページを逆引き',
      '・削除ページ(要注意) … D操作を行うページのみ抽出',
    ];
    desc.forEach((s, i) => { const c = ws.getCell(14 + i, 1); c.value = s; c.font = F.DATA; });

    const counts = areaCounts(pages);
    ws.getCell(19, 1).value = '─── 業務領域別ページ数 ───'; ws.getCell(19, 1).font = F.BOLD;
    const areas = Object.keys(counts).sort((a, b) => a.localeCompare(b, 'ja'));
    areas.forEach((a, i) => {
      const r = 20 + i;
      ws.getCell(r, 1).value = a; ws.getCell(r, 1).font = F.DATA;
      ws.getCell(r, 2).value = counts[a]; ws.getCell(r, 2).font = F.DATA;
    });
    const last = 20 + areas.length;
    ws.getCell(last, 1).value = '合計'; ws.getCell(last, 1).font = F.BOLD;
    ws.getCell(last, 2).value = areas.reduce((a, k) => a + counts[k], 0); ws.getCell(last, 2).font = F.BOLD;

    const dr = last + 2;
    ws.mergeCells(dr, 1, dr, 2);
    const dc = ws.getCell(dr, 1);
    dc.value = DISCLAIMER;
    dc.font = F.ITALIC_GRAY;
    dc.alignment = { wrapText: true, vertical: 'top' };
    ws.getRow(dr).height = 60;
  }

  function buildMatrix(wb, pages) {
    const ws = wb.addWorksheet('CRUDマトリクス');
    const widths = { 1: 5, 2: 30, 3: 16, 4: 24, 5: 36, 6: 28, 7: 20, 8: 42, 9: 18, 10: 35 };
    for (const k in widths) ws.getColumn(Number(k)).width = widths[k];

    ws.getCell(1, 1).value = 'ページ別CRUD一覧'; ws.getCell(1, 1).font = F.BOLD;
    ws.getCell(2, 1).value = 'C=作成(緑) / R=参照 / U=更新(黄) / D=削除(橙)　※書き込み系を色分け';
    ws.getCell(2, 1).font = F.ITALIC_GRAY;

    const headers = ['No', 'ページ名', '種別', 'C 作成', 'R 参照', 'U 更新', 'D 削除', '→ 遷移先', '業務領域', '注記'];
    headers.forEach((h, i) => hdr(ws, 4, i + 1, h));

    pages.forEach((p, idx) => {
      const r = 5 + idx;
      const cStr = p.C.length ? p.C.join(', ') : null;
      const rStr = p.R.length ? p.R.join(', ') : null;
      const uStr = p.U.length ? p.U.join(', ') : null;
      const dStr = p.D.length ? p.D.join(', ') : null;
      const trStr = p.transitions.length ? p.transitions.join('、') : null;
      const rowData = [p.no, p.name, 'ページ', cStr, rStr, uStr, dStr, trStr, p.area, p.note || null];
      rowData.forEach((v, ci) => {
        const cell = ws.getCell(r, ci + 1);
        cell.value = v;
        cell.font = F.DATA;
        cell.alignment = WRAP_TOP;
      });
      if (cStr) ws.getCell(r, 4).fill = fill(C.C);
      if (uStr) ws.getCell(r, 6).fill = fill(C.U);
      if (dStr) ws.getCell(r, 7).fill = fill(C.D);
      ws.getCell(r, 9).fill = fill(C.AREA);
      ws.getCell(r, 9).font = F.AREA;
      if (p.note) { ws.getCell(r, 10).fill = fill(C.NOTE); ws.getCell(r, 10).font = F.NOTE; }
    });

    ws.autoFilter = `A4:J${4 + pages.length}`;
    ws.views = [{ state: 'frozen', ySplit: 4 }];
  }

  function buildTableIndex(wb, pages) {
    const ws = wb.addWorksheet('テーブル別逆引き');
    [30, 45, 45, 30, 50].forEach((w, i) => ws.getColumn(i + 1).width = w);

    ws.getCell(1, 1).value = 'テーブル別 CRUD逆引き（このテーブルを操作するページ）'; ws.getCell(1, 1).font = F.BOLD;
    const headers = ['テーブル名', 'C 追加するページ', 'U 更新するページ', 'D 削除するページ', 'R 参照するページ'];
    headers.forEach((h, i) => hdr(ws, 2, i + 1, h));

    const tc = {}, tu = {}, td = {}, tr = {};
    const push = (obj, key, val) => { (obj[key] = obj[key] || []).push(val); };
    for (const p of pages) {
      for (const t of p.C) push(tc, t, p.name);
      for (const t of p.U) push(tu, t, p.name);
      for (const t of p.D) push(td, t, p.name);
      const rOnly = p.R.filter((t) => !p.C.includes(t) && !p.U.includes(t) && !p.D.includes(t));
      for (const t of rOnly) push(tr, t, p.name);
    }
    const all = Array.from(new Set([].concat(Object.keys(tc), Object.keys(tu), Object.keys(td), Object.keys(tr))))
      .sort((a, b) => a.localeCompare(b, 'ja'));

    all.forEach((tbl, idx) => {
      const r = 3 + idx;
      const rowData = [
        tbl,
        (tc[tbl] || []).join('、') || null,
        (tu[tbl] || []).join('、') || null,
        (td[tbl] || []).join('、') || null,
        (tr[tbl] || []).join('、') || null,
      ];
      rowData.forEach((v, ci) => {
        const cell = ws.getCell(r, ci + 1);
        cell.value = v; cell.font = F.DATA; cell.alignment = WRAP_TOP;
      });
      if (td[tbl]) ws.getCell(r, 4).fill = fill(C.D);
    });
  }

  function buildImpact(wb, pages) {
    const ws = wb.addWorksheet('改修インパクト');
    [34, 7, 7, 7, 11, 13, 30, 8, 16].forEach((w, i) => ws.getColumn(i + 1).width = w);
    ws.getCell(1, 1).value = 'テーブル別 改修インパクト（このテーブルを直すと何に響くか）'; ws.getCell(1, 1).font = F.BOLD;
    ws.getCell(2, 1).value = '影響度 = (C+U)×1 + D×3 + 参照×0.2。複数領域がさわる／削除されるテーブルは要注意。'; ws.getCell(2, 1).font = F.ITALIC_GRAY;
    ['テーブル名', 'C', 'U', 'D', 'R参照のみ', '操作ページ(のべ)', '業務領域', '影響度', 'リスク'].forEach((h, i) => hdr(ws, 4, i + 1, h));

    const tc = {}, tu = {}, td = {}, tr = {}, areas = {};
    const bump = (o, t) => { o[t] = (o[t] || 0) + 1; };
    for (const p of pages) {
      const touched = new Set([].concat(p.C, p.U, p.D, p.R));
      for (const t of p.C) bump(tc, t);
      for (const t of p.U) bump(tu, t);
      for (const t of p.D) bump(td, t);
      const rOnly = p.R.filter((t) => !p.C.includes(t) && !p.U.includes(t) && !p.D.includes(t));
      for (const t of rOnly) bump(tr, t);
      for (const t of touched) { (areas[t] = areas[t] || new Set()).add(p.area); }
    }
    const all = Array.from(new Set([].concat(Object.keys(tc), Object.keys(tu), Object.keys(td), Object.keys(tr))));
    const rows = all.map((t) => {
      const c = tc[t] || 0, u = tu[t] || 0, d = td[t] || 0, r = tr[t] || 0;
      const aset = areas[t] || new Set();
      const score = Math.ceil((c + u) * 1 + d * 3 + r * 0.2);
      let risk;
      if (d > 0) risk = '高（削除あり）';
      else if (aset.size >= 2) risk = '中（複数領域）';
      else if (c + u >= 3) risk = '中（多書込み）';
      else risk = '低';
      return { t, c, u, d, r, total: c + u + d + r, areas: Array.from(aset).sort((a, b) => a.localeCompare(b, 'ja')), score, risk };
    });
    rows.sort((a, b) => b.score - a.score || a.t.localeCompare(b.t, 'ja'));

    rows.forEach((row, idx) => {
      const rN = 5 + idx;
      const vals = [row.t, row.c || null, row.u || null, row.d || null, row.r || null, row.total, row.areas.join('、'), row.score, row.risk];
      vals.forEach((v, ci) => { const cell = ws.getCell(rN, ci + 1); cell.value = v; cell.font = F.DATA; cell.alignment = WRAP_TOP; });
      if (row.d > 0) ws.getCell(rN, 4).fill = fill(C.D);
      ws.getCell(rN, 7).fill = fill(C.AREA); ws.getCell(rN, 7).font = F.AREA;
      const rc = ws.getCell(rN, 9);
      if (row.risk.startsWith('高')) { rc.fill = fill(C.D); rc.font = F.BOLD; }
      else if (row.risk.startsWith('中')) rc.fill = fill(C.U);
    });
    ws.autoFilter = `A4:I${4 + rows.length}`;
    ws.views = [{ state: 'frozen', ySplit: 4 }];
  }

  function buildDeletePages(wb, pages) {
    const ws = wb.addWorksheet('削除リスク台帳');
    [5, 34, 18, 34, 14, 10].forEach((w, i) => ws.getColumn(i + 1).width = w);
    ws.getCell(1, 1).value = '削除リスク台帳 ― レコード削除(D)を行う箇所。改修時は特に注意'; ws.getCell(1, 1).font = F.BOLD;
    ws.getCell(2, 1).value = '確認ダイアログ「なし」＝誤操作で即削除される高リスク。上位に並べています。'; ws.getCell(2, 1).font = F.ITALIC_GRAY;
    ['No', 'ページ名', '業務領域', '削除テーブル', '確認ダイアログ', 'リスク'].forEach((h, i) => hdr(ws, 4, i + 1, h));

    const rows = [];
    for (const p of pages) {
      for (const t of p.D) {
        const nc = (p.Dnoconfirm || []).includes(t);
        rows.push({ no: p.no, name: p.name, area: p.area, table: t, confirm: nc ? 'なし' : 'あり', risk: nc ? '高' : '中', nc });
      }
    }
    rows.sort((a, b) => (a.nc !== b.nc ? (a.nc ? -1 : 1) : (a.area.localeCompare(b.area, 'ja') || a.name.localeCompare(b.name, 'ja'))));

    rows.forEach((row, idx) => {
      const rN = 5 + idx;
      const vals = [row.no, row.name, row.area, row.table, row.confirm, row.risk];
      vals.forEach((v, ci) => { const cell = ws.getCell(rN, ci + 1); cell.value = v; cell.font = F.DATA; cell.alignment = WRAP_TOP; });
      if (row.nc) {
        ws.getCell(rN, 5).fill = fill(C.D); ws.getCell(rN, 5).font = F.BOLD;
        ws.getCell(rN, 6).fill = fill(C.D); ws.getCell(rN, 6).font = F.BOLD;
      } else {
        ws.getCell(rN, 6).fill = fill(C.U);
      }
    });
    ws.autoFilter = `A4:F${4 + rows.length}`;
    ws.views = [{ state: 'frozen', ySplit: 4 }];
  }

  // ── ① 関連図（Mermaid Markdown）──────────────────────────────
  function mmNode(s) { return String(s).replace(/[^0-9A-Za-z\u3040-\u30FF\u4E00-\u9FFF]/g, '_'); }
  function mmLabel(s) { return String(s).replace(/"/g, '＂'); }

  function buildMermaidDoc(pages, extra) {
    extra = extra || {};
    const today = new Date().toISOString().slice(0, 10);
    const nameToArea = {};
    for (const p of pages) nameToArea[p.name] = p.area;

    // (A) 業務領域フロー: ページ遷移を領域→領域に集約
    const areaEdges = {};
    for (const p of pages) {
      for (const tgt of p.transitions) {
        const ta = nameToArea[tgt];
        if (!ta || ta === p.area) continue;
        const key = p.area + '\u0000' + ta;
        areaEdges[key] = (areaEdges[key] || 0) + 1;
      }
    }
    const allAreas = Array.from(new Set(pages.map((p) => p.area))).sort((a, b) => a.localeCompare(b, 'ja'));

    let flowA = 'flowchart LR\n';
    for (const a of allAreas) flowA += `  ${mmNode(a)}["${mmLabel(a)}"]\n`;
    const edgesA = Object.entries(areaEdges).sort((x, y) => y[1] - x[1]);
    for (const [k, n] of edgesA) {
      const [f, t] = k.split('\u0000');
      flowA += `  ${mmNode(f)} -->|${n}| ${mmNode(t)}\n`;
    }

    // (B) 共有テーブル: 2領域以上がさわるテーブルと、関与領域の関係
    const tblAreas = {};
    for (const p of pages) {
      for (const t of new Set([].concat(p.C, p.U, p.D, p.R))) {
        (tblAreas[t] = tblAreas[t] || new Set()).add(p.area);
      }
    }
    const shared = Object.entries(tblAreas).filter(([, s]) => s.size >= 2)
      .sort((a, b) => b[1].size - a[1].size || a[0].localeCompare(b[0], 'ja'));

    let flowB = 'flowchart LR\n';
    if (shared.length === 0) {
      flowB += '  none["複数領域で共有されるテーブルはありません"]\n';
    } else {
      const usedAreas = new Set();
      for (const [t, s] of shared) for (const a of s) usedAreas.add(a);
      for (const a of usedAreas) flowB += `  ${mmNode('A_' + a)}["${mmLabel(a)}"]\n`;
      for (const [t, s] of shared) {
        flowB += `  ${mmNode('T_' + t)}(["${mmLabel(t)}"]):::shared\n`;
        for (const a of s) flowB += `  ${mmNode('A_' + a)} --- ${mmNode('T_' + t)}\n`;
      }
      flowB += '  classDef shared fill:#FFE08A,stroke:#B5840C,stroke-width:2px;\n';
    }

    // (C) 領域別 主要テーブル（テキスト）
    const ownedByArea = {};
    for (const [t, s] of Object.entries(tblAreas)) {
      if (s.size === 1) { const a = Array.from(s)[0]; (ownedByArea[a] = ownedByArea[a] || []).push(t); }
    }

    let md = `# データ関連図\n\n生成日: ${today} ／ CRUD一覧ジェネレーター\n\n`;
    md += `> ${DISCLAIMER}\n\n`;
    md += `> GitHub上やMermaid対応ビューアでそのまま図として表示されます。\n\n`;
    md += `## 1. 業務フロー（画面遷移を業務領域でまとめたもの）\n\n`;
    md += '```mermaid\n' + flowA + '```\n\n';
    md += `## 2. 共有テーブル（複数の業務領域がさわる＝改修要注意）\n\n`;
    md += '黄色のノードが共有テーブル。1つ直すと複数領域に影響します。\n\n';
    md += '```mermaid\n' + flowB + '```\n\n';
    md += `## 3. 業務領域ごとの専有テーブル（その領域のみが操作）\n\n`;
    for (const a of allAreas) {
      const list = (ownedByArea[a] || []).sort((x, y) => x.localeCompare(y, 'ja'));
      md += `### ${a}\n`;
      md += list.length ? list.map((t) => `- ${t}`).join('\n') + '\n\n' : '（専有テーブルなし）\n\n';
    }
    md += `## 共有テーブル一覧（関与領域つき）\n\n`;
    if (shared.length === 0) md += '（なし）\n';
    else for (const [t, s] of shared) md += `- **${t}** … ${Array.from(s).sort((a, b) => a.localeCompare(b, 'ja')).join('、')}\n`;
    md += '\n';

    // 未使用テーブル（TABLES定義はあるが、どのページも操作しない）
    if (extra.definedTables && extra.definedTables.length) {
      md += `## 未使用テーブル候補（定義はあるが操作なし）\n\n`;
      if (!extra.namesAlign) {
        md += '> テーブル定義名とページの参照名が一致しないため、未使用判定は保留しています。\n\n';
      } else if (!extra.orphanTables || extra.orphanTables.length === 0) {
        md += '（なし。定義済みテーブルはすべていずれかのページで操作されています）\n\n';
      } else {
        md += `定義テーブル ${extra.definedTables.length} 件中、どのページからも操作されていない ${extra.orphanTables.length} 件：\n\n`;
        for (const t of extra.orphanTables) md += `- ${t}\n`;
        md += '\n';
      }
    }
    return md;
  }

  // =====================================================================
  // 7. 対応診断（バージョン互換チェック）
  //    ページJSONを再帰走査し、全$type・UpdateType内訳・構造の有無を集計
  // =====================================================================

  const KNOWN_CMD = new Set(['UpdateDataTableCommand', 'ConditionCommand', 'NavigateCommand', 'ShowPopupCommand', 'QueryCommand', 'LoopCommand']);
  const KNOWN_CELL = new Set(['ButtonCellType', 'DropDownListCellType', 'ComboBoxCellType']);

  function walkTypes(node, acc, depth) {
    if (depth > 200 || node == null) return;
    if (Array.isArray(node)) { for (const x of node) walkTypes(x, acc, depth + 1); return; }
    if (typeof node === 'object') {
      const t = node['$type'];
      if (t) {
        const s = shortType(t);
        acc.types[s] = (acc.types[s] || 0) + 1;
        if (s === 'UpdateDataTableCommand') {
          const ut = node.UpdateType == null ? '(未指定)' : String(node.UpdateType);
          acc.updateTypes[ut] = (acc.updateTypes[ut] || 0) + 1;
        }
      }
      for (const k in node) walkTypes(node[k], acc, depth + 1);
    }
  }

  async function diagnose(arrayBuffer, onProgress) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const pageEntries = [];
    const tableEntries = [];
    const folders = {}; // 直下フォルダ名 → JSON件数
    let anyJson = 0;
    zip.forEach((p, f) => {
      if (f.dir) return;
      const isJson = /\.json$/i.test(p);
      if (isJson) {
        anyJson++;
        const top = pathParts(p)[0] || '(直下)';
        folders[top] = (folders[top] || 0) + 1;
      }
      if (isPageJson(p)) pageEntries.push(f);
      if (isTableJson(p)) tableEntries.push(f);
    });
    pageEntries.sort((a, b) => a.name.localeCompare(b.name, 'ja'));

    const acc = { types: {}, updateTypes: {} };
    let parseOk = 0, withSig = 0, noAttach = 0;
    const parseErrors = [];
    const total = pageEntries.length;
    let samplePageKeys = [];
    let i = 0;
    for (const f of pageEntries) {
      const name = baseName(f.name);
      const raw = await f.async('string');
      if (raw.lastIndexOf('}//') > 0) withSig++;
      let d;
      try { d = loadJsonText(raw); parseOk++; }
      catch (e) { parseErrors.push({ page: name, error: String(e) }); i++; continue; }
      if (samplePageKeys.length === 0 && d && typeof d === 'object') samplePageKeys = Object.keys(d);
      if (!d.AttachInfos || typeof d.AttachInfos !== 'object') noAttach++;
      walkTypes(d, acc, 0);
      i++;
      if (onProgress && (i % 20 === 0 || i === total)) { onProgress(i, total); await new Promise((r) => setTimeout(r, 0)); }
    }

    // テーブルJSONの構造サンプル（最初の1件の最上位キーのみ。値は出さない）
    let sampleTableKeys = [];
    if (tableEntries.length) {
      try {
        const td = loadJsonText(await tableEntries[0].async('string'));
        if (td && typeof td === 'object') sampleTableKeys = Object.keys(td);
      } catch (e) { /* ignore */ }
    }

    const cmds = {}, cells = {}, others = {};
    for (const [t, n] of Object.entries(acc.types)) {
      if (t.endsWith('Command')) cmds[t] = n;
      else if (t.endsWith('CellType')) cells[t] = n;
      else others[t] = n;
    }
    const unknownCmds = Object.keys(cmds).filter((t) => !KNOWN_CMD.has(t)).sort();
    const knownCmds = Object.keys(cmds).filter((t) => KNOWN_CMD.has(t)).sort();
    const unknownCells = Object.keys(cells).filter((t) => !KNOWN_CELL.has(t)).sort();
    const knownCells = Object.keys(cells).filter((t) => KNOWN_CELL.has(t)).sort();

    // 総合判定
    let verdict, verdictText;
    if (total === 0) {
      verdict = 'ng';
      const folderList = Object.keys(folders).join(' / ') || '(なし)';
      verdictText = anyJson > 0
        ? `ページのJSON（PAGE / PAGES / Pages フォルダ）が見つかりません。検出フォルダ: ${folderList}。フォルダ名が異なる可能性があります。`
        : 'JSONが見つかりません。.fgcp ではない、または未対応の形式の可能性があります。';
    } else if (parseOk === 0) {
      verdict = 'ng';
      verdictText = 'ページJSONを1件もパースできませんでした。構造が大きく異なる可能性があります。';
    } else if (noAttach >= Math.ceil(total * 0.5)) {
      verdict = 'ng';
      verdictText = '多くのページで AttachInfos が見つかりません。この版ではセルの格納構造が異なり、抽出できない可能性があります。最上位キーを開発者に共有してください。';
    } else if (unknownCmds.length > 0 || (acc.updateTypes['(未指定)'] || 0) > 0 || noAttach > 0 || parseErrors.length > 0) {
      verdict = 'warn';
      verdictText = 'おおむね使えますが、未対応のコマンドやUpdateType未指定などがあり、一部の書き込みが拾えていない可能性があります。出力件数を実態と照合してください。';
    } else {
      verdict = 'ok';
      verdictText = '認識できる構造です。問題なく使えると見込まれます。';
    }

    return {
      verdict, verdictText,
      pagesFound: total, anyJson, parseOk, parseErrors,
      withSignature: withSig, noAttach,
      folders, tablesFound: tableEntries.length,
      samplePageKeys, sampleTableKeys,
      knownCmds, unknownCmds, cmdCounts: cmds,
      knownCells, unknownCells, cellCounts: cells,
      otherTypes: others,
      updateTypes: acc.updateTypes,
    };
  }

  // =====================================================================
  // 8. オーケストレーター
  // =====================================================================

  async function generate(opts) {
    const { fgcpBuffer, fgdocBuffer, onProgress } = opts;
    const report = (msg) => { if (onProgress) onProgress({ stage: 'message', message: msg }); };

    report('.fgcp を解凍中…');
    const zip = await JSZip.loadAsync(fgcpBuffer);
    const pages = await pagesFromZip(zip, (done, total) => {
      if (onProgress) onProgress({ stage: 'parse', done, total });
    });
    const definedTables = collectTableNames(zip);

    let noMap = {};
    if (fgdocBuffer) {
      report('Forguncyドキュメントからページ番号を取得中…');
      noMap = await loadPageNumbers(fgdocBuffer);
      for (const p of pages) if (noMap[p.name] != null) p.no = noMap[p.name];
    }

    // ページ番号→名前でソート（番号なしは後ろ）
    pages.sort((a, b) => {
      const an = a.no == null, bn = b.no == null;
      if (an !== bn) return an ? 1 : -1;
      if (!an && !bn && a.no !== b.no) return a.no - b.no;
      return a.name.localeCompare(b.name, 'ja');
    });

    // 未使用テーブル判定（TABLES定義はあるがどのページも操作しない）
    const touched = new Set();
    for (const p of pages) for (const t of new Set([].concat(p.C, p.U, p.D, p.R))) touched.add(t);
    const overlap = definedTables.filter((t) => touched.has(t)).length;
    const namesAlign = definedTables.length > 0 && touched.size > 0 && (overlap / touched.size) >= 0.3;
    const orphanTables = namesAlign ? definedTables.filter((t) => !touched.has(t)) : [];

    report('Excel を生成中…');
    const wb = buildWorkbook(pages);
    const mermaid = buildMermaidDoc(pages, { definedTables, orphanTables, namesAlign });

    const stats = {
      total: pages.length,
      write: pages.filter((p) => p.C.length || p.U.length || p.D.length).length,
      del: pages.filter((p) => p.D.length).length,
      numbered: Object.keys(noMap).length,
      definedTables: definedTables.length,
      orphanTables: orphanTables.length,
      tableNamesAlign: namesAlign,
      areas: areaCounts(pages),
    };
    return { workbook: wb, pages, stats, mermaid, definedTables, orphanTables };
  }

  return {
    loadJsonText, shortType, describeCmd, parsePageJson, classifyArea,
    SSC_NOTES, getNote, extractPagesFromFgcp, loadPageNumbers,
    buildWorkbook, buildImpact, buildMermaidDoc, generate, diagnose, walkTypes,
  };
});
