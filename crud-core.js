/* =====================================================================
 * crud-core.js — CAS CRUD一覧 生成コア (generate_crud.py のJS移植)
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

    const cSet = new Set(), uSet = new Set(), dSet = new Set(), rSet = new Set(), trSet = new Set();

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
      if ((m = line.match(/・DB削除: \[([^\]]+)\]/))) dSet.add(m[1].trim());
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
  // 4. .fgcp 解凍 → ページ抽出
  // =====================================================================

  async function extractPagesFromFgcp(arrayBuffer, onProgress) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    // Pages 直下の *.json のみ（サブフォルダ除外）
    const entries = [];
    zip.forEach((relPath, file) => {
      if (/^Pages\/[^/]+\.json$/i.test(relPath) && !file.dir) entries.push(file);
    });
    if (entries.length === 0) {
      throw new Error('Pages フォルダ内の *.json が見つかりません。正しい .fgcp ファイルか確認してください。');
    }
    entries.sort((a, b) => a.name.localeCompare(b.name, 'ja'));

    const pages = [];
    let i = 0;
    for (const file of entries) {
      const pageName = file.name.replace(/^Pages\//i, '').replace(/\.json$/i, '');
      const text = await file.async('string');
      const crud = parsePageJson(text, pageName);
      const allTables = [].concat(crud.C, crud.R, crud.U, crud.D).join(' ');
      pages.push({
        no: null,
        name: pageName,
        C: crud.C, R: crud.R, U: crud.U, D: crud.D,
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
    wb.creator = 'CAS CRUD PWA';
    buildSummary(wb, pages);
    buildMatrix(wb, pages);
    buildTableIndex(wb, pages);
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

    ws.getCell(1, 1).value = 'CASシステム ページ別CRUD一覧'; ws.getCell(1, 1).font = F.BOLD;
    const today = new Date().toISOString().slice(0, 10);
    ws.getCell(2, 1).value = `生成日: ${today}  (CAS CRUD PWA)`; ws.getCell(2, 1).font = F.ITALIC_GRAY;

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
  }

  function buildMatrix(wb, pages) {
    const ws = wb.addWorksheet('CRUDマトリクス');
    const widths = { 1: 5, 2: 30, 3: 16, 4: 24, 5: 36, 6: 28, 7: 20, 8: 42, 9: 18, 10: 35 };
    for (const k in widths) ws.getColumn(Number(k)).width = widths[k];

    ws.getCell(1, 1).value = 'CASシステム ページ別CRUD一覧'; ws.getCell(1, 1).font = F.BOLD;
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

  function buildDeletePages(wb, pages) {
    const ws = wb.addWorksheet('削除ページ(要注意)');
    [5, 38, 40].forEach((w, i) => ws.getColumn(i + 1).width = w);

    ws.getCell(1, 1).value = 'レコード削除(D)を行うページ ― 改修時は特に注意'; ws.getCell(1, 1).font = F.BOLD;
    ['No', 'ページ名', '削除対象テーブル'].forEach((h, i) => hdr(ws, 2, i + 1, h));

    const del = pages.filter((p) => p.D.length);
    del.forEach((p, idx) => {
      const r = 3 + idx;
      const rowData = [p.no, p.name, p.D.join(', ')];
      rowData.forEach((v, ci) => {
        const cell = ws.getCell(r, ci + 1);
        cell.value = v; cell.font = F.DATA; cell.alignment = WRAP_TOP; cell.fill = fill(C.DELROW);
      });
    });
  }

  // =====================================================================
  // 7. オーケストレーター
  // =====================================================================

  async function generate(opts) {
    const { fgcpBuffer, fgdocBuffer, onProgress } = opts;
    const report = (msg) => { if (onProgress) onProgress({ stage: 'message', message: msg }); };

    report('.fgcp を解凍中…');
    const pages = await extractPagesFromFgcp(fgcpBuffer, (done, total) => {
      if (onProgress) onProgress({ stage: 'parse', done, total });
    });

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

    report('Excel を生成中…');
    const wb = buildWorkbook(pages);

    const stats = {
      total: pages.length,
      write: pages.filter((p) => p.C.length || p.U.length || p.D.length).length,
      del: pages.filter((p) => p.D.length).length,
      numbered: Object.keys(noMap).length,
      areas: areaCounts(pages),
    };
    return { workbook: wb, pages, stats };
  }

  return {
    loadJsonText, shortType, describeCmd, parsePageJson, classifyArea,
    SSC_NOTES, getNote, extractPagesFromFgcp, loadPageNumbers,
    buildWorkbook, generate,
  };
});
