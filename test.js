/* Node smoke test: 合成.fgcpでコアを検証 */
global.JSZip = require('jszip');
global.ExcelJS = require('exceljs');
const CrudCore = require('./crud-core.js');
const fs = require('fs');

// --- ページJSON を作るヘルパ（末尾に }//署名 を付ける） ---
function pageJson(obj) {
  return JSON.stringify(obj) + '//SIGNATURE_abcdef0123456789';
}

// ボタン: add/update/delete + 画面遷移、Condition入れ子、ListView参照
const orderPage = {
  PageInfo: { PageLoadCommandList: [] },
  AttachInfos: {
    'A1': {
      ListViewInfo: { TableName: 'オーダー行情報' },
    },
    'B2': {
      CellType: {
        '$type': 'ButtonCellType, Forguncy',
        CommandList: [
          { '$type': 'UpdateDataTableCommand, Forguncy', UpdateType: 'add', TableName: 'オーダー行情報', ShowConfirm: false },
          { '$type': 'ConditionCommand, Forguncy', ConditionAndCommandPairList: [
              { CommandList: [
                { '$type': 'UpdateDataTableCommand, Forguncy', UpdateType: 'update', TableName: '販売計画データ' },
              ] }
            ],
            ElseCommandList: [
              { '$type': 'UpdateDataTableCommand, Forguncy', UpdateType: 'delete', TableName: 'オーダー行情報', ShowConfirm: true },
            ]
          },
          { '$type': 'NavigateCommand, Forguncy', PageName: '配車（振り分け）' },
        ]
      }
    },
    'C3': {
      CellType: { '$type': 'DropDownListCellType, Forguncy', DataSource: '品種マスタ' }
    }
  }
};

// 5テーブル書込みで「要注意」注記が出るページ
const heavyPage = {
  AttachInfos: {
    'X': { CellType: { '$type': 'ButtonCellType', CommandList: [
      { '$type': 'UpdateDataTableCommand', UpdateType: 'update', TableName: 'T1' },
      { '$type': 'UpdateDataTableCommand', UpdateType: 'update', TableName: 'T2' },
      { '$type': 'UpdateDataTableCommand', UpdateType: 'update', TableName: 'T3' },
      { '$type': 'UpdateDataTableCommand', UpdateType: 'add',    TableName: 'T4' },
      { '$type': 'UpdateDataTableCommand', UpdateType: 'delete', TableName: 'T5' },
    ] } }
  }
};

// 参照のみ
const refPage = {
  AttachInfos: { 'L': { BindingInfo: { TableName: '物件マスタ' } } }
};

// 表示専用（操作なし）
const blankPage = { AttachInfos: {} };

async function buildFgcp() {
  const zip = new JSZip();
  zip.file('Pages/オーダー入力.json', pageJson(orderPage));
  zip.file('Pages/粗利明細Excel出力.json', pageJson(heavyPage));
  zip.file('Pages/物件マスタ_一覧.json', pageJson(refPage));
  zip.file('Pages/スタート.json', pageJson(blankPage));
  zip.file('SomeOther/ignore.json', '{}'); // Pages外は無視されるはず
  zip.file('manifest.txt', 'dummy');
  return await zip.generateAsync({ type: 'nodebuffer' });
}

(async () => {
  const fgcp = await buildFgcp();
  fs.writeFileSync('/tmp/test.fgcp', fgcp);

  const result = await CrudCore.generate({
    fgcpBuffer: fgcp,
    fgdocBuffer: null,
    onProgress: (e) => { if (e.stage === 'message') console.log('  …', e.message); }
  });

  console.log('\n=== 抽出結果 ===');
  for (const p of result.pages) {
    console.log(`[${p.name}] area=${p.area}`);
    console.log(`   C=${JSON.stringify(p.C)} U=${JSON.stringify(p.U)} D=${JSON.stringify(p.D)}`);
    console.log(`   R=${JSON.stringify(p.R)} →${JSON.stringify(p.transitions)} note="${p.note}"`);
  }
  console.log('\n=== stats ===', JSON.stringify(result.stats, null, 2));

  // xlsx を書き出して再読込で検証
  await result.workbook.xlsx.writeFile('/tmp/test_output.xlsx');
  const verify = new ExcelJS.Workbook();
  await verify.xlsx.readFile('/tmp/test_output.xlsx');
  console.log('\n=== xlsx 検証 ===');
  console.log('シート:', verify.worksheets.map(w => w.name).join(' / '));
  const mat = verify.getWorksheet('CRUDマトリクス');
  console.log('マトリクス行数:', mat.rowCount, ' autoFilter:', mat.autoFilter);
  console.log('R1C1:', mat.getCell(1,1).value);

  // --- アサーション ---
  const order = result.pages.find(p => p.name === 'オーダー入力');
  const assert = (cond, msg) => { if (!cond) { console.error('❌ FAIL:', msg); process.exitCode = 1; } else console.log('✅', msg); };
  console.log('\n=== アサーション ===');
  assert(order.C.includes('オーダー行情報'), 'add → C にオーダー行情報');
  assert(order.U.includes('販売計画データ'), 'Condition内 update → U に販売計画データ');
  assert(order.D.includes('オーダー行情報'), 'Else内 delete → D にオーダー行情報');
  assert(order.transitions.includes('配車（振り分け）'), 'NavigateCommand → 遷移先');
  assert(order.R.includes('品種マスタ'), 'DropDown DataSource → R');
  assert(order.R.includes('オーダー行情報'), 'C/U/D テーブルも R に含む');
  const heavy = result.pages.find(p => p.name === '粗利明細Excel出力');
  assert(heavy.note.includes('SSC'), 'SSC注記が付与される');
  assert(heavy.note.includes('要注意'), '5テーブル書込みで要注意');
  assert(heavy.area === 'F 粗利分析', '粗利→F 粗利分析: ' + heavy.area);
  assert(result.stats.del >= 1, '削除ページ数 >= 1');
  assert(verify.worksheets.length === 4, '4シート生成');
  console.log('\n完了。/tmp/test_output.xlsx を生成しました。');
})().catch(e => { console.error(e); process.exitCode = 1; });
