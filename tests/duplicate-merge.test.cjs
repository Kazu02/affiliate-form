// 同じ人×同じ案件の重複申請を1件にまとめる処理を、実際の GAS とローカルの偽シートで確かめる。
// 実アカウント・本番APIへの通信は行わない。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const dir = path.resolve(__dirname, '..', 'gas-project');
const source = ['Code.gs', 'AspReconcile.gs', 'DuplicateMerge.gs'].map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');

// 2次元配列をシートに見立てる（1始まり）
function fakeSheet(name, grid) {
  const cell = (r, c) => { while (grid.length < r) grid.push([]); return grid[r - 1][c - 1] ?? ''; };
  const set = (r, c, v) => { while (grid.length < r) grid.push([]); grid[r - 1][c - 1] = v; };
  const lastRow = () => { for (let i = grid.length; i > 0; i--) if ((grid[i - 1] || []).some(v => v !== '' && v != null)) return i; return 0; };
  const lastCol = () => Math.max(0, ...grid.map(r => { for (let j = r.length; j > 0; j--) if (r[j - 1] !== '' && r[j - 1] != null) return j; return 0; }));
  const sheet = {
    grid, deleted: [],
    getName: () => name,
    getLastRow: lastRow,
    getLastColumn: lastCol,
    getDataRange: () => sheet.getRange(1, 1, Math.max(1, lastRow()), Math.max(1, lastCol())),
    deleteRow: r => { grid.splice(r - 1, 1); sheet.deleted.push(r); },
    getRange(r, c, nr = 1, nc = 1) {
      const block = f => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => f(r + i, c + j)));
      return {
        getValues: () => block(cell),
        getValue: () => cell(r, c),
        getNumberFormats: () => block(() => 'General'),
        getFormulas: () => block(() => ''),
        setValues(v) { v.forEach((row, i) => row.forEach((x, j) => set(r + i, c + j, x))); return this; },
        setValue(v) { set(r, c, v); return this; },
        setNumberFormats() { return this; }, setNumberFormat() { return this; },
        setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; }
      };
    },
    setFrozenRows() {}
  };
  return sheet;
}

const H = ['フォーム記号', '受信日時', 'クリック日時', '送信日時', 'お名前', '電話番号', '紹介者名', 'スクショURL', '承認', '代理店'];
function caseGrid(rows) {
  const conf = [['フォーム記号', 'nomukomu'], ['フォーム名', 'ノムコム'], ['代理店コード', '']];
  const g = [];
  const n = Math.max(rows.length + 1, conf.length);
  for (let i = 0; i < n; i++) {
    const left = conf[i] ? [conf[i][0], conf[i][1], '', '', '', ''] : ['', '', '', '', '', ''];
    const right = i === 0 ? H : (rows[i - 1] || H.map(() => ''));
    g.push(left.concat(right));
  }
  return g;
}
const row = (recv, name, tel, ref, shot, appr) => ['nomukomu', recv, '', '', name, tel, ref, shot, appr, ''];

function harness(grid, advRows = []) {
  const main = fakeSheet('設定_ノムコム', grid);
  const sheets = [main];
  const ss = {
    getSheets: () => sheets,
    getSheetByName: n => sheets.find(s => s.getName() === n) || null,
    insertSheet: n => { const s = fakeSheet(n, []); sheets.push(s); return s; }
  };
  const adv = fakeSheet('202608', [['タイトル'], ['受信日時', '広告名', 'お名前', '紹介者名', 'スクショURL', 'トラッキング漏れ', '承認'], ...advRows]);
  const ctx = {
    Logger: { log() {} }, console,
    Utilities: { formatDate: (d) => d.toISOString() },
    SpreadsheetApp: { flush() {}, openById: () => ({ getSheets: () => [adv], getSheetByName: n => (n === '202608' ? adv : null) }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty() {} }) }
  };
  vm.createContext(ctx);
  vm.runInContext(source + '\nvar AGENCY_COLUMN_LABEL = "代理店";', ctx);
  ctx.getOrCreateSpreadsheet = () => ss;
  ctx.syncFormSheetToAgency = () => {};
  ctx.getOrCreateAgencySpreadsheet = () => ({});
  ctx.updateManagementSheet = () => {};
  ctx.buildApplicationStatusSheet = () => {};
  ctx.buildAgencySummarySheet = () => {};
  return { ctx, main, ss, adv };
}

test('空白・敬称の違いは同じ人×同じ案件は1件にまとめ、承認は強いほうを残す', () => {
  const grid = caseGrid([
    row('2026/08/01 10:00:00', '山田 太郎', '', '菅原貴博', 'https://a', ''),
    row('2026/08/01 10:05:00', '別人 花子', '', '村井亮介', 'https://b', '⭕'),
    row('2026/08/01 10:10:00', 'ヤマダタロウ', '', '菅原貴博', 'https://c', ''),
    row('2026/08/01 10:20:00', '山田太郎様', '', '菅原貴博', '', '⭕️'),
    row('2026/08/02 09:00:00', 'やまだたろう', '', '菅原貴博', '', '重複')
  ]);
  // 「山田 太郎」「山田太郎」「山田太郎様」は同じ人。「やまだたろう」は漢字と照合しない（読みは分からない）
  grid[3][6 + 4] = '山田太郎';
  const { ctx, main, adv } = harness(grid, [
    ['2026/08/01 10:00:00', 'ノムコム', '山田 太郎', '菅原貴博', 'https://a', false, false],
    ['2026/08/01 10:10:00', 'ノムコム', '山田太郎', '菅原貴博', 'https://c', false, false],
    ['2026/08/01 10:20:00', 'ノムコム', '山田太郎様', '菅原貴博', '', false, true]
  ]);

  const preview = ctx.mergeDuplicateApplicationsCore_(true);
  assert.equal(preview.groups, 1);
  assert.equal(preview.removed, 2);
  assert.equal(main.grid[2][6 + 4], '別人 花子', '下見では何も変えない');

  const r = ctx.mergeDuplicateApplicationsCore_(false);
  assert.equal(r.removed, 2);
  const names = main.grid.slice(1).map(x => x[6 + 4]).filter(Boolean);
  assert.deepEqual(names, ['山田 太郎', '別人 花子', 'やまだたろう']);
  assert.equal(main.grid[1][6 + 8], '⭕️', '後の申請の承認を引き継ぐ');
  assert.equal(main.grid[0][0], 'フォーム記号', 'A〜F列の設定は動かさない');
  assert.equal(main.grid[1][0], 'フォーム名');
  // 広告主シートは外した2行が消え、残した行に承認が付く
  assert.equal(adv.grid.length, 3);
  assert.equal(adv.grid[2][6], true);
});

test('電話番号が食い違う同名は別人として残す', () => {
  const grid = caseGrid([
    row('2026/08/01 10:00:00', '佐藤一郎', '090-1111-2222', '菅原貴博', 'https://a', ''),
    row('2026/08/03 10:00:00', '佐藤 一郎', '08033334444', '菅原貴博', 'https://b', '')
  ]);
  const { ctx } = harness(grid);
  assert.equal(ctx.mergeDuplicateApplicationsCore_(true).groups, 0);
});

test('申請時: 同じ人の2回目は行を足さず、最初の行の空欄だけ埋めて退避タブへ残す', () => {
  const grid = caseGrid([row('2026/08/01 10:00:00', '鈴木花子', '', '岩本拓也', '', '')]);
  const { ctx, main, ss } = harness(grid);
  const config = { fields: [{ id: 'name', label: 'お名前' }, { id: 'tel', label: '電話番号' }, { id: 'referrer', label: '紹介者名' }] };
  const rowData = ['nomukomu', '2026/08/05 12:00:00', '', '', '鈴木 花子', '', '岩本拓也', 'https://new', '', ''];
  const merged = ctx.dupMergeIncoming_(ss, main, 'nomukomu', rowData, config);
  assert.ok(merged);
  assert.equal(merged.row, 2);
  assert.equal(main.grid[1][6 + 7], 'https://new', '空のスクショ欄を埋める');
  assert.equal(main.grid[1][6 + 1], '2026/08/01 10:00:00', '最初の受信日時は変えない');
  assert.equal(main.getLastRow(), 3, '行は増えない（設定行ぶんの3行のまま）');
  const arch = ss.getSheetByName('重複申請_統合済み');
  assert.equal(arch.getLastRow(), 2);

  const other = ['nomukomu', '2026/08/05 12:00:00', '', '', '田中次郎', '', '岩本拓也', 'https://x', '', ''];
  assert.equal(ctx.dupMergeIncoming_(ss, main, 'nomukomu', other, config), null);
});

test('カタカナとひらがなの違いは同じ名前として扱う', () => {
  const grid = caseGrid([
    row('2026/08/01 10:00:00', 'ヤマダ タロウ', '', '菅原貴博', 'https://a', ''),
    row('2026/08/01 10:03:00', 'やまだたろう', '', '菅原貴博', 'https://b', '')
  ]);
  const { ctx } = harness(grid);
  assert.equal(ctx.mergeDuplicateApplicationsCore_(true).removed, 1);
});

test('広告主シートの日時は「2026/07/27 14:24:15」でも String(Date) の形でも同じ瞬間に読む', () => {
  const { ctx } = harness(caseGrid([]));
  const a = ctx.dupTimeMs_('2026/07/27 14:24:15');
  const b = ctx.dupTimeMs_('Mon Jul 27 2026 14:24:15 GMT+0900 (日本標準時)');
  assert.ok(a > 0);
  assert.equal(a, b);
});

test('広告主シート: 20秒差の2回申請でも最初の申請の行を残す・別の月のシートの重複も消す', () => {
  const grid = caseGrid([
    row('2026/08/01 10:00:00', '高橋 三郎', '', '菅原貴博', 'https://first', ''),
    row('2026/08/01 10:00:20', '高橋三郎', '', '菅原貴博', 'https://second', '')
  ]);
  const { ctx, adv } = harness(grid, [
    ['2026/08/01 10:00:00', 'ノムコム', '高橋 三郎', '菅原貴博', 'https://first', false, false],
    ['2026/08/01 10:00:20', 'ノムコム', '高橋三郎', '菅原貴博', 'https://second', false, false],
    ['2026/08/01 10:00:20', 'ノムコム', '高橋三郎', '菅原貴博', 'https://second', false, false]
  ]);
  const r = ctx.mergeDuplicateApplicationsCore_(false);
  assert.equal(r.advertiserRemoved, 2);
  assert.equal(adv.grid.length, 3);
  assert.equal(adv.grid[2][4], 'https://first');
});

test('広告主シート: 残す申請の行が無く後の申請の行だけ残っていたら、最初の申請の内容に直す', () => {
  const { ctx, adv } = harness(caseGrid([]), [
    ['2026/08/01 10:00:20', 'ノムコム', '高橋三郎', '菅原貴博', 'https://second', false, false]
  ]);
  const result = { advertiserRemoved: 0, advertiserMissing: 0 };
  ctx.formatJST = () => '2026/08/01 10:00:00';
  ctx.dupMergeAdvertiserSheets_([{ caseName: 'ノムコム', name: '高橋 三郎', keptMs: ctx.dupTimeMs_('2026/08/01 10:00:00'),
    keptApproval: '⭕', keptUrl: 'https://first', removeMs: [ctx.dupTimeMs_('2026/08/01 10:00:20')] }], result);
  assert.equal(result.advertiserFixed, 1);
  assert.equal(result.advertiserRemoved, 0);
  assert.deepEqual(adv.grid[2], ['2026/08/01 10:00:00', 'ノムコム', '高橋三郎', '菅原貴博', 'https://first', false, true]);
});
