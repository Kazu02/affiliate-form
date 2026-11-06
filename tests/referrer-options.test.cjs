// 自社フォームの紹介者選択肢の同期を、実際の GAS とローカルの偽シートで確かめる。
// 実アカウント・本番APIへの通信は行わない。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const source = fs.readFileSync(path.resolve(__dirname, '..', 'gas-project', 'Code.gs'), 'utf8');

function harness(sheets, applied) {
  const props = { REFERRER_OPTIONS_APPLIED: applied };
  const wrap = (name, rows) => ({
    getName: () => name,
    getDataRange: () => ({ getValues: () => rows }),
    getRange: (row, col) => ({ setValue(v) { rows[row - 1][col - 1] = v; return this; } })
  });
  const ss = { getSheets: () => Object.entries(sheets).map(([n, rows]) => wrap(n, rows)) };
  const ctx = {
    Logger: { log() {} },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => props[k] ?? null,
      setProperty: (k, v) => { props[k] = v; }
    }) }
  };
  vm.createContext(ctx);
  vm.runInContext(source, ctx);
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, props };
}

const formSheet = agency => [
  ['フォーム名', 'テスト案件'],
  ['代理店コード', agency],
  ['フィールドID', 'ラベル', '種類', '必須', '', '選択肢'],
  ['name', '氏名', 'text', 'TRUE', '', ''],
  ['referrer', '紹介者', 'select', 'TRUE', '', '柳沢悠貴']
];

test('自社フォームの紹介者欄に営業名簿と社外紹介者2社を並べ、代理店フォームは変えない', () => {
  const own = formSheet('');
  const agency = formSheet('ag01');
  const { ctx, props } = harness({ '設定_test': own, '設定_ag': agency, '回答_test': formSheet('') }, 'old');
  const result = ctx.applyReferrerSelectToJishaSheets(false);
  assert.equal(result.updated, 1);
  const options = own[4][5].split(',');
  const roster = vm.runInContext('JISHA_REFERRER_OPTIONS', ctx).split(',');
  assert.deepEqual(options, roster.concat(['株式会社Wonder Craft', '新市場開発株式会社']));
  assert.equal(roster.length, 10);
  assert.equal(agency[4][5], '柳沢悠貴');
  assert.equal(props.REFERRER_OPTIONS_APPLIED, own[4][5]);
  // 営業名簿そのものには入れない（担当タブ・ダッシュボードは増えない）
  assert.ok(!roster.includes('株式会社Wonder Craft'));
  assert.ok(!roster.includes('新市場開発株式会社'));
});

test('適用済みの選択肢と同じなら起動時の再適用をしない', () => {
  const own = formSheet('');
  const first = harness({ '設定_test': own }, null);
  first.ctx.applyReferrerSelectToJishaSheets(false);
  const again = harness({ '設定_test': own }, first.props.REFERRER_OPTIONS_APPLIED);
  assert.equal(again.ctx.applyReferrerSelectToJishaSheets(false).skipped, true);
});
