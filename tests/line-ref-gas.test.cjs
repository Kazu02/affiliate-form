// LineRef.gs: 公式LINE経由の申請を別シートへ記録し、スナップショット用に読み出す（2026-11-16）。
// 実アカウント・本番APIへの通信は行わない。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const dir = path.resolve(__dirname, '..', 'gas-project');
const source = fs.readFileSync(path.join(dir, 'LineRef.gs'), 'utf8');
const REF = 'b'.repeat(32) + 'fedcba9876543210';

function fakeSpreadsheet() {
  const sheets = {};
  const makeSheet = name => {
    const rows = [];
    const sheet = {
      rows,
      appendRow: r => rows.push(r),
      getLastRow: () => rows.length,
      getRange: (r, c, nr = 1, nc = 1) => ({
        setValues: v => { v.forEach((row, i) => { rows[r - 1 + i] = row; }); return this; },
        setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; },
        getValues: () => rows.slice(r - 1, r - 1 + nr).map(row => row.slice(c - 1, c - 1 + nc)),
      }),
      setFrozenRows() {}, setColumnWidth() {},
    };
    return sheet;
  };
  return {
    sheets,
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => (sheets[n] = makeSheet(n)),
  };
}

function load() {
  const ctx = { Logger: { log() {} }, formatJST: () => '2026/11/16 12:00:00', toDisplayDate_: v => String(v || '') };
  vm.createContext(ctx);
  vm.runInContext(source, ctx);
  return ctx;
}

test('records only a well-formed reference and never throws', () => {
  const g = load();
  const ss = fakeSpreadsheet();
  assert.equal(g.recordLineSubmission_(ss, 'nomukomu', REF), true);
  assert.equal(g.recordLineSubmission_(ss, 'nomukomu', 'not-a-ref'), false);
  assert.equal(g.recordLineSubmission_(ss, '', REF), false);
  assert.equal(g.recordLineSubmission_({ getSheetByName() { throw new Error('boom'); } }, 'nomukomu', REF), false);
  const sh = JSON.parse(JSON.stringify(ss.sheets['LINE経由申請']));
  assert.deepEqual(sh.rows[0], ['受付日時', 'フォーム記号', 'LINE参照']);
  assert.deepEqual(sh.rows.slice(1), [['2026/11/16 12:00:00', 'nomukomu', REF]]);
});

test('snapshot rows carry case code, reference and date; bad rows are skipped', () => {
  const g = load();
  const ss = fakeSpreadsheet();
  assert.deepEqual(JSON.parse(JSON.stringify(g.buildLineSubmissions_(ss))), { rows: [], truncated: false });
  g.recordLineSubmission_(ss, 'nomukomu', REF);
  ss.sheets['LINE経由申請'].rows.push(['2026/11/16 13:00:00', 'bad code!', REF], ['2026/11/16 13:00:00', 'doda', 'x']);
  const out = JSON.parse(JSON.stringify(g.buildLineSubmissions_(ss)));
  assert.deepEqual(out, { rows: [{ caseCode: 'nomukomu', lineRef: REF, submittedOn: '2026/11/16' }], truncated: false });
});

test('same case, reference and day collapse to one row; order stays oldest first', () => {
  const g = load();
  const ss = fakeSpreadsheet();
  g.recordLineSubmission_(ss, 'nomukomu', REF);
  const rows = ss.sheets['LINE経由申請'].rows;
  rows.push(['2026/11/16 15:00:00', 'nomukomu', REF], ['2026/11/17 09:00:00', 'nomukomu', REF], ['2026/11/17 10:00:00', 'doda', REF]);
  const out = JSON.parse(JSON.stringify(g.buildLineSubmissions_(ss)));
  assert.deepEqual(out, { truncated: false, rows: [
    { caseCode: 'nomukomu', lineRef: REF, submittedOn: '2026/11/16' },
    { caseCode: 'nomukomu', lineRef: REF, submittedOn: '2026/11/17' },
    { caseCode: 'doda', lineRef: REF, submittedOn: '2026/11/17' },
  ] });
});

test('keeps the newest rows and reports truncation above the cap', () => {
  const g = load();
  const ss = fakeSpreadsheet();
  g.recordLineSubmission_(ss, 'c0', REF);
  const rows = ss.sheets['LINE経由申請'].rows;
  const max = g.LINE_REF_MAX_ROWS ?? vm.runInContext('LINE_REF_MAX_ROWS', g);
  for (let i = 1; i <= max; i++) rows.push(['2026/11/16 12:00:00', 'c' + i, REF]);
  const out = g.buildLineSubmissions_(ss);
  assert.equal(out.truncated, true);
  assert.equal(out.rows.length, max);
  assert.equal(out.rows[0].caseCode, 'c1');
  assert.equal(out.rows[out.rows.length - 1].caseCode, 'c' + max);
});
