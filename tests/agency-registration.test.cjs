// 実際の HTML / GAS を使い、スプレッドシートとメールだけをローカルの偽物に置き換える。
// 実アカウント・本番APIへの通信は行わない。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const { chromium } = require('playwright');

const project = path.resolve(__dirname, '..');
const source = ['Code.gs', 'CaseAgency.gs', 'AgencyStatus.gs', 'AspReconcile.gs', 'ApprovalGap.gs',
  'SalesDashboard.gs', 'SalesLineReview.gs', 'CustomerLineStats.gs', 'CustomerMenu.gs'].map(name =>
  fs.readFileSync(path.join(project, 'gas-project', name), 'utf8')).join('\n');
const html = fs.readFileSync(path.join(project, 'agency.html'), 'utf8');
const headers = ['代理店コード', '代理店名', '担当者名', 'メールアドレス', 'トークン', '状態', '登録日時', 'リンク集URL'];
const referralHeader = '紹介元の代理店';
const treeHeaders = ['担当営業', '上位代理店コード'];
const signupKey = 'test-only-signup-key';
const input = { action: 'registerAgency', signupKey, agencyName: 'テスト登録会社', personName: 'テスト担当', email: 'new@example.test', salesRep: '岩本拓也' };
const legacy = ['ag05', '既存テスト会社', '既存テスト担当', 'old@example.test', 'test-token-5', '稼働', '2026/09/01 10:00:00', 'https://kazu02.github.io/affiliate-form/links.html?t=test-token-5'];
const copy = value => JSON.parse(JSON.stringify(value));

class Sheet {
  constructor(rows = []) {
    this.rows = copy(rows);
    this.maxColumns = Math.max(8, ...rows.map(row => row.length));
    this.writes = [];
  }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0, ...this.rows.map(row => row.length)); }
  getMaxColumns() { return this.maxColumns; }
  insertColumnsAfter(col, count) {
    assert.equal(col, this.maxColumns, '列の追加は末尾のみ');
    this.maxColumns += count;
  }
  setFrozenRows() {}
  setColumnWidth() {}
  getRange(row, col, height = 1, width = 1) {
    assert.ok(col + width - 1 <= this.maxColumns, '実際のシート幅を超えていない');
    const sheet = this;
    return {
      getValues() {
        return Array.from({ length: height }, (_, r) => Array.from({ length: width }, (_, c) => {
          const value = sheet.rows[row - 1 + r]?.[col - 1 + c] ?? '';
          // Sheets の入力用アポストロフィは読み取り時の文字列に含まれない。
          return typeof value === 'string' && value.startsWith("'") ? value.slice(1) : value;
        }));
      },
      setValues(values) {
        assert.equal(values.length, height);
        for (const line of values) assert.equal(line.length, width);
        sheet.writes.push({ row, col, values: copy(values) });
        values.forEach((line, r) => {
          const target = sheet.rows[row - 1 + r] ||= [];
          line.forEach((value, c) => { target[col - 1 + c] = value; });
        });
        return this;
      },
      setValue(value) { return this.setValues([[value]]); },
      setFontWeight() { return this; },
      setBackground() { return this; },
      setFontColor() { return this; }
    };
  }
}

function harness({ rows = [headers, legacy], key = signupKey, mailFails = false, extraSheets = {}, snapshotKey = 'test-only-snapshot-key', spreadsheetId = 'test-sheet' } = {}) {
  let sheet = rows === null ? null : new Sheet(rows);
  let sequence = 0;
  const mails = [], customerTabs = [];
  let matrixSyncs = 0;
  let locked = false;
  const namedSheets = Object.fromEntries(Object.entries(extraSheets).map(([name, values]) => [name, new Sheet(values)]));
  class FixedDate extends Date {
    constructor(...args) { super(...(args.length ? args : ['2026-09-21T10:00:00Z'])); }
    static now() { return Date.parse('2026-09-21T10:00:00Z'); }
  }
  const context = vm.createContext({
    Date: FixedDate,
    SpreadsheetApp: { flush() { assert.equal(locked, true, '書き込みを反映してからロックを解放'); } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: name => name === 'AGENCY_SIGNUP_KEY' ? key : name === 'AFFILIATE_SNAPSHOT_SECRET' ? snapshotKey : name === 'SPREADSHEET_ID' ? spreadsheetId : null }) },
    LockService: { getScriptLock: () => ({
      waitLock() { assert.equal(locked, false); locked = true; },
      releaseLock() { assert.equal(locked, true); locked = false; }
    }) },
    Utilities: { getUuid: () => 'test-uuid-' + (++sequence) },
    MailApp: { sendEmail(mail) { assert.equal(locked, false, 'メール送信前にロックを解放'); if (mailFails) throw new Error('test-mail-failure'); mails.push(copy(mail)); } },
    Logger: { log() {} },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput(text) { return { text, setMimeType() { return this; } }; }
    }
  });
  vm.runInContext(source, context, { filename: 'agency-gas-under-test.js' });
  Object.assign(context, {
    getOrCreateSpreadsheet: () => ({
      getSheetByName: name => name === '代理店マスタ' ? sheet : namedSheets[name] || null,
      insertSheet: name => name === '代理店マスタ' ? (sheet = new Sheet()) : (namedSheets[name] = new Sheet())
    }),
    _verifyCustomerLineBridgeRequest: () => ({ present: false }),
    ensureAgencyCustomerTab_: person => customerTabs.push(person),
    syncAgencyCaseMatrix: () => { matrixSyncs++; },
    agencyCaseSelection_: () => ['case-one'],
    listActiveCases_: () => [{ code: 'case-one', name: 'テスト案件' }, { code: 'case-two', name: '対象外案件' }],
    countAgencyApplicationsByName_: () => ({})
  });
  context.SpreadsheetApp.openById = () => context.getOrCreateSpreadsheet();
  return {
    context, mails, customerTabs, namedSheets,
    get locked() { return locked; },
    get sheet() { return sheet; },
    get matrixSyncs() { return matrixSyncs; },
    post(payload) { return JSON.parse(context.doPost({ postData: { contents: JSON.stringify(payload) } }).text); },
    agencies() { return copy(context.readAgencies_()); }
  };
}

test('紹介元を保存・読取・応答し、既存行とメール・リンク集案内を保持する', () => {
  const h = harness();
  const response = h.post({ ...input, referringAgency: '  テスト紹介代理店  ' });
  assert.equal(response.result, 'success');
  assert.equal(response.referringAgency, 'テスト紹介代理店');
  assert.equal(response.agencyCode, 'ag06');
  assert.equal(response.caseCount, 1);
  assert.equal(response.mailSent, true);
  assert.deepEqual(h.sheet.rows[1], legacy);
  assert.deepEqual(h.sheet.rows[0], [...headers, referralHeader, ...treeHeaders]);
  const saved = h.agencies().find(agency => agency.code === response.agencyCode);
  assert.equal(saved.referringAgency, response.referringAgency);
  assert.equal(saved.name, input.agencyName);
  assert.equal(saved.person, input.personName);
  assert.equal(saved.email, input.email);
  assert.equal(saved.links, response.linksUrl);
  assert.deepEqual(h.customerTabs, [input.personName]);
  assert.equal(h.matrixSyncs, 1);
  assert.equal(h.mails[0].to, input.email);
  assert.ok(h.mails[0].htmlBody.includes(response.linksUrl));
  assert.ok(h.mails[0].htmlBody.includes('?form=case-one&ag=ag06'));
  assert.ok(!h.mails[0].htmlBody.includes('対象外案件'));
  const plain = harness();
  plain.post(input);
  assert.deepEqual(h.mails, plain.mails, '紹介元を追加してもメール本文・宛先は従来どおり');
});

for (const [label, extra] of [['旧フォーム（項目なし）', {}], ['空文字', { referringAgency: '' }], ['空白', { referringAgency: '　 \t ' }], ['null', { referringAgency: null }]]) {
  test('紹介元が任意: ' + label, () => {
    const h = harness();
    const response = h.post({ ...input, ...extra });
    assert.equal(response.result, 'success');
    assert.equal(response.referringAgency, '');
    assert.equal(h.agencies()[1].referringAgency, '');
    assert.equal(h.mails.length, 1);
  });
}

test('既存マスタの参照だけでは列を追加せず、古い登録内容も読める', () => {
  const h = harness();
  const agency = h.agencies()[0];
  assert.equal(agency.referringAgency, '');
  assert.equal(agency.token, legacy[4]);
  assert.equal(h.sheet.writes.length, 0);
});

test('空の環境でも従来の8列と紹介元を作成できる', () => {
  const h = harness({ rows: null });
  const response = h.post({ ...input, referringAgency: '初回紹介元' });
  assert.equal(response.result, 'success');
  assert.equal(response.agencyCode, 'ag01');
  assert.equal(h.agencies()[0].referringAgency, '初回紹介元');
});

test('運用で足した列を保持し、その後ろへ紹介元を追加する', () => {
  const rows = [[...headers, '運用メモ'], [...legacy, '保持する既存メモ']];
  const h = harness({ rows });
  h.post({ ...input, referringAgency: '紹介会社' });
  assert.deepEqual(h.sheet.rows[1], rows[1]);
  assert.deepEqual(h.sheet.rows[0], [...headers, '運用メモ', referralHeader, ...treeHeaders]);
  assert.equal(h.agencies()[1].referringAgency, '紹介会社');
  h.post({ ...input, email: 'another@example.test', referringAgency: '別の紹介会社' });
  assert.equal(h.sheet.rows[0].filter(label => label === referralHeader).length, 1);
});

test('紹介元列の右に別の列があっても、紹介元を正しい列へ保存する', () => {
  const rows = [[...headers, referralHeader, '運用メモ'], [...legacy, '既存紹介元', '保持するメモ']];
  const h = harness({ rows });
  h.post({ ...input, referringAgency: '新規紹介元' });
  assert.deepEqual(h.sheet.rows[1], rows[1]);
  assert.equal(h.agencies()[1].referringAgency, '新規紹介元');
  assert.equal(h.sheet.rows[0].length, 12);
});

test('同じメールでは重複せず、未入力の紹介元だけ補完し、保存済みは空欄・別名で上書きしない', () => {
  const h = harness();
  const repeat = { ...input, email: legacy[3].toUpperCase() };
  const first = h.post({ ...repeat, referringAgency: '最初の紹介元' });
  assert.equal(first.agencyCode, legacy[0]);
  assert.deepEqual(h.sheet.rows[1].slice(0, 8), legacy);
  for (const extra of [{}, { referringAgency: '' }, { referringAgency: '別の紹介元' }]) {
    const response = h.post({ ...repeat, ...extra });
    assert.equal(response.referringAgency, '最初の紹介元');
    assert.equal(response.linksUrl, legacy[7]);
  }
  assert.equal(h.sheet.getLastRow(), 2);
  assert.equal(h.agencies()[0].referringAgency, '最初の紹介元');
  assert.equal(h.mails.length, 4);
});

for (const value of ['=1+1', '+SUM(1,2)', "'先頭引用符", '<img src=x onerror=alert(1)>']) {
  test('紹介元を文字列で保存して読み戻す: ' + value, () => {
    const h = harness();
    const response = h.post({ ...input, referringAgency: value });
    assert.equal(response.referringAgency, value);
    assert.equal(h.agencies()[1].referringAgency, value);
    assert.equal(h.sheet.rows[2][8], "'" + value, '数式として書き込まない');
  });
}

test('合言葉が不正・未指定・未設定の場合は列追加も登録もメールも行わない', () => {
  for (const [key, submitted] of [[signupKey, 'wrong'], [signupKey, undefined], ['', signupKey]]) {
    const h = harness({ key });
    const response = h.post({ ...input, signupKey: submitted, referringAgency: '紹介会社' });
    assert.equal(response.result, 'error');
    assert.equal(h.sheet.writes.length, 0);
    assert.equal(h.mails.length, 0);
    assert.equal(h.matrixSyncs, 0);
  }
});

test('既存必須項目と紹介元の型・長さをサーバーでも検証する', () => {
  for (const extra of [{ agencyName: '' }, { personName: '' }, { email: 'invalid' }, { referringAgency: {} }, { referringAgency: 123 }, { referringAgency: 'あ'.repeat(201) }]) {
    const h = harness();
    const response = h.post({ ...input, ...extra });
    assert.equal(response.result, 'error');
    assert.equal(h.sheet.writes.length, 0);
    assert.equal(h.mails.length, 0);
  }
  const h = harness();
  assert.equal(h.post({ ...input, referringAgency: 'あ'.repeat(200) }).result, 'success');
});

test('メール失敗時も登録内容とリンク集URLを返す', () => {
  const h = harness({ mailFails: true });
  const response = h.post({ ...input, referringAgency: '紹介会社' });
  assert.equal(response.result, 'success');
  assert.equal(response.mailSent, false);
  assert.ok(response.linksUrl.includes('links.html?t='));
  assert.equal(h.agencies()[1].referringAgency, '紹介会社');
});

test('ブラウザから実GAS経路へ入力・未入力を送信し、保存された紹介元と既存案内を表示する', async t => {
  // このワークスペースのインストール済み Chrome を、一時プロファイルで起動する。
  const browser = await chromium.launch({ headless: true, channel: 'chrome' });
  t.after(() => browser.close());
  const scenarios = [
    { name: '入力あり', value: ' テスト紹介代理店 ', expected: 'テスト紹介代理店' },
    { name: '未入力', value: '', expected: '未入力' },
    { name: '表示はテキスト', value: '<img src=x onerror=alert(1)>', expected: '<img src=x onerror=alert(1)>' },
    { name: 'メール失敗', value: '紹介会社', expected: '紹介会社', mailFails: true },
    { name: '保存済みを表示', value: '別の紹介元', expected: '保存済み紹介元', existing: true }
  ];
  for (const scenario of scenarios) {
    await t.test(scenario.name, async () => {
      const h = harness(scenario.existing
        ? { rows: [[...headers, referralHeader], [...legacy, '保存済み紹介元']] }
        : { mailFails: scenario.mailFails });
      const page = await browser.newPage({ viewport: { width: 390, height: 844 } });
      const errors = [], requests = [];
      page.on('pageerror', error => errors.push(error.message));
      await page.route('**/*', async route => {
        const url = new URL(route.request().url());
        assert.equal(url.origin, 'https://agency.example.test', '外部通信しない');
        if (url.pathname === '/agency.html') return route.fulfill({ contentType: 'text/html', body: html });
        if (url.pathname === '/config.js') return route.fulfill({ contentType: 'application/javascript', body: 'var CONFIG = { gasUrl: "https://agency.example.test/register" };' });
        if (url.pathname === '/register') {
          if (route.request().method() === 'GET') {
            return route.fulfill({ contentType: 'application/json', body: JSON.stringify(h.context.agencyRegistrationPayload_(url.searchParams.get('parent'))) });
          }
          const payload = route.request().postDataJSON();
          requests.push(payload);
          assert.ok(route.request().headers()['content-type'].startsWith('text/plain'));
          return route.fulfill({ contentType: 'application/json', body: JSON.stringify(h.post(payload)) });
        }
        return route.abort();
      });
      try {
        await page.goto('https://agency.example.test/agency.html?k=' + signupKey);
        assert.equal(await page.locator('#referringAgency').getAttribute('required'), null);
        await page.locator('#agencyName').fill(input.agencyName);
        await page.locator('#personName').fill(input.personName);
        await page.locator('#email').fill(scenario.existing ? legacy[3] : input.email);
        await page.locator('#salesRep').selectOption(input.salesRep);
        await page.getByLabel('紹介元の代理店（任意）').fill(scenario.value);
        assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true);
        await page.locator('#submitBtn').click();
        await page.locator('#doneCard.show').waitFor();
        assert.equal(requests.length, 1);
        assert.equal(requests[0].signupKey, signupKey);
        assert.equal(requests[0].referringAgency, scenario.value.trim());
        assert.equal(await page.locator('#doneReferringAgency').textContent(), '登録済みの紹介元の代理店：' + scenario.expected);
        assert.equal(await page.locator('#doneReferringAgency img').count(), 0);
        assert.ok((await page.locator('#doneBody').textContent()).includes(scenario.mailFails ? 'メールの送信ができませんでした' : 'へ専用リンク集のURLをお送りしました'));
        const saved = h.agencies().find(agency => agency.email === (scenario.existing ? legacy[3] : input.email));
        assert.equal(await page.locator('#linksBtn').getAttribute('href'), saved.links);
        assert.equal(saved.referringAgency || '未入力', scenario.expected);
        assert.deepEqual(errors, []);
      } finally { await page.close(); }
    });
  }
});

const appHeaders = ['受信日時', '案件名', '顧客名', '紹介者名', '代理店', '承認', 'スクショURL'];
const agencyRow = (code, salesRep = '岩本拓也', parent = '', extra = {}) => [
  code, extra.name || '作り物代理店' + code, extra.person || '作り物担当' + code,
  code + '@example.test', 'private-token-' + code, extra.status || '稼働',
  '2026/08/01', 'https://agency.example.test/links.html?t=private-token-' + code, salesRep, parent
];
const treeRows = (...rows) => [[...headers, ...treeHeaders], ...rows];
const application = (agency, person, approval = '', date = '2026/09/12 10:00:00') =>
  [date, 'テスト案件', '作り物顧客', person, agency, approval, 'https://shot.example.test/one'];

test('担当営業・親コードの列をヘッダーで読み書きし、不足列だけを作る（再実行は非破壊）', () => {
  const rows = [[...headers, '運用メモ', treeHeaders[1], referralHeader, treeHeaders[0]],
    [...legacy, '保持', '', '過去紹介元', '']];
  const h = harness({ rows });
  const before = copy(h.sheet.rows);
  h.context.ensureAgencyTreeColumns();
  h.context.ensureAgencyTreeColumns();
  assert.deepEqual(h.sheet.rows, before);
  h.context.setAgencyRelationship('ag05', '菅原貴博', '');
  assert.equal(h.agencies()[0].salesRep, '菅原貴博');
  assert.equal(h.sheet.rows[1][8], '保持');
  assert.equal(h.agencies()[0].referringAgency, '過去紹介元');
  const old = harness();
  old.context.ensureAgencyTreeColumns();
  assert.deepEqual(old.sheet.rows[0], [...headers, ...treeHeaders]);
  assert.deepEqual(old.sheet.rows[1], legacy);
  assert.equal(old.locked, false);
});

test('担当営業はサーバーでも必須・名簿外と自由入力を拒否する', () => {
  for (const salesRep of ['', undefined, '自由入力の担当者', { name: '岩本拓也' }]) {
    const h = harness();
    assert.equal(h.post({ ...input, salesRep }).result, 'error');
    assert.equal(h.sheet.writes.length, 0);
    assert.equal(h.mails.length, 0);
  }
});

for (const [label, rows, code, parentCode, message] of [
  ['自分自身', [agencyRow('ag01')], 'ag01', 'ag01', /自分自身/],
  ['A→B→A', [agencyRow('ag01'), agencyRow('ag02', '菅原貴博', 'ag01')], 'ag01', 'ag02', /循環/],
  ['親の既存循環', [agencyRow('ag01', '岩本拓也', 'ag02'), agencyRow('ag02', '岩本拓也', 'ag01'), agencyRow('ag03')], 'ag03', 'ag01', /循環/],
  ['6階層目', Array.from({ length: 6 }, (_, i) => agencyRow('ag0' + (i + 1), '岩本拓也', i > 0 && i < 5 ? 'ag0' + i : '')), 'ag06', 'ag05', /5階層/],
  ['移動で子孫が6階層目', [agencyRow('ag01'), agencyRow('ag02', '岩本拓也', 'ag01'), agencyRow('ag03', '岩本拓也', 'ag02'), agencyRow('ag04', '岩本拓也', 'ag03'), agencyRow('ag05'), agencyRow('ag06', '岩本拓也', 'ag05')], 'ag05', 'ag04', /5階層/]
]) {
  test('親関係を保存前に拒否: ' + label, () => {
    const h = harness({ rows: treeRows(...rows) });
    const before = copy(h.sheet.rows);
    assert.throws(() => h.context.setAgencyRelationship(code, '岩本拓也', parentCode), message);
    assert.deepEqual(h.sheet.rows, before);
    assert.equal(h.sheet.writes.length, 0);
    assert.equal(h.locked, false);
    const response = h.post({ ...input, email: code + '@example.test', parentCode });
    assert.equal(response.result, 'error');
    assert.match(response.message, message);
    assert.deepEqual(h.sheet.rows, before);
    assert.equal(h.mails.length, 0);
    assert.equal(h.matrixSyncs, 0);
  });
}

test('5階層目の新規登録は成功し、6階層目の新規登録は拒否する', () => {
  const h = harness({ rows: treeRows(...Array.from({ length: 4 }, (_, i) => agencyRow('ag0' + (i + 1), '岩本拓也', i ? 'ag0' + i : ''))) });
  const fifth = h.post({ ...input, parentCode: 'ag04' });
  assert.equal(fifth.result, 'success');
  assert.equal(fifth.agencyCode, 'ag05');
  const before = copy(h.sheet.rows);
  assert.equal(h.post({ ...input, email: 'sixth@example.test', parentCode: 'ag05' }).result, 'error');
  assert.deepEqual(h.sheet.rows, before);
});

test('未知・無効・停止中の親でも登録できる。次の採番と同じ未知コードも親なし', () => {
  for (const parentCode of ['unknown', 'ag06', '<invalid>', 'stopped', 'invalid-status']) {
    const h = harness({ rows: treeRows(agencyRow('ag05'), agencyRow('stopped', '岩本拓也', '', { status: '停止' }), agencyRow('invalid-status', '岩本拓也', '', { status: '無効' })) });
    const result = h.post({ ...input, parentCode });
    assert.equal(result.result, 'success', parentCode);
    assert.equal(h.agencies().find(a => a.code === result.agencyCode).parentCode, '');
    assert.equal(result.parentName, null);
  }
});

test('再登録では保存済みの担当・親を上書きせず、未設定なら補完する', () => {
  const h = harness({ rows: treeRows(agencyRow('ag01', '岩本拓也'), agencyRow('ag02', '菅原貴博'), agencyRow('ag03', '', '')) });
  const repeat = { ...input, email: 'ag03@example.test', parentCode: 'ag01' };
  assert.equal(h.post(repeat).result, 'success');
  const saved = h.agencies()[2];
  assert.equal(saved.parentCode, 'ag01');
  assert.equal(saved.salesRep, '岩本拓也');
  h.post({ ...repeat, parentCode: 'ag02', salesRep: '菅原貴博' });
  assert.equal(h.agencies()[2].parentCode, 'ag01');
  assert.equal(h.agencies()[2].salesRep, '岩本拓也');
});

const duplicateAgencies = () => treeRows(
  agencyRow('ag01'),
  agencyRow('ag07', '菅原貴博', 'ag01', { name: '同名代理店', person: '担当甲' }),
  agencyRow('ag08', '村井亮介', '', { name: '同名代理店', person: '担当乙' }),
  agencyRow('ag09', '村井亮介', 'ag07'),
  agencyRow('ag10', '岩本拓也', 'ag09')
);
const statSheets = () => ({
  '申請状況一覧': [appHeaders,
    application('作り物代理店ag01', '作り物担当ag01', '⭕'),
    application('同名代理店', '担当甲', '○'), application('同名代理店', '担当乙', '×'),
    application('ag08', '担当乙', '◯'), application('同名代理店', '担当甲', '⭕', '2026/08/31 23:59:59'),
    application('作り物代理店ag09', '作り物担当ag09'), application('', '岩本', '○')],
  '営業マスタ': [['営業担当', '合言葉（さわらないでください）', '状態', '発行日', 'ダッシュボードURL'],
    ['岩本拓也', 'test-sales-token-aaaaaaaa', '稼働', '', 'private-sales-url'],
    ['菅原貴博', 'test-sales-token-bbbbbbbb', '稼働', '', 'private-sales-url-b']],
  '営業メモ': [['営業担当', '顧客名', 'メモ', '更新日時']],
  '承認漏れ管理': [['型']],
  '案件マスタ': [['案件コード', '案件名', '稼働', '最終更新', '備考'], ['case-one', 'テスト案件', true, '', '']]
});

test('公開の名前解決は指定1件だけ。担当者・連絡先・他社一覧を返さない', () => {
  const h = harness({ rows: duplicateAgencies() });
  for (const parent of ['ag07', 'ag08']) {
    const data = copy(h.context.agencyRegistrationPayload_(parent));
    assert.deepEqual(Object.keys(data), ['salesReps', 'parentName']);
    assert.equal(data.parentName, '同名代理店');
    assert.equal(data.salesReps.length, 10);
    assert.ok(!JSON.stringify(data).includes('担当甲'));
  }
  assert.equal(h.context.agencyRegistrationPayload_('unknown').parentName, null);
  assert.equal(h.context.agencyRegistrationPayload_('').parentName, null);
  assert.equal(h.sheet.writes.length, 0);
});

test('同名ag07/ag08をコードで結び、他担当の子・孫も含め、無関係な枝は出さない', () => {
  const h = harness({ rows: duplicateAgencies(), extraSheets: statSheets() });
  const registered = h.post({ ...input, parentCode: 'ag08' });
  assert.equal(h.agencies().find(a => a.code === registered.agencyCode).parentCode, 'ag08');
  // 新規登録で別枝に増えた自分の担当も、独立した起点になる。
  const result = copy(h.context.salesDashboardPayload_('test-sales-token-aaaaaaaa'));
  assert.deepEqual(result.agencyTree.map(a => [a.code, a.depth]), [['ag01', 0], ['ag07', 1], ['ag09', 2], ['ag10', 3], [registered.agencyCode, 0]]);
  assert.ok(!result.agencyTree.some(a => a.code === 'ag08'));
  assert.equal(result.agencyTree.filter(a => a.code === 'ag10').length, 1);
  assert.deepEqual(result.agencyTree.map(a => a.applications), [1, 1, 1, 0, 0]);
  assert.deepEqual(result.salesTotals.find(r => r.name === '岩本拓也'), { name: '岩本拓也', applications: 2, approved: 2 });
  assert.deepEqual(result.salesTotals.find(r => r.name === '菅原貴博'), { name: '菅原貴博', applications: 1, approved: 1 });
  assert.deepEqual(result.salesTotals.find(r => r.name === '村井亮介'), { name: '村井亮介', applications: 3, approved: 1 });
  assert.equal(result.salesTotals.length, 10);
  for (const secret of ['@example.test', 'private-token-', 'private-sales-url', 'test-sales-token-']) assert.ok(!JSON.stringify(result).includes(secret));
  assert.ok(h.context.salesDashboardPayload_('invalid').error);
});

test('agencyTreeはURLからは取れない。成功応答のキーは契約どおり・全期間の直接件数のみ', () => {
  const h = harness({ rows: duplicateAgencies(), extraSheets: statSheets() });
  // **合言葉をURLに載せる口は塞いである。** GASのリダイレクト先まで秘密値が付いて回るため、
  // 受け渡しは doPost の本文（customerMenuVerifyRequest_）に一本化した。
  const viaUrl = h.context.doGet({ parameter: { action: 'agencyTree', key: 'test-only-snapshot-key' } });
  assert.ok(!String(viaUrl.text).includes('"agencies"'), 'GETで代理店の一覧を返してはいけない');
  // vm の中で作られた配列は外の Array と別物なので、JSON を通してから比べる。
  const result = JSON.parse(JSON.stringify(h.context.agencyTreePayload_()));
  assert.deepEqual(Object.keys(result), ['ok', 'generatedAt', 'agencies']);
  assert.equal(result.generatedAt, '2026/09/21 19:00:00');
  for (const row of result.agencies) {
    assert.deepEqual(Object.keys(row), ['code', 'name', 'contact', 'salesRep', 'parentCode', 'status', 'registeredOn', 'applications', 'approved']);
    assert.equal(row.registeredOn, '2026/08/01');
  }
  assert.equal(result.agencies[0].parentCode, null);
  assert.deepEqual(result.agencies.map(a => [a.code, a.applications, a.approved]), [['ag01', 1, 1], ['ag07', 2, 2], ['ag08', 2, 1], ['ag09', 1, 0], ['ag10', 0, 0]]);
  assert.deepEqual(result, JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures/agencyTree.example.json'), 'utf8')));
  assert.equal(h.sheet.writes.length, 0);
  Object.values(h.namedSheets).forEach(sh => assert.equal(sh.writes.length, 0));
});

test('代理店も担当者も同名で特定不能なら誤集計を返さず、画面では確認中にする', () => {
  const rows = duplicateAgencies();
  rows[3][2] = '担当甲';
  const h = harness({ rows, extraSheets: statSheets() });
  const payload = h.context.agencyTreePayload_();
  assert.equal(payload.ok, false);
  assert.equal(payload.status, 422);
  const dashboard = h.context.salesDashboardPayload_('test-sales-token-aaaaaaaa');
  assert.match(dashboard.agencyNotice, /判別できない/);
  assert.equal(dashboard.agencyTree.find(a => a.code === 'ag07').applications, null);
});

test('列順の異なる申請状況一覧と日付型・月境界を扱い、シート欠落は0件としない', () => {
  const sheets = statSheets();
  const order = [5, 4, 2, 0, 3, 6, 1];
  sheets['申請状況一覧'] = [appHeaders,
    application('ag01', '', '⭕', '2026-08-31T15:00:00Z'),
    application('ag01', '', '⭕', '2026-08-31T14:59:59Z'),
    application('ag01', '', '○', '2026/09/30'),
    application('ag01', '', '', '2026-10-01 00:00:00')].map(r => order.map(i => r[i]));
  const h = harness({ rows: duplicateAgencies(), extraSheets: sheets });
  const result = h.context.agencyApplicationStats_(h.context.readAgencies_(true), '2026/09');
  assert.equal(result.counts.ag01.applications, 2);
  assert.equal(result.counts.ag01.approved, 2);
  delete h.namedSheets['申請状況一覧'];
  assert.equal(h.context.agencyTreePayload_().status, 503);
  assert.match(h.context.sdAgencySummary_('岩本拓也', '2026/09').agencyNotice, /0件ではありません/);
});

test('旧8列のマスタをagencyTreeで読んでも不足列を作らず、担当と親は空として返す', () => {
  const h = harness({ extraSheets: statSheets() });
  const result = h.context.agencyTreePayload_();
  assert.equal(result.ok, true);
  assert.equal(result.agencies[0].salesRep, '');
  assert.equal(result.agencies[0].parentCode, null);
  assert.equal(h.sheet.writes.length, 0);
});

test('agencyTreeは接続失敗・設定欠落でも新しいスプレッドシートやプロパティを作らない', () => {
  for (const spreadsheetId of ['', 'unreachable-sheet']) {
    const h = harness({ spreadsheetId });
    h.context.getOrCreateSpreadsheet = () => { assert.fail('作成しうる共通関数を呼んではいけない'); };
    let reads = 0;
    h.context.SpreadsheetApp.openById = () => { reads++; throw new Error('not accessible'); };
    assert.equal(h.context.agencyTreePayload_().status, 503);
    assert.equal(reads, 1); // プロパティに関係なく定数のメインSSだけを開く
    assert.equal(h.sheet.writes.length, 0);
    assert.equal(h.mails.length, 0);
  }
});

test('紹介URLを既存の合言葉から実行時に発行し、リンク集の案件URLと件数は保つ', () => {
  const h = harness();
  const response = h.context.agencyLinksPayload_(legacy[4]);
  const referral = new URL(response.referralUrl);
  assert.equal(referral.searchParams.get('parent'), legacy[0]);
  assert.equal(referral.searchParams.get('k'), signupKey);
  assert.equal(referral.searchParams.get('t'), null);
  assert.equal(response.cases[0].url, 'https://kazu02.github.io/affiliate-form/?form=case-one&ag=ag05');
  assert.equal(response.totalCount, 0);
});

test('既存snapshotと通常GETのJSONは変更前とバイト単位で等しい', () => {
  const { execFileSync } = require('node:child_process');
  const h = harness({ rows: duplicateAgencies(), extraSheets: statSheets() });
  const baseline = harness({ rows: duplicateAgencies(), extraSheets: statSheets() });
  const beforeCode = execFileSync('git', ['show', 'HEAD:gas-project/Code.gs'], { cwd: project, encoding: 'utf8' });
  const oldGet = beforeCode.slice(beforeCode.indexOf('function doGet(e)'), beforeCode.indexOf('// ---- POST:'));
  vm.runInContext(oldGet, baseline.context);
  for (const context of [h.context, baseline.context]) Object.assign(context, {
    advertiserAdminKey_: () => '', getFirstFormCode: () => 'case-one',
    readConfig: () => ({ title: '作り物フォーム', fields: [] }), isCaseActive_: () => true
  });
  for (const parameter of [{}, { action: 'snapshot' }, { form: 'case-one' }]) {
    assert.equal(h.context.doGet({ parameter }).text, baseline.context.doGet({ parameter }).text);
  }
  // 既存snapshotの**中身**が変わっていないこと。2026-09-21に、合言葉の確認だけを
  // customerMenuVerifyRequest_ へ切り出して代理店の系図と共用にしたので、
  // ソースのバイト一致では見ない（切り出しただけで落ちてしまうため）。
  assert.equal(JSON.stringify(h.context.buildCustomerMenuSnapshot_()), JSON.stringify(baseline.context.buildCustomerMenuSnapshot_()));
  // 合言葉が未設定なら、**組み立てる前に**断る（fail-closed）。系図も同じ確認を通る。
  h.context.buildCustomerMenuSnapshot_ = () => assert.fail('確認の前に組み立ててはいけない');
  assert.deepEqual(JSON.parse(h.context.handleCustomerMenuSnapshot_({}).text), { ok: false, error: 'not_configured' });
  assert.deepEqual(
    JSON.parse(h.context.customerMenuVerifyRequest_({
      issuedAt: new Date().toISOString(), nonce: 'nonce-1234', token: 'x'.repeat(32)
    }).text),
    { ok: false, error: 'not_configured' });
});

test('スマホの紹介URLから担当営業を選び、既知・未知・停止中の親で実登録する', async t => {
  const browser = await chromium.launch({ headless: true, channel: 'chrome' });
  t.after(() => browser.close());
  for (const parentCode of ['ag07', 'ag08', 'unknown', 'stopped']) {
    await t.test(parentCode, async () => {
      const rows = duplicateAgencies();
      rows.push(agencyRow('stopped', '岩本拓也', '', { status: '停止' }));
      const h = harness({ rows });
      const page = await browser.newPage({ viewport: { width: 375, height: 812 } });
      const errors = [], posts = [];
      page.on('pageerror', err => errors.push(err.message));
      await page.route('**/*', async route => {
        const url = new URL(route.request().url());
        assert.equal(url.origin, 'https://agency.example.test');
        if (url.pathname === '/agency.html') return route.fulfill({ contentType: 'text/html', body: html });
        if (url.pathname === '/config.js') return route.fulfill({ contentType: 'application/javascript', body: 'var CONFIG = {gasUrl: "https://agency.example.test/gas"};' });
        if (url.pathname === '/gas') {
          if (route.request().method() === 'GET') return route.fulfill({ contentType: 'application/json', body: h.context.doGet({ parameter: Object.fromEntries(url.searchParams) }).text });
          const payload = route.request().postDataJSON(); posts.push(payload);
          return route.fulfill({ contentType: 'application/json', body: JSON.stringify(h.post(payload)) });
        }
        return route.abort();
      });
      try {
        await page.goto('https://agency.example.test/agency.html?parent=' + parentCode + '&k=' + signupKey);
        await page.waitForFunction(() => !document.getElementById('salesRep').disabled);
        assert.equal(await page.locator('#salesRep').getAttribute('required'), '');
        assert.equal(await page.locator('#referringAgency').isVisible(), false);
        assert.equal(await page.locator('body').innerText().then(text => text.includes(parentCode)), false);
        if (parentCode.startsWith('ag')) assert.equal(await page.locator('#parentInfo').innerText(), '紹介元: 同名代理店（変更できません）');
        else assert.equal(await page.locator('#parentInfo').isVisible(), false);
        await page.locator('#agencyName').fill(input.agencyName);
        await page.locator('#personName').fill(input.personName);
        await page.locator('#email').fill(input.email);
        await page.locator('#submitBtn').click();
        assert.match(await page.locator('#errBanner').innerText(), /担当営業を選択/);
        assert.equal(posts.length, 0);
        await page.locator('#salesRep').selectOption(input.salesRep);
        await page.locator('#submitBtn').click();
        await page.locator('#doneCard.show').waitFor();
        const saved = h.agencies().find(a => a.email === input.email);
        assert.equal(saved.parentCode, parentCode.startsWith('ag') ? parentCode : '');
        assert.equal(posts[0].parentCode, parentCode);
        const referral = new URL(await page.locator('#referralLink').getAttribute('href'));
        assert.equal(referral.searchParams.get('parent'), saved.code);
        assert.equal(referral.searchParams.get('k'), signupKey);
        assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true);
        assert.deepEqual(errors, []);
      } finally { await page.close(); }
    });
  }
});

test('営業画面で4タブ・階層・全員集計・最終申込・未実施を確認し、秘密情報や他社を出さない', async t => {
  const h = harness({ rows: duplicateAgencies(), extraSheets: statSheets() });
  h.context.listActiveCases_ = () => [{ code: 'case-one', name: 'テスト案件' }, { code: 'todo', name: '未実施の稼働案件' }];
  const data = copy(h.context.salesDashboardPayload_('test-sales-token-aaaaaaaa'));
  data.lineEnabled = true;
  data.agencyTree.find(a => a.code === 'ag09').name = '<img src=x onerror=alert(1)>';
  const browser = await chromium.launch({ headless: true, channel: 'chrome' });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 375, height: 812 } });
  const errors = [];
  page.on('pageerror', err => errors.push(err.message));
  await page.route('**/*', async route => {
    const url = new URL(route.request().url());
    assert.equal(url.origin, 'https://sales.example.test');
    if (url.pathname === '/sales.html') return route.fulfill({ contentType: 'text/html', body: fs.readFileSync(path.join(project, 'sales.html'), 'utf8') });
    if (url.pathname === '/config.js') return route.fulfill({ contentType: 'application/javascript', body: 'var CONFIG = {gasUrl: "https://sales.example.test/gas"};' });
    if (url.pathname === '/gas') return route.fulfill({ contentType: 'application/json', body: JSON.stringify(url.searchParams.has('sales_line') ? { available: true, pending: [], customers: [] } : data) });
    return route.abort();
  });
  await page.goto('https://sales.example.test/sales.html?k=test-sales-token-aaaaaaaa');
  await page.locator('#paneCust:not(.hidden)').waitFor();
  assert.equal(await page.locator('.tab:visible').count(), 4);
  assert.match(await page.locator('#custList').innerText(), /最終の申込 2026\/09\/12/);
  assert.match(await page.locator('#custList').innerText(), /承認されました/);
  assert.match(await page.locator('#custList').innerText(), /未実施の稼働案件/);
  await page.locator('#tabAgency').click();
  assert.equal(await page.locator('#paneAgency').isVisible(), true);
  assert.equal(await page.locator('.agency-row').count(), 4);
  assert.deepEqual(await page.locator('.agency-row').evaluateAll(rows => rows.map(row => row.style.getPropertyValue('--depth'))), ['0', '1', '2', '3']);
  assert.match(await page.locator('#agencyTree').innerText(), /今月の申込 1 件 \/ 承認 1 件/);
  assert.ok(!(await page.locator('#agencyTree').innerText()).includes('ag08'));
  assert.equal(await page.locator('#salesTotals tbody tr').count(), 10);
  assert.equal(await page.locator('#agencyTree img').count(), 0);
  for (const secret of ['@example.test', 'private-token-', 'private-sales-url', 'test-sales-token-', 'test-only-signup-key']) assert.ok(!(await page.locator('body').innerText()).includes(secret));
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true);
  await page.locator('#tabLine').click();
  assert.equal(await page.locator('#paneLine').isVisible(), true);
  await page.locator('#tabCheck').click();
  assert.equal(await page.locator('#paneCheck').isVisible(), true);
  assert.deepEqual(errors, []);
});

test('既存代理店のリンク集に紹介URLとコピー操作が出る', async t => {
  const h = harness();
  const data = copy(h.context.agencyLinksPayload_(legacy[4]));
  const browser = await chromium.launch({ headless: true, channel: 'chrome' });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 375, height: 812 } });
  await page.addInitScript(() => {
    Object.defineProperty(navigator, 'clipboard', { value: { writeText: text => { window.copiedText = text; return Promise.resolve(); } } });
  });
  await page.route('**/*', route => {
    const url = new URL(route.request().url());
    assert.equal(url.origin, 'https://links.example.test');
    if (url.pathname === '/links.html') return route.fulfill({ contentType: 'text/html', body: fs.readFileSync(path.join(project, 'links.html'), 'utf8') });
    if (url.pathname === '/config.js') return route.fulfill({ contentType: 'application/javascript', body: 'var CONFIG = {gasUrl: "https://links.example.test/gas"};' });
    if (url.pathname === '/gas') return route.fulfill({ contentType: 'application/json', body: JSON.stringify(data) });
    return route.abort();
  });
  await page.goto('https://links.example.test/links.html?t=' + legacy[4]);
  await page.locator('#referralCard:not(.hidden)').waitFor();
  assert.equal(await page.locator('#referralUrl').getAttribute('href'), data.referralUrl);
  assert.equal(await page.locator('#list .card').count(), 1);
  await page.locator('#copyReferral').click();
  assert.equal(await page.evaluate(() => window.copiedText), data.referralUrl);
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth), true);
});
