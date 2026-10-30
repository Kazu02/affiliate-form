// 停止案件の保存専用。元表・設定・案件マスタ・既存集計には書き込まない。
// 毎月20日ごろの確認時に archiveStoppedApplications() を1回実行する。
const STOPPED_ARCHIVE_PROPERTY_ = 'STOPPED_APPLICATIONS_ARCHIVE_ID';
const STOPPED_ARCHIVE_TAB_ = '停止案件_過去申込';
const STOPPED_ARCHIVE_HEADERS_ = ['申込月（JST）', '案件コード', '案件名', '元行照合キー', '受信日時（JST）', '申込者', '紹介者', '承認原値', '取得日時（JST）'];

function archiveStoppedApplications() {
  return stoppedArchiveRun_(false);
}

// 初回検証用。手動実行1回の中で2回同期し、件数・承認原値・元表不変を検証する。
function archiveStoppedApplicationsAndVerify() {
  return stoppedArchiveRun_(true);
}

function stoppedArchiveAssert_(ok, message) {
  if (!ok) throw new Error('停止案件保存: ' + message);
}

function stoppedArchiveHash_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(value), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function stoppedArchiveMonth_(value) {
  if (value === '' || value == null) return '月不明（受信日時なし）';
  if (value instanceof Date) {
    stoppedArchiveAssert_(!isNaN(value.getTime()), '不正な受信日時');
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM');
  }
  const text = String(value).trim();
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:[ T]|$)/.test(text) && !/(?:Z|[+-]\d{2}:?\d{2})$/.test(text)) {
    const m = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    const month = Number(m[2]), day = Number(m[3]);
    stoppedArchiveAssert_(month >= 1 && month <= 12 && day >= 1 && day <= 31, '不正なJST日付');
    return m[1] + '-' + ('0' + month).slice(-2);
  }
  const parsed = new Date(text);
  stoppedArchiveAssert_(!isNaN(parsed.getTime()), '解釈できない受信日時');
  return Utilities.formatDate(parsed, 'Asia/Tokyo', 'yyyy-MM');
}

function stoppedArchiveSnapshot_(main) {
  const sheets = main.getSheets().map(function (sheet) {
    const range = sheet.getDataRange();
    return { id: sheet.getSheetId(), name: sheet.getName(), rows: sheet.getLastRow(), columns: sheet.getLastColumn(),
      sha256: stoppedArchiveHash_([range.getValues(), range.getFormulas()]) };
  });
  return { sha256: stoppedArchiveHash_(sheets), sheets: sheets };
}

function stoppedArchiveCollect_(main) {
  // getCaseMasterSheet_ / getOrCreateSpreadsheet は初期化の書込を持つため使わない。
  const master = main.getSheetByName(CASE_MASTER_SHEET);
  stoppedArchiveAssert_(!!master, '案件マスタがありません');
  const cases = {}, counts = {}, months = {}, sources = [], rows = [], seen = {}, emptyUnlistedSheets = [];
  master.getDataRange().getValues().slice(1).forEach(function (r) {
    const code = String(r[CM_COL_CODE - 1] || '').trim();
    if (!code) return;
    stoppedArchiveAssert_(!Object.prototype.hasOwnProperty.call(cases, code), '案件コード重複');
    cases[code] = { name: String(r[CM_COL_NAME - 1] || code), stopped: r[CM_COL_ACTIVE - 1] !== true };
    if (cases[code].stopped) counts[code] = 0;
  });
  const acquired = new Date(Math.floor(Date.now() / 1000) * 1000);
  let agencyRows = 0, unknownMonthRows = 0;
  main.getSheets().forEach(function (sheet) {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const table = sheet.getDataRange().getValues();
    const config = table.find(function (r) { return String(r[0]) === FORM_CODE_HEADER; });
    const code = config && String(config[1] || '').trim();
    stoppedArchiveAssert_(!!code, '設定シートの案件コードなし: ' + sheet.getName());
    if (!cases[code]) {
      // 旧代理店専用タブは案件マスタ対象外。空であることを確認して記録する。
      stoppedArchiveAssert_(!table.slice(1).some(function (r) { return r.slice(ANSWER_START_COL - 1).some(function (v) { return v !== '' && v != null; }); }), '案件マスタ外の元表に回答あり: ' + sheet.getName());
      emptyUnlistedSheets.push({name:sheet.getName(), code:code, rows:0});
      return;
    }
    if (!cases[code].stopped) return;
    const h = table[0].slice(ANSWER_START_COL - 1);
    const required = ['受信日時', 'お名前', '紹介者名', '承認'];
    required.forEach(function (name) { stoppedArchiveAssert_(h.filter(function (v) { return v === name; }).length === 1, '必須列不正: ' + sheet.getName() + '/' + name); });
    const idx = function (name) { return h.indexOf(name); };
    const duplicates = {};
    let count = 0;
    table.slice(1).forEach(function (full, index) {
      const r = full.slice(ANSWER_START_COL - 1);
      if (!r.some(function (v) { return v !== '' && v != null; })) return;
      const received = r[idx('受信日時')], applicant = r[idx('お名前')], referrer = r[idx('紹介者名')];
      // 承認値はキーに含めない。完全同一の申込も出現回数で別行として保存する。
      const identity = [received, applicant, referrer, r[idx('クリック日時')] || '', r[idx('送信日時')] || ''];
      if (!received && !applicant && !referrer) identity.push('元行', index + 2);
      const digest = stoppedArchiveHash_(identity);
      duplicates[digest] = (duplicates[digest] || 0) + 1;
      const key = main.getId() + ':' + sheet.getSheetId() + ':' + digest + ':' + duplicates[digest];
      stoppedArchiveAssert_(!seen[key], '元行照合キー重複'); seen[key] = true;
      const month = stoppedArchiveMonth_(received);
      rows.push([month, code, cases[code].name, key, received, applicant, referrer, r[idx('承認')], acquired]);
      counts[code]++; count++; months[month] = (months[month] || 0) + 1;
      if (month.indexOf('月不明') === 0) unknownMonthRows++;
      const agency = r[idx(AGENCY_COLUMN_LABEL)];
      if (agency && agency !== AGENCY_DEFAULT) agencyRows++;
    });
    sources.push({sheetId:sheet.getSheetId(), name:sheet.getName(), code:code, rows:count});
  });
  Object.keys(counts).forEach(function (code) { stoppedArchiveAssert_(sources.some(function (s) { return s.code === code; }), '停止案件の元表なし: ' + code); });
  rows.sort(function (a,b) { return String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1])) || a[3].localeCompare(b[3]); });
  return {rows:rows, byCase:counts, byMonth:months, sources:sources, agencyRows:agencyRows, unknownMonthRows:unknownMonthRows, emptyUnlistedSheets:emptyUnlistedSheets};
}

function stoppedArchivePrivacy_(id) {
  const file = DriveApp.getFileById(id);
  stoppedArchiveAssert_(file.getOwner().getEmail() === Session.getEffectiveUser().getEmail(), '保存先の所有者不一致');
  stoppedArchiveAssert_(file.getSharingAccess() === DriveApp.Access.PRIVATE && file.getEditors().length === 0 && file.getViewers().length === 0, '保存先が所有者限定ではありません');
  return {ownerOnly:true, sharingAccess:'PRIVATE', editors:0, viewers:0};
}

function stoppedArchiveDestination_(main) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(STOPPED_ARCHIVE_PROPERTY_);
  let dest;
  if (id) { stoppedArchiveAssert_(id !== main.getId(), '保存先が元表です'); dest = SpreadsheetApp.openById(id); }
  else {
    dest = SpreadsheetApp.create('アフィリンク 停止案件の過去申込（月別保存）');
    props.setProperty(STOPPED_ARCHIVE_PROPERTY_, dest.getId());
    dest.getSheets()[0].setName(STOPPED_ARCHIVE_TAB_);
  }
  stoppedArchivePrivacy_(dest.getId());
  stoppedArchiveAssert_(dest.getSheets().length === 1 && !!dest.getSheetByName(STOPPED_ARCHIVE_TAB_), '保存先タブ構成不一致');
  dest.setSpreadsheetTimeZone('Asia/Tokyo');
  return dest;
}

function stoppedArchiveSync_(dest, collected) {
  const sheet = dest.getSheetByName(STOPPED_ARCHIVE_TAB_);
  const existing = sheet.getLastRow() ? sheet.getDataRange().getValues() : [];
  if (existing.length) stoppedArchiveAssert_(JSON.stringify(existing[0]) === JSON.stringify(STOPPED_ARCHIVE_HEADERS_), '保存先ヘッダー不一致');
  const body = existing.slice(1), positions = {};
  body.forEach(function (r,i) { stoppedArchiveAssert_(r[3] && positions[r[3]] === undefined, '保存先キー不正'); positions[r[3]] = i; });
  let inserted = 0, updated = 0;
  collected.rows.forEach(function (r) {
    if (positions[r[3]] === undefined) { positions[r[3]] = body.length; body.push(r); inserted++; }
    else { body[positions[r[3]]] = r; updated++; }
  });
  const all = [STOPPED_ARCHIVE_HEADERS_].concat(body);
  if (sheet.getMaxRows() < all.length) sheet.insertRowsAfter(sheet.getMaxRows(), all.length - sheet.getMaxRows());
  const range = sheet.getRange(1,1,all.length,STOPPED_ARCHIVE_HEADERS_.length);
  // 年月を日付へ自動変換させず、月の絞込値を文字列で保持する。
  sheet.getRange(1,1,all.length,1).setNumberFormat('@');
  range.setValues(all);
  // = で始まる文字列も原値として保存し、数式にしない。
  all.forEach(function (r,i) { r.forEach(function (v,j) {
    if (typeof v === 'string' && v.charAt(0) === '=') sheet.getRange(i+1,j+1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(v).build());
  }); });
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,STOPPED_ARCHIVE_HEADERS_.length).setFontWeight('bold');
  if (body.length) {
    sheet.getRange(2,5,body.length,1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sheet.getRange(2,9,body.length,1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  }
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  range.createFilter();
  SpreadsheetApp.flush();
  const actual = sheet.getDataRange().getValues().slice(1), keys = {};
  actual.forEach(function (r) { stoppedArchiveAssert_(!keys[r[3]], '書込後キー重複'); keys[r[3]] = r; });
  collected.rows.forEach(function (r, rowIndex) {
    const saved = keys[r[3]];
    stoppedArchiveAssert_(!!saved, '書込後に照合キーが見つかりません');
    r.forEach(function (v, columnIndex) {
      const actual = saved[columnIndex];
      // Sheetsの日時シリアルとJSミリ秒の往復には最大1msの丸めがある。
      const same = v instanceof Date && actual instanceof Date
        ? Math.abs(v.getTime() - actual.getTime()) <= 1
        : stoppedArchiveHash_(v) === stoppedArchiveHash_(actual);
      stoppedArchiveAssert_(same, '書込後の原値不一致: 対象行' + (rowIndex + 1) + '/列' + (columnIndex + 1));
    });
  });
  return {sourceRows:collected.rows.length, archiveRows:actual.length, inserted:inserted, updated:updated, uniqueKeys:actual.length, approvalAndAllValuesMatch:true};
}

function stoppedArchiveRun_(verifyTwice) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    stoppedArchiveAssert_(Session.getEffectiveUser().getEmail() === 'shinhogle@gmail.com', '会社アカウント不一致');
    const main = SpreadsheetApp.openById(REP_STATUS_MAIN_ID);
    const before = stoppedArchiveSnapshot_(main), collected = stoppedArchiveCollect_(main);
    const triggers = ScriptApp.getProjectTriggers().map(function (t) { return [t.getUniqueId(),t.getHandlerFunction()]; }).sort();
    const dest = stoppedArchiveDestination_(main);
    const first = stoppedArchiveSync_(dest, collected);
    let second = null;
    if (verifyTwice) {
      const again = stoppedArchiveCollect_(main);
      second = stoppedArchiveSync_(dest, again);
      stoppedArchiveAssert_(second.inserted === 0 && first.archiveRows === second.archiveRows, '再実行で件数変化');
    }
    const after = stoppedArchiveSnapshot_(main);
    stoppedArchiveAssert_(before.sha256 === after.sha256, '実行中に元表が変わりました（元表は本関数から書込なし）');
    stoppedArchiveAssert_(JSON.stringify(triggers) === JSON.stringify(ScriptApp.getProjectTriggers().map(function (t) { return [t.getUniqueId(),t.getHandlerFunction()]; }).sort()), 'トリガー変化');
    const report = {at:new Date().toISOString(),url:dest.getUrl(),tab:STOPPED_ARCHIVE_TAB_,first:first,second:second,
      byCase:collected.byCase,byMonth:collected.byMonth,sources:collected.sources,agencyRows:collected.agencyRows,unknownMonthRows:collected.unknownMonthRows,emptyUnlistedSheets:collected.emptyUnlistedSheets,
      privacy:stoppedArchivePrivacy_(dest.getId()),sourceBefore:before,sourceAfter:after,sourceUnchanged:true,triggersUnchanged:true,
      monthBoundaryChecks:stoppedArchiveMonth_(new Date('2026-08-31T14:59:59Z')) === '2026-08' && stoppedArchiveMonth_(new Date('2026-08-31T15:00:00Z')) === '2026-09'};
    // 実行ログ1件の上限を避け、JSONを番号付きの短い断片で出す。
    const json = JSON.stringify(report), size = 4000;
    for (let offset = 0; offset < json.length; offset += size) console.log('STOPPED_ARCHIVE_REPORT ' + (offset / size) + ' ' + json.slice(offset, offset + size));
    return report;
  } finally { lock.releaseLock(); }
}
