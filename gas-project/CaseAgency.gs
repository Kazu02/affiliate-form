// =============================================
// 案件マスタ / 代理店マスタ / 代理店リンク集
// （2026-08-19 追加）
//
// 設計の要点:
// - 1案件 = 設定タブ1枚。代理店ごとにタブを複製しない。代理店の識別は回答行の
//   「代理店」列で行うため、案件のアフィリエイトURLは全代理店で共有する。
// - 稼働/停止は案件マスタのチェックボックスが唯一の入力。ここを起点に
//   (1) 設定タブの表示/非表示 (2) 代理店リンク集の掲載可否 が同時に決まる。
// - リンク集ページは開くたびにGASへ問い合わせるので、稼働状況の変更は即座に反映される。
// =============================================

const CASE_MASTER_SHEET   = "案件マスタ";
const AGENCY_MASTER_SHEET = "代理店マスタ";
const AGENCY_LINKS_PAGE   = "https://kazu02.github.io/affiliate-form/links.html";
const AGENCY_CODE_PREFIX  = "ag";
const AGENCY_COLUMN_LABEL = "代理店";

// 案件マスタの列（1-based）
const CM_COL_CODE    = 1; // 案件コード（フォーム記号）
const CM_COL_NAME    = 2; // 案件名
const CM_COL_ACTIVE  = 3; // 稼働（チェックボックス）← これが「ボタン」
const CM_COL_UPDATED = 4; // 最終更新
const CM_COL_NOTE    = 5; // 備考
const CM_HEADERS = ["案件コード", "案件名", "稼働", "最終更新", "備考"];

// 代理店マスタの列（1-based）
const AM_COL_CODE   = 1; // 代理店コード
const AM_COL_NAME   = 2; // 代理店名
const AM_COL_PERSON = 3; // 担当者名（＝紹介者名として回答に記録される）
const AM_COL_EMAIL  = 4; // メールアドレス
const AM_COL_TOKEN  = 5; // トークン（リンク集URLの鍵）
const AM_COL_STATUS = 6; // 状態（稼働/停止）
const AM_COL_REGAT  = 7; // 登録日時
const AM_COL_LINKS  = 8; // リンク集URL
const AM_HEADERS = ["代理店コード", "代理店名", "担当者名", "メールアドレス", "トークン", "状態", "登録日時", "リンク集URL"];

const AGENCY_STATUS_ACTIVE = "稼働";
const AGENCY_STATUS_STOP   = "停止";

// =============================================
// 案件マスタ
// =============================================

function getCaseMasterSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(CASE_MASTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CASE_MASTER_SHEET, 1);
    const h = sh.getRange(1, 1, 1, CM_HEADERS.length);
    h.setValues([CM_HEADERS]);
    h.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(CM_COL_CODE, 160);
    sh.setColumnWidth(CM_COL_NAME, 280);
    sh.setColumnWidth(CM_COL_ACTIVE, 70);
    sh.setColumnWidth(CM_COL_UPDATED, 170);
    sh.setColumnWidth(CM_COL_NOTE, 320);
  }
  return sh;
}

// 自社（代理店コード無し）の設定シートだけを列挙する。
// 旧・代理店専用タブ（設定_○○（小中様）など）は案件として扱わない。
//
// **設定はシート先頭の数行にしかないので、そこだけ読む。**
// 以前は getFormCodeFromSheet / getAgencyCode / 表示名の3回とも
// sheet.getDataRange().getValues() を呼んでおり、回答データを含む全行×全列を
// シート31枚ぶん3回読んでいた。代理店登録が56秒かかり、ブラウザ側が
// エラーに見える状態になっていた（2026-08-20）。
const CASE_CONFIG_SCAN_ROWS = 14; // 設定は8行目まで。余裕を見て14行。
let _caseSheetsCache = null;      // 1回の実行の中だけ使う

function listCaseSheets_(ss) {
  if (_caseSheetsCache) return _caseSheetsCache;
  const out = [];
  ss.getSheets().forEach(function (sheet) {
    const sheetName = sheet.getName();
    if (!sheetName.startsWith(CONFIG_PREFIX)) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;
    const rows = Math.min(lastRow, CASE_CONFIG_SCAN_ROWS);

    let code = "", agency = "", displayName = "";
    try {
      const conf = sheet.getRange(1, 1, rows, 2).getValues();
      for (let i = 0; i < conf.length; i++) {
        const k = String(conf[i][0]);
        const v = conf[i][1];
        if (k === FORM_CODE_HEADER && v) code = String(v).trim();
        else if (k === AGENCY_KEY)      agency = v == null ? "" : String(v).trim();
        else if (k === FORM_NAME_KEY && v) displayName = String(v).trim();
      }
    } catch (e) { return; }

    // フォーム記号が無いシートはシート名から補う（getFormCodeFromSheet と同じ規則）
    if (!code) code = sheetName.replace(CONFIG_PREFIX, "");
    if (!code) return;
    if (agency && agency !== AGENCY_DEFAULT) return;
    if (!displayName) displayName = code;

    out.push({ code: code, name: displayName, sheet: sheet });
  });
  _caseSheetsCache = out;
  return out;
}

// 案件マスタを設定タブから同期する。稼働・最終更新・備考の既存値は保持する。
function syncCaseMaster() {
  const ss = getOrCreateSpreadsheet();
  const sh = getCaseMasterSheet_();

  const prev = {};
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, CM_HEADERS.length).getValues().forEach(function (r) {
      const code = String(r[CM_COL_CODE - 1] || "").trim();
      if (!code) return;
      prev[code] = {
        active:  r[CM_COL_ACTIVE - 1] === true,
        updated: r[CM_COL_UPDATED - 1] || "",
        note:    r[CM_COL_NOTE - 1] || ""
      };
    });
  }

  const cases = listCaseSheets_(ss);
  const body = cases.map(function (c) {
    const p = prev[c.code] || { active: false, updated: "", note: "" };
    return [c.code, c.name, p.active, p.updated, p.note];
  });

  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, CM_HEADERS.length).clearContent();
    sh.getRange(2, CM_COL_ACTIVE, lastRow - 1, 1).clearDataValidations();
  }
  if (body.length) {
    sh.getRange(2, 1, body.length, CM_HEADERS.length).setValues(body);
    const cb = sh.getRange(2, CM_COL_ACTIVE, body.length, 1);
    cb.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    cb.setHorizontalAlignment("center");
  }
  return { count: body.length };
}

function readCaseActiveMap_() {
  const sh = getCaseMasterSheet_();
  const map = {};
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return map;
  sh.getRange(2, 1, lastRow - 1, CM_HEADERS.length).getValues().forEach(function (r) {
    const code = String(r[CM_COL_CODE - 1] || "").trim();
    if (code) map[code] = r[CM_COL_ACTIVE - 1] === true;
  });
  return map;
}

function isCaseActive_(formCode) {
  const map = readCaseActiveMap_();
  return map[String(formCode || "").trim()] === true;
}

// 稼働中の案件（コード・名前）を返す。
// **案件マスタだけを1回読む。** 案件マスタに案件コードと案件名の両方があるので、
// 設定シートを1枚ずつ開く必要がない。リンク集は代理店が開くたびに呼ばれるため、
// ここが重いとページが十数秒かかる（2026-08-20 に実測して差し替えた）。
function listActiveCases_() {
  const sh = getCaseMasterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rows = sh.getRange(2, 1, lastRow - 1, CM_HEADERS.length).getValues();
  const out = [];
  rows.forEach(function (r) {
    const code = String(r[CM_COL_CODE - 1] || "").trim();
    if (!code) return;
    if (r[CM_COL_ACTIVE - 1] !== true) return;
    out.push({ code: code, name: String(r[CM_COL_NAME - 1] || code).trim() });
  });
  return out;
}

// 稼働状態を設定タブの表示/非表示へ反映する。
// 旧・代理店専用タブは 1案件1タブへ集約したため常に非表示にする。
function applyCaseVisibility() {
  const ss = getOrCreateSpreadsheet();
  const active = readCaseActiveMap_();
  const result = { shown: [], hidden: [], hiddenAgencyTabs: [] };

  ss.getSheets().forEach(function (sheet) {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;

    // 設定は先頭数行にしかない。回答データまで読まない（listCaseSheets_ と同じ理由）。
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;
    let code = "", agency = "";
    try {
      const conf = sheet.getRange(1, 1, Math.min(lastRow, CASE_CONFIG_SCAN_ROWS), 2).getValues();
      for (let i = 0; i < conf.length; i++) {
        const k = String(conf[i][0]);
        if (k === FORM_CODE_HEADER && conf[i][1]) code = String(conf[i][1]).trim();
        else if (k === AGENCY_KEY) agency = conf[i][1] == null ? "" : String(conf[i][1]).trim();
      }
    } catch (e) { return; }
    if (!code) code = name.replace(CONFIG_PREFIX, "");
    if (!code) return;

    if (agency && agency !== AGENCY_DEFAULT) {
      if (!sheet.isSheetHidden()) sheet.hideSheet();
      result.hiddenAgencyTabs.push(name);
      return;
    }
    if (active[code] === true) {
      if (sheet.isSheetHidden()) sheet.showSheet();
      result.shown.push(name);
    } else {
      if (!sheet.isSheetHidden()) sheet.hideSheet();
      result.hidden.push(name);
    }
  });
  return result;
}

// 「稼働」チェックボックスを押した瞬間に反映する（＝ボタンの実体）。
// 単純トリガーなので、同じスプレッドシートの表示切替だけを行い認可が要る処理はしない。
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CASE_MASTER_SHEET) return;
    if (e.range.getColumn() !== CM_COL_ACTIVE) return;
    const row = e.range.getRow();
    if (row < 2) return;

    const ss   = sheet.getParent();
    const code = String(sheet.getRange(row, CM_COL_CODE).getValue() || "").trim();
    if (!code) return;
    const isActive = sheet.getRange(row, CM_COL_ACTIVE).getValue() === true;

    const target = getConfigSheetByCode(ss, code);
    if (target) {
      if (isActive) { if (target.isSheetHidden()) target.showSheet(); }
      else          { if (!target.isSheetHidden()) target.hideSheet(); }
    }
    sheet.getRange(row, CM_COL_UPDATED).setValue(formatJST(new Date()));
  } catch (err) {
    Logger.log("onEdit(案件マスタ): " + err);
  }
}

function syncCaseMasterAndApply() {
  const r1 = syncCaseMaster();
  const r2 = applyCaseVisibility();
  SpreadsheetApp.getUi().alert(
    "案件マスタを同期しました。\n\n" +
    "案件数: " + r1.count + "\n" +
    "表示: " + r2.shown.length + " 件\n" +
    "非表示: " + r2.hidden.length + " 件\n" +
    "旧・代理店専用タブ（常時非表示）: " + r2.hiddenAgencyTabs.length + " 件"
  );
}

function applyCaseVisibilityFromMenu() {
  const r = applyCaseVisibility();
  SpreadsheetApp.getUi().alert(
    "稼働状況をシート表示へ反映しました。\n\n" +
    "表示: " + r.shown.length + " 件\n" +
    "非表示: " + r.hidden.length + " 件\n" +
    "旧・代理店専用タブ: " + r.hiddenAgencyTabs.length + " 件（常時非表示）"
  );
}

// =============================================
// 代理店マスタ
// =============================================

function getAgencyMasterSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(AGENCY_MASTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AGENCY_MASTER_SHEET, 2);
    const h = sh.getRange(1, 1, 1, AM_HEADERS.length);
    h.setValues([AM_HEADERS]);
    h.setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(AM_COL_NAME, 220);
    sh.setColumnWidth(AM_COL_PERSON, 140);
    sh.setColumnWidth(AM_COL_EMAIL, 240);
    sh.setColumnWidth(AM_COL_TOKEN, 250);
    sh.setColumnWidth(AM_COL_REGAT, 170);
    sh.setColumnWidth(AM_COL_LINKS, 380);
  }
  return sh;
}

function readAgencies_() {
  const sh = getAgencyMasterSheet_();
  const out = [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  sh.getRange(2, 1, lastRow - 1, AM_HEADERS.length).getValues().forEach(function (r, i) {
    const code = String(r[AM_COL_CODE - 1] || "").trim();
    if (!code) return;
    out.push({
      row:    i + 2,
      code:   code,
      name:   String(r[AM_COL_NAME - 1] || "").trim(),
      person: String(r[AM_COL_PERSON - 1] || "").trim(),
      email:  String(r[AM_COL_EMAIL - 1] || "").trim(),
      token:  String(r[AM_COL_TOKEN - 1] || "").trim(),
      status: String(r[AM_COL_STATUS - 1] || "").trim(),
      links:  String(r[AM_COL_LINKS - 1] || "").trim()
    });
  });
  return out;
}

function findAgencyByCode_(code) {
  const key = String(code || "").trim();
  if (!key) return null;
  const list = readAgencies_();
  for (let i = 0; i < list.length; i++) if (list[i].code === key) return list[i];
  return null;
}

function findAgencyByToken_(token) {
  const key = String(token || "").trim();
  if (!key) return null;
  const list = readAgencies_();
  for (let i = 0; i < list.length; i++) if (list[i].token && list[i].token === key) return list[i];
  return null;
}

function nextAgencyCode_() {
  const list = readAgencies_();
  let max = 0;
  list.forEach(function (a) {
    const m = /^ag(\d+)$/.exec(a.code);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  const n = max + 1;
  return AGENCY_CODE_PREFIX + (n < 10 ? "0" + n : String(n));
}

function agencyLinksUrl_(token) {
  return AGENCY_LINKS_PAGE + "?t=" + encodeURIComponent(token);
}

// 代理店1件の「稼働中案件リンク一式」を組み立てる。
// 稼働中の案件だけを対象にするので、停止した案件のリンクは自動的に消える。
function buildAgencyLinkList_(agencyCode) {
  // 代理店別取扱案件のチェックで絞る。未設定（列が無い）なら全稼働案件を渡す。
  let allowed = null;
  try { allowed = agencyCaseSelection_(agencyCode); } catch (e) { allowed = null; }

  return listActiveCases_()
    .filter(function (c) { return allowed === null || allowed.indexOf(c.code) >= 0; })
    .map(function (c) {
      return {
        caseCode: c.code,
        caseName: c.name,
        url: FORM_BASE_URL + "?form=" + encodeURIComponent(c.code) + "&ag=" + encodeURIComponent(agencyCode)
      };
    });
}

function isValidEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || "").trim());
}

// 代理店登録フォームの合言葉。ScriptProperties に置く。
// このリポジトリは GitHub Pages の配信元で public なので、値をソースへ書かない。
// 未設定のあいだは登録フォームを一切通さない（fail-closed）。
// 保護が要る理由: 無防備だと誰でも任意のアドレス宛にオーナーのGmailから
// メールを送らせられる＝送信の踏み台になる。
const AGENCY_SIGNUP_KEY_PROPERTY = "AGENCY_SIGNUP_KEY";

function agencySignupKey_() {
  return PropertiesService.getScriptProperties().getProperty(AGENCY_SIGNUP_KEY_PROPERTY) || "";
}

// 代理店を登録し、リンク集URLをメールで送る。
// 冪等性: 同じメールアドレスが既にあれば新規作成せず、その代理店のリンク集を送り直す。
// fromForm=true（公開フォーム経由）のときだけ合言葉を検証する。
// スプレッドシートのメニューからの登録は、操作者が編集権限を持つ時点で検証済みとみなす。
function registerAgency_(data, fromForm) {
  if (fromForm) {
    const key = agencySignupKey_();
    if (!key || String((data && data.signupKey) || "") !== key) {
      throw new Error("この登録フォームは現在ご利用いただけません。");
    }
  }

  const name   = String((data && data.agencyName)   || "").trim();
  const person = String((data && data.personName)   || "").trim();
  const email  = String((data && data.email)        || "").trim();

  if (!name)   throw new Error("会社名（代理店名）を入力してください。");
  if (!person) throw new Error("担当者名を入力してください。");
  if (!isValidEmail_(email)) throw new Error("メールアドレスの形式が正しくありません。");

  const sh = getAgencyMasterSheet_();
  const existing = readAgencies_().filter(function (a) {
    return a.email.toLowerCase() === email.toLowerCase();
  })[0];

  let agency;
  if (existing) {
    // 既存の再送。トークンが無ければ発行する。
    agency = existing;
    if (!agency.token) {
      agency.token = Utilities.getUuid().replace(/-/g, "");
      sh.getRange(agency.row, AM_COL_TOKEN).setValue(agency.token);
    }
    agency.links = agencyLinksUrl_(agency.token);
    sh.getRange(agency.row, AM_COL_LINKS).setValue(agency.links);
    if (agency.status !== AGENCY_STATUS_ACTIVE) {
      sh.getRange(agency.row, AM_COL_STATUS).setValue(AGENCY_STATUS_ACTIVE);
      agency.status = AGENCY_STATUS_ACTIVE;
    }
  } else {
    const code  = nextAgencyCode_();
    const token = Utilities.getUuid().replace(/-/g, "");
    const links = agencyLinksUrl_(token);
    const row = [code, name, person, email, token, AGENCY_STATUS_ACTIVE, formatJST(new Date()), links];
    sh.getRange(sh.getLastRow() + 1, 1, 1, AM_HEADERS.length).setValues([row]);
    agency = { code: code, name: name, person: person, email: email, token: token,
               status: AGENCY_STATUS_ACTIVE, links: links };
  }

  // 代理店の顧客も一元管理するため、顧客管理SSに担当者タブを用意する。
  try { ensureAgencyCustomerTab_(agency.person); }
  catch (e) { Logger.log("代理店の顧客管理タブ作成に失敗: " + e); }

  // 代理店別取扱案件に列を用意する（新規は全稼働案件にチェックが入る）。
  // ここで作っておかないと、リンク集が「未設定＝全案件」で動き続けて
  // 案件を絞りたくなったときに設定場所が無い、という状態になる。
  try { syncAgencyCaseMatrix(); }
  catch (e) { Logger.log("代理店別取扱案件の同期に失敗: " + e); }

  const cases = buildAgencyLinkList_(agency.code);
  let mailSent = false, mailError = "";
  try {
    sendAgencyLinksMail_(agency, cases);
    mailSent = true;
  } catch (e) {
    mailError = String(e);
    Logger.log("代理店リンク集メール送信エラー: " + e);
  }

  return {
    result: "success",
    agencyCode: agency.code,
    linksUrl: agency.links,
    caseCount: cases.length,
    mailSent: mailSent,
    mailError: mailError
  };
}

function sendAgencyLinksMail_(agency, cases) {
  const rows = cases.map(function (c) {
    return '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;">' + escapeHtml_(c.caseName) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;">' +
      '<a href="' + c.url + '" style="color:#4f46e5;">申請フォームを開く</a></td>' +
      '</tr>';
  }).join("");

  const html =
    '<div style="font-family:sans-serif;line-height:1.7;color:#111827;">' +
    '<p>' + escapeHtml_(agency.name) + '<br>' + escapeHtml_(agency.person) + ' 様</p>' +
    '<p>お世話になっております。<br>代理店登録が完了しましたので、専用のリンク集をお送りします。</p>' +
    '<p style="margin:20px 0;">' +
    '<a href="' + agency.links + '" style="background:#4f46e5;color:#fff;padding:12px 20px;' +
    'border-radius:6px;text-decoration:none;display:inline-block;">専用リンク集を開く</a></p>' +
    '<p style="font-size:13px;color:#6b7280;">' +
    'このページは開くたびに最新の状態を読み込みます。<br>' +
    '取扱いを終了した案件は自動的に消え、新しく始まった案件は自動的に増えます。<br>' +
    '<strong>ブックマークしてお使いください。</strong></p>' +
    '<p style="font-size:13px;color:#6b7280;word-break:break-all;">リンクが開けない場合はこちら:<br>' +
    agency.links + '</p>' +
    (cases.length
      ? '<p style="margin-top:24px;">現在ご案内できる案件（' + cases.length + '件）</p>' +
        '<table style="border-collapse:collapse;font-size:14px;">' + rows + '</table>'
      : '<p style="margin-top:24px;">現在ご案内できる案件はありません。稼働を開始しましたらリンク集に反映されます。</p>') +
    '</div>';

  MailApp.sendEmail({
    to: agency.email,
    subject: "【市場作り】代理店専用リンク集のご案内",
    htmlBody: html
  });
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

// 代理店の担当者タブを顧客管理SSへ用意する（無ければ作る）。
function ensureAgencyCustomerTab_(personName) {
  const person = String(personName || "").trim();
  if (!person) return;
  const css = getCustomerManagementSS();
  if (!css) return;
  const exists = css.getSheets().some(function (s) {
    return normalizeName(s.getName()) === normalizeName(person);
  });
  if (exists) return;

  // 既存タブを見本にして同じ列構成で作る（案件列を含む）
  let template = null;
  css.getSheets().forEach(function (s) {
    if (!template || s.getLastColumn() > template.getLastColumn()) template = s;
  });
  const sheet = css.insertSheet(person);
  if (template && template.getLastColumn() >= 1) {
    const headers = template.getRange(1, 1, 1, template.getLastColumn()).getValues();
    const r = sheet.getRange(1, 1, 1, headers[0].length);
    r.setValues(headers);
    r.setFontWeight("bold").setBackground("#312e81").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
}

// メニューから代理店を手動登録する
function showAgencyRegisterPrompt() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt("代理店登録 (1/3)", "会社名（代理店名）を入力してください。", ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const b = ui.prompt("代理店登録 (2/3)", "担当者名を入力してください。", ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const c = ui.prompt("代理店登録 (3/3)", "メールアドレスを入力してください。", ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  try {
    const res = registerAgency_({
      agencyName: a.getResponseText(),
      personName: b.getResponseText(),
      email:      c.getResponseText()
    });
    ui.alert(
      "代理店を登録しました。\n\n" +
      "代理店コード: " + res.agencyCode + "\n" +
      "稼働中の案件: " + res.caseCount + " 件\n" +
      "メール送信: " + (res.mailSent ? "成功" : "失敗（" + res.mailError + "）") + "\n\n" +
      "リンク集URL:\n" + res.linksUrl
    );
  } catch (e) {
    ui.alert("登録に失敗しました。\n\n" + e);
  }
}

// 全代理店へリンク集を送り直す（稼働案件が変わったことを知らせたいとき）
function resendAllAgencyLinks() {
  const ui = SpreadsheetApp.getUi();
  const list = readAgencies_().filter(function (a) { return a.status === AGENCY_STATUS_ACTIVE && a.email; });
  if (!list.length) { ui.alert("稼働中の代理店がありません。"); return; }
  const ok = ui.alert("稼働中の代理店 " + list.length + " 件へリンク集を送り直します。よろしいですか。",
                      ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  let sent = 0, failed = 0;
  list.forEach(function (a) {
    try { sendAgencyLinksMail_(a, buildAgencyLinkList_(a.code)); sent++; }
    catch (e) { failed++; Logger.log("再送失敗 " + a.code + ": " + e); }
  });
  ui.alert("送信しました。\n\n成功: " + sent + " 件\n失敗: " + failed + " 件");
}

// =============================================
// 回答シートの「代理店」列
// =============================================

// 各案件タブの回答ヘッダー末尾に「代理店」列を足す（既にあれば何もしない）。
// 末尾に足すので既存の列位置は動かない＝既存データと既存処理に影響しない。
function ensureAgencyColumnOnAllSheets() {
  const ss = getOrCreateSpreadsheet();
  const added = [];
  ss.getSheets().forEach(function (sheet) {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const lastCol = sheet.getLastColumn();
    if (lastCol < ANSWER_START_COL) return;
    const count = lastCol - ANSWER_START_COL + 1;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, count).getValues()[0].map(String);
    if (headers.indexOf(AGENCY_COLUMN_LABEL) >= 0) return;
    if (headers.indexOf("承認") < 0) return; // ヘッダー未初期化のシートは触らない
    const col = lastCol + 1;
    const r = sheet.getRange(1, col);
    r.setValue(AGENCY_COLUMN_LABEL);
    r.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    sheet.setColumnWidth(col, 140);
    added.push(sheet.getName());
  });
  return added;
}

function ensureAgencyColumnFromMenu() {
  const added = ensureAgencyColumnEverywhere_();
  SpreadsheetApp.getUi().alert("「代理店」列を追加しました。\n\n追加: " + added.length + " シート");
}

function ensureAgencyColumnEverywhere_() {
  return ensureAgencyColumnOnAllSheets();
}

// =============================================
// 公開エンドポイント用のペイロード
// =============================================

// リンク集ページが開くたびに呼ぶ。稼働状況はここで毎回読むのでリアルタイムに反映される。
function agencyLinksPayload_(token) {
  const agency = findAgencyByToken_(token);
  if (!agency) return { error: "リンクが無効です。" };
  if (agency.status !== AGENCY_STATUS_ACTIVE) {
    return { error: "このリンクは現在ご利用いただけません。" };
  }
  // 実績（そのリンクから何人申請したか）を添える。リンク集を一枚ものの管理画面にするため。
  // 件数は申請状況一覧（日次生成）から引くので、最大1日ぶん遅れる。
  let counts = {};
  try { counts = countAgencyApplicationsByName_(agency.name); } catch (e) { counts = {}; }

  const cases = buildAgencyLinkList_(agency.code).map(function (c) {
    return {
      caseCode: c.caseCode,
      caseName: c.caseName,
      url: c.url,
      count: counts[c.caseName] || 0
    };
  });
  let total = 0;
  Object.keys(counts).forEach(function (k) { total += counts[k]; });

  return {
    agencyName: agency.name,
    personName: agency.person,
    cases: cases,
    totalCount: total,
    updatedAt: formatJST(new Date())
  };
}

// 新しい権限（メール送信）を承認するために一度だけ実行する関数。
// 何も送信せず、残quotaを読むだけ。実行すると承認画面が出るのでそこで許可する。
function authorizeNewScopes() {
  const q = MailApp.getRemainingDailyQuota();
  Logger.log("メール送信の残quota: " + q);
  return q;
}

// =============================================
// 初回セットアップ（1回だけ実行する）
// =============================================

// 2026-08-19 時点で稼働中と確認された案件（ユーザー判断）。
// 初回セットアップでのみ使う種データで、以後の稼働管理は案件マスタの
// チェックボックスが正になる。ここを後から書き換えても反映されない。
const INITIAL_ACTIVE_CASE_CODES = [
  "nomukomu",     // ノムコム
  "sezonbpamex",  // セゾンビジネスプラチナアメックスカード
  "sezonmoney",   // セゾンマネーカード
  "sbishoken",    // SBI証券
  "hokenmammoth", // 保険マンモス
  "livingmuch",   // リビンマッチ
  "reguide",      // RE-Guide(リガイド)
  "iekatsu"       // いえカツLIFE
];

// 旧・代理店（小中様）。稼働していないので停止として記録し、タブを隠す。
const LEGACY_AGENCY_CODE = "01kn";
const LEGACY_AGENCY_NAME = "小中様";

// これ1本を実行すれば、新機能の初期状態が全部そろう。
// 実行時にメール送信の権限承認が求められるので、そこで許可すること。
function setupCaseAgencyFeature() {
  const log = [];

  // 1. メール送信の権限をこの実行に含める（承認画面を出すため）
  let quota = -1;
  try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { log.push("メール権限: " + e); }
  log.push("メール送信の残quota: " + quota);

  // 2. 案件マスタを作って設定タブから同期
  const synced = syncCaseMaster();
  log.push("案件マスタ: " + synced.count + " 件");

  // 3. 稼働中の案件に印を付ける（既にtrueのものはそのまま）
  const sh = getCaseMasterSheet_();
  const lastRow = sh.getLastRow();
  let marked = 0;
  if (lastRow >= 2) {
    const codes = sh.getRange(2, CM_COL_CODE, lastRow - 1, 1).getValues();
    for (let i = 0; i < codes.length; i++) {
      const code = String(codes[i][0] || "").trim();
      if (INITIAL_ACTIVE_CASE_CODES.indexOf(code) >= 0) {
        sh.getRange(i + 2, CM_COL_ACTIVE).setValue(true);
        sh.getRange(i + 2, CM_COL_UPDATED).setValue(formatJST(new Date()));
        marked++;
      }
    }
  }
  log.push("稼働に設定: " + marked + " 件");

  // 4. 稼働状況をシートの表示/非表示へ反映
  const vis = applyCaseVisibility();
  log.push("表示: " + vis.shown.length + " / 非表示: " + vis.hidden.length +
           " / 旧代理店タブ: " + vis.hiddenAgencyTabs.length);

  // 5. 回答シートに「代理店」列を追加
  const added = ensureAgencyColumnOnAllSheets();
  log.push("「代理店」列を追加: " + added.length + " シート");

  // 6. 代理店登録フォームの合言葉を発行（未設定のときだけ）。
  //    値はソースへ書かず ScriptProperties に置く。このリポジトリは public のため。
  try {
    const props = PropertiesService.getScriptProperties();
    let key = props.getProperty(AGENCY_SIGNUP_KEY_PROPERTY) || "";
    if (!key) {
      key = Utilities.getUuid().replace(/-/g, "");
      props.setProperty(AGENCY_SIGNUP_KEY_PROPERTY, key);
      log.push("代理店登録フォームの合言葉を新規発行しました");
    } else {
      log.push("代理店登録フォームの合言葉は設定済み");
    }
    log.push("代理店へ渡す登録フォームURL:");
    log.push("https://kazu02.github.io/affiliate-form/agency.html?k=" + key);
  } catch (e) {
    log.push("合言葉の設定に失敗: " + e);
  }

  // 7. 小中様を停止中の代理店として記録（メールは送らない）
  try {
    const am = getAgencyMasterSheet_();
    if (!findAgencyByCode_(LEGACY_AGENCY_CODE)) {
      am.getRange(am.getLastRow() + 1, 1, 1, AM_HEADERS.length).setValues([[
        LEGACY_AGENCY_CODE, LEGACY_AGENCY_NAME, "", "", "",
        AGENCY_STATUS_STOP, formatJST(new Date()), ""
      ]]);
      log.push("代理店マスタ: 小中様を停止として登録");
    } else {
      log.push("代理店マスタ: 小中様は登録済み");
    }
  } catch (e) {
    log.push("代理店マスタの初期化に失敗: " + e);
  }

  const msg = log.join("\n");
  Logger.log(msg);
  return msg;
}

// メニューから実行したとき用（結果をダイアログで見せる）
function setupCaseAgencyFeatureFromMenu() {
  const msg = setupCaseAgencyFeature();
  SpreadsheetApp.getUi().alert("初回セットアップが完了しました。\n\n" + msg);
}

// =============================================
// 代理店の削除
// =============================================

// 代理店を1件消す。代理店マスタの行だけでなく、登録時に作られた
// 顧客管理SSの担当者タブも片付ける（データが入っていれば残す）。
// dryRun=true なら何も消さずに、消す対象だけを返す。
function deleteAgencyCore_(code, dryRun) {
  const key = String(code || "").trim();
  if (!key) throw new Error("代理店コードを指定してください。");

  const agency = findAgencyByCode_(key);
  if (!agency) throw new Error("代理店コード「" + key + "」は見つかりません。");

  // 顧客管理SSの担当者タブの状態を調べる
  let tabState = "なし";
  let tabSheet = null;
  const css = getCustomerManagementSS();
  if (css && agency.person) {
    css.getSheets().forEach(function (sh) {
      if (normalizeName(sh.getName()) === normalizeName(agency.person)) tabSheet = sh;
    });
    if (tabSheet) {
      // 1行目はヘッダー。2行目以降に値があればデータ有りとみなす。
      tabState = (tabSheet.getLastRow() <= 1) ? "空なので削除" : "データがあるので残す";
    }
  }

  const plan = {
    code: agency.code, name: agency.name, person: agency.person,
    email: agency.email, row: agency.row, customerTab: tabState
  };
  if (dryRun) return plan;

  // 顧客管理タブ（空のときだけ消す）
  if (tabSheet && tabSheet.getLastRow() <= 1) {
    css.deleteSheet(tabSheet);
  }

  // 代理店マスタの行
  const sh = getAgencyMasterSheet_();
  sh.deleteRow(agency.row);
  SpreadsheetApp.flush();

  // 消えたことを確認する（フィルタが掛かっていると deleteRow は無言で失敗する）
  if (findAgencyByCode_(key)) {
    throw new Error("行の削除に失敗しました。代理店マスタにフィルタが掛かっていないか確認してください。");
  }
  return plan;
}

function showAgencyDeletePrompt() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt("代理店の削除", "削除する代理店コードを入力してください（例: ag01）。",
                      ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const code = a.getResponseText();

  let plan;
  try {
    plan = deleteAgencyCore_(code, true);
  } catch (e) {
    ui.alert("削除できません。\n\n" + e);
    return;
  }

  const ok = ui.alert(
    "次の代理店を削除します。よろしいですか。\n\n" +
    "代理店コード: " + plan.code + "\n" +
    "代理店名: " + plan.name + "\n" +
    "担当者名: " + plan.person + "\n" +
    "メール: " + plan.email + "\n" +
    "顧客管理の担当者タブ: " + plan.customerTab,
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  try {
    const done = deleteAgencyCore_(code, false);
    ui.alert("削除しました。\n\n" + done.code + " / " + done.name +
             "\n顧客管理の担当者タブ: " + done.customerTab);
  } catch (e) {
    ui.alert("削除に失敗しました。\n\n" + e);
  }
}

// テストで作った代理店 ag01 を消すための直接実行用。
// エディタから1回実行すれば済むよう、確認ダイアログを挟まない。
function deleteTestAgencyAg01() {
  const msg = JSON.stringify(deleteAgencyCore_("ag01", false));
  Logger.log(msg);
  return msg;
}
