// =============================================
// 申請状況の一元表示（自社＋全代理店）と、代理店への新規申請通知
// （2026-08-19 追加）
//
// 案件ごとにタブが分かれていると、代理店別の承認状況を横断で見られない。
// 全案件タブの回答を1枚に集め、代理店で色分けする。生成物なので手で編集しない。
// =============================================

const APP_STATUS_SHEET = "申請状況一覧";
const APP_STATUS_HEADERS = ["受信日時", "案件名", "顧客名", "紹介者名", "代理店", "承認", "スクショURL"];

// 代理店ごとの行の色。自社は白のまま。足りなくなったら先頭から使い回す。
const AGENCY_ROW_COLORS = [
  "#e0f2fe", "#dcfce7", "#fef9c3", "#fae8ff", "#ffe4e6", "#e0e7ff", "#ccfbf1", "#ffedd5"
];

function getAppStatusSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(APP_STATUS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(APP_STATUS_SHEET, 4);
  }
  return sh;
}

// 受信日時は古い行が Date、新しい行が文字列のことがあるので両方を扱う。
function toDateForSort_(v) {
  if (v instanceof Date) return v.getTime();
  const s = String(v || "").trim();
  if (!s) return 0;
  const t = Date.parse(s.replace(/\//g, "-").replace(" ", "T") + "+09:00");
  if (!isNaN(t)) return t;
  const t2 = Date.parse(s);
  return isNaN(t2) ? 0 : t2;
}

function toDisplayDate_(v) {
  if (v instanceof Date) return formatJST(new Date(v.getTime() - 9 * 60 * 60 * 1000));
  return String(v || "");
}

// 全案件タブの回答を1枚に集める。
function collectApplications_() {
  const ss = getOrCreateSpreadsheet();
  const out = [];

  listCaseSheets_(ss).forEach(function (c) {
    const sheet = c.sheet;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const width = lastCol - ANSWER_START_COL + 1;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const idx = function (label) { return headers.indexOf(label); };

    const iRecv = idx("受信日時");
    const iName = idx("お名前");
    const iRef  = idx("紹介者名");
    const iAgcy = idx(AGENCY_COLUMN_LABEL);
    const iAppr = idx("承認");
    const iShot = idx("スクショURL");
    if (iRecv < 0) return;

    const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, width).getValues();
    data.forEach(function (row) {
      // 空行を飛ばす（受信日時も名前も無ければデータではない）
      const hasRecv = row[iRecv] !== "" && row[iRecv] !== null && row[iRecv] !== undefined;
      const hasName = iName >= 0 && String(row[iName] || "").trim() !== "";
      if (!hasRecv && !hasName) return;

      out.push({
        sortKey: toDateForSort_(row[iRecv]),
        values: [
          toDisplayDate_(row[iRecv]),
          c.name,
          iName >= 0 ? row[iName] : "",
          iRef  >= 0 ? row[iRef]  : "",
          iAgcy >= 0 ? String(row[iAgcy] || "") : "",
          iAppr >= 0 ? row[iAppr] : "",
          iShot >= 0 ? row[iShot] : ""
        ]
      });
    });
  });

  out.sort(function (a, b) { return b.sortKey - a.sortKey; }); // 新しい順
  return out;
}

// 申請状況一覧を作り直す。生成物なので既存内容は全て置き換える。
function buildApplicationStatusSheet() {
  const sh = getAppStatusSheet_();
  const rows = collectApplications_();

  sh.clear();

  const hr = sh.getRange(1, 1, 1, APP_STATUS_HEADERS.length);
  hr.setValues([APP_STATUS_HEADERS]);
  hr.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 150);
  sh.setColumnWidth(6, 70);
  sh.setColumnWidth(7, 260);

  if (!rows.length) return { count: 0, agencies: 0 };

  sh.getRange(2, 1, rows.length, APP_STATUS_HEADERS.length)
    .setValues(rows.map(function (r) { return r.values; }));

  // 代理店ごとに行を色分けする。自社（代理店が空）は白のまま。
  const colorOf = {};
  let next = 0;
  rows.forEach(function (r) {
    const ag = String(r.values[4] || "").trim();
    if (!ag) return;
    if (!colorOf[ag]) {
      colorOf[ag] = AGENCY_ROW_COLORS[next % AGENCY_ROW_COLORS.length];
      next++;
    }
  });
  const bg = rows.map(function (r) {
    const ag = String(r.values[4] || "").trim();
    const c = ag ? colorOf[ag] : "#ffffff";
    return APP_STATUS_HEADERS.map(function () { return c; });
  });
  sh.getRange(2, 1, rows.length, APP_STATUS_HEADERS.length).setBackgrounds(bg);

  return { count: rows.length, agencies: Object.keys(colorOf).length };
}

function buildApplicationStatusSheetFromMenu() {
  const r = buildApplicationStatusSheet();
  SpreadsheetApp.getUi().alert(
    "申請状況一覧を作り直しました。\n\n" +
    "申請: " + r.count + " 件\n" +
    "色分けした代理店: " + r.agencies + " 件\n\n" +
    "このシートは生成物です。手で編集しても次の再生成で消えます。"
  );
}

// =============================================
// 新規申請を代理店へ通知する
// =============================================

// 代理店経由の申請が入ったとき、その代理店へメールで知らせる。
// 自社経由（agencyName が空）のときは何もしない。
// doPost から呼ぶので、失敗しても申請の記録は止めない（呼び出し側で try/catch する）。
function notifyAgencyOfApplication_(agencyName, caseName, customerName) {
  const name = String(agencyName || "").trim();
  if (!name) return;

  const agency = readAgencies_().filter(function (a) {
    return a.name === name && a.status === AGENCY_STATUS_ACTIVE && a.email;
  })[0];
  if (!agency) return;

  // 顧客の氏名はメールに載せない（宛先を間違えたときの被害を小さくするため）。
  const html =
    '<div style="font-family:sans-serif;line-height:1.7;color:#111827;">' +
    '<p>' + escapeHtml_(agency.name) + '<br>' + escapeHtml_(agency.person) + ' 様</p>' +
    '<p>貴社のリンク経由で新しい申請が入りました。</p>' +
    '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:16px 0;">' +
    '<p style="margin:0;"><strong>案件</strong>: ' + escapeHtml_(caseName) + '</p>' +
    '<p style="margin:6px 0 0;"><strong>受付</strong>: ' + formatJST(new Date()) + '</p>' +
    '</div>' +
    '<p style="margin:20px 0;">' +
    '<a href="' + agency.links + '" style="background:#4f46e5;color:#fff;padding:12px 20px;' +
    'border-radius:6px;text-decoration:none;display:inline-block;">申請数を確認する</a></p>' +
    '<p style="font-size:13px;color:#6b7280;">' +
    '承認の可否は広告主の確認後に確定します。確定までお時間をいただきます。</p>' +
    '</div>';

  MailApp.sendEmail({
    to: agency.email,
    subject: "【市場作り】新しい申請が入りました（" + caseName + "）",
    htmlBody: html
  });
}

// 申請状況一覧を毎日作り直す。手で押さないと古いまま、を避けるため。
// 申請のたびに作り直すと数百行の再生成が送信のたびに走るので、日次にしている。
// 申請状況一覧と代理店別サマリーをまとめて作り直す rebuildAgencyReports を回す。
// 旧版は buildApplicationStatusSheet を直接登録していたので、見つけたら差し替える。
const APP_STATUS_TRIGGER_FN     = "rebuildAgencyReports";
const APP_STATUS_TRIGGER_FN_OLD = "buildApplicationStatusSheet";

function ensureAppStatusTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === APP_STATUS_TRIGGER_FN_OLD) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  const exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === APP_STATUS_TRIGGER_FN;
  });
  if (exists) return removed ? "旧トリガーを削除しました（新は登録済み）" : "既に登録済み";
  ScriptApp.newTrigger(APP_STATUS_TRIGGER_FN)
    .timeBased().everyDays(1).atHour(7).create();
  return "日次トリガーを登録しました（毎日7時台）" + (removed ? "。旧トリガーは削除" : "");
}
