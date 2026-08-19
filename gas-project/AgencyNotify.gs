// =============================================
// 代理店への稼働変更通知（メール）
// （2026-08-19 追加）
//
// 稼働状況を変えても、誰かが手でメニューを押さない限り代理店には伝わらない。
// それだと「止めた案件のリンクを代理店が配り続ける」「始めた案件が使われない」が起きるので、
// 日次で稼働案件の集合を前回と突き合わせ、変わったときだけメールする。
//
// LINE を使わずメールにしているのは、(1) 送信の仕組みが既に動いていて追加費用がない
// (2) LINEのプッシュは従量で無料枠が小さい (3) GASの doPost はHTTPヘッダーを読めず
// X-Line-Signature を検証できないため、LINE を足すには統合アプリ側の変更が要る、の3点。
// LINE化する場合の設計は DECISIONS.md 2026-08-19 を参照。
// =============================================

const NOTIFIED_CASES_PROPERTY = "AGENCY_NOTIFIED_CASE_CODES";
const AGENCY_NOTIFY_TRIGGER_FN = "notifyAgenciesIfCasesChanged";

function readNotifiedCaseCodes_() {
  const raw = PropertiesService.getScriptProperties().getProperty(NOTIFIED_CASES_PROPERTY);
  if (!raw) return null; // 未記録（初回）
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : null;
  } catch (e) {
    return null;
  }
}

function writeNotifiedCaseCodes_(codes) {
  PropertiesService.getScriptProperties()
    .setProperty(NOTIFIED_CASES_PROPERTY, JSON.stringify(codes.slice().sort()));
}

// 稼働案件が前回から変わっていたら、稼働中の全代理店へメールする。
// 初回（未記録）のときは送らずに現状を記録するだけ。
// これをしないと「全案件が新規」という誤った内容を一斉送信してしまう。
function notifyAgenciesIfCasesChanged() {
  const current = listActiveCases_();
  const currentCodes = current.map(function (c) { return c.code; }).sort();
  const prevCodes = readNotifiedCaseCodes_();

  if (prevCodes === null) {
    writeNotifiedCaseCodes_(currentCodes);
    Logger.log("初回のため現状を記録しonly（送信なし）: " + currentCodes.join(","));
    return { firstRun: true, sent: 0, added: [], removed: [] };
  }

  const prevSet = {};
  prevCodes.forEach(function (c) { prevSet[c] = true; });
  const currSet = {};
  currentCodes.forEach(function (c) { currSet[c] = true; });

  const addedCodes   = currentCodes.filter(function (c) { return !prevSet[c]; });
  const removedCodes = prevCodes.filter(function (c) { return !currSet[c]; });

  if (!addedCodes.length && !removedCodes.length) {
    Logger.log("稼働案件に変化なし");
    return { changed: false, sent: 0, added: [], removed: [] };
  }

  // 案件コード → 名前。終了した案件は稼働一覧に居ないので設定タブから引く。
  const nameOf = {};
  listCaseSheets_(getOrCreateSpreadsheet()).forEach(function (c) { nameOf[c.code] = c.name; });
  const added   = addedCodes.map(function (c) { return nameOf[c] || c; });
  const removed = removedCodes.map(function (c) { return nameOf[c] || c; });

  const agencies = readAgencies_().filter(function (a) {
    return a.status === AGENCY_STATUS_ACTIVE && a.email;
  });

  let sent = 0, failed = 0;
  agencies.forEach(function (a) {
    try {
      sendAgencyChangeMail_(a, buildAgencyLinkList_(a.code), added, removed);
      sent++;
    } catch (e) {
      failed++;
      Logger.log("代理店への変更通知に失敗 " + a.code + ": " + e);
    }
  });

  // 1件でも送れていれば記録を進める。全滅なら次回に再試行させたいので進めない。
  if (sent > 0 || agencies.length === 0) writeNotifiedCaseCodes_(currentCodes);

  Logger.log("稼働変更を通知: 追加=" + added.join("/") + " 終了=" + removed.join("/") +
             " 送信=" + sent + " 失敗=" + failed);
  return { changed: true, sent: sent, failed: failed, added: added, removed: removed };
}

function sendAgencyChangeMail_(agency, cases, added, removed) {
  const rows = cases.map(function (c) {
    return '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;">' + escapeHtml_(c.caseName) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;">' +
      '<a href="' + c.url + '" style="color:#4f46e5;">申請フォームを開く</a></td>' +
      '</tr>';
  }).join("");

  let changeHtml = '';
  if (added.length) {
    changeHtml += '<p style="margin:4px 0;"><strong>新しく取扱いできる案件</strong><br>' +
                  escapeHtml_(added.join("、")) + '</p>';
  }
  if (removed.length) {
    changeHtml += '<p style="margin:4px 0;"><strong>取扱いを終了した案件</strong><br>' +
                  escapeHtml_(removed.join("、")) +
                  '<br><span style="font-size:12px;color:#6b7280;">' +
                  'これらのリンクは配布をお止めください。開いても申請できません。</span></p>';
  }

  const html =
    '<div style="font-family:sans-serif;line-height:1.7;color:#111827;">' +
    '<p>' + escapeHtml_(agency.name) + '<br>' + escapeHtml_(agency.person) + ' 様</p>' +
    '<p>お世話になっております。取扱い案件に変更がありましたのでお知らせします。</p>' +
    '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:16px 0;">' +
    changeHtml + '</div>' +
    '<p style="margin:20px 0;">' +
    '<a href="' + agency.links + '" style="background:#4f46e5;color:#fff;padding:12px 20px;' +
    'border-radius:6px;text-decoration:none;display:inline-block;">最新のリンク集を開く</a></p>' +
    '<p style="font-size:13px;color:#6b7280;">' +
    'リンク集は常に最新の状態を表示します。ブックマークしてお使いください。</p>' +
    (cases.length
      ? '<p style="margin-top:24px;">現在ご案内できる案件（' + cases.length + '件）</p>' +
        '<table style="border-collapse:collapse;font-size:14px;">' + rows + '</table>'
      : '<p style="margin-top:24px;">現在ご案内できる案件はありません。</p>') +
    '</div>';

  MailApp.sendEmail({
    to: agency.email,
    subject: "【市場作り】取扱い案件の変更のお知らせ",
    htmlBody: html
  });
}

// 日次トリガーを用意する（重複作成しない）
function ensureAgencyNotifyTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === AGENCY_NOTIFY_TRIGGER_FN;
  });
  if (exists) return "既に登録済み";
  ScriptApp.newTrigger(AGENCY_NOTIFY_TRIGGER_FN)
    .timeBased().everyDays(1).atHour(9).create();
  return "日次トリガーを登録しました（毎日9時台）";
}

// メニュー用: いま変更があれば通知する
function notifyAgencyCaseChangesFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = notifyAgenciesIfCasesChanged();
    if (r.firstRun) {
      ui.alert("現在の稼働案件を基準として記録しました。\n\n" +
               "次回以降、変更があったときに代理店へ自動でメールします。");
      return;
    }
    if (!r.changed) {
      ui.alert("前回の通知から稼働案件に変化はありません。\n\n送信していません。");
      return;
    }
    ui.alert("代理店へ変更を通知しました。\n\n" +
             "新しく取扱い: " + (r.added.join("、") || "なし") + "\n" +
             "取扱い終了: " + (r.removed.join("、") || "なし") + "\n" +
             "送信: " + r.sent + " 件 / 失敗: " + (r.failed || 0) + " 件");
  } catch (e) {
    ui.alert("通知に失敗しました。\n\n" + e);
  }
}

function ensureAgencyNotifyTriggerFromMenu() {
  SpreadsheetApp.getUi().alert(ensureAgencyNotifyTrigger());
}
