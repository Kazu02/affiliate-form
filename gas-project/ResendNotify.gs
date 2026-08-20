// =============================================
// 新規申請通知の再送（2026-08-20 追加）
//
// スクショURLがLINE通知から抜けていた期間（2026-08-19 に「代理店」列を末尾へ足してから
// 2026-08-20 の修正まで）の申請を、スクショ付きで送り直すために作った。
// 通知が飛ばなかった・内容が欠けた分を後から補うのは今後も起きるので恒久の口として残す。
//
// 見出しは【再通知】にする。元の通知と同じ【新規申請】だと新しい申請と誤認され、
// 対応する人が二重に動くため。
// =============================================

const RESEND_DEFAULT_COUNT = 5;
const RESEND_MAX_COUNT     = 20;

// 直近n件を新しい順に集める。行はシートの実ヘッダーで読むので、列構成が変わる前の
// 古い行が混ざっていても取り違えない。
function collectRecentApplicationsForResend_(n) {
  const ss = getOrCreateSpreadsheet();
  const out = [];

  listCaseSheets_(ss).forEach(function (c) {
    const sheet   = c.sheet;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const width   = lastCol - ANSWER_START_COL + 1;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const iRecv   = headers.indexOf("受信日時");
    if (iRecv < 0) return;

    const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, width).getValues();
    data.forEach(function (row) {
      const recv = row[iRecv];
      if (recv === "" || recv === null || recv === undefined) return;
      out.push({ sortKey: toDateForSort_(recv), caseName: c.name, headers: headers, row: row });
    });
  });

  out.sort(function (a, b) { return b.sortKey - a.sortKey; });
  return out.slice(0, n);
}

// 元の新規申請通知と同じ並びで組み立てる。項目はヘッダー名で引くので、
// 案件ごとに項目が違っても、あとから列が増えても追随する。
function buildResendMessage_(caseName, headers, row) {
  const BASE_FIXED = 4; // フォーム記号 / 受信日時 / クリック日時 / 送信日時
  const iRecv = headers.indexOf("受信日時");
  const iShot = headers.lastIndexOf("スクショURL");
  const iAgcy = headers.indexOf(AGENCY_COLUMN_LABEL);

  const lines = ["【再通知】" + caseName];
  if (iRecv >= 0) lines.push("受信日時: " + toDisplayDate_(row[iRecv]));

  // 固定4列とスクショURLの間が、その案件の入力項目。
  const end = iShot >= 0 ? iShot : headers.length;
  for (let i = BASE_FIXED; i < end; i++) {
    const v = row[i];
    if (v !== "" && v !== null && v !== undefined) lines.push(headers[i] + ": " + v);
  }
  if (iAgcy >= 0 && String(row[iAgcy] || "").trim()) {
    lines.push(AGENCY_COLUMN_LABEL + ": " + String(row[iAgcy]).trim());
  }

  // スクショはフォームで必須なので「行が無い」＝異常。状態にかかわらず必ず1行出す。
  const shot = iShot >= 0 ? String(row[iShot] || "").trim() : "";
  if (shot.startsWith("http"))  lines.push("スクショ: " + shot);
  else if (shot)                lines.push("スクショ: 保存に失敗しました（" + shot.substring(0, 120) + "）");
  else                          lines.push("スクショ: 取得できませんでした（要確認）");

  return lines.join("\n");
}

// 直近n件ぶんの本文を、古い順（申請された順）に並べて返す。組み立てと送信を分けてあるのは、
// 送る前に中身を確認できるようにするため。
function buildRecentResendMessages_(n) {
  const count = Math.min(Math.max(parseInt(n, 10) || RESEND_DEFAULT_COUNT, 1), RESEND_MAX_COUNT);
  return collectRecentApplicationsForResend_(count)
    .reverse()
    .map(function (it) { return buildResendMessage_(it.caseName, it.headers, it.row); });
}

// ---- メニューから再送（送る前に中身を出して確認する）----
function resendRecentApplicationsFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const messages = buildRecentResendMessages_(RESEND_DEFAULT_COUNT);
  if (!messages.length) {
    ui.alert("再送できる申請がありませんでした。");
    return;
  }
  const answer = ui.alert(
    "直近" + messages.length + "件をLINEへ再通知します",
    messages.join("\n\n--------------------\n\n") + "\n\n送信しますか？",
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  messages.forEach(function (m) { notifyLineGroup(m); });
  ui.alert(messages.length + "件を再通知しました。");
}

