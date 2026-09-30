// =============================================
// 月間送信数の上限で届かなかった新規申請通知の再送（2026-09-30 追加・9月分の1回限り）
//
// 公式LINEの月間送信数が 2026-09-29 夕方に上限へ達し、それ以降の通知は
// notifyLineGroup() が muteHttpExceptions で応答を見ないため、黙って捨てられていた。
// 通数は月初（日本時間 10/1 0:00）に戻るので、その直後に本文をシートから組み立て直して送る。
//
// グループへの push は「グループ人数 × 1通」で数えられる。1件ずつ送ると件数倍になるので、
// 1回の push に吹き出しを5つまで詰めて送る（10月の通数を使い過ぎないため）。
// 本文は ResendNotify.gs の buildResendMessage_()（見出し【再通知】・受信日時つき）を使う。
// =============================================

// アポ共有グループ（同じ公式LINE）でボットの通知が最後に届いた時刻。これより後の申請が対象。
const QR_SINCE        = new Date("2026-09-29T18:31:36+09:00");
// 通数が戻る時刻。これ以降の申請は通常どおり届くので対象にしない。
const QR_UNTIL        = new Date("2026-10-01T00:00:00+09:00");
const QR_FIRST_RUN_AT = new Date("2026-10-01T00:01:00+09:00");
const QR_HANDLER      = "runQuotaResend202609";
const QR_PROGRESS_KEY = "QUOTA_RESEND_202609_SENT_PUSHES";
const QR_ATTEMPT_KEY  = "QUOTA_RESEND_202609_ATTEMPTS";
const QR_MAX_ATTEMPTS = 18;   // 10分おきに3時間まで。まだ上限なら人が見る
const QR_BUBBLES_PER_PUSH = 5; // LINE の push 1回に入れられる吹き出しの上限

function collectQuotaResendItems_() {
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

    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, width).getValues().forEach(function (row) {
      const t = toDateForSort_(row[iRecv]);
      if (t > QR_SINCE.getTime() && t < QR_UNTIL.getTime()) {
        out.push({ sortKey: t, caseName: c.name, headers: headers, row: row });
      }
    });
  });
  out.sort(function (a, b) { return a.sortKey - b.sortKey; });
  return out;
}

// 送る吹き出しを push 単位に分けて返す。先頭に事情の説明を1つ置く。
function buildQuotaResendPushes_() {
  const bodies = collectQuotaResendItems_().map(function (it) {
    return buildResendMessage_(it.caseName, it.headers, it.row);
  });
  if (!bodies.length) return [];
  const intro = "【未送信分の再送】\n" +
    "公式LINEの月間送信数が上限に達していたため、9/29 夕方〜9/30 の新規申請通知 " + bodies.length +
    "件が届いていませんでした。受信順にまとめて送ります。\n" +
    "新しい申請ではありません（受信日時をご確認ください）。";
  const bubbles = [intro].concat(bodies);
  const pushes = [];
  for (let i = 0; i < bubbles.length; i += QR_BUBBLES_PER_PUSH) {
    pushes.push(bubbles.slice(i, i + QR_BUBBLES_PER_PUSH));
  }
  return pushes;
}

function qrPush_(bubbles) {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty("LINE_CHANNEL_TOKEN");
  const groupId = props.getProperty("LINE_GROUP_ID");
  if (!token || !groupId) return { code: -1, body: "LINE設定なし" };
  const res = UrlFetchApp.fetch(LINE_PUSH_API, {
    method: "post",
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
    payload: JSON.stringify({
      to: groupId,
      messages: bubbles.map(function (t) {
        return { type: "text", text: t.length > 4990 ? t.substring(0, 4990) + "..." : t };
      })
    }),
    muteHttpExceptions: true
  });
  return { code: res.getResponseCode(), body: res.getContentText().substring(0, 300) };
}

function qrDeleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === QR_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

// ---- 送らずに中身だけ確かめる（GASエディタから実行）----
function previewQuotaResend202609() {
  const pushes = buildQuotaResendPushes_();
  Logger.log("push回数: " + pushes.length + " / 送信済み: " +
    (PropertiesService.getScriptProperties().getProperty(QR_PROGRESS_KEY) || "0"));
  pushes.forEach(function (p, i) {
    p.forEach(function (b, j) { Logger.log((i + 1) + "-" + (j + 1) + ": " + b.split("\n").slice(0, 2).join(" / ")); });
  });
}

// ---- 10/1 0:01 に1回だけ動く予約を入れる（GASエディタから1回実行）----
function scheduleQuotaResend202609() {
  qrDeleteTriggers_();
  ScriptApp.newTrigger(QR_HANDLER).timeBased().at(QR_FIRST_RUN_AT).create();
  Logger.log("予約しました: " + QR_FIRST_RUN_AT);
  previewQuotaResend202609();
}

// ---- 本体。まだ上限なら10分後に自分を予約し直す ----
function runQuotaResend202609() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    if (new Date() < QR_UNTIL) { Logger.log("月が替わる前なので送らない"); return; }
    const props  = PropertiesService.getScriptProperties();
    const pushes = buildQuotaResendPushes_();
    let done = parseInt(props.getProperty(QR_PROGRESS_KEY) || "0", 10);
    for (let i = done; i < pushes.length; i++) {
      const r = qrPush_(pushes[i]);
      Logger.log("push " + (i + 1) + "/" + pushes.length + ": HTTP " + r.code + " " + r.body);
      if (r.code !== 200) {
        const attempts = parseInt(props.getProperty(QR_ATTEMPT_KEY) || "0", 10) + 1;
        props.setProperty(QR_ATTEMPT_KEY, String(attempts));
        qrDeleteTriggers_();
        if (attempts < QR_MAX_ATTEMPTS) {
          ScriptApp.newTrigger(QR_HANDLER).timeBased().after(10 * 60 * 1000).create();
        }
        return;
      }
      props.setProperty(QR_PROGRESS_KEY, String(i + 1));
    }
    qrDeleteTriggers_();
    Logger.log("再送完了: push " + pushes.length + "回");
  } finally {
    lock.releaseLock();
  }
}
