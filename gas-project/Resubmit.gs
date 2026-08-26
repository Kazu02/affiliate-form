// =============================================
// スクショの再提出（2026-08-26 追加）
//
// 承認漏れの確認で、営業担当が「サンクスメールのスクショに申込日時が写っていない」と
// 判断した申請について、顧客へ渡す再提出用URLを発行する。
// 顧客がそのURLから撮り直した画像を送ると、元の申請行のスクショURLが差し替わる。
//
// 設計の判断（2026-08-26 ユーザー承認）:
//   - 仕様は「日時が写っているか」の仮説で固定する。広告主の反応を待たずに作る
//   - **古いスクショは消さない。** 差し替えではなく、旧URLを台帳へ残したうえで更新する
//   - トークンは代理店リンク集と同じ作り（UUID をシートに持ち `?rs=` で開く）
//
// 安全側に倒している点:
//   - **期限つき（既定14日）・1回使い切り。** 漏れても窓が短く、上書きは1回だけ
//   - **公開ページに顧客名を出さない。** 案件名と申込日だけ返す（URLが漏れたときの被害を抑える）
//   - 書き込む前に顧客名と受信日時を照合する。行は編集で動くので行番号を信じない
// =============================================

const RESUBMIT_SHEET      = "再提出トークン";
const RESUBMIT_PAGE       = "https://kazu02.github.io/affiliate-form/resubmit.html";
const RESUBMIT_VALID_DAYS = 14;
const AG_RESUBMIT_PREFIX  = "再提出_";   // SS2 の担当別タブ

const RS_HEADERS = [
  "トークン", "案件名", "顧客名", "紹介者（営業）", "受信日時",
  "対象シート", "対象行", "発行日時", "期限", "状態", "使用日時",
  "旧スクショURL", "新スクショURL", "再提出URL"
];
const RSC_TOKEN = 1, RSC_CASE = 2, RSC_NAME = 3, RSC_REF = 4, RSC_RECV = 5,
      RSC_SHEET = 6, RSC_ROW = 7, RSC_ISSUED = 8, RSC_EXPIRE = 9, RSC_STATE = 10,
      RSC_USED = 11, RSC_OLDURL = 12, RSC_NEWURL = 13, RSC_URL = 14;

const RS_STATE_OPEN = "未使用";
const RS_STATE_DONE = "使用済み";
const RS_STATE_VOID = "無効";

function rsSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(RESUBMIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RESUBMIT_SHEET);
    const hr = sh.getRange(1, 1, 1, RS_HEADERS.length);
    hr.setValues([RS_HEADERS]);
    hr.setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff").setWrap(true);
    sh.setFrozenRows(1);
    sh.setColumnWidth(RSC_TOKEN, 90);
    sh.setColumnWidth(RSC_CASE, 200);
    sh.setColumnWidth(RSC_URL, 320);
  }
  return sh;
}

function rsUrl_(token) {
  return RESUBMIT_PAGE + "?rs=" + encodeURIComponent(token);
}

// 期限切れを状態へ反映しつつ、生きているトークン行を返す（無ければ null）
function rsFindByToken_(token) {
  const key = String(token || "").trim();
  if (!key) return null;
  const sh = rsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, RS_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][RSC_TOKEN - 1] || "").trim() !== key) continue;
    return { row: i + 2, values: data[i], sheet: sh };
  }
  return null;
}

function rsIsExpired_(values) {
  const ms = agDateToMillis_(values[RSC_EXPIRE - 1]);
  if (!ms) return false;
  // 期限は「その日いっぱい」まで有効にする（JSTの0時が入るので1日足す）
  return Date.now() > ms + 86400000;
}

// ---------------------------------------------------------------
// 発行: 承認漏れ管理 の「要再取得」行にトークンを配る
// ---------------------------------------------------------------
function issueResubmitTokens() {
  const ss = getOrCreateSpreadsheet();
  const ag = ss.getSheetByName(AG_MANAGE_SHEET);
  if (!ag) throw new Error(AG_MANAGE_SHEET + " がありません。先に棚卸しを実行してください。");
  const last = ag.getLastRow();
  if (last < 2) return { issued: 0, reused: 0, rows: [] };
  const agData = ag.getRange(2, 1, last - 1, AG_HEADERS.length).getValues();

  const rs = rsSheet_();
  const rsLast = rs.getLastRow();
  const rsData = rsLast >= 2 ? rs.getRange(2, 1, rsLast - 1, RS_HEADERS.length).getValues() : [];

  // 既に生きているトークンがある申請は作り直さない（顧客へ送ったURLを無効にしないため）
  const liveByKey = {};
  rsData.forEach(function (r) {
    if (String(r[RSC_STATE - 1] || "") !== RS_STATE_OPEN) return;
    if (rsIsExpired_(r)) return;
    liveByKey[agKeyOf_(r[RSC_CASE - 1], r[RSC_NAME - 1], aspToMillis_(r[RSC_RECV - 1]))] =
      String(r[RSC_TOKEN - 1] || "");
  });

  const now = new Date();
  const expire = new Date(now.getTime() + RESUBMIT_VALID_DAYS * 86400000);
  const add = [], out = [];
  let reused = 0;

  agData.forEach(function (r) {
    if (String(r[AGC_SALESCHK - 1] || "").trim() !== "要再取得") return;
    const key = agKeyOf_(r[AGC_CASE - 1], r[AGC_NAME - 1], aspToMillis_(r[AGC_RECV - 1]));
    if (liveByKey[key]) {
      reused++;
      out.push({ ref: r[AGC_REF - 1], caseName: r[AGC_CASE - 1], name: r[AGC_NAME - 1],
                 recv: r[AGC_RECV - 1], url: rsUrl_(liveByKey[key]),
                 expire: "", state: RS_STATE_OPEN });
      return;
    }
    const token = Utilities.getUuid().replace(/-/g, "");
    add.push([
      token, r[AGC_CASE - 1], r[AGC_NAME - 1], r[AGC_REF - 1], r[AGC_RECV - 1],
      r[AGC_SHEET - 1], r[AGC_ROW - 1],
      formatJST(now), formatJST(expire).slice(0, 10), RS_STATE_OPEN, "",
      r[AGC_SHOT - 1], "", rsUrl_(token)
    ]);
    out.push({ ref: r[AGC_REF - 1], caseName: r[AGC_CASE - 1], name: r[AGC_NAME - 1],
               recv: r[AGC_RECV - 1], url: rsUrl_(token),
               expire: formatJST(expire).slice(0, 10), state: RS_STATE_OPEN });
  });

  if (add.length) {
    rs.getRange(Math.max(rsLast + 1, 2), 1, add.length, RS_HEADERS.length).setValues(add);
    SpreadsheetApp.flush();
  }
  return { issued: add.length, reused: reused, rows: out };
}

// 担当別に「再提出_<担当>」タブへ書き出す。営業担当はここからURLを顧客へ送る。
function pushResubmitTabs_(rows) {
  const byRep = {};
  rows.forEach(function (r) {
    const rep = String(r.ref || "").trim() || "（担当なし）";
    if (!byRep[rep]) byRep[rep] = [];
    byRep[rep].push([r.caseName, r.name, r.recv, r.url, r.expire, r.state]);
  });
  let outSS;
  try { outSS = SpreadsheetApp.openById(REP_STATUS_SS2_ID); }
  catch (e) { throw new Error("SS2（担当別ステータス表）を開けません: " + e); }

  const head = ["案件名", "顧客名", "申込日時", "再提出URL（お客様へ送る）", "期限", "状態"];
  let reps = 0;
  Object.keys(byRep).forEach(function (rep) {
    const name = AG_RESUBMIT_PREFIX + rep;
    let tab = outSS.getSheetByName(name);
    if (!tab) tab = outSS.insertSheet(name);
    tab.clear();
    const hr = tab.getRange(1, 1, 1, head.length);
    hr.setValues([head]);
    hr.setFontWeight("bold").setBackground("#b45309").setFontColor("#ffffff");
    tab.setFrozenRows(1);
    const list = byRep[rep];
    tab.getRange(2, 1, list.length, head.length).setValues(list);
    tab.setColumnWidth(1, 200); tab.setColumnWidth(2, 120);
    tab.setColumnWidth(3, 150); tab.setColumnWidth(4, 340);
    tab.getRange(list.length + 3, 1).setValue(
      "このURLをお客様へお送りください。お客様が申込完了メールのスクリーンショットを" +
      "撮り直して送ると、こちらの記録が自動で差し替わります。" +
      "URLは1回だけ使えます（期限を過ぎたものは再発行が必要です）。" +
      "お願いするときは「メールの受信日時が画面に写るように撮ってください」と添えてください。");
    reps++;
  });
  return reps;
}

// ---------------------------------------------------------------
// 公開ページからの参照（doGet）: 案件名と申込日だけ返す
// **顧客名は返さない。** URLが漏れたときに誰の何かが分かってしまうため。
// ---------------------------------------------------------------
function resubmitPayload_(token) {
  const hit = rsFindByToken_(token);
  if (!hit) return { ok: false, reason: "notfound" };
  const v = hit.values;
  const state = String(v[RSC_STATE - 1] || "");
  if (state === RS_STATE_DONE) return { ok: false, reason: "used" };
  if (state === RS_STATE_VOID) return { ok: false, reason: "void" };
  if (rsIsExpired_(v)) return { ok: false, reason: "expired" };
  return {
    ok: true,
    caseName: String(v[RSC_CASE - 1] || ""),
    appliedAt: String(v[RSC_RECV - 1] || "").slice(0, 10),
    expire: String(v[RSC_EXPIRE - 1] || "")
  };
}

// ---------------------------------------------------------------
// 公開ページからの送信（doPost）
// ---------------------------------------------------------------
function handleResubmit_(data) {
  const hit = rsFindByToken_(data.token);
  if (!hit) return { result: "error", message: "このURLは無効です。" };
  const v = hit.values, sh = hit.sheet, row = hit.row;

  const state = String(v[RSC_STATE - 1] || "");
  if (state === RS_STATE_DONE) return { result: "error", message: "このURLは既に使用されています。" };
  if (state === RS_STATE_VOID) return { result: "error", message: "このURLは無効です。" };
  if (rsIsExpired_(v))         return { result: "error", message: "このURLは有効期限が切れています。" };
  if (!data.screenshot)        return { result: "error", message: "画像が添付されていません。" };

  // **行番号だけを信じて書かない。** 案件シートの行は編集で動く。
  const ss = getOrCreateSpreadsheet();
  const target = ss.getSheetByName(String(v[RSC_SHEET - 1] || ""));
  if (!target) return { result: "error", message: "対象の記録が見つかりませんでした。" };
  const targetRow = Number(v[RSC_ROW - 1] || 0);
  if (!targetRow) return { result: "error", message: "対象の記録が見つかりませんでした。" };

  const lastCol = target.getLastColumn();
  const width = lastCol - ANSWER_START_COL + 1;
  const headers = target.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
  const iName = headers.indexOf("お名前");
  const iRecv = headers.indexOf("受信日時");
  const iShot = headers.lastIndexOf("スクショURL");
  if (iShot < 0) return { result: "error", message: "対象の記録が見つかりませんでした。" };

  const cur = target.getRange(targetRow, ANSWER_START_COL, 1, width).getValues()[0];
  if (iName >= 0 && String(cur[iName] || "").trim() !== String(v[RSC_NAME - 1] || "").trim()) {
    return { result: "error", message: "対象の記録が見つかりませんでした。" };
  }
  if (iRecv >= 0) {
    const a = aspToMillis_(cur[iRecv]), b = aspToMillis_(v[RSC_RECV - 1]);
    if (a && b && Math.abs(a - b) > 60000) {
      return { result: "error", message: "対象の記録が見つかりませんでした。" };
    }
  }

  const oldUrl = String(cur[iShot] || "");
  const url = saveScreenshot(
    data.screenshot,
    data.screenshotName || "resubmit.png",
    { name: String(v[RSC_NAME - 1] || "再提出"), clickTime: null }
  );
  if (String(url).indexOf("保存エラー") === 0) {
    return { result: "error", message: "画像を保存できませんでした。時間をおいて再度お試しください。" };
  }

  // 申請行を更新（**旧URLは台帳へ残す**。証拠を消さない）
  target.getRange(targetRow, ANSWER_START_COL + iShot).setValue(url);

  sh.getRange(row, RSC_STATE).setValue(RS_STATE_DONE);
  sh.getRange(row, RSC_USED).setValue(formatJST(new Date()));
  sh.getRange(row, RSC_OLDURL).setValue(oldUrl);
  sh.getRange(row, RSC_NEWURL).setValue(url);

  // 承認漏れ管理 側も、あれば追随させる（次の棚卸しを待たずに営業へ戻す）
  try {
    const ag = ss.getSheetByName(AG_MANAGE_SHEET);
    if (ag && ag.getLastRow() >= 2) {
      const agData = ag.getRange(2, 1, ag.getLastRow() - 1, AG_HEADERS.length).getValues();
      const key = agKeyOf_(v[RSC_CASE - 1], v[RSC_NAME - 1], aspToMillis_(v[RSC_RECV - 1]));
      for (let i = 0; i < agData.length; i++) {
        const k = agKeyOf_(agData[i][AGC_CASE - 1], agData[i][AGC_NAME - 1],
                           aspToMillis_(agData[i][AGC_RECV - 1]));
        if (k !== key) continue;
        ag.getRange(i + 2, AGC_SHOT).setValue(url);
        ag.getRange(i + 2, AGC_SALESCHK).setValue("未確認");   // 撮り直したので営業が見直す
        ag.getRange(i + 2, AGC_SALESCMT).setValue("再提出あり（" + formatJST(new Date()) + "）");
        break;
      }
    }
  } catch (e) {
    Logger.log("承認漏れ管理への追随に失敗（再提出自体は成功）: " + e);
  }

  SpreadsheetApp.flush();

  // 受け取ったことを社内へ通知する。黙って溜まると誰も見に行かない。
  try {
    notifyLineGroup("【スクショ再提出】\n案件: " + String(v[RSC_CASE - 1] || "") +
                    "\nお客様: " + String(v[RSC_NAME - 1] || "") +
                    "\n担当: " + String(v[RSC_REF - 1] || "") +
                    "\nスクショ: " + url +
                    "\n\n営業確認を「未確認」に戻しました。日時が写っているかご確認ください。");
  } catch (e) {
    Logger.log("再提出のLINE通知に失敗（再提出自体は成功）: " + e);
  }

  return { result: "success" };
}

// ---------------------------------------------------------------
// メニュー
// ---------------------------------------------------------------
function issueResubmitFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = issueResubmitTokens();
    if (!r.issued && !r.reused) {
      ui.alert("再提出URLを発行する対象がありませんでした。\n\n" +
               "「営業確認」が『要再取得』の行が対象です。" +
               "先に「営業の確認結果を取り込む」を実行してください。");
      return;
    }
    const reps = pushResubmitTabs_(r.rows);
    ui.alert("再提出URLを発行しました。\n\n" +
             "新規発行: " + r.issued + " 件\n" +
             "既存URLを再掲: " + r.reused + " 件（送信済みのURLは無効にしません）\n" +
             "担当別タブ: " + reps + " 名分\n\n" +
             "担当別ステータス表（SS2）の「" + AG_RESUBMIT_PREFIX + "＜担当名＞」タブに" +
             "URLが入っています。担当者からお客様へお送りください。\n" +
             "有効期限は発行から " + RESUBMIT_VALID_DAYS + " 日です。");
  } catch (e) {
    ui.alert("発行できませんでした。\n\n" + e);
  }
}

// 誤って発行したURLを止める（顧客へ送る前／送った後どちらでも）
function voidResubmitToken(token) {
  const hit = rsFindByToken_(token);
  if (!hit) throw new Error("トークンが見つかりません: " + token);
  hit.sheet.getRange(hit.row, RSC_STATE).setValue(RS_STATE_VOID);
  SpreadsheetApp.flush();
  return true;
}
