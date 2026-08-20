// =============================================
// ASP獲得ログとの突合
// （2026-08-20 追加）
//
// ASP（FPR）の獲得ログには顧客の氏名が無いが、クリック日時がある。
// 自社フォームもクリック日時を記録しているので、これをキーに突き合わせる。
//
// 2026-08-20 の実測では、クリック日時での一致は 217/451。一致するときは中央値0秒
// （完全一致111件・±2秒183件）でキーとしては正しいが、半分は端末時計のズレ等で拾えない。
// **したがって自動で一括更新はしない。** 突合結果を出し、人がチェックした行だけ反映する。
//
// 書き戻すのは「ASPが承認しているのに自社が承認になっていない」方向だけにする。
// 逆方向（自社が承認・ASPが承認待ち/否認）は、承認を消す判断が業務判断になるため報告のみ。
//
// ASPはセッションCookieでログインするためGASからCSVを取得できない。
// ユーザーがASPの獲得ログからCSVをダウンロードし、ASP獲得ログシートへ取り込む。
// =============================================

const ASP_LOG_SHEET    = "ASP獲得ログ";
const ASP_RESULT_SHEET = "ASP突合結果";
const ASP_MATCH_TOLERANCE_SEC = 120;

const ASP_RESULT_HEADERS = [
  "適用", "判定", "ASPクリック日時", "ASP広告名", "ASP報酬額", "ASPステータス",
  "自社案件", "自社顧客名", "自社の承認", "ズレ秒", "対象シート", "対象行"
];
const ASPR_COL_APPLY = 1;

function getAspLogSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(ASP_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ASP_LOG_SHEET);
    const head = ["承認日時", "クリック日時", "注文日時", "広告ID", "広告名", "サイト名", "OS",
                  "リファラ", "報酬額", "ステータス", "セッションID",
                  "CL付加情報1", "CL付加情報2", "CL付加情報3", "CL付加情報4", "CL付加情報5"];
    const r = sh.getRange(1, 1, 1, head.length);
    r.setValues([head]);
    r.setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.getRange(3, 1).setValue(
      "ASPの獲得ログからCSVをダウンロードし、このシートへ取り込んでください（ファイル > インポート > 現在のシートを置換）。"
    );
  }
  return sh;
}

// 日時セルは Date でも文字列でも来る。ミリ秒に正規化する。
function aspToMillis_(v) {
  if (v instanceof Date) return v.getTime();
  const s = String(v || "").trim();
  if (!s || s.indexOf("0000-00-00") === 0) return 0;
  const m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(s);
  if (!m) return 0;
  // JSTとして解釈する（ASPも自社シートもJST表記）
  return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4] - 9, +m[5], +(m[6] || 0));
}

function aspApprovalOf_(v) {
  const f = getAdvertiserApprovalFlags(v);
  if (f.approved) return "承認";
  if (f.trackingMissing) return "トラッキング漏れ";
  return "未記入";
}

// 自社の全回答行を {ms, case, name, appr, sheetName, row} で集める
function collectOwnClickRows_() {
  const ss = getOrCreateSpreadsheet();
  const out = [];
  listCaseSheets_(ss).forEach(function (c) {
    const sheet = c.sheet;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;
    const width = lastCol - ANSWER_START_COL + 1;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const iClick = headers.indexOf("クリック日時");
    if (iClick < 0) return;
    const iName = headers.indexOf("お名前");
    const iAppr = headers.indexOf("承認");
    const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, width).getValues();
    for (let k = 0; k < data.length; k++) {
      const ms = aspToMillis_(data[k][iClick]);
      if (!ms) continue;
      out.push({
        ms: ms,
        caseName: c.name,
        name: iName >= 0 ? String(data[k][iName] || "") : "",
        appr: iAppr >= 0 ? data[k][iAppr] : "",
        apprCol: iAppr >= 0 ? ANSWER_START_COL + iAppr : 0,
        sheetName: sheet.getName(),
        row: k + 2
      });
    }
  });
  out.sort(function (a, b) { return a.ms - b.ms; });
  return out;
}

// ASP獲得ログを読む
function collectAspLogRows_() {
  const sh = getAspLogSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return [];
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const iClick = headers.indexOf("クリック日時");
  const iAd    = headers.indexOf("広告名");
  const iRw    = headers.indexOf("報酬額");
  const iSt    = headers.indexOf("ステータス");
  if (iClick < 0 || iSt < 0) {
    throw new Error("ASP獲得ログの見出しに「クリック日時」「ステータス」がありません。CSVを取り込み直してください。");
  }
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = [];
  data.forEach(function (r) {
    const ms = aspToMillis_(r[iClick]);
    if (!ms) return;
    out.push({
      ms: ms,
      ad: iAd >= 0 ? String(r[iAd] || "") : "",
      reward: iRw >= 0 ? r[iRw] : "",
      status: String(r[iSt] || "").trim()
    });
  });
  return out;
}

// ASPの広告名と自社の案件名が同じものを指しているか。
// ASP側は「【CTIA様専用】…」「…（特別単価）」のような修飾が付くので落としてから比べる。
// 例: イエイ ↔ 不動産一括査定イエイ / ポケットリサーチ（特別単価）↔ ポケットリサーチ
function aspNameMatches_(adName, caseName) {
  const strip = function (s) {
    return String(s || "")
      .replace(/【[^】]*】/g, "")
      .replace(/（[^）]*）/g, "")
      .replace(/\([^)]*\)/g, "")
      .replace(/\s+/g, "")
      .toLowerCase();
  };
  const a = strip(adName), c = strip(caseName);
  if (!a || !c) return false;
  return a === c || a.indexOf(c) >= 0 || c.indexOf(a) >= 0;
}

// クリック日時が近い自社行を探す。
// **時刻だけで決めない。** 許容範囲内の候補のうち案件名が対応するものを優先する。
// 2026-08-20 の検証で、案件の違う行に68秒差で誤マッチする例が実在したため。
function nearestOwn_(own, ms, adName) {
  let lo = 0, hi = own.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (own[mid].ms < ms) lo = mid + 1; else hi = mid - 1;
  }
  const from = Math.max(0, lo - 12), to = Math.min(own.length, lo + 12);
  let named = null, any = null;
  for (let i = from; i < to; i++) {
    const d = Math.abs(own[i].ms - ms);
    if (any === null || d < any.d) any = { d: d, row: own[i] };
    if (adName && aspNameMatches_(adName, own[i].caseName)) {
      if (named === null || d < named.d) named = { d: d, row: own[i], nameOk: true };
    }
  }
  // 案件名が合う候補が許容範囲内にあればそれを採る
  if (named && named.d <= ASP_MATCH_TOLERANCE_SEC * 1000) return named;
  if (any) { any.nameOk = false; return any; }
  return null;
}

// 突合して ASP突合結果 シートを作る（書き戻しはしない）
function reconcileAspLog() {
  const asp = collectAspLogRows_();
  if (!asp.length) {
    throw new Error("ASP獲得ログにデータがありません。先にCSVを取り込んでください。");
  }
  const own = collectOwnClickRows_();

  const rows = [];
  const tally = { fix: 0, ok: 0, review: 0, nomatch: 0 };

  asp.forEach(function (a) {
    const near = nearestOwn_(own, a.ms, a.ad);
    const inRange = near && near.d <= ASP_MATCH_TOLERANCE_SEC * 1000;
    const hit = inRange ? near.row : null;
    const delta = near ? Math.round(near.d / 1000) : "";

    if (!hit) {
      tally.nomatch++;
      rows.push([false, "未一致（自社に該当なし）", new Date(a.ms), a.ad, a.reward, a.status,
                 "", "", "", delta, "", ""]);
      return;
    }
    const ownAppr = aspApprovalOf_(hit.appr);

    // 案件名が対応しない行は、時刻がたまたま近いだけの別人の可能性がある。
    // 承認を書き込む対象にはせず、人が見る「要確認」に落とす。
    if (!near.nameOk) {
      tally.review++;
      rows.push([false, "要確認: 時刻は近いが案件名が対応しない", new Date(a.ms), a.ad, a.reward, a.status,
                 hit.caseName, hit.name, ownAppr, delta, "", ""]);
      return;
    }

    let verdict;
    if (a.status === "承認" && ownAppr !== "承認") {
      verdict = "要修正: ASPは承認・自社は" + ownAppr;
      tally.fix++;
    } else if (a.status !== "承認" && ownAppr === "承認") {
      verdict = "要確認: 自社は承認・ASPは" + a.status;
      tally.review++;
    } else {
      verdict = "一致";
      tally.ok++;
    }
    rows.push([false, verdict, new Date(a.ms), a.ad, a.reward, a.status,
               hit.caseName, hit.name, ownAppr, delta, hit.sheetName, hit.row]);
  });

  // 要修正 → 要確認 → 未一致 → 一致 の順に並べる
  const order = function (v) {
    if (v.indexOf("要修正") === 0) return 0;
    if (v.indexOf("要確認") === 0) return 1;
    if (v.indexOf("未一致") === 0) return 2;
    return 3;
  };
  rows.sort(function (x, y) {
    const d = order(x[1]) - order(y[1]);
    return d !== 0 ? d : (y[2] - x[2]);
  });

  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(ASP_RESULT_SHEET);
  if (!sh) sh = ss.insertSheet(ASP_RESULT_SHEET);
  sh.clear();

  const hr = sh.getRange(1, 1, 1, ASP_RESULT_HEADERS.length);
  hr.setValues([ASP_RESULT_HEADERS]);
  hr.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff").setWrap(true);
  sh.setFrozenRows(1);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, ASP_RESULT_HEADERS.length).setValues(rows);
    const cb = sh.getRange(2, ASPR_COL_APPLY, rows.length, 1);
    cb.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    cb.setHorizontalAlignment("center");
    // 要修正の行だけ色を付ける
    const bg = rows.map(function (r) {
      const c = String(r[1]).indexOf("要修正") === 0 ? "#fef9c3"
              : String(r[1]).indexOf("要確認") === 0 ? "#ffe4e6" : "#ffffff";
      return ASP_RESULT_HEADERS.map(function () { return c; });
    });
    sh.getRange(2, 1, rows.length, ASP_RESULT_HEADERS.length).setBackgrounds(bg);
  }
  sh.setColumnWidth(2, 240);
  sh.setColumnWidth(3, 150);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(7, 200);

  return { total: asp.length, fix: tally.fix, review: tally.review, ok: tally.ok, nomatch: tally.nomatch };
}

// チェックの付いた「要修正」行だけ、自社シートの承認欄へ ⭕ を書く。
// 逆方向（承認を消す）は業務判断なので、この関数では一切行わない。
function applyAspReconciliation() {
  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(ASP_RESULT_SHEET);
  if (!sh) throw new Error("ASP突合結果 がありません。先に突合を実行してください。");
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { applied: 0, skipped: 0 };

  const data = sh.getRange(2, 1, lastRow - 1, ASP_RESULT_HEADERS.length).getValues();
  let applied = 0, skipped = 0;

  data.forEach(function (r) {
    if (r[0] !== true) return;
    const verdict = String(r[1] || "");
    const sheetName = String(r[10] || "");
    const row = Number(r[11] || 0);
    if (verdict.indexOf("要修正") !== 0 || !sheetName || !row) { skipped++; return; }

    const target = ss.getSheetByName(sheetName);
    if (!target) { skipped++; return; }
    const lastCol = target.getLastColumn();
    const width = lastCol - ANSWER_START_COL + 1;
    const headers = target.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const iAppr = headers.indexOf("承認");
    if (iAppr < 0) { skipped++; return; }
    target.getRange(row, ANSWER_START_COL + iAppr).setValue("⭕");
    applied++;
  });
  SpreadsheetApp.flush();
  return { applied: applied, skipped: skipped };
}

function reconcileAspLogFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = reconcileAspLog();
    ui.alert(
      "ASP獲得ログと突合しました。" + "\n\n" +
      "ASPの行数: " + r.total + "\n" +
      "要修正（ASPは承認・自社は未承認）: " + r.fix + " 件\n" +
      "要確認（自社は承認・ASPは未承認）: " + r.review + " 件\n" +
      "一致: " + r.ok + " 件\n" +
      "未一致（自社に該当なし）: " + r.nomatch + " 件\n\n" +
      "ASP突合結果 シートで内容を確認し、直す行の「適用」にチェックを入れてから" +
      "「ASP突合の修正を反映」を実行してください。"
    );
  } catch (e) {
    ui.alert("突合できませんでした。" + "\n\n" + e);
  }
}

function applyAspReconciliationFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert(
    "「適用」にチェックの付いた要修正行について、自社の承認欄へ ⭕ を書き込みます。" + "\n" +
    "承認を消す操作は行いません。よろしいですか。",
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  try {
    const r = applyAspReconciliation();
    ui.alert("反映しました。" + "\n\n" +
             "書き込み: " + r.applied + " 件\n" +
             "スキップ: " + r.skipped + " 件（要修正以外、または対象が特定できない行）" + "\n\n" +
             "広告主シートへ提出する前に、対象月を再生成してください。");
  } catch (e) {
    ui.alert("反映に失敗しました。" + "\n\n" + e);
  }
}
