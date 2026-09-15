// =============================================
// 案件の稼働/停止を外から切り替える口（2026-09-15 追加）
//
// ワークスペース側の「案件ステータス監視」（非公開）が、LINE の停止/再開の連絡を読んで
// 案件マスタの「稼働」を切り替えるために叩く。人がスプレッドシートのチェックボックスを
// 押すのと同じこと（setCaseActive_ → applyCaseVisibility）しかしない。
//
// 認証は CustomerMenu.gs と同じ型。**こちらは SHA-256 しか持たない**
// （ScriptProperties の CASE_STATUS_SECRET_SHA256）。平文は監視側の private/ だけが持つ。
// 5分より古い issuedAt は受けない。nonce は使い捨て。未設定のあいだは誰も通らない。
// なお秘密値はソースにも書かない（このリポジトリは GitHub Pages の配信元で public）。
//
//   POST <本番/execのURL>
//   { "action": "caseStatus", "issuedAt": "2026-09-15T04:00:00.000Z",
//     "nonce": "<uuid>", "token": "<平文の合言葉>", "op": "list" }
//       → { ok:true, cases:[{code,name,active,updatedAt,note}] }
//   { ..., "op": "set", "caseCode": "acom", "active": false, "note": "FPR停止連絡" }
//       → 稼働を書き換えて表示を反映し、変更後の一覧を返す
//
// **できるのは案件マスタの「稼働」と「備考」の書き換えだけ。** 案件の追加・削除・URL の
// 変更はできない。備考は追記で、直近の分だけを残す。
// =============================================

const CASE_STATUS_SECRET_PROPERTY = "CASE_STATUS_SECRET_SHA256";
const CASE_STATUS_NOTE_MAX        = 200;  // 1回に受け取る備考の長さ
const CASE_STATUS_NOTE_KEEP       = 400;  // 備考セルに残す長さ（超えたら古い側を落とす）

function caseStatusSecret_() {
  return String(PropertiesService.getScriptProperties().getProperty(CASE_STATUS_SECRET_PROPERTY) || "").toLowerCase();
}

function caseStatusJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 案件マスタをそのまま一覧にする（読み取り）。
function caseStatusList_() {
  const sh = getCaseMasterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rows = sh.getRange(2, 1, lastRow - 1, CM_HEADERS.length).getValues();
  const out = [];
  rows.forEach(function (r) {
    const code = String(r[CM_COL_CODE - 1] || "").trim();
    if (!code) return;
    out.push({
      code: code,
      name: String(r[CM_COL_NAME - 1] || code).trim(),
      active: r[CM_COL_ACTIVE - 1] === true,
      updatedAt: r[CM_COL_UPDATED - 1] == null ? "" : String(r[CM_COL_UPDATED - 1]),
      note: r[CM_COL_NOTE - 1] == null ? "" : String(r[CM_COL_NOTE - 1])
    });
  });
  return out;
}

// 備考へ追記する。人が書いた既存の備考は消さず、長すぎるときだけ古い側を落とす。
function caseStatusAppendNote_(row, note) {
  const sh = getCaseMasterSheet_();
  const cell = sh.getRange(row, CM_COL_NOTE);
  const prev = cell.getValue() == null ? "" : String(cell.getValue()).trim();
  let next = prev ? prev + " / " + note : note;
  if (next.length > CASE_STATUS_NOTE_KEEP) next = next.slice(next.length - CASE_STATUS_NOTE_KEEP);
  // 先頭に = や + が来ると数式として保存されるので、文字列であることを固定する
  cell.setValue("'" + next.replace(/^'+/, ""));
}

/**
 * doPost から呼ばれる入口。検証して一覧を返すか、稼働を書き換える。
 * 検証に落ちた理由は呼び出し側へ返すが、**秘密値そのものは何も返さない。**
 */
function handleCaseStatus_(data) {
  const expected = caseStatusSecret_();
  if (expected.length !== 64) return customerMenuReject_("not_configured");

  const issuedAt = String((data && data.issuedAt) || "");
  const nonce    = String((data && data.nonce) || "");
  const token    = String((data && data.token) || "");
  const op       = String((data && data.op) || "");
  if (!issuedAt || nonce.length < 8 || token.length < 16) return customerMenuReject_("invalid_request");

  const issuedMs = Date.parse(issuedAt);
  if (!issuedMs || Math.abs(Date.now() - issuedMs) > CUSTOMER_MENU_SKEW_MS) return customerMenuReject_("stale");

  if (!_customerLineBridgeConstantTimeEqual(expected, customerMenuSha256_(token))) {
    return customerMenuReject_("invalid_token");
  }
  if (!customerMenuReserveNonce_(nonce)) return customerMenuReject_("replayed");

  if (op === "list") {
    return caseStatusJson_({ ok: true, cases: caseStatusList_() });
  }

  if (op === "set") {
    const code = String((data && data.caseCode) || "").trim();
    if (!code || typeof data.active !== "boolean") return customerMenuReject_("invalid_request");
    const note = String((data && data.note) || "").trim().slice(0, CASE_STATUS_NOTE_MAX);

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return customerMenuReject_("busy");
    try {
      const before = readCaseActiveMap_();
      if (!(code in before)) return customerMenuReject_("unknown_case");
      const changed = before[code] !== data.active;
      const r = setCaseActive_(code, data.active);
      if (!r.found) return customerMenuReject_("unknown_case");
      if (note) caseStatusAppendNote_(r.row, note);

      // 表示の切り替えは見やすさのためだけ。失敗しても稼働の書き換えは成立している。
      let visibilityApplied = false;
      try { applyCaseVisibility(); visibilityApplied = true; }
      catch (e) { Logger.log("caseStatus: applyCaseVisibility に失敗: " + e); }

      return caseStatusJson_({
        ok: true,
        caseCode: code,
        active: data.active,
        changed: changed,
        row: r.row,
        visibilityApplied: visibilityApplied,
        cases: caseStatusList_()
      });
    } finally {
      lock.releaseLock();
    }
  }

  return customerMenuReject_("unknown_op");
}
