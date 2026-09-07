// =============================================
// 顧客向け公式LINE リッチメニュー用の読み取り口
// （2026-09-07 追加）
//
// 統合アプリ（Vercel・プロジェクト「アプリ化」）が、顧客のLINEメニューを組み立てるために
// 「いま稼働している案件」と「顧客ごとの申込・承認状況」をここから取る。
//
// **読み取りだけ。** このファイルは1行もシートへ書かない。
//
// 認証は合言葉。**こちら側は SHA-256 しか持たない**（ScriptProperties の
// CUSTOMER_MENU_SECRET_SHA256）。平文を持つのは統合アプリの環境変数だけ。
//
// **プロジェクトの設定画面はプロパティの値をそのまま表示する。** 平文を置くと、画面を開いた人・
// スクリーンショット・自動操作のログに残る。ハッシュなら漏れてもこの口は叩けない。
// なお秘密値はソースにも書かない（このリポジトリは GitHub Pages の配信元で public）。
// 未設定のあいだは誰も通らない（fail-closed）。
//
//   POST <本番/execのURL>
//   { "action": "customerMenuSnapshot",
//     "issuedAt": "2026-09-07T12:00:00.000Z",
//     "nonce": "<uuid>",
//     "token": "<平文の合言葉>" }
//
// 5分より古い issuedAt は受けない。nonce は使い捨て（CacheService で10分覚える）。
// =============================================

const CUSTOMER_MENU_SECRET_PROPERTY = "CUSTOMER_MENU_SECRET_SHA256";
const CUSTOMER_MENU_SKEW_MS         = 5 * 60 * 1000;
const CUSTOMER_MENU_NONCE_TTL_SEC   = 600;
const CUSTOMER_MENU_NONCE_PREFIX    = "CUSTOMER_MENU_NONCE_";

// 申請状況一覧が将来ふくらんでも応答が壊れないよう上限を置く。
// 打ち切ったときは truncated を立てる。**黙って減らさない。**
const CUSTOMER_MENU_MAX_ROWS = 6000;

function customerMenuSecret_() {
  return PropertiesService.getScriptProperties().getProperty(CUSTOMER_MENU_SECRET_PROPERTY) || "";
}

function customerMenuHex_(bytes) {
  return bytes.map(function (b) {
    const v = b < 0 ? b + 256 : b;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

function customerMenuSha256_(value) {
  return customerMenuHex_(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
  );
}

// nonce の使い回しを弾く。取れなかった（キャッシュ障害）ときは通さない。
function customerMenuReserveNonce_(nonce) {
  const cache = CacheService.getScriptCache();
  if (!cache) return false;
  const key = CUSTOMER_MENU_NONCE_PREFIX + Utilities.base64EncodeWebSafe(nonce).slice(0, 200);
  if (cache.get(key)) return false;
  cache.put(key, "1", CUSTOMER_MENU_NONCE_TTL_SEC);
  return true;
}

function customerMenuReject_(code) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: code }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doPost から呼ばれる入口。検証してスナップショットを返す。
 * 検証に落ちた理由は呼び出し側へ返すが、**秘密値そのものは何も返さない。**
 */
function handleCustomerMenuSnapshot_(data) {
  const expected = customerMenuSecret_().toLowerCase();
  if (expected.length !== 64) return customerMenuReject_("not_configured");

  const issuedAt = String((data && data.issuedAt) || "");
  const nonce    = String((data && data.nonce) || "");
  const token    = String((data && data.token) || "");
  if (!issuedAt || nonce.length < 8 || token.length < 16) return customerMenuReject_("invalid_request");

  const issuedMs = Date.parse(issuedAt);
  if (!issuedMs || Math.abs(Date.now() - issuedMs) > CUSTOMER_MENU_SKEW_MS) return customerMenuReject_("stale");

  if (!_customerLineBridgeConstantTimeEqual(expected, customerMenuSha256_(token))) {
    return customerMenuReject_("invalid_token");
  }
  if (!customerMenuReserveNonce_(nonce)) return customerMenuReject_("replayed");

  let payload;
  try {
    payload = buildCustomerMenuSnapshot_();
  } catch (err) {
    Logger.log("customerMenuSnapshot エラー: " + err);
    return customerMenuReject_("unavailable");
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 案件マスタ（稼働の正）と申請状況一覧（申込の正）から、顧客メニューに要るものだけを作る。
 *
 * **案件シート（設定_*）は読まない。** 全行スキャンは56秒かかった前科がある。
 */
function buildCustomerMenuSnapshot_() {
  const ss = getOrCreateSpreadsheet();

  // --- 案件マスタ ---
  const cases = [];
  const codeByCaseName = {};
  const masterSheet = ss.getSheetByName(CASE_MASTER_SHEET);
  if (!masterSheet) throw new Error(CASE_MASTER_SHEET + " が見つかりません");
  const masterLastRow = masterSheet.getLastRow();
  if (masterLastRow >= 2) {
    masterSheet.getRange(2, 1, masterLastRow - 1, CM_HEADERS.length).getValues().forEach(function (r) {
      const code = String(r[CM_COL_CODE - 1] || "").trim();
      const name = String(r[CM_COL_NAME - 1] || "").trim();
      if (!code || !name) return;
      cases.push({ code: code, name: name, active: r[CM_COL_ACTIVE - 1] === true });
      codeByCaseName[normalizeName(name)] = code;
    });
  }

  // --- 申請状況一覧 ---
  const applications = [];
  let truncated = false;
  const statusSheet = ss.getSheetByName(APP_STATUS_SHEET);
  if (!statusSheet) throw new Error(APP_STATUS_SHEET + " が見つかりません");
  const statusLastRow = statusSheet.getLastRow();
  if (statusLastRow >= 2) {
    let rowCount = statusLastRow - 1;
    if (rowCount > CUSTOMER_MENU_MAX_ROWS) { rowCount = CUSTOMER_MENU_MAX_ROWS; truncated = true; }
    // 新しい行が上に来る作りなので、先頭から必要数だけ読む。
    const values = statusSheet.getRange(2, 1, rowCount, APP_STATUS_HEADERS.length).getValues();
    const repMemo = {};

    values.forEach(function (r) {
      const customerName = String(r[2] || "").trim();
      const caseName     = String(r[1] || "").trim();
      if (!customerName || !caseName) return;

      const rawRep = String(r[3] || "").trim();
      if (repMemo[rawRep] === undefined) repMemo[rawRep] = rawRep ? agCanonicalRep_(rawRep) : "";
      const flags = getAdvertiserApprovalFlags(r[5]);

      applications.push({
        customerName: customerName,
        repName:      repMemo[rawRep] || rawRep,
        caseName:     caseName,
        caseCode:     codeByCaseName[normalizeName(caseName)] || null,
        appliedOn:    String(toDisplayDate_(r[0]) || "").substring(0, 10),
        status:       flags.approved ? "承認" : (flags.trackingMissing ? "非承認" : "確認中")
      });
    });
  }

  return {
    ok: true,
    generatedAt: formatJST(new Date()),
    truncated: truncated,
    cases: cases,
    applications: applications
  };
}
