// =============================================
// 顧客向け公式LINE → 市場作り管理BOT のグループ通知
// （2026-09-08 追加）
//
// 統合アプリ（Vercel・プロジェクト「アプリ化」）が、顧客の動き（申し込みたい／質問／
// メッセージ）を**アフィリエイト報告用のグループ**へ知らせるために叩く。
//
// **なぜここを経由するのか。** 通知先のグループにいるのは「市場作り管理BOT」であって、
// 顧客向け公式アカウント（@730cphak）ではない。顧客向けの側からグループへ送るには、
// あちらをグループへ招待する設定が要り、招待するとグループ内の発言まで顧客の受信経路へ
// 流れ込む。**既にこのBOTがそのグループにいる**ので、ここから出すのが素直で副作用が無い。
//
// 送信の実体は Code.gs の notifyLineGroup()。トークンとグループIDは
// ScriptProperties（LINE_CHANNEL_TOKEN / LINE_GROUP_ID）にあり、このファイルは持たない。
//
// 認証は CustomerMenu.gs と同じ合言葉（CUSTOMER_MENU_SECRET_SHA256）を使う。
// **この合言葉は「読み取りだけ」ではなくなった。** 持っている相手はグループへ1通投げられる。
// 平文を持つのは統合アプリの環境変数だけで、こちらは SHA-256 しか持たない。
//
//   POST <本番/execのURL>
//   { "action": "customerLineNotify",
//     "issuedAt": "2026-09-08T12:00:00.000Z",
//     "nonce": "<uuid>",
//     "token": "<平文の合言葉>",
//     "text": "本文" }
//
// 5分より古い issuedAt は受けない。nonce は使い捨て。未設定のあいだは誰も通らない。
// =============================================

// LINEの1通は5,000文字までだが、グループの通知にそこまで要らない。
// 長い本文を投げられて読めなくなるほうが困るので、こちら側でも切る。
const CUSTOMER_NOTIFY_MAX_TEXT = 1000;

/**
 * doPost から呼ばれる入口。検証してグループへ1通送る。
 *
 * **顧客の氏名・本文は載せない前提**で呼ばれる（統合アプリ側が種別と対応画面のURLだけを送る）。
 * ここでは中身を検査しないが、上限だけは掛ける。
 */
function handleCustomerLineNotify_(data) {
  const expected = customerMenuSecret_().toLowerCase();
  if (expected.length !== 64) return customerMenuReject_("not_configured");

  const issuedAt = String((data && data.issuedAt) || "");
  const nonce    = String((data && data.nonce) || "");
  const token    = String((data && data.token) || "");
  const text     = String((data && data.text) || "").trim();
  if (!issuedAt || nonce.length < 8 || token.length < 16) return customerMenuReject_("invalid_request");
  if (!text) return customerMenuReject_("empty_text");

  const issuedMs = Date.parse(issuedAt);
  if (!issuedMs || Math.abs(Date.now() - issuedMs) > CUSTOMER_MENU_SKEW_MS) return customerMenuReject_("stale");

  if (!_customerLineBridgeConstantTimeEqual(expected, customerMenuSha256_(token))) {
    return customerMenuReject_("invalid_token");
  }
  if (!customerMenuReserveNonce_(nonce)) return customerMenuReject_("replayed");

  // 送信先が未設定なら、通ったふりをしない。呼び出し側が記録して気づけるようにする。
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("LINE_CHANNEL_TOKEN") || !props.getProperty("LINE_GROUP_ID")) {
    return customerMenuReject_("group_not_configured");
  }

  try {
    notifyLineGroup(text.slice(0, CUSTOMER_NOTIFY_MAX_TEXT));
  } catch (err) {
    Logger.log("customerLineNotify エラー: " + err);
    return customerMenuReject_("unavailable");
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
