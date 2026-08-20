// =============================================
// 顧客向け公式LINE：営業担当ごとの顧客登録数
// （2026-08-20 追加）
//
// 日次レポートに「各営業マンの顧客が何人LINE登録したか」を足すためのもの。
//
// データは統合アプリ（Vercel + Neon・プロジェクト「アプリ化」）が持っている。
// GASからNeonへは直接つなげないので、統合アプリの内部APIを叩く。
//   GET <URL>/api/internal/customer-line-stats?date=YYYY-MM-DD
//   Authorization: Bearer <専用の秘密値>
//
// **この秘密値の平文を持つのはここ（ScriptProperties）だけ。** 統合アプリ側は
// CUSTOMER_LINE_STATS_SECRET_SHA256 として SHA-256 しか保存していないので、
// 万一サーバー側の環境変数が漏れても、この口を叩くことはできない。
// 権限も最小で、件数の読み取りしかできない（ジョブ実行はできない）。
//
// **URLと秘密値はソースに書かない。** このリポジトリは GitHub Pages の配信元で public。
// ScriptProperties に置き、未設定なら日次レポートからこの節を黙って省く（fail-closed）。
//   CUSTOMER_LINE_STATS_URL    : 統合アプリのオリジン（例 https://xxxx.vercel.app）
//   CUSTOMER_LINE_STATS_SECRET : 平文。SHA-256 を統合アプリの環境変数へ登録しておく
// =============================================

const CL_STATS_URL_PROPERTY    = "CUSTOMER_LINE_STATS_URL";
const CL_STATS_SECRET_PROPERTY = "CUSTOMER_LINE_STATS_SECRET";

function customerLineStatsConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    origin: (props.getProperty(CL_STATS_URL_PROPERTY) || "").replace(/\/+$/, ""),
    secret: props.getProperty(CL_STATS_SECRET_PROPERTY) || ""
  };
}

// 統合アプリから集計を取る。取れなければ null（日次レポートは止めない）。
function fetchCustomerLineStats_(dateStr) {
  const cfg = customerLineStatsConfig_();
  if (!cfg.origin || !cfg.secret) return null;

  const url = cfg.origin + "/api/internal/customer-line-stats" +
              (dateStr ? "?date=" + encodeURIComponent(dateStr) : "");
  let res;
  try {
    res = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { Authorization: "Bearer " + cfg.secret },
      muteHttpExceptions: true,
      followRedirects: true
    });
  } catch (e) {
    Logger.log("顧客LINE集計の取得に失敗: " + e);
    return null;
  }

  const code = res.getResponseCode();
  if (code !== 200) {
    Logger.log("顧客LINE集計の取得に失敗: HTTP " + code + " " + res.getContentText().slice(0, 200));
    return null;
  }
  try {
    const body = JSON.parse(res.getContentText());
    return body && body.stats ? body.stats : null;
  } catch (e) {
    Logger.log("顧客LINE集計の解析に失敗: " + e);
    return null;
  }
}

// 日次レポートへ差し込む文面を作る。取れなければ空文字。
function buildCustomerLineStatsSection_(dateStr) {
  const stats = fetchCustomerLineStats_(dateStr);
  if (!stats || !stats.rows) return "";

  const rows = stats.rows;
  const t = stats.totals || { daily: 0, monthly: 0, linked: 0 };

  // 昨日の登録が誰も無いときは、担当を全員並べても情報量が無いので1行にまとめる。
  if (!t.daily) {
    return "【顧客LINE登録】" + stats.date + "\n" +
           "・昨日の新規登録: 0件\n" +
           "・今月の登録: " + t.monthly + "件 / 紐づけ済み累計: " + t.linked + "件";
  }

  const lines = rows
    .filter(function (r) { return r.daily > 0 || r.monthly > 0; })
    .map(function (r) {
      return "・" + r.staff + ": " + r.daily + "件（月計: " + r.monthly + "件 / 紐づけ済み: " + r.linked + "件）";
    });

  return "【顧客LINE登録】" + stats.date + "\n" +
         lines.join("\n") + "\n" +
         "昨日合計: " + t.daily + "件 / 今月累計: " + t.monthly + "件";
}

// 設定と疎通の確認用（GASエディタから手動実行）。秘密値は出さない。
function checkCustomerLineStats() {
  const cfg = customerLineStatsConfig_();
  const lines = [];
  lines.push("URL設定: " + (cfg.origin ? "あり（" + cfg.origin + "）" : "未設定"));
  lines.push("秘密値設定: " + (cfg.secret ? "あり（" + cfg.secret.length + "文字）" : "未設定"));
  if (!cfg.origin || !cfg.secret) {
    lines.push("→ 未設定のため日次レポートには顧客LINEの節が出ません。");
    const msg = lines.join("\n");
    Logger.log(msg);
    return msg;
  }
  const stats = fetchCustomerLineStats_(null);
  if (!stats) {
    lines.push("→ 取得に失敗しました。実行ログを確認してください。");
  } else {
    lines.push("対象日: " + stats.date);
    lines.push("担当行数: " + (stats.rows ? stats.rows.length : 0));
    lines.push("昨日: " + stats.totals.daily + "件 / 今月: " + stats.totals.monthly +
               "件 / 紐づけ済み累計: " + stats.totals.linked + "件");
    lines.push("");
    lines.push("--- 日次レポートに出る文面 ---");
    lines.push(buildCustomerLineStatsSection_(null));
  }
  const msg = lines.join("\n");
  Logger.log(msg);
  return msg;
}

function checkCustomerLineStatsFromMenu() {
  SpreadsheetApp.getUi().alert("顧客LINE登録集計の確認\n\n" + checkCustomerLineStats());
}

// 統合アプリのURLと秘密値を設定する（GASエディタから手動実行）。
// 引数で渡した値はソースにもGitにも残らない。実行後は引数を消しておくこと。
function setCustomerLineStatsConfig(origin, secret) {
  if (!origin || !secret) throw new Error("origin と secret の両方を渡してください。");
  if (String(secret).length < 32) throw new Error("秘密値が短すぎます（32文字以上）。");
  PropertiesService.getScriptProperties().setProperties({
    CUSTOMER_LINE_STATS_URL: String(origin).replace(/\/+$/, ""),
    CUSTOMER_LINE_STATS_SECRET: String(secret)
  });
  Logger.log("設定しました。checkCustomerLineStats() で疎通を確認してください。");
  return "設定しました。checkCustomerLineStats() で疎通を確認してください。";
}
