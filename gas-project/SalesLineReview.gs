// =============================================
// 営業ダッシュボード：顧客の公式LINE追加の確認
// （2026-08-31 追加）
//
// お客様が公式LINEを友だち追加して氏名・生年月日・担当営業を送ると、
// 「この人はどのお客様か」を担当営業が確かめる工程が生まれる。
// これまでは統合アプリ（Googleログインが要る管理画面）でしかできなかったので、
// 営業ダッシュボードからも押せるようにする。
//
// データはこのスプレッドシートに無い。統合アプリ（Vercel + Neon・プロジェクト「アプリ化」）
// が正データで、GASからNeonへは直接つなげないので内部APIを叩く。
//   GET  <URL>/api/internal/customer-line/sales-review?staff=<担当名>
//   POST <URL>/api/internal/customer-line/sales-review
//   Authorization: Bearer <専用の秘密値>
//
// **この口でできるのは営業段階の確認だけ。** 最終承認（管理者）はアプリ側でしかできない。
// 合言葉つきURLが漏れても、お客様との紐づけが確定することは無い（二段階確認は崩さない）。
//
// **秘密値は日次レポート用（CUSTOMER_LINE_STATS_SECRET）と別にする。** あちらは
// 「件数を読むだけ」と決めて渡したもので、書ける口で使い回すとその約束が崩れる。
//
// **URLと秘密値はソースに書かない。** このリポジトリは GitHub Pages の配信元で public。
// ScriptProperties に置き、未設定ならこの機能は画面に出ない（fail-closed）。
//   CUSTOMER_LINE_APP_URL  : 統合アプリのオリジン。未設定なら CUSTOMER_LINE_STATS_URL を使う
//   SALES_REVIEW_SECRET    : 平文。SHA-256 を統合アプリの SALES_DASHBOARD_SECRET_SHA256 へ登録
// =============================================

const SLR_APP_URL_PROPERTY = "CUSTOMER_LINE_APP_URL";
const SLR_SECRET_PROPERTY  = "SALES_REVIEW_SECRET";
const SLR_PATH             = "/api/internal/customer-line/sales-review";

// 画面へ出す文言。ITに不慣れな読み手が前提なので、原因ではなく次の動作を書く。
const SLR_MSG_UNAVAILABLE = "いまLINEの確認ができません。事務までご連絡ください。";
const SLR_MSG_NETWORK     = "通信できませんでした。電波の良いところで、もう一度お試しください。";
const SLR_MSG_CONFLICT    = "この方はすでに確認済みか、状況が変わりました。「最新の状態にする」を押してご確認ください。";

function slrConfig_() {
  const props = PropertiesService.getScriptProperties();
  const origin = String(
    props.getProperty(SLR_APP_URL_PROPERTY) ||
    props.getProperty(CL_STATS_URL_PROPERTY) || ""
  ).replace(/\/+$/, "");
  return { origin: origin, secret: props.getProperty(SLR_SECRET_PROPERTY) || "" };
}

// 設定が揃っているか。**揃っていなければ画面にタブ自体を出さない。**
// 中途半端に出すと「押しても何も起きない」になり、営業の信頼を落とす。
function slrEnabled_() {
  const cfg = slrConfig_();
  return !!(cfg.origin && cfg.secret);
}

// 統合アプリを叩く。戻りは { ok, status, body }。例外は投げない（画面を落とさない）。
function slrFetch_(method, query, payload) {
  const cfg = slrConfig_();
  if (!cfg.origin || !cfg.secret) return { ok: false, status: 0, body: null, reason: "unconfigured" };

  const url = cfg.origin + SLR_PATH + (query || "");
  const options = {
    method: method,
    headers: { Authorization: "Bearer " + cfg.secret },
    muteHttpExceptions: true,
    followRedirects: true
  };
  if (payload) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }

  let res;
  try {
    res = UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("LINE確認の通信に失敗: " + e);
    return { ok: false, status: 0, body: null, reason: "network" };
  }
  const status = res.getResponseCode();
  let body = null;
  try {
    body = JSON.parse(res.getContentText());
  } catch (e) {
    body = null;
  }
  // 秘密値は絶対にログへ出さない。URLにも入れていない（Authorizationヘッダのみ）。
  if (status !== 200) Logger.log("LINE確認が失敗: HTTP " + status + " " + JSON.stringify(body));
  return { ok: status === 200, status: status, body: body, reason: "" };
}

// ---------------------------------------------------------------
// 画面へ返すデータ（doGet ?sales_line=<合言葉>）
// ---------------------------------------------------------------

// **本体の画面とは別に読む。** 統合アプリが遅い・落ちているときに、
// 承認確認と顧客一覧まで巻き込んで開けなくなるのを避けるため。
function salesLinePayload_(token) {
  const rep = sdFindByToken_(token);
  if (!rep || rep.status !== SD_STATUS_ACTIVE) {
    return { available: false };
  }
  if (!slrEnabled_()) return { available: false };

  const r = slrFetch_("get", "?staff=" + encodeURIComponent(rep.name), null);
  if (!r.ok) {
    // **理由をそのまま画面へ出さない。** 「対応表に登録されていません」は社内の言い方で、
    // 受け取る営業には何をすればよいのか分からない（この画面はITに不慣れな人が前提）。
    // 原因はログへ残し、画面には次の動作だけを出す。原因を見るときはメニューの
    // 「LINE確認の疎通を確かめる」で担当ごとに確かめられる。
    if (r.status === 404 && r.body && r.body.message) {
      Logger.log("LINE確認: 担当を引き当てられません（" + rep.name + "）: " + r.body.message);
    }
    return { available: true, pending: [], customers: [], notice: SLR_MSG_UNAVAILABLE };
  }

  const body = r.body || {};
  return {
    available: true,
    notice: "",
    pending: body.pending || [],
    customers: body.customers || [],
    customersTruncated: !!body.customersTruncated
  };
}

// ---------------------------------------------------------------
// 画面からの書き込み（doPost action=salesLineReview）
// ---------------------------------------------------------------

function handleSalesLineReview_(data) {
  const rep = sdFindByToken_(data && data.token);
  if (!rep || rep.status !== SD_STATUS_ACTIVE) {
    return { result: "error", message: "このURLでは操作できません。担当者へご連絡ください。" };
  }
  if (!slrEnabled_()) {
    return { result: "error", message: SLR_MSG_UNAVAILABLE };
  }

  const registrationId  = String((data && data.registrationId) || "").trim();
  const decision        = String((data && data.decision) || "").trim();
  const customerId      = String((data && data.customerId) || "").trim();
  const expectedUpdated = String((data && data.expectedUpdatedAt) || "").trim();

  if (!registrationId || !expectedUpdated) {
    return { result: "error", message: "対象が特定できませんでした。画面を開き直してください。" };
  }
  // new = 台帳に居ない方。統合アプリ側で、登録時のお名前・生年月日から顧客を作って承認する
  if (decision !== "approved" && decision !== "rejected" && decision !== "new") {
    return { result: "error", message: "答えの種類が正しくありません。画面を開き直してください。" };
  }
  if (decision === "approved" && !customerId) {
    return { result: "error", message: "どのお客様かを選んでから押してください。" };
  }

  // **担当名はここで入れる。** 画面から送らせると、他人の名前を名乗って
  // 他人の未確認分を操作できてしまう（身元は合言葉だけが決める）。
  const payload = {
    staff: rep.name,
    registrationId: registrationId,
    decision: decision === "new" ? "approved_as_new" : decision,
    expectedUpdatedAt: expectedUpdated
  };
  if (decision === "approved") payload.candidateCustomerId = customerId;

  const r = slrFetch_("post", "", payload);
  if (r.ok) return { result: "ok" };

  if (r.reason === "network") return { result: "error", message: SLR_MSG_NETWORK };
  if (r.status === 409) {
    // 新規登録は「顧客を作る」と「承認する」の2段構え。顧客だけができた場合は
    // **作り直させない**（もう一度押されると同じ人が2件できる）。
    // 開き直せば一覧に出てくるので、選んで承認してもらう。
    if (r.body && r.body.error === "customer_created_review_failed") {
      return {
        result: "error",
        message: "お客様の登録はできましたが、確認の記録に失敗しました。" +
                 "「最新の状態にする」を押すと一覧にその方が出てきますので、" +
                 "選んでから「このお客様で間違いありません」を押してください。"
      };
    }
    return { result: "error", message: SLR_MSG_CONFLICT };
  }
  if (r.status === 404) {
    // 担当が引き当てられない＝設定の話。対象が見つからない＝先に誰かが答えた。
    // どちらも社内の言い方は出さず、次の動作だけを伝える。
    if (r.body && r.body.error === "staff_unresolved") {
      Logger.log("LINE確認: 担当を引き当てられません（" + rep.name + "）: " + (r.body.message || ""));
      return { result: "error", message: SLR_MSG_UNAVAILABLE };
    }
    return { result: "error", message: SLR_MSG_CONFLICT };
  }
  if (r.status === 400) {
    return { result: "error", message: "そのお客様は選べませんでした。画面を開き直してお試しください。" };
  }
  return { result: "error", message: SLR_MSG_UNAVAILABLE };
}

// ---------------------------------------------------------------
// 設定と確認（GASエディタから手動実行）
// ---------------------------------------------------------------

// 統合アプリのURLと秘密値を設定する。
// 引数で渡した値はソースにもGitにも残らない。実行後は引数を消しておくこと。
function setSalesReviewConfig(origin, secret) {
  if (!secret) throw new Error("secret を渡してください。");
  if (String(secret).length < 32) throw new Error("秘密値が短すぎます（32文字以上）。");
  const props = {};
  props[SLR_SECRET_PROPERTY] = String(secret);
  if (origin) props[SLR_APP_URL_PROPERTY] = String(origin).replace(/\/+$/, "");
  PropertiesService.getScriptProperties().setProperties(props);
  Logger.log("設定しました。checkSalesLineReview() で疎通を確認してください。");
  return "設定しました。checkSalesLineReview() で疎通を確認してください。";
}

// 秘密値をこの場で作り直し、**SHA-256だけ**を返す。
//
// 平文をこちらから外へ渡さないための入口。統合アプリ側は SHA-256 しか保存しないので、
// 「平文はGASだけが持つ」を、設定作業の途中でも崩さずに済む。
// 返した SHA-256 は秘密ではないので、そのまま Vercel の
// `SALES_DASHBOARD_SECRET_SHA256` へ入れてよい。
function rotateSalesReviewSecret() {
  const secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ""); // 64文字
  PropertiesService.getScriptProperties().setProperty(SLR_SECRET_PROPERTY, secret);
  return slrSecretSha256_();
}

function slrSecretSha256_() {
  const secret = PropertiesService.getScriptProperties().getProperty(SLR_SECRET_PROPERTY) || "";
  if (!secret) return "";
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return ("0" + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join("");
}

// メニュー用。秘密値を作り直し、コピーできる形で SHA-256 を見せる。
function showSalesReviewSecretHash() {
  const hash = rotateSalesReviewSecret();
  const html =
    '<style>body{font-family:-apple-system,"Hiragino Sans","Noto Sans JP",sans-serif;font-size:13px;' +
    'color:#0f172a;margin:0;padding:14px;line-height:1.8;}' +
    'input{width:100%;font-size:12px;padding:8px;border:1px solid #cbd5e1;border-radius:6px;' +
    'font-family:ui-monospace,Consolas,monospace;}' +
    'p{margin:0 0 10px;}b{font-weight:700;}' +
    'button{margin-top:10px;font-size:12px;padding:8px 14px;border:none;border-radius:6px;' +
    'background:#4f46e5;color:#fff;cursor:pointer;}</style>' +
    '<p>営業ダッシュボードの「LINEの確認」用に、<b>新しい秘密の合言葉を作り直しました。</b></p>' +
    '<p>下の文字列は<b>秘密ではありません</b>（合言葉そのものではなく、その指紋です）。' +
    'これを統合アプリ側の設定 <b>SALES_DASHBOARD_SECRET_SHA256</b> に入れてください。</p>' +
    '<input type="text" readonly value="' + hash + '" onclick="this.select()">' +
    '<button type="button" onclick="var i=document.querySelector(\'input\');i.select();' +
    'document.execCommand(\'copy\');this.textContent=\'コピーしました\';">コピー</button>' +
    '<p style="margin-top:12px;color:#991b1b;">入れ替えるまでのあいだ、営業の画面では' +
    '「いまLINEの確認ができません」と出ます（安全側に倒しています）。</p>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(560).setHeight(330),
    "LINE確認の秘密値を作り直しました"
  );
}

// 疎通と権限の確認。**読み取りだけ**で、承認は行わない。
// 秘密値そのものは出さず、桁数だけを出す。
function checkSalesLineReview() {
  const cfg = slrConfig_();
  const lines = [];
  lines.push("URL設定: " + (cfg.origin ? "あり（" + cfg.origin + "）" : "未設定"));
  lines.push("秘密値設定: " + (cfg.secret ? "あり（" + cfg.secret.length + "文字）" : "未設定"));
  if (!cfg.origin || !cfg.secret) {
    lines.push("→ 未設定のため、営業ダッシュボードに「LINEの確認」は出ません。");
    const msg = lines.join("\n");
    Logger.log(msg);
    return msg;
  }

  const reps = sdReadReps_().filter(function (r) { return r.status === SD_STATUS_ACTIVE; });
  lines.push("");
  lines.push("--- 担当ごとの確認待ち ---");
  reps.forEach(function (rep) {
    const r = slrFetch_("get", "?staff=" + encodeURIComponent(rep.name), null);
    if (r.ok) {
      const b = r.body || {};
      lines.push(rep.name + ": 確認待ち " + ((b.pending || []).length) +
                 " 件 / 選べるお客様 " + ((b.customers || []).length) + " 人");
    } else if (r.status === 404 && r.body && r.body.message) {
      lines.push(rep.name + ": " + r.body.message);
    } else {
      lines.push(rep.name + ": 取得できません（HTTP " + r.status + "）");
    }
  });

  const msg = lines.join("\n");
  Logger.log(msg);
  return msg;
}

function checkSalesLineReviewFromMenu() {
  SpreadsheetApp.getUi().alert("LINE確認の疎通確認\n\n" + checkSalesLineReview());
}
