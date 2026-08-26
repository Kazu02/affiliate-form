// =============================================
// クリック時刻をサーバー側で記録する（2026-08-26 追加）
//
// **これまでクリック日時は利用者の端末時計で記録していた。**
// index.html のボタン押下時に `new Date()` をブラウザで実行していたため、
// 端末の時計がずれていればそのままずれた値がシートへ入る。
//
// ASP獲得ログとの突合はこのクリック日時が唯一のキーなので、ここがずれると当たらない。
// 2026-08-20 の実測で一致は 217/451 = 48%。一致するときは中央値0秒なので
// **キーとしては正しく、ずれているのは時計**という切り分けができている。
//
// そこで、アフィリエイトボタンを押した瞬間にGASへ小さな知らせを飛ばし、
// **サーバー側の時刻**を記録する。ASP側もASPのサーバー時刻で記録するので、
// サーバー同士の比較になり端末時計の影響が消える。
//
// 設計上の約束:
//   - **失敗しても申請は通す。** 知らせが届かなければ従来どおり端末の時刻を使う。
//     申請フォームは2026-08-05 に2日間止めた前科があるので、ここで止めない。
//   - **端末が申告した時刻も残す。** どちらを使ったかを台帳で追えるようにする。
// =============================================

const CLICK_LOG_SHEET = "クリックログ";
const CLICK_LOG_HEADERS = [
  "クリックID", "サーバー記録日時", "フォーム記号", "代理店", "端末申告日時", "申請で使用"
];
const CL_COL_ID = 1, CL_COL_SERVER = 2, CL_COL_FORM = 3, CL_COL_AGENCY = 4,
      CL_COL_CLIENT = 5, CL_COL_USED = 6;

// クリックIDの見た目の検査。**素性の分からない文字列をシートへ書かない。**
// 公開フォームから来るので、長さと文字種を絞る。
function isValidClickId_(v) {
  return /^[A-Za-z0-9_-]{8,64}$/.test(String(v || "").trim());
}

function clickLogSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(CLICK_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CLICK_LOG_SHEET);
    const hr = sh.getRange(1, 1, 1, CLICK_LOG_HEADERS.length);
    hr.setValues([CLICK_LOG_HEADERS]);
    hr.setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(CL_COL_ID, 200);
    sh.setColumnWidth(CL_COL_SERVER, 160);
    sh.setColumnWidth(CL_COL_CLIENT, 160);
  }
  return sh;
}

// アフィリエイトボタンが押されたときの知らせ。**追記だけ。**
function handleClickPing_(data) {
  const id = String(data.clickId || "").trim();
  if (!isValidClickId_(id)) return { result: "error", message: "bad id" };
  clickLogSheet_().appendRow([
    id, formatJST(new Date()),
    String(data.formName || "").slice(0, 60),
    String(data.agencyCode || "").slice(0, 30),
    data.clientTime ? formatJST(new Date(data.clientTime)) : "",
    ""
  ]);
  return { result: "ok" };
}

// クリックIDに対応するサーバー記録時刻を返す（無ければ ""）。
// 同じIDが複数あれば**最初の1件**を使う（押し直しても最初のクリックが成果に紐づくため）。
function findServerClickTime_(clickId) {
  const id = String(clickId || "").trim();
  if (!isValidClickId_(id)) return { time: "", row: 0 };
  const sh = clickLogSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { time: "", row: 0 };
  const data = sh.getRange(2, 1, last - 1, CLICK_LOG_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][CL_COL_ID - 1] || "").trim() !== id) continue;
    return { time: String(data[i][CL_COL_SERVER - 1] || ""), row: i + 2 };
  }
  return { time: "", row: 0 };
}

// 申請行に書くクリック日時を決める。
// **サーバー記録があればそれを使う。** 無ければ従来どおり端末の申告値。
function resolveClickAt_(data) {
  try {
    const hit = findServerClickTime_(data.clickId);
    if (hit.time) {
      // どちらを使ったか台帳に残す（あとで「なぜこの時刻か」を追える）
      try { clickLogSheet_().getRange(hit.row, CL_COL_USED).setValue("使用"); } catch (e) {}
      return hit.time;
    }
  } catch (e) {
    Logger.log("サーバー記録のクリック時刻を引けませんでした（端末値を使う）: " + e);
  }
  return data.clickTime ? formatJST(new Date(data.clickTime)) : "";
}

// 端末時計とサーバー記録がどれだけずれているかを見る（効果の確認用）
function checkClickTimeDrift() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sh = clickLogSheet_();
    const last = sh.getLastRow();
    if (last < 2) { ui.alert("クリックログにデータがありません。"); return; }
    const data = sh.getRange(2, 1, last - 1, CLICK_LOG_HEADERS.length).getValues();
    const diffs = [];
    data.forEach(function (r) {
      const s = aspToMillis_(r[CL_COL_SERVER - 1]);
      const c = aspToMillis_(r[CL_COL_CLIENT - 1]);
      if (s && c) diffs.push(Math.round((c - s) / 1000));
    });
    if (!diffs.length) { ui.alert("端末申告日時が入った行がまだありません。"); return; }
    diffs.sort(function (a, b) { return Math.abs(a) - Math.abs(b); });
    const over = diffs.filter(function (d) { return Math.abs(d) > 120; }).length;
    ui.alert("端末時計とサーバー記録のずれ\n\n" +
             "件数: " + diffs.length + "\n" +
             "中央値: " + diffs[Math.floor(diffs.length / 2)] + " 秒\n" +
             "最大: " + diffs[diffs.length - 1] + " 秒\n" +
             "±120秒を超えた件数: " + over + " 件（" +
             Math.round(over * 100 / diffs.length) + "%）\n\n" +
             "この割合が、これまで突合できていなかった分にあたります。");
  } catch (e) {
    ui.alert("確認できませんでした。\n\n" + e);
  }
}
