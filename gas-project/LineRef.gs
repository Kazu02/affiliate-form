// =============================================
// 公式LINE経由の申請を、名前に頼らず記録する（2026-11-16 追加）
//
// これまで公式LINE（統合アプリ）は、申請状況一覧の「お客さんの名前」と台帳の名前を
// 突き合わせて「誰がどの案件をやったか」を推定していた。表記ゆれ・あだ名・代理店経由で
// 当たらない行が多い（11/16 実測で申込857件中599件が台帳と一致しない）。
//
// そこで、公式LINEから申請フォームを開くURLに `lc=<参照>` を付け、申請時にその参照と
// 案件記号だけをこのシートへ1行足す。参照は統合アプリだけが検証できる番号
// （顧客ID 32桁＋署名 16桁の16進）で、**名前・電話番号などの個人情報は入っていない。**
// 偽の参照は統合アプリ側の署名検査で捨てる。ここでは形だけを確かめる。
//
// 約束:
//   - **既存の案件シート・申請状況一覧・代理店シートの列は変えない。** 別シートに足すだけ。
//   - **ここで失敗しても申請は通す。** 記録できなければ従来どおり名前の照合に戻るだけ。
//   - 統合アプリが `includeLineSubmissions: true` を付けて頼んだときだけスナップショットに載せる
//     （古い統合アプリは応答の形を厳密に検査しているため、頼まれない限り形を変えない）。
// =============================================

const LINE_REF_SHEET = "LINE経由申請";
const LINE_REF_HEADERS = ["受付日時", "フォーム記号", "LINE参照"];
const LINE_REF_PATTERN = /^[a-f0-9]{48}$/;
// 返す行の上限は統合アプリの受け入れ上限（20,000）に合わせる。同じ日・同じ案件・同じ参照は1行にまとめる
// ので、上限に届くのは「顧客×案件×申請日」が2万通りを超えたとき（2026-11 の規模なら数十年先）。
const LINE_REF_MAX_ROWS = 20000;
// 読むのは新しい行から最大この行数まで（GAS の実行時間を守るため）。
const LINE_REF_SCAN_ROWS = 60000;

function isValidLineRef_(v) {
  return LINE_REF_PATTERN.test(String(v || "").trim());
}

function lineRefSheet_(ss) {
  let sh = ss.getSheetByName(LINE_REF_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LINE_REF_SHEET);
    const hr = sh.getRange(1, 1, 1, LINE_REF_HEADERS.length);
    hr.setValues([LINE_REF_HEADERS]);
    hr.setFontWeight("bold").setBackground("#06C755").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 380);
  }
  return sh;
}

// 申請を受け付けたときに呼ぶ。**追記だけ。例外は外へ出さない。**
function recordLineSubmission_(ss, formName, lineRef) {
  try {
    const ref = String(lineRef || "").trim();
    const code = String(formName || "").trim();
    if (!isValidLineRef_(ref) || !code) return false;
    lineRefSheet_(ss).appendRow([formatJST(new Date()), code.slice(0, 60), ref]);
    return true;
  } catch (e) {
    Logger.log("LINE経由申請の記録に失敗（申請は通常どおり記録）: " + e);
    return false;
  }
}

// スナップショット用。新しい行から読み、同じ「案件・参照・申請日」は1行にまとめて最大 LINE_REF_MAX_ROWS 件。
// 返す順は古い→新しい。読み切れなかった行・上限で落とした行があれば truncated。
function buildLineSubmissions_(ss) {
  const sh = ss.getSheetByName(LINE_REF_SHEET);
  if (!sh) return { rows: [], truncated: false };
  const last = sh.getLastRow();
  if (last < 2) return { rows: [], truncated: false };
  let first = 2;
  let truncated = false;
  if (last - 1 > LINE_REF_SCAN_ROWS) { first = last - LINE_REF_SCAN_ROWS + 1; truncated = true; }
  const values = sh.getRange(first, 1, last - first + 1, LINE_REF_HEADERS.length).getValues();
  const rows = [];
  const seen = {};
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    const ref = String(r[2] || "").trim();
    const caseCode = String(r[1] || "").trim();
    if (!isValidLineRef_(ref) || !/^[A-Za-z0-9][A-Za-z0-9._-]{0,119}$/.test(caseCode)) continue;
    const submittedOn = String(toDisplayDate_(r[0]) || "").substring(0, 10);
    const key = caseCode + "|" + ref + "|" + submittedOn;
    if (seen[key]) continue;
    if (rows.length >= LINE_REF_MAX_ROWS) { truncated = true; break; }
    seen[key] = true;
    rows.push({ caseCode: caseCode, lineRef: ref, submittedOn: submittedOn });
  }
  rows.reverse();
  return { rows: rows, truncated: truncated };
}
