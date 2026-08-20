// =============================================
// 代理店別の取扱案件（チェックで選ぶ）と実績カウント
// （2026-08-19 追加）
//
// 代理店によって扱いたい案件が違うため、全稼働案件を一律に見せるとずれる。
// 「代理店別取扱案件」シートで 案件×代理店 のチェックを持ち、リンク集はそれで絞る。
//
// 既定は「全部チェック」。代理店を登録した時点で稼働中の案件すべてに印を付けるので、
// 何もしなければ従来どおり全案件が渡る。外したい案件だけチェックを外す運用。
// =============================================

const AGENCY_CASE_MATRIX_SHEET = "代理店別取扱案件";
const ACM_FIXED_COLS = 2; // A=案件コード, B=案件名。C列以降が代理店。

function getAgencyCaseMatrixSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(AGENCY_CASE_MATRIX_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AGENCY_CASE_MATRIX_SHEET, 3);
    const h = sh.getRange(1, 1, 1, ACM_FIXED_COLS);
    h.setValues([["案件コード", "案件名"]]);
    h.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setFrozenColumns(ACM_FIXED_COLS);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 260);
  }
  return sh;
}

// 現在の状態を {caseCode: {agencyCode: true/false}} で読む
function readAgencyCaseMatrix_() {
  const sh = getAgencyCaseMatrixSheet_();
  const out = { byCase: {}, agencyCodes: [] };
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol <= ACM_FIXED_COLS) return out;

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const codes = [];
  for (let c = ACM_FIXED_COLS; c < lastCol; c++) {
    // ヘッダーは「代理店名（コード）」。括弧内のコードを正とする。
    const m = /（([^（）]+)）\s*$/.exec(String(header[c] || ""));
    codes.push(m ? m[1] : "");
  }
  out.agencyCodes = codes;

  if (lastRow < 2) return out;
  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  body.forEach(function (row) {
    const caseCode = String(row[0] || "").trim();
    if (!caseCode) return;
    const per = {};
    for (let c = ACM_FIXED_COLS; c < lastCol; c++) {
      const ag = codes[c - ACM_FIXED_COLS];
      if (ag) per[ag] = row[c] === true;
    }
    out.byCase[caseCode] = per;
  });
  return out;
}

// 案件（稼働中のみ）×代理店（稼働中のみ）で作り直す。既存のチェックは保持する。
// 新しく現れた組み合わせは true（＝渡す）で作る。
function syncAgencyCaseMatrix() {
  const sh = getAgencyCaseMatrixSheet_();
  const prev = readAgencyCaseMatrix_();

  const cases = listActiveCases_();
  const agencies = readAgencies_().filter(function (a) {
    return a.status === AGENCY_STATUS_ACTIVE;
  });

  // いったん全消し（書式ごと）してから作り直す
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow >= 1 && lastCol >= 1) {
    sh.getRange(1, 1, Math.max(lastRow, 1), Math.max(lastCol, 1)).clear();
  }

  const header = ["案件コード", "案件名"].concat(agencies.map(function (a) {
    return a.name + "（" + a.code + "）";
  }));
  const hr = sh.getRange(1, 1, 1, header.length);
  hr.setValues([header]);
  hr.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setWrap(true);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(ACM_FIXED_COLS);

  if (!cases.length || !agencies.length) {
    return { cases: cases.length, agencies: agencies.length };
  }

  const body = cases.map(function (c) {
    const row = [c.code, c.name];
    agencies.forEach(function (a) {
      const prevRow = prev.byCase[c.code];
      const had = prevRow && Object.prototype.hasOwnProperty.call(prevRow, a.code);
      row.push(had ? prevRow[a.code] === true : true); // 新しい組み合わせは既定で渡す
    });
    return row;
  });

  sh.getRange(2, 1, body.length, header.length).setValues(body);
  const cb = sh.getRange(2, ACM_FIXED_COLS + 1, body.length, agencies.length);
  cb.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  cb.setHorizontalAlignment("center");
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 260);
  for (let i = 0; i < agencies.length; i++) sh.setColumnWidth(ACM_FIXED_COLS + 1 + i, 130);

  return { cases: cases.length, agencies: agencies.length };
}

// この代理店に渡す案件コードの配列を返す。
// マトリクスに情報が無ければ null を返し、呼び出し側は「全稼働案件」として扱う。
function agencyCaseSelection_(agencyCode) {
  const m = readAgencyCaseMatrix_();
  if (m.agencyCodes.indexOf(agencyCode) < 0) return null; // 列が無い＝未設定
  const picked = [];
  Object.keys(m.byCase).forEach(function (caseCode) {
    if (m.byCase[caseCode][agencyCode] === true) picked.push(caseCode);
  });
  return picked;
}

function syncAgencyCaseMatrixFromMenu() {
  const r = syncAgencyCaseMatrix();
  SpreadsheetApp.getUi().alert(
    "代理店別の取扱案件を同期しました。\n\n" +
    "稼働案件: " + r.cases + " 件\n" +
    "稼働代理店: " + r.agencies + " 件\n\n" +
    "新しい組み合わせは「渡す」で作られます。渡したくない案件のチェックを外してください。" +
    (r.agencies === 0 ? "\n\n※稼働中の代理店がまだ無いため、代理店の列はありません。" : "")
  );
}

// =============================================
// 代理店の実績カウント（リンク集の一枚もの用）
// =============================================

// 代理店ごとの申請件数を {案件名: 件数} で返す。
//
// **申請状況一覧（生成物）を1回読むだけにする。** 以前は案件シートを1枚ずつ開いて
// 「代理店」列を数えており、リンク集の表示に十数秒かかっていた（2026-08-20 実測）。
// 申請状況一覧は日次トリガーで作り直されるので、件数は最大1日ぶん遅れる。
// 代理店へ見せる目安の数字なので、速さを優先してよいと判断した。
// 案件名で持つのは、申請状況一覧が案件コードを持たないため。
function countAgencyApplicationsByName_(agencyName) {
  const name = String(agencyName || "").trim();
  const counts = {};
  if (!name) return counts;

  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(APP_STATUS_SHEET);
  if (!sh) return counts;                    // まだ生成されていない
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return counts;

  // B列=案件名 / E列=代理店 だけを読む
  const vals = sh.getRange(2, 2, lastRow - 1, 4).getValues();  // B〜E
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][3] || "").trim() !== name) continue;     // E列
    const caseName = String(vals[i][0] || "").trim();           // B列
    if (!caseName) continue;
    counts[caseName] = (counts[caseName] || 0) + 1;
  }
  return counts;
}
